#!/usr/bin/env bash
# Swagger 2.0 JSON -> Excel. Requires Bash and Python 3.8+ (standard library only).
set -euo pipefail

if [[ -n "${PYTHON:-}" ]]; then
    python_cmd="$PYTHON"
elif command -v python3 >/dev/null 2>&1 && python3 -c 'import sys; sys.exit(sys.version_info < (3,8))' >/dev/null 2>&1; then
    python_cmd=python3
elif command -v python >/dev/null 2>&1 && python -c 'import sys; sys.exit(sys.version_info < (3,8))' >/dev/null 2>&1; then
    python_cmd=python
else
    printf '%s\n' 'Error: Python 3.8 or newer is required. Install Python or set PYTHON to its executable path.' >&2
    exit 1
fi

exec "$python_cmd" - "$@" <<'PYTHON_CODE'
import argparse
import json
import math
import os
from pathlib import Path
import re
import sys
import tempfile
from urllib.parse import unquote
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED

def xml(value):
    value = str(value)
    if re.search(r'[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]', value):
        raise ValueError('Source contains a character that is not valid in XML.')
    if len(value) > 32767:
        raise ValueError('Source text exceeds the Excel cell text limit.')
    return escape(value)

def main():
    parser = argparse.ArgumentParser(description='Export Swagger 2.0 JSON to an Excel workbook with Requests and Responses tabs.')
    parser.add_argument('swagger_path', help='Input Swagger JSON file')
    parser.add_argument('output_path', nargs='?', default='SwaggerFields.xlsx')
    parser.add_argument('--exclude-request-headers', action='store_true')
    parser.add_argument('--force', action='store_true', help='Replace an existing output file')
    args = parser.parse_args()
    source, output = Path(args.swagger_path).resolve(), Path(args.output_path).resolve()
    if output.suffix.lower() != '.xlsx':
        raise ValueError('Output path must end with .xlsx.')
    if source == output:
        raise ValueError('Input and output paths must differ.')
    if output.exists() and not args.force:
        raise ValueError('Output already exists. Use --force to replace it.')
    with source.open(encoding='utf-8-sig') as f:
        document = json.load(f)
    if document.get('swagger') != '2.0':
        raise ValueError('Only Swagger 2.0 JSON is supported; OpenAPI 3.x and YAML are not supported.')

    def resolve(node, chain=()):
        ref = node.get('$ref')
        if not ref:
            return node
        if ref in chain:
            raise ValueError('Circular reference alias: ' + ref)
        if not ref.startswith('#/'):
            raise ValueError('External references are not supported: ' + ref)
        target = document
        try:
            for part in unquote(ref[2:]).split('/'):
                target = target[part.replace('~1', '/').replace('~0', '~')]
        except (KeyError, TypeError):
            raise ValueError('Unresolved reference: ' + ref)
        return resolve(target, chain + (ref,))

    def row(rows, field, datatype='', required='', style=0):
        rows.append((str(field), str(datatype), str(required), style))

    def schema(rows, node, name, required='', depth=0, stack=(), direction='request'):
        if depth > 60:
            raise ValueError('Schema exceeds supported nesting depth at ' + name)
        prefix = '  ' * depth
        ref = node.get('$ref')
        if ref in stack:
            row(rows, prefix + name, 'recursive reference: ' + ref, required)
            return
        if ref:
            stack += (ref,)
        node = resolve(node)
        for keyword in ('allOf', 'oneOf', 'anyOf'):
            if keyword in node:
                raise ValueError("Schema composition '%s' is not supported at %s. Flatten it first." % (keyword, name))
        datatype = node.get('type') or ('object' if 'properties' in node or 'additionalProperties' in node else 'any')
        if datatype in ('object', 'array'):
            row(rows, prefix + '<<' + name + '>>', datatype, required, 2)
            if datatype == 'object':
                for key, child in node.get('properties', {}).items():
                    if direction == 'request' and resolve(child).get('readOnly') is True:
                        continue
                    status = 'Required' if key in node.get('required', []) else 'Optional'
                    schema(rows, child, key, status, depth + 1, stack, direction)
                additional = node.get('additionalProperties')
                if isinstance(additional, dict):
                    schema(rows, additional, '{additional key}', 'Optional', depth + 1, stack, direction)
                elif additional is True:
                    row(rows, prefix + '  {additional key}', 'any', 'Optional')
            else:
                if 'items' not in node:
                    raise ValueError('Array is missing items at ' + name)
                schema(rows, node['items'], name + '[]', '', depth + 1, stack, direction)
            row(rows, prefix + '<</' + name + '>>', style=2)
        else:
            if node.get('format'):
                datatype += ' (' + node['format'] + ')'
            row(rows, prefix + name, datatype, required)

    def name(node, fallback):
        return unquote(node['$ref'].rsplit('/', 1)[-1]).replace('~1', '/').replace('~0', '~') if '$ref' in node else fallback

    requests, responses = [], []
    for rows in (requests, responses):
        row(rows, 'Source Fields', 'Fields DataType', 'Required/Optional', 1)
    count = 0
    for path, item in document.get('paths', {}).items():
        item = resolve(item)
        for method in ('get', 'put', 'post', 'delete', 'options', 'head', 'patch'):
            if method not in item:
                continue
            operation = item[method]
            count += 1
            for rows in (requests, responses):
                row(rows, method.upper() + ' ' + path, style=3)
            parameters = {}
            for param in item.get('parameters', []) + operation.get('parameters', []):
                param = resolve(param)
                parameters[(param['in'], param['name'])] = param
            for param in parameters.values():
                if args.exclude_request_headers and param['in'] == 'header':
                    continue
                status = 'Required' if param.get('required') or param['in'] == 'path' else 'Optional'
                if param['in'] == 'body':
                    schema(requests, param['schema'], name(param['schema'], param['name']), status)
                else:
                    schema(requests, param, param['name'] + ' [' + param['in'] + ']', status)
            for code, response in operation.get('responses', {}).items():
                response = resolve(response)
                row(responses, 'HTTP ' + code + ': ' + response.get('description', ''), style=3)
                if 'schema' in response:
                    schema(responses, response['schema'], name(response['schema'], 'body'), direction='response')
                else:
                    row(responses, '(No response body schema)')
                for key, header in response.get('headers', {}).items():
                    schema(responses, header, key + ' [response header]', direction='response')
    if not count:
        raise ValueError('No supported HTTP operations were found in paths.')

    def sheet_xml(rows):
        if len(rows) > 1048576:
            raise ValueError('Worksheet exceeds the Excel row limit.')
        parts = ['<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/><cols><col min="1" max="1" width="70" customWidth="1"/><col min="2" max="2" width="28" customWidth="1"/><col min="3" max="3" width="23" customWidth="1"/></cols><sheetData>']
        for index, values in enumerate(rows, 1):
            height = min(409, max(22, max(16 * math.ceil(len(v) / width) for v, width in zip(values[:3], (65, 25, 21)))))
            parts.append('<row r="%d" ht="%d" customHeight="1">' % (index, height))
            for col, value in zip('ABC', values[:3]):
                parts.append('<c r="%s%d" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (col, index, values[3], xml(value)))
            parts.append('</row>')
        return ''.join(parts) + '</sheetData></worksheet>'

    main_ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    office_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    content_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
    parts = {
        '[Content_Types].xml': '<Types xmlns="%s"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>%s</Types>' % (content_ns, ''.join('<Override PartName="/xl/%s" ContentType="%s%s+xml"/>' % (part, mime, kind) for part, kind in [('workbook.xml', 'sheet.main'), ('styles.xml', 'styles'), ('worksheets/sheet1.xml', 'worksheet'), ('worksheets/sheet2.xml', 'worksheet')])),
        '_rels/.rels': '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/></Relationships>' % (rel_ns, office_ns),
        'xl/workbook.xml': '<workbook xmlns="%s" xmlns:r="%s"><sheets><sheet name="Requests" sheetId="1" r:id="rId1"/><sheet name="Responses" sheetId="2" r:id="rId2"/></sheets></workbook>' % (main_ns, office_ns),
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="%s">%s</Relationships>' % (rel_ns, ''.join('<Relationship Id="rId%d" Type="%s/%s" Target="%s"/>' % (i, office_ns, kind, target) for i, kind, target in [(1, 'worksheet', 'worksheets/sheet1.xml'), (2, 'worksheet', 'worksheets/sheet2.xml'), (3, 'styles', 'styles.xml')])),
        'xl/worksheets/sheet1.xml': sheet_xml(requests),
        'xl/worksheets/sheet2.xml': sheet_xml(responses),
    }
    fonts = '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
    fills = '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F607A"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8F0F4"/><bgColor indexed="64"/></patternFill></fill></fills>'
    borders = '<borders count="2"><border/><border>' + ''.join('<%s style="thin"><color rgb="FF7D8D96"/></%s>' % (edge, edge) for edge in ('left', 'right', 'top', 'bottom')) + '</border></borders>'
    styles = '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4">' + ''.join('<xf numFmtId="49" fontId="%d" fillId="%d" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' % ids for ids in [(0, 0), (1, 2), (2, 0), (2, 3)]) + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    parts['xl/styles.xml'] = '<styleSheet xmlns="%s">%s%s%s%s</styleSheet>' % (main_ns, fonts, fills, borders, styles)
    output.parent.mkdir(parents=True, exist_ok=True)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix='.tmp', delete=False) as temp:
            temp_path = Path(temp.name)
        with ZipFile(temp_path, 'w', compression=ZIP_DEFLATED) as archive:
            for path, content in parts.items():
                archive.writestr(path, content.encode('utf-8'))
        if args.force:
            os.replace(temp_path, output)
        else:
            # Exclusive creation prevents overwriting a file created during the export.
            with output.open('xb') as dest, temp_path.open('rb') as src:
                import shutil
                shutil.copyfileobj(src, dest)
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()
    print('Created %s (%d operation(s); %d request rows; %d response rows).' % (output, count, len(requests) - 1, len(responses) - 1))

if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError, KeyError, TypeError, RecursionError) as exc:
        print('Error: ' + str(exc), file=sys.stderr)
        sys.exit(1)
PYTHON_CODE
