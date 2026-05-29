"""Phase 1-2 smoke test: convert a PDF to HTML and report results.

Usage:
    python tests/pdf_poc_test.py <path-to-pdf> [output-dir]

The output defaults to /tmp/pdf_poc_out/.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'backend'))

import pdf_converter


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_pdf = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else '/tmp/pdf_poc_out'

    if not os.path.isfile(input_pdf):
        print(f'ERROR: file not found: {input_pdf}')
        sys.exit(1)

    print(f'Input : {input_pdf}')
    print(f'Output: {output_dir}')
    print()

    result = pdf_converter.convert(input_pdf, output_dir)

    print(f'Success   : {result["success"]}')
    print(f'Duration  : {result["duration_sec"]}s')
    print(f'Exit code : {result["exit_code"]}')
    if result['stderr']:
        print(f'stderr    :\n{result["stderr"].strip()}')
    if result['html_path']:
        size = os.path.getsize(result['html_path'])
        print(f'HTML path : {result["html_path"]} ({size:,} bytes)')
        print()
        print('Open in browser:')
        print(f'  open "{result["html_path"]}"')

    sys.exit(0 if result['success'] else 1)


if __name__ == '__main__':
    main()
