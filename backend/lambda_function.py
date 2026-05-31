import os
import subprocess
import tempfile
import boto3
import logging
import json
from urllib.parse import unquote_plus

LOG = logging.getLogger('lambda')
LOG.setLevel(logging.INFO)

s3 = boto3.client('s3')
RENDERER_TIMEOUT_SEC = int(os.environ.get('RENDERER_TIMEOUT_SEC', '120'))

def download_s3(bucket, key, dest_path):
    s3.download_file(bucket, key, dest_path)

def upload_s3(path, bucket, key):
    s3.upload_file(path, bucket, key)

def run_renderer(input_path, output_path, use_ai=False):
    # Prefer calling the existing docx_renderer.py if it handles HTML output
    if os.path.exists('/var/task/docx_renderer.py'):
        # Build args for renderer with optional AI flag
        args = ['python3', '/var/task/docx_renderer.py', input_path, output_path]
        if use_ai:
            args.append('--use-ai')
        subprocess.run(args, check=True, timeout=RENDERER_TIMEOUT_SEC)
    else:
        # Fallback: use libreoffice to convert to PDF
        subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(output_path), input_path],
            check=True,
            timeout=RENDERER_TIMEOUT_SEC,
        )

def make_presigned_url(bucket, key, expires=3600):
    return s3.generate_presigned_url('get_object', Params={'Bucket': bucket, 'Key': key}, ExpiresIn=expires)

def lambda_handler(event, context):
    # Expect event: {"bucket": "my-bucket", "key": "path/to/file.docx", "useAI": true}
    LOG.info('Event: %s', event)
    bucket = event.get('bucket') or event.get('Records', [{}])[0].get('s3', {}).get('bucket', {}).get('name')
    key = event.get('key') or event.get('Records', [{}])[0].get('s3', {}).get('object', {}).get('key')
    use_ai = event.get('useAI', False)
    
    if not bucket or not key:
        raise ValueError('Missing bucket or key in event')
    key = unquote_plus(key)

    base, ext = os.path.splitext(os.path.basename(key))
    # Preserve UUID path so presign_handler can find the result by replacing uploads/ prefix
    key_dir = os.path.dirname(key)  # e.g. uploads/<uuid>
    out_dir = key_dir.replace('uploads/', 'converted/', 1) if key_dir.startswith('uploads/') else f"converted/{key_dir}"
    out_key = f"{out_dir}/{base}.html"

    with tempfile.TemporaryDirectory() as tmp:
        inp = os.path.join(tmp, 'input' + ext)
        out = os.path.join(tmp, base + '.html')
        download_s3(bucket, key, inp)

        try:
            run_renderer(inp, out, use_ai=use_ai)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
                # try LibreOffice PDF fallback
                LOG.exception('Renderer failed; attempting LibreOffice PDF fallback')
                pdf_out = os.path.join(tmp, base + '.pdf')
                subprocess.run(
                    ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, inp],
                    check=True,
                    timeout=RENDERER_TIMEOUT_SEC,
                )
                out = pdf_out
                out_key = f"converted/{base}.pdf"

        upload_s3(out, bucket, out_key)

    url = make_presigned_url(bucket, out_key)
    return {"bucket": bucket, "key": out_key, "url": url}
