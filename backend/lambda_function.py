import os
import subprocess
import tempfile
import boto3
import logging
from urllib.parse import unquote_plus

LOG = logging.getLogger('lambda')
LOG.setLevel(logging.INFO)

s3 = boto3.client('s3')

def download_s3(bucket, key, dest_path):
    s3.download_file(bucket, key, dest_path)

def upload_s3(path, bucket, key):
    s3.upload_file(path, bucket, key)

def run_renderer(input_path, output_path):
    # Prefer calling the existing docx_renderer.py if it handles HTML output
    if os.path.exists('/var/task/docx_renderer.py'):
        subprocess.check_call(['python3', '/var/task/docx_renderer.py', input_path, output_path])
    else:
        # Fallback: use libreoffice to convert to PDF
        subprocess.check_call(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(output_path), input_path])

def make_presigned_url(bucket, key, expires=3600):
    return s3.generate_presigned_url('get_object', Params={'Bucket': bucket, 'Key': key}, ExpiresIn=expires)

def lambda_handler(event, context):
    # Expect event: {"bucket": "my-bucket", "key": "path/to/file.docx"}
    LOG.info('Event: %s', event)
    bucket = event.get('bucket') or event.get('Records', [{}])[0].get('s3', {}).get('bucket', {}).get('name')
    key = event.get('key') or event.get('Records', [{}])[0].get('s3', {}).get('object', {}).get('key')
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
            run_renderer(inp, out)
        except subprocess.CalledProcessError:
            # try LibreOffice PDF fallback
            LOG.exception('Renderer failed; attempting LibreOffice PDF fallback')
            pdf_out = os.path.join(tmp, base + '.pdf')
            subprocess.check_call(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, inp])
            out = pdf_out
            out_key = f"converted/{base}.pdf"

        upload_s3(out, bucket, out_key)

    url = make_presigned_url(bucket, out_key)
    return {"bucket": bucket, "key": out_key, "url": url}
