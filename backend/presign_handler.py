import os
import json
import uuid
import sys
import importlib
import logging
import boto3
from urllib.parse import parse_qs
from botocore.exceptions import ClientError

logger = logging.getLogger()
logger.setLevel(logging.INFO)

_region = os.environ.get('AWS_DEFAULT_REGION', 'eu-west-1')
s3 = boto3.client('s3', region_name=_region,
                  endpoint_url=f'https://s3.{_region}.amazonaws.com')
lambda_client = boto3.client('lambda', region_name=_region)

BUCKET = os.environ.get('PRESIGN_BUCKET')
if not BUCKET:
    raise RuntimeError('PRESIGN_BUCKET environment variable must be set')

LAMBDA_FUNCTION = os.environ.get('RENDERER_LAMBDA_FUNCTION', 'paperless-renderer')
URL_EXPIRATION = int(os.environ.get('PRESIGN_EXPIRATION', '300'))


def make_response(status_code, body_obj):
    return {
        'statusCode': status_code,
        'headers': {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
            'Access-Control-Allow-Headers': '*',
        },
        'body': json.dumps(body_obj)
    }


def get_query_param(event, name):
    qs = event.get('rawQueryString') if 'rawQueryString' in event else None
    if qs:
        parsed = parse_qs(qs)
        return parsed.get(name, [None])[0]

    params = event.get('queryStringParameters') or {}
    return params.get(name)


def parse_json_body(event):
    body = event.get('body') or ''
    if event.get('isBase64Encoded'):
        import base64
        body = base64.b64decode(body).decode('utf-8')
    if not body:
        return None, 'invalid JSON body'
    try:
        return json.loads(body), None
    except json.JSONDecodeError:
        return None, 'invalid JSON body'


def _ensure_backend_import_paths():
    candidates = [
        '/var/task',
        os.path.dirname(__file__),
    ]
    for path in candidates:
        if path and os.path.isdir(path) and path not in sys.path:
            sys.path.insert(0, path)


def _import_backend_module(module_name):
    _ensure_backend_import_paths()
    package = (__package__ or '').rpartition('.')[0]
    candidates = [module_name]
    if package:
        candidates.insert(0, f'{package}.{module_name}')

    seen = set()
    last_error = None
    for name in candidates:
        if name in seen:
            continue
        seen.add(name)
        try:
            return importlib.import_module(name)
        except ModuleNotFoundError as exc:
            last_error = exc
    if last_error is not None:
        raise last_error
    raise ModuleNotFoundError(f'Unable to import backend module: {module_name}')


def converted_html_key(source_key):
    base_no_ext = os.path.splitext(source_key)[0]
    if base_no_ext.startswith('uploads/'):
        return base_no_ext.replace('uploads/', 'converted/', 1) + '.html'
    return f'converted/{base_no_ext}.html'


def extracted_fields_key(source_key):
    base_no_ext = os.path.splitext(source_key)[0]
    if base_no_ext.startswith('uploads/'):
        return base_no_ext.replace('uploads/', 'extracted/', 1) + '.json'
    return f'extracted/{base_no_ext}.json'


def _s3_user_message(exc: ClientError, context: str) -> str:
    code = exc.response.get('Error', {}).get('Code', '')
    if code in ('404', 'NoSuchKey', 'NotFound'):
        return f'{context}: file not found — try re-uploading'
    if code in ('403', 'AccessDenied'):
        return f'{context}: access denied — check bucket permissions'
    return f'{context}: storage error ({code})'


def object_exists(bucket, key):
    try:
        s3.head_object(Bucket=bucket, Key=key)
        return True
    except ClientError as exc:
        error_code = exc.response.get('Error', {}).get('Code')
        if error_code in {'404', 'NoSuchKey', 'NotFound'}:
            return False
        raise


def list_converted_keys():
    paginator = s3.get_paginator('list_objects_v2')
    keys: set[str] = set()

    for page in paginator.paginate(Bucket=BUCKET, Prefix='converted/'):
        for obj in page.get('Contents', []):
            key = obj.get('Key')
            if key:
                keys.add(key)

    return keys


def list_uploaded_files():
    paginator = s3.get_paginator('list_objects_v2')
    items = []
    converted_keys = list_converted_keys()

    for page in paginator.paginate(Bucket=BUCKET, Prefix='uploads/'):
        for obj in page.get('Contents', []):
            key = obj.get('Key')
            if not key:
                continue

            file_name = key.rsplit('/', 1)[-1]
            html_key = converted_html_key(key)
            items.append({
                'key': key,
                'fileName': file_name,
                'lastModified': obj.get('LastModified').isoformat() if obj.get('LastModified') else None,
                'status': 'ready' if html_key in converted_keys else 'pending',
            })

    items.sort(key=lambda item: item.get('lastModified') or '', reverse=True)
    return items


def lambda_handler(event, context):
    # Support Function URL (v2) and simple proxy shapes
    method = None
    if 'requestContext' in event and 'http' in event['requestContext']:
        method = event['requestContext']['http'].get('method')
    elif 'httpMethod' in event:
        method = event['httpMethod']

    # CORS preflight
    if method == 'OPTIONS':
        return make_response(204, {})

    try:
        action = get_query_param(event, 'action')

        if method == 'GET' and action == 'list':
            return make_response(200, {'items': list_uploaded_files()})

        if method == 'POST' and action == 'convert':
            # Invoke the renderer Lambda with useAI flag
            # Expect body: {"key": "uploads/uuid/file.docx", "useAI": true}
            payload, error = parse_json_body(event)
            if error:
                return make_response(400, {'error': error})
            key = payload.get('key')
            use_ai = payload.get('useAI', False)

            if not key:
                return make_response(400, {'error': 'key is required'})

            try:
                invoke_payload = {
                    'bucket': BUCKET,
                    'key': key,
                    'useAI': use_ai,
                }
                response = lambda_client.invoke(
                    FunctionName=LAMBDA_FUNCTION,
                    InvocationType='Event',  # Async invocation
                    Payload=json.dumps(invoke_payload),
                )
                return make_response(202, {'status': 'conversion_started', 'key': key})
            except Exception as e:
                return make_response(500, {'error': f'Failed to invoke renderer: {str(e)}'})

        if method == 'POST' and action == 'export-docx':
            # Body: {"originalKey": "uploads/uuid/file.docx", "patches": [...]}
            payload, error = parse_json_body(event)
            if error:
                return make_response(400, {'error': error})
            original_key = payload.get('originalKey')
            patches = payload.get('patches', [])
            if not original_key:
                return make_response(400, {'error': 'originalKey is required'})

            try:
                import tempfile, os
                apply_patches = _import_backend_module('docx_writer').apply_patches

                with tempfile.TemporaryDirectory() as tmp:
                    src = os.path.join(tmp, 'original.docx')
                    try:
                        s3.download_file(BUCKET, original_key, src)
                    except ClientError as exc:
                        return make_response(404 if exc.response['Error']['Code'] in ('404', 'NoSuchKey') else 502,
                                             {'error': _s3_user_message(exc, 'Download')})
                    with open(src, 'rb') as fh:
                        docx_bytes = fh.read()

                out_bytes = apply_patches(docx_bytes, patches)

                base = os.path.splitext(os.path.basename(original_key))[0]
                key_dir = os.path.dirname(original_key).replace('uploads/', 'exported/', 1)
                out_key = f"{key_dir}/{base}-exported.docx"
                s3.put_object(
                    Bucket=BUCKET,
                    Key=out_key,
                    Body=out_bytes,
                    ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                )
                download_url = s3.generate_presigned_url(
                    'get_object',
                    Params={'Bucket': BUCKET, 'Key': out_key},
                    ExpiresIn=URL_EXPIRATION,
                )
                return make_response(200, {'url': download_url, 'key': out_key})
            except ModuleNotFoundError as exc:
                logger.error('export-docx: module import failed: %s', exc)
                return make_response(500, {'error': 'Server configuration error — deployment may be incomplete'})
            except Exception as exc:
                logger.error('export-docx failed: %s', exc, exc_info=True)
                return make_response(500, {'error': f'Export failed: {exc}'})

        if method == 'POST' and action == 'extract':
            # Body: {"key":"uploads/uuid/file.docx","useAI":true,"force":false}
            payload, error = parse_json_body(event)
            if error:
                return make_response(400, {'error': error})
            key = payload.get('key')
            use_ai = payload.get('useAI', True)
            force = payload.get('force', False)

            if not key:
                return make_response(400, {'error': 'key is required'})

            cache_key = extracted_fields_key(key)
            if not force and object_exists(BUCKET, cache_key):
                obj = s3.get_object(Bucket=BUCKET, Key=cache_key)
                cached = json.loads(obj['Body'].read().decode('utf-8'))
                return make_response(200, {'cached': True, 'key': cache_key, 'fields': cached.get('fields', [])})

            try:
                import tempfile, os
                build_docx_ast = _import_backend_module('docx_ast').build_docx_ast
                field_extractor = _import_backend_module('field_extractor')
                extract_fields_from_ast = field_extractor.extract_fields_from_ast
                extract_fields_from_ast_with_ai = field_extractor.extract_fields_from_ast_with_ai

                with tempfile.TemporaryDirectory() as tmp:
                    src = os.path.join(tmp, 'original.docx')
                    try:
                        s3.download_file(BUCKET, key, src)
                    except ClientError as exc:
                        return make_response(404 if exc.response['Error']['Code'] in ('404', 'NoSuchKey') else 502,
                                             {'error': _s3_user_message(exc, 'Download')})
                    try:
                        ast = build_docx_ast(src)
                    except Exception as exc:
                        logger.error('extract: AST build failed for key=%s: %s', key, exc)
                        return make_response(422, {'error': f'Could not parse document: {exc}'})

                if use_ai:
                    extracted = extract_fields_from_ast_with_ai(ast)
                else:
                    extracted = extract_fields_from_ast(ast)

                s3.put_object(
                    Bucket=BUCKET,
                    Key=cache_key,
                    Body=json.dumps(extracted).encode('utf-8'),
                    ContentType='application/json',
                )
                return make_response(200, {'cached': False, 'key': cache_key, 'fields': extracted.get('fields', [])})
            except ModuleNotFoundError as exc:
                logger.error('extract: module import failed: %s', exc)
                return make_response(500, {'error': 'Server configuration error — deployment may be incomplete'})
            except Exception as exc:
                logger.error('extract failed: %s', exc, exc_info=True)
                return make_response(500, {'error': f'Extraction failed: {exc}'})

        if method == 'POST':
            # expect JSON body with fileName and optional contentType
            payload, error = parse_json_body(event)
            if error:
                return make_response(400, {'error': error})
            file_name = payload.get('fileName')
            content_type = payload.get('contentType', 'application/octet-stream')
            if not file_name:
                return make_response(400, {'error': 'fileName is required'})

            uid = str(uuid.uuid4())
            key = f"uploads/{uid}/{file_name}"

            params = {'Bucket': BUCKET, 'Key': key}
            # include content type so S3 enforces it on PUT
            params_put = params.copy()
            params_put['ContentType'] = content_type

            upload_url = s3.generate_presigned_url('put_object', Params=params_put, ExpiresIn=URL_EXPIRATION)

            return make_response(200, {'uploadUrl': upload_url, 'key': key})

        elif method == 'GET':
            # status check: query string contains key
            qs = event.get('rawQueryString') if 'rawQueryString' in event else None
            key = None
            if qs:
                # parse rawQueryString
                parsed = parse_qs(qs)
                key = parsed.get('key', [None])[0]
            else:
                # fallback for API Gateway v1
                params = event.get('queryStringParameters') or {}
                key = params.get('key')

            if not key:
                return make_response(400, {'error': 'key query parameter is required'})

            out_key = converted_html_key(key)

            if not object_exists(BUCKET, out_key):
                return make_response(404, {'status': 'pending'})

            # object exists - return presigned GET for both converted HTML and original upload
            get_url = s3.generate_presigned_url('get_object', Params={'Bucket': BUCKET, 'Key': out_key}, ExpiresIn=URL_EXPIRATION)
            original_url = s3.generate_presigned_url('get_object', Params={'Bucket': BUCKET, 'Key': key}, ExpiresIn=URL_EXPIRATION)
            return make_response(200, {'status': 'ready', 'resultUrl': get_url, 'originalUrl': original_url, 'key': out_key})

        else:
            return make_response(405, {'error': 'method not allowed'})

    except Exception as e:
        logger.error('Unhandled error: %s', e, exc_info=True)
        return make_response(500, {'error': str(e)})
