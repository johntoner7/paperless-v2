import os
import json
import uuid
import boto3
from urllib.parse import parse_qs
from botocore.exceptions import ClientError

_region = os.environ.get('AWS_DEFAULT_REGION', 'eu-west-1')
s3 = boto3.client('s3', region_name=_region,
                  endpoint_url=f'https://s3.{_region}.amazonaws.com')

BUCKET = os.environ.get('PRESIGN_BUCKET')
if not BUCKET:
    raise RuntimeError('PRESIGN_BUCKET environment variable must be set')

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


def converted_html_key(source_key):
    base_no_ext = os.path.splitext(source_key)[0]
    if base_no_ext.startswith('uploads/'):
        return base_no_ext.replace('uploads/', 'converted/', 1) + '.html'
    return f'converted/{base_no_ext}.html'


def object_exists(bucket, key):
    try:
        s3.head_object(Bucket=bucket, Key=key)
        return True
    except ClientError as exc:
        error_code = exc.response.get('Error', {}).get('Code')
        if error_code in {'404', 'NoSuchKey', 'NotFound'}:
            return False
        raise


def list_uploaded_files():
    paginator = s3.get_paginator('list_objects_v2')
    items = []

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
                'status': 'ready' if object_exists(BUCKET, html_key) else 'pending',
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

        if method == 'POST':
            # expect JSON body with fileName and optional contentType
            body = event.get('body') or ''
            if event.get('isBase64Encoded'):
                import base64
                body = base64.b64decode(body).decode('utf-8')
            payload = json.loads(body)
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
        return make_response(500, {'error': str(e)})
