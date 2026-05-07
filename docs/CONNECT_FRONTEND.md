# Connecting the Frontend to AWS Lambda

## What exists

| Component | Status | Details |
|-----------|--------|---------|
| `paperless-presign` Lambda | **Deployed** | eu-west-1 |
| API Gateway HTTP API | **Live** | `zadr4m6dej.execute-api.eu-west-1.amazonaws.com` |
| `paperless-renderer` Lambda | Pending | Needs ECR build + deploy (see below) |
| Frontend (`/frontend`) | Ready | Next.js app, wired to `NEXT_PUBLIC_PRESIGN_URL` |
| S3 bucket | `paperless-fe` | Used for uploads/ and converted/ prefixes |

## End-to-end flow

```
User picks .docx
    → POST /presign-lambda  { fileName, contentType }
    ← { uploadUrl, key }
    → PUT uploadUrl  (file goes direct to S3)
    → S3 fires paperless-renderer Lambda
    → renderer writes converted/<uuid>/file.html to S3
    → frontend polls GET /presign-lambda?key=uploads/<uuid>/file.docx
    ← { status: "ready", resultUrl }
    → fetch resultUrl → display HTML in editor
```

## Deployed resources

| Resource | Value |
|----------|-------|
| API Gateway endpoint | `https://zadr4m6dej.execute-api.eu-west-1.amazonaws.com` |
| S3 bucket | `paperless-fe` (eu-west-1) |
| IAM role | `arn:aws:iam::165044447372:role/paperless-lambda-role` |

## Step 1 — Configure the frontend (done once)

Create `frontend/.env.local`:

```
NEXT_PUBLIC_PRESIGN_URL=https://zadr4m6dej.execute-api.eu-west-1.amazonaws.com
```

Then run the dev server:

```bash
cd frontend
npm run dev
```

## Step 2 — Deploy the renderer Lambda (container)

The renderer converts DOCX → HTML inside a Docker container with LibreOffice.

```bash
# From repo root
acct=165044447372
region=eu-west-1
repo_uri="$acct.dkr.ecr.$region.amazonaws.com/paperless-be:latest"

# Create ECR repo (idempotent)
aws ecr create-repository --repository-name paperless-be --region $region || true

# Authenticate Docker to ECR
aws ecr get-login-password --region $region \
  | docker login --username AWS --password-stdin $acct.dkr.ecr.$region.amazonaws.com

# Build and push (linux/amd64 required for Lambda)
docker buildx build \
  --platform linux/amd64 \
  --provenance=false \
  --sbom=false \
  --push \
  -f Dockerfile.lambda \
  -t $repo_uri .

# Create or update the Lambda
aws lambda create-function \
  --function-name paperless-renderer \
  --package-type Image \
  --code ImageUri=$repo_uri \
  --role arn:aws:iam::$acct:role/paperless-lambda-role \
  --timeout 300 \
  --memory-size 2048 \
  --region $region \
|| aws lambda update-function-code \
  --function-name paperless-renderer \
  --image-uri $repo_uri \
  --region $region
```

## Step 3 — Wire S3 → renderer Lambda trigger

Run once after both the bucket and renderer Lambda exist:

```bash
acct=165044447372
region=eu-west-1
bucket=paperless-fe
renderer_arn=arn:aws:lambda:$region:$acct:function:paperless-renderer

# Allow S3 to invoke the renderer
aws lambda add-permission \
  --function-name paperless-renderer \
  --statement-id allow-s3-trigger \
  --action lambda:InvokeFunction \
  --principal s3.amazonaws.com \
  --source-arn arn:aws:s3:::$bucket \
  --region $region

# Set S3 notification on uploads/ prefix
cat > /tmp/notify.json <<'JSON'
{
  "LambdaFunctionConfigurations": [
    {
      "Id": "paperless-upload-trigger",
      "LambdaFunctionArn": "RENDERER_ARN",
      "Events": ["s3:ObjectCreated:Put"],
      "Filter": {
        "Key": {
          "FilterRules": [{"Name": "prefix", "Value": "uploads/"}]
        }
      }
    }
  ]
}
JSON
# Replace placeholder with real ARN
sed -i '' "s|RENDERER_ARN|$renderer_arn|" /tmp/notify.json

aws s3api put-bucket-notification-configuration \
  --bucket $bucket \
  --notification-configuration file:///tmp/notify.json
```

## Step 4 — Test end-to-end

```bash
# Invoke presign Lambda to get an upload URL
curl -s -X POST \
  https://zadr4m6dej.execute-api.eu-west-1.amazonaws.com/ \
  -H 'Content-Type: application/json' \
  -d '{"fileName":"test.docx","contentType":"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}'

# You should get back: {"uploadUrl": "https://s3...", "key": "uploads/<uuid>/test.docx"}

# Upload directly to the presigned URL
curl -X PUT "<uploadUrl>" \
  -H 'Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document' \
  --data-binary @invoice.docx

# Poll for result (replace KEY with the key returned above)
curl -s "https://zadr4m6dej.execute-api.eu-west-1.amazonaws.com/?key=uploads/..."
# Returns {"status":"pending"} until conversion is done, then {"status":"ready","resultUrl":"..."}
```

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `PRESIGN_BUCKET env var not set` | Re-deploy presign Lambda with correct env |
| Upload PUT returns 403 | S3 bucket policy or CORS may block; check bucket CORS config |
| Renderer Lambda not triggering | Check S3 notification config and Lambda resource policy |
| Conversion timeout | Raise renderer Lambda `--timeout` (max 900s) and `--memory-size` |
| Cold start slow | First invocation may take 10-30s; subsequent calls are fast |

## Updating the presign Lambda

```bash
cd /Users/johntoner/paperless-be
zip -j presign.zip backend/presign_handler.py
aws lambda update-function-code \
  --function-name paperless-presign \
  --zip-file fileb://presign.zip \
  --region eu-west-1
```
