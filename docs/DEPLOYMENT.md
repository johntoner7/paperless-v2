# Deployment guidance — Lambda container and alternatives

This document covers deploying the `paperless-be` renderer as a serverless function on AWS. It also outlines tradeoffs and an alternative using AWS Fargate (ECS).

## Recommended: Lambda Container (for small 1–2 page docs)
# Deployment guidance — Lambda container (detailed)

This document explains a step-by-step Lambda container deployment for the `paperless-be` renderer. It assumes you have small input documents (1–2 pages). If you later need heavier processing or longer time, see the Fargate notes at the end.

Prerequisites
- AWS account and credentials configured via `aws configure`.
- `docker` installed and working locally.
- `aws` CLI v2 installed.
- The repo contains `Dockerfile.lambda` and `lambda_function.py` (added).

High-level flow
1. Build container image locally (includes LibreOffice and your code).
2. Push image to ECR.
3. Create an IAM role for Lambda that allows S3 access and CloudWatch logging (and VPC ENI creation if using EFS).
4. Create or update the Lambda function with the container image.
5. (Optional) Mount EFS for extra disk, or configure S3 event / API Gateway trigger.
6. Test by invoking Lambda with a sample S3 object or direct invoke.

Step-by-step commands

1) Build the image locally

```bash
# from repo root
docker build -f Dockerfile.lambda -t paperless-be:lambda .
```

2) Create ECR repo (once) and push image

```bash
# create ECR repo (idempotent)
aws ecr create-repository --repository-name paperless-be || true

# helper vars
acct=$(aws sts get-caller-identity --query Account --output text)
region=$(aws configure get region)
repo_uri="$acct.dkr.ecr.$region.amazonaws.com/paperless-be:latest"

# login and push
aws ecr get-login-password --region $region | docker login --username AWS --password-stdin $acct.dkr.ecr.$region.amazonaws.com
# Force a single-image push. buildx can publish an OCI index unless provenance/SBOM are disabled.
docker buildx build --platform linux/amd64 --provenance=false --sbom=false --push -f Dockerfile.lambda -t $repo_uri .
```

3) Create an IAM role for Lambda

You can create a role in the console or via CLI. The role needs a trust policy that allows Lambda to assume the role, and an inline policy granting S3 & CloudWatch permissions. If you mount EFS (Lambda in VPC), also attach `AWSLambdaVPCAccessExecutionRole` or the equivalent permissions.

Trust policy (save as `trust.json`):

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Principal": {"Service": "lambda.amazonaws.com"},
      "Action": "sts:AssumeRole"
    }
  ]
}
```

Basic permissions policy (save as `lambda-s3-policy.json`):

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {"Effect":"Allow","Action":["s3:GetObject","s3:PutObject","s3:ListBucket"],"Resource":["arn:aws:s3:::your-bucket","arn:aws:s3:::your-bucket/*"]},
    {"Effect":"Allow","Action":["logs:CreateLogGroup","logs:CreateLogStream","logs:PutLogEvents"],"Resource":"*"},
    {"Effect":"Allow","Action":["ec2:CreateNetworkInterface","ec2:DescribeNetworkInterfaces","ec2:DeleteNetworkInterface"],"Resource":"*"}
  ]
}
```

Create role and attach policy:

```bash
aws iam create-role --role-name paperless-lambda-role --assume-role-policy-document file://trust.json
aws iam put-role-policy --role-name paperless-lambda-role --policy-name paperless-s3-policy --policy-document file://lambda-s3-policy.json
# Optionally attach AWSLambdaVPCAccessExecutionRole managed policy for VPC/EFS
aws iam attach-role-policy --role-name paperless-lambda-role --policy-arn arn:aws:iam::aws:policy/service-role/AWSLambdaVPCAccessExecutionRole
```

4) Create or update the Lambda function

```bash
# first time create
aws lambda create-function \
  --function-name paperless-renderer \
  --package-type Image \
  --code ImageUri=$repo_uri \
  --role arn:aws:iam::$acct:role/paperless-lambda-role \
  --timeout 300 \
  --memory-size 2048 || \
aws lambda update-function-code --function-name paperless-renderer --image-uri $repo_uri
```

Notes:
- Set `--timeout` to a value that covers your conversion time (max 900 seconds). For 1–2 pages, 60–120s is usually enough but start conservative.
- `--memory-size` affects CPU allocation; try `2048` and raise if the renderer is slow.
- If Lambda still reports an unsupported media type, re-push with `--provenance=false --sbom=false` and verify ECR shows `application/vnd.docker.distribution.manifest.v2+json` or `application/vnd.oci.image.manifest.v1+json`, not an OCI index.

5) (Optional) Mount EFS for more disk

- Create an EFS filesystem and one or more mount targets in the same VPC/subnets as your Lambda functions.
- Create an EFS access point and note its ARN.
- Configure security groups: Lambda function's security group must be able to talk to EFS mount targets over NFS (port 2049).

Mount EFS to Lambda (example):

```bash
# after creating EFS access point, attach when updating function config
aws lambda update-function-configuration --function-name paperless-renderer \
  --file-system-configs Arn=arn:aws:elasticfilesystem:REGION:ACCOUNT_ID:access-point/AP_ID,LocalMountPath=/mnt/efs \
  --vpc-config SubnetIds="subnet-aaa,subnet-bbb",SecurityGroupIds="sg-xxxx"
```

6) Configure triggers

- S3: Create an S3 event notification on the bucket/prefix to invoke the Lambda on `s3:ObjectCreated:*`.
- API Gateway: Create an API (HTTP API) that invokes the Lambda for on-demand conversions.
- SQS: Use SQS events for queued/batched processing; have a small Lambda that polls SQS and runs conversions (helps smoothing traffic).

---

## Presigned S3 + presign Lambda (recommended for production)

This flow is robust and scales: the frontend uploads directly to S3 using a presigned PUT URL, and the main `paperless-renderer` Lambda is triggered by S3 `ObjectCreated` events on the `uploads/` prefix. The renderer writes outputs to `converted/` so the frontend can fetch the result via a presigned GET URL.

1) Deploy a small presign Lambda (simple Python zip) that exposes a Function URL. It returns a presigned PUT URL for `uploads/<uuid>/<filename>` on POST and reports conversion status on GET `?key=uploads/<uuid>/<filename>`.

Example (from repo root):

```bash
# package handler + pdf modules + pymupdf (linux/x86_64 wheel for Lambda runtime)
rm -rf /tmp/presign-pkg && mkdir /tmp/presign-pkg
pip3 install \
  --platform manylinux2014_x86_64 \
  --target /tmp/presign-pkg \
  --python-version 3.11 \
  --only-binary=:all: \
  "pymupdf>=1.24.0"
cp backend/presign_handler.py backend/docx_writer.py backend/docx_node_ids.py \
   backend/docx_ast.py backend/field_extractor.py \
   backend/pdf_field_extractor.py backend/pdf_exporter.py \
   /tmp/presign-pkg/
cd /tmp/presign-pkg && zip -r "$OLDPWD/presign.zip" . -x "*.pyc" -x "__pycache__/*" -x "*.dist-info/*" && cd -

# create function (reuse `paperless-lambda-role` or create a minimal role that allows writing/reading the bucket and logs)
acct=$(aws sts get-caller-identity --query Account --output text)
role_arn=arn:aws:iam::$acct:role/paperless-lambda-role
aws lambda create-function \
  --function-name paperless-presign \
  --runtime python3.11 \
  --handler presign_handler.lambda_handler \
  --zip-file fileb://presign.zip \
  --role $role_arn \
  --environment Variables={PRESIGN_BUCKET=your-bucket,PRESIGN_EXPIRATION=300} || \
aws lambda update-function-configuration --function-name paperless-presign --environment Variables={PRESIGN_BUCKET=your-bucket,PRESIGN_EXPIRATION=300}

# create a public (or authenticated) Function URL and enable CORS for testing
aws lambda create-function-url-config --function-name paperless-presign --auth-type NONE --cors AllowOrigins='*',AllowMethods='GET,POST,OPTIONS',AllowHeaders='*'

# get the function URL
aws lambda get-function-url-config --function-name paperless-presign --query 'FunctionUrl' --output text
```

Notes:
- `PRESIGN_BUCKET` must be set to the bucket you will use for `uploads/` and `converted/` prefixes.
- The presign Lambda only needs permission to generate presigned URLs. In practice, reusing the `paperless-lambda-role` (which already has `s3:GetObject`/`s3:PutObject`) is the simplest setup.

2) Configure S3 notification to trigger the renderer on `uploads/` puts (or use the S3 console):

```bash
# sample notification JSON saved to notify.json
cat > notify.json <<'JSON'
{
  "LambdaFunctionConfigurations": [
    {
      "Id": "paperless-upload-trigger",
      "LambdaFunctionArn": "arn:aws:lambda:REGION:ACCOUNT_ID:function:paperless-renderer",
      "Events": ["s3:ObjectCreated:Put"],
      "Filter": {"Key": {"FilterRules": [{"Name":"prefix","Value":"uploads/"}]}}
    }
  ]
}

aws s3api put-bucket-notification-configuration --bucket your-bucket --notification-configuration file://notify.json
```

3) Converter Lambda responsibility
- When invoked (with the S3 `bucket` and `key`), the renderer should read `uploads/.../name.docx`, process it, and write the result to `converted/.../name.html` (same subpath under `converted/`). This is what `presign_handler.py` expects when generating presigned GET URLs.

4) Frontend configuration
- Set `NEXT_PUBLIC_PRESIGN_URL` in your frontend environment (`.env.local` for local dev, or AWS Amplify Environment variables for the hosted build) to the Function URL produced above. The frontend will POST `{ fileName, contentType }` to get `{ uploadUrl, key }`, then PUT the file to `uploadUrl`, then poll `GET ${PRESIGN_URL}?key=${encodeURIComponent(key)}` until ready.
- In Amplify, add the variable before building the app and redeploy after changing it, because this repo’s Next.js frontend is statically exported and reads the value at build time.

Example `.env.local` (frontend folder):

```
NEXT_PUBLIC_PRESIGN_URL=https://<your-function-url>
```

5) Security and production notes
- For production, do NOT leave the Function URL open (`--auth-type NONE`) — require auth (IAM authorizer, JWT via API Gateway, or a short-lived token). For quick testing, `NONE` + CORS is convenient.
- Ensure the renderer Lambda's IAM role has `s3:GetObject`, `s3:PutObject` for `arn:aws:s3:::your-bucket/uploads/*` and `arn:aws:s3:::your-bucket/converted/*`.
- Consider lifecycle rules to expire `uploads/` and `converted/` objects after a retention period.

---

7) Test the function

Upload a sample file to S3 and invoke the function directly with a test payload:

```bash
# upload
aws s3 cp sample.docx s3://your-bucket/path/sample.docx

# invoke by passing JSON with bucket and key
aws lambda invoke --function-name paperless-renderer --payload '{"bucket":"your-bucket","key":"path/sample.docx"}' response.json
cat response.json

# tail logs while running
aws logs tail /aws/lambda/paperless-renderer --follow
```

Local testing tips
- Run the container locally to validate LibreOffice and renderer availability:

```bash
docker run --rm -v "$PWD":/var/task -v "$PWD/test-data":/tmp/input paperless-be:lambda \
  python3 -c "import subprocess; subprocess.check_call(['python3','/var/task/docx_renderer.py','/tmp/input/sample.docx','/tmp/input/out.html']); print('done')"
```

Troubleshooting
- Check CloudWatch logs: `aws logs tail /aws/lambda/<name> --follow`.
- If conversions fail due to missing fonts, add fonts packages (e.g., `fonts-dejavu-core`) to `Dockerfile.lambda`.
- If Lambda times out, raise `--timeout` and/or `--memory-size`.
- Large container images increase cold-start time — prune unnecessary packages and locales.

IAM & security notes
- Use least-privilege: replace `your-bucket` in the policy with the specific bucket ARNs.
- For production, consider creating a dedicated service role with minimal S3 scope and logging permissions.

Observability & cleanup
- Configure CloudWatch log retention, and S3 lifecycle rules to clean up `converted/` prefixes after a retention period.

Alternative: Fargate
- If you need longer running jobs or more disk/CPU, build the same Docker image, push to ECR, and run on Fargate. Fargate removes the 15 minute Lambda limit and is easier to scale for heavy workloads.

If you'd like, I can now:
- create a simple `deploy.sh` that runs the ECR build/push and Lambda create/update commands, or
- walk you through creating the IAM role in the console step-by-step.
