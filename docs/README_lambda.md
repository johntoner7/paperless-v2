# Lambda container: paperless-be renderer

This document explains how to build and push a Lambda container image that includes LibreOffice and the repository's Python renderer, and how to deploy it as an AWS Lambda function.

Files added for this workflow:
- `Dockerfile.lambda` — container image that installs LibreOffice and the app.
- `lambda_function.py` — Lambda handler that accepts an `event` with `bucket` and `key` or a vanilla S3 event, downloads the docx, runs the renderer, uploads output, and returns a presigned URL.

Quick local build (for testing with Docker):

```bash
# build image
docker build -f Dockerfile.lambda -t paperless-be:lambda .

# run a container for quick sanity (you can mount a local input dir)
docker run --rm -v "$PWD":/var/task paperless-be:lambda python3 -c "import sys; print('image OK')"
```

Push to ECR and deploy to Lambda (see `DEPLOYMENT.md` for full commands).

Notes:
- Lambda has a 15 minute timeout and 512 MB `/tmp` by default. For larger working storage or long conversions, mount EFS to the Lambda function.
- Keep the image small to avoid slow cold starts: remove unneeded locale packs and tools if you can.
