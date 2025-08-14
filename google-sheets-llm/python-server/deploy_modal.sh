#!/usr/bin/env bash
set -euo pipefail

# Usage: ./deploy_modal.sh [ENV_FILE]
# Defaults to ./modal.env in the same directory

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ENV_FILE="${1:-$SCRIPT_DIR/modal.env}"

if [[ ! -f "$ENV_FILE" ]]; then
  echo "Env file not found: $ENV_FILE"
  echo "Create one with MODAL_TOKEN_ID and MODAL_TOKEN_SECRET variables."
  exit 1
fi

# shellcheck source=/dev/null
source "$ENV_FILE"

if [[ -z "${MODAL_TOKEN_ID:-}" || -z "${MODAL_TOKEN_SECRET:-}" ]]; then
  echo "MODAL_TOKEN_ID or MODAL_TOKEN_SECRET is empty. Please set them in $ENV_FILE"
  exit 1
fi

# Set modal token locally (writes to ~/.modal)
modal token set --token-id "$MODAL_TOKEN_ID" --token-secret "$MODAL_TOKEN_SECRET"

# Deploy the Modal server
modal deploy "$SCRIPT_DIR/python-server/modal_server.py"

echo "Deployment complete. Update your Chrome extension background.js with the endpoint shown by 'modal app list' or deploy output."
