#!/bin/zsh
set -euo pipefail

ROOT="/Users/sebbo/Desktop/Newsletter"
GENERATOR="$ROOT/build_newsletter.py"
DEFAULT_SHEET_URL="https://docs.google.com/spreadsheets/d/1ukXTu8PXWHGe4Fzg5rA424BmN6Bi3FPAWwA12QybNpI/edit"

SHEET_URL="${NEWSLETTER_SHEET_URL:-$DEFAULT_SHEET_URL}"
WATCH_MODE="false"
OPEN_AFTER="false"
INTERVAL_SECONDS="${NEWSLETTER_WATCH_INTERVAL:-20}"

usage() {
  cat <<EOF
Usage:
  $ROOT/generate_newsletter.sh [options]

Options:
  --sheet URL       Google Sheet URL or ID (default: configured project sheet)
  --watch           Regenerate continuously every --interval seconds
  --interval N      Watch interval in seconds (default: 20, minimum: 5)
  --open            Open newsletter.html after the first successful generation
  --help            Show this help

Examples:
  $ROOT/generate_newsletter.sh
  $ROOT/generate_newsletter.sh --watch --interval 15
  $ROOT/generate_newsletter.sh --sheet "https://docs.google.com/spreadsheets/d/<ID>/edit"
EOF
}

run_once() {
  python3 "$GENERATOR" --google-sheet "$SHEET_URL"
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --sheet)
      if [[ $# -lt 2 ]]; then
        echo "Error: --sheet requires a value."
        exit 1
      fi
      SHEET_URL="$2"
      shift 2
      ;;
    --watch)
      WATCH_MODE="true"
      shift
      ;;
    --interval)
      if [[ $# -lt 2 ]]; then
        echo "Error: --interval requires a value."
        exit 1
      fi
      INTERVAL_SECONDS="$2"
      shift 2
      ;;
    --open)
      OPEN_AFTER="true"
      shift
      ;;
    --help|-h)
      usage
      exit 0
      ;;
    *)
      echo "Error: unknown option '$1'."
      usage
      exit 1
      ;;
  esac
done

if ! [[ "$INTERVAL_SECONDS" =~ ^[0-9]+$ ]]; then
  echo "Error: --interval must be a number (seconds)."
  exit 1
fi

if (( INTERVAL_SECONDS < 5 )); then
  echo "Error: --interval must be at least 5 seconds."
  exit 1
fi

if [[ "$WATCH_MODE" == "true" ]]; then
  echo "Watch mode started."
  echo "Sheet: $SHEET_URL"
  echo "Interval: ${INTERVAL_SECONDS}s"
  echo "Stop with Ctrl+C."
  while true; do
    if run_once; then
      echo "[$(date '+%Y-%m-%d %H:%M:%S')] Newsletter regenerated."
      if [[ "$OPEN_AFTER" == "true" ]]; then
        open "$ROOT/newsletter.html"
        OPEN_AFTER="false"
      fi
    else
      echo "[$(date '+%Y-%m-%d %H:%M:%S')] Generation failed."
    fi
    sleep "$INTERVAL_SECONDS"
  done
else
  run_once
  echo "Output: $ROOT/newsletter.html"
  if [[ "$OPEN_AFTER" == "true" ]]; then
    open "$ROOT/newsletter.html"
  fi
fi
