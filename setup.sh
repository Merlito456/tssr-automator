#!/bin/bash
set -e
python -m playwright install chromium
python -m playwright install-deps chromium || true
