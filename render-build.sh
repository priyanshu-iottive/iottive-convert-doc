#!/usr/bin/env bash
# exit on error
set -o errexit

npm install --include=dev
npx tsx script/build.ts
pip install python-docx
