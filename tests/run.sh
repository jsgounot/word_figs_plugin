#!/bin/sh
# Runs the core unit tests with Node if available, else macOS JavaScriptCore.
cd "$(dirname "$0")/.." || exit 1
JSC=/System/Library/Frameworks/JavaScriptCore.framework/Versions/Current/Helpers/jsc
if command -v node >/dev/null 2>&1; then
  node -e "require('./core.js'); require('./tests/core.test.js');"
elif [ -x "$JSC" ]; then
  "$JSC" core.js tests/core.test.js
else
  echo "Need node or JavaScriptCore (jsc) to run tests" >&2
  exit 1
fi
