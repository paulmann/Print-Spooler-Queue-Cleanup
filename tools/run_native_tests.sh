#!/usr/bin/env bash
# Run platform-independent unit tests on the host compiler (Linux/macOS/MinGW).
set -euo pipefail
cd "$(dirname "$0")/.."
CXX=${CXX:-g++}
OUT=$(mktemp -d)/psqc_tests
$CXX -std=c++20 -O1 -Wall -Wextra -Wno-unused-parameter -finput-charset=UTF-8 \
   -I src tests/native/test_pure.cpp src/pure.cpp src/i18n_tables_a.cpp src/i18n_tables_b.cpp src/i18n_tables_c.cpp \
   -o "$OUT"
"$OUT"
