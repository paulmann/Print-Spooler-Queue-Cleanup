#!/usr/bin/env bash
# Cross-compile SpoolerCleanup.exe with MinGW-w64 (Linux CI / dev box).
# Produces dist/SpoolerCleanup.exe — fully static, zero runtime dependencies.
set -euo pipefail
cd "$(dirname "$0")/.."
CXX=${CXX:-x86_64-w64-mingw32-g++-posix}
WINDRES=${WINDRES:-x86_64-w64-mingw32-windres}
STD=${STD:-c++20}
mkdir -p dist
SRC=$(ls src/*.cpp)
FLAGS="-std=$STD -O2 -s -Wall -Wextra -Wno-unused-parameter -finput-charset=UTF-8 -fwide-exec-charset=UTF-16LE"
DEFS="-DUNICODE -D_UNICODE -DWIN32_LEAN_AND_MEAN -DNOMINMAX -D_WIN32_WINNT=0x0A00"
LIBS="-municode -mwindows -static -static-libgcc -static-libstdc++ \
      -lcomctl32 -lshell32 -lshlwapi -luxtheme -ldwmapi -lwinspool -lwinhttp \
      -ladvapi32 -lole32 -luuid -loleaut32 -lgdi32 -luser32 -lkernel32"
echo "[rc]  res/app.rc"
$WINDRES -I res -O coff -o dist/app.res.o res/app.rc
echo "[cxx] $SRC"
$CXX $FLAGS $DEFS -I src $SRC dist/app.res.o -o dist/SpoolerCleanup.exe $LIBS
echo "[ok]  dist/SpoolerCleanup.exe  $(stat -c%s dist/SpoolerCleanup.exe) bytes"
