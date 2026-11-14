#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
export DOTNET_ROOT="${DOTNET_ROOT:-$HOME/.dotnet}"
export PATH="$DOTNET_ROOT:$PATH"

OUT="$ROOT/artifacts/CvGenerator"
cd "$ROOT"

if [[ "$(uname -s)" != MINGW* && "$(uname -s)" != CYGWIN* && "$(uname -s)" != MSYS* && "${OS:-}" != "Windows_NT" ]]; then
  echo "Self-contained win-x64 publish must run on Windows (WPF)."
  echo "Running library tests on this OS instead."
  dotnet test tests/CvGenerator.Core.Tests/CvGenerator.Core.Tests.csproj -c Release --nologo
  dotnet test tests/CvGenerator.Docx.Tests/CvGenerator.Docx.Tests.csproj -c Release --nologo
  exit 1
fi

dotnet restore CvGenerator.sln
dotnet test tests/CvGenerator.Core.Tests/CvGenerator.Core.Tests.csproj -c Release --nologo
dotnet test tests/CvGenerator.Docx.Tests/CvGenerator.Docx.Tests.csproj -c Release --nologo
dotnet publish src/CvGenerator.App/CvGenerator.App.csproj \
  -c Release \
  -r win-x64 \
  --self-contained true \
  -p:PublishSingleFile=true \
  -p:IncludeNativeLibrariesForSelfExtract=true \
  -p:EnableCompressionInSingleFile=true \
  -o "$OUT"

mkdir -p "$OUT/Templates" "$OUT/Output"
cp -n "$ROOT/Templates/"*.docx "$OUT/Templates/" 2>/dev/null || true
echo "Published self-contained single-file win-x64 to $OUT"
echo "Copy the CvGenerator folder to other Windows PCs (no .NET install required)."
echo "Requires: WebView2 Evergreen + LibreOffice for PDF."
