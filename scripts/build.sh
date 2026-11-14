#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
export DOTNET_ROOT="${DOTNET_ROOT:-$HOME/.dotnet}"
export PATH="$DOTNET_ROOT:$PATH"

OUT="$ROOT/artifacts/CvGenerator-framework"
cd "$ROOT"

if [[ "$(uname -s)" != MINGW* && "$(uname -s)" != CYGWIN* && "$(uname -s)" != MSYS* && "${OS:-}" != "Windows_NT" ]]; then
  echo "Note: WPF app publish requires Windows. Running library tests only on this OS."
  dotnet test tests/CvGenerator.Core.Tests/CvGenerator.Core.Tests.csproj -c Release --nologo
  dotnet test tests/CvGenerator.Docx.Tests/CvGenerator.Docx.Tests.csproj -c Release --nologo
  echo "Re-run this script on Windows to publish the UI."
  exit 0
fi

dotnet restore CvGenerator.sln
dotnet test tests/CvGenerator.Core.Tests/CvGenerator.Core.Tests.csproj -c Release --nologo
dotnet test tests/CvGenerator.Docx.Tests/CvGenerator.Docx.Tests.csproj -c Release --nologo
dotnet publish src/CvGenerator.App/CvGenerator.App.csproj \
  -c Release \
  -r win-x64 \
  --self-contained false \
  -o "$OUT"

mkdir -p "$OUT/Templates" "$OUT/Output"
cp -n "$ROOT/Templates/"*.docx "$OUT/Templates/" 2>/dev/null || true
echo "Published (framework-dependent win-x64) to $OUT"
