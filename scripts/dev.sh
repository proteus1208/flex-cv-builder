#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
export DOTNET_ROOT="${DOTNET_ROOT:-$HOME/.dotnet}"
export PATH="$DOTNET_ROOT:$PATH"

cd "$ROOT"

if [[ "$(uname -s)" != MINGW* && "$(uname -s)" != CYGWIN* && "$(uname -s)" != MSYS* && "${OS:-}" != "Windows_NT" ]]; then
  echo "CvGenerator.App is a Windows WPF app. Run this script on Windows."
  echo "On Linux you can run: dotnet test tests/CvGenerator.Core.Tests && dotnet test tests/CvGenerator.Docx.Tests"
  exit 1
fi

dotnet restore CvGenerator.sln
dotnet run --project src/CvGenerator.App/CvGenerator.App.csproj -c Debug
