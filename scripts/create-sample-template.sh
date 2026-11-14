#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
export DOTNET_ROOT="${DOTNET_ROOT:-$HOME/.dotnet}"
export PATH="$DOTNET_ROOT:$PATH"

TMP="$ROOT/artifacts/tmp-template-gen"
mkdir -p "$TMP" "$ROOT/Templates"

cat > "$TMP/Program.cs" <<'EOF'
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

var outPath = args[0];
using var doc = WordprocessingDocument.Create(outPath, WordprocessingDocumentType.Document);
var main = doc.AddMainDocumentPart();
main.Document = new Document(new Body(
    new Paragraph(new Run(new RunProperties(new Bold()), new Text("${name}"))),
    new Paragraph(new Run(new Text("${role}"))),
    new Paragraph(new Run(new Text("${linkedin}"))),
    new Paragraph(new Run(new Text("${summary}"))),
    new Paragraph(new Run(new Text("Experience"))),
    new Paragraph(new Run(new Text("${ArrayStart experience}"))),
    new Paragraph(new Run(new Text("${experience_i.title} — ${experience_i.company} (${experience_i.years})"))),
    new Paragraph(new Run(new Text("${experience_i.description}"))),
    new Paragraph(new Run(new Text("${ArrayEnd experience}")))
));
main.Document.Save();
Console.WriteLine("Wrote " + outPath);
EOF

cat > "$TMP/TemplateGen.csproj" <<'EOF'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
  </ItemGroup>
</Project>
EOF

dotnet run --project "$TMP/TemplateGen.csproj" -c Release -- "$ROOT/Templates/modern.docx"
echo "Sample template ready: Templates/modern.docx"
