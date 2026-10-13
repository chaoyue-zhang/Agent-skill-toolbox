# Export Word documents to PDF for visual verification (Word COM; Windows + MS Word).
#
# Usage:
#   pwsh -NoProfile -File export_pdf_preview.ps1 -Docs "D:\产物\01 项目建议书.docx","D:\产物\02 任务书.doc" [-OutDir "D:\preview"]
#   Default OutDir: each document's own folder. Accepts .doc and .docx.
#
# Production pitfalls handled here:
# - ExportAsFixedFormat SILENTLY FAILS when the target PDF already exists (no
#   error raised, stale file kept, so you end up verifying the previous version)
#   -> the old PDF is deleted first.
# - Documents open ReadOnly, so a user session holding the file (or a leftover
#   ~$ lock file) does not block the export.
# - When generating a PS1 like this from Python, write it with
#   codecs.open(path, 'w', 'utf-8-sig') (BOM) or PowerShell may misparse Chinese.
#
# Render pages afterwards (PyMuPDF) and inspect with the read_image tool:
#   import fitz; d = fitz.open(pdf)
#   for i, p in enumerate(d): p.get_pixmap(dpi=150).save('page-%02d.png' % (i + 1))
param(
    [Parameter(Mandatory = $true)][string[]]$Docs,
    [string]$OutDir
)

$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($d in $Docs) {
        $full = (Resolve-Path -LiteralPath $d).Path
        $dir = if ($OutDir) { $OutDir } else { Split-Path -Path $full -Parent }
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $pdf = Join-Path $dir ([IO.Path]::GetFileNameWithoutExtension($full) + '.pdf')
        if (Test-Path -LiteralPath $pdf) {
            Remove-Item -LiteralPath $pdf -Force   # same-name overwrite fails silently
        }
        $doc = $word.Documents.Open($full, $false, $true)   # ConfirmConversions=false, ReadOnly=true
        $doc.ExportAsFixedFormat($pdf, 17)                  # 17 = wdExportFormatPDF
        $doc.Close($false)
        Write-Host "OK -> $pdf"
    }
}
finally {
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
