$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
$pptPath = Join-Path $project 'presentation\soju_regional_rank_2026.pptx'
$pdfPath = Join-Path $env:TEMP 'soju_regional_rank_2026_qa.pdf'
$app = New-Object -ComObject PowerPoint.Application
try {
  $deck = $app.Presentations.Open($pptPath, $true, $false, $false)
  if ($deck.Slides.Count -ne 10) { throw "Expected 10 slides, found $($deck.Slides.Count)" }
  $allText = foreach ($slide in $deck.Slides) {
    foreach ($shape in $slide.Shapes) {
      if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $shape.TextFrame.TextRange.Text }
    }
  }
  $text = $allText -join "`n"
  foreach ($required in @('참이슬 후레쉬','경상권','진로','가정용 소주류 구매액 순위')) {
    if (-not $text.Contains($required)) { throw "PPTX missing required text: $required" }
  }
  $deck.SaveAs($pdfPath, 32)
  Write-Output "PPTX opened in PowerPoint; $($deck.Slides.Count) slides; PDF render: $pdfPath"
  $deck.Close()
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
