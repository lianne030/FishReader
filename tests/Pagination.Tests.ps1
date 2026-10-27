param([string]$SourcePath = (Join-Path $PSScriptRoot '..\FishReader.ps1'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($SourcePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors -join "`n") }
$names = @('Format-PageText','Get-PageViewport','Test-PageFit','Get-PageEnd','Current-Offset','Reset-PageLayout','Get-PreviousPageStart','Render-Page','Next-Page','Prev-Page','Refresh-FromOffset')
foreach ($fn in $ast.FindAll({param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $false)) {
  if ($fn.Name -in $names) { . ([scriptblock]::Create($fn.Extent.Text)) }
}
function Save-Settings {}
function Save-CurrentProgress([int]$offset) {}
$settings = @{ ParagraphSpacing = 1 }
$window = [pscustomobject]@{ Width = 300.0; Height = 180.0 }
$shell = New-Object Windows.Controls.Border
$shell.Padding = New-Object Windows.Thickness(14)
$textBlock = New-Object Windows.Controls.TextBlock
$textBlock.FontFamily = 'Microsoft YaHei UI'
$textBlock.FontSize = 22
$textBlock.TextWrapping = 'Wrap'
$textBlock.LineStackingStrategy = 'MaxHeight'
$textBlock.LineHeight = 34.76
$shell.Child = $textBlock
function Set-TestSize([double]$width, [double]$height) {
  $window.Width = $width; $window.Height = $height
  $shell.Measure((New-Object Windows.Size($width, $height)))
  $shell.Arrange((New-Object Windows.Rect(0, 0, $width, $height)))
}
function Assert-Visible([int]$start, [int]$end) {
  $rendered = Format-PageText $bookText.Substring($start, $end - $start)
  $view = New-Object Windows.Controls.TextBlock
  $view.FontFamily = $textBlock.FontFamily; $view.FontSize = $textBlock.FontSize
  $view.LineHeight = $textBlock.LineHeight; $view.LineStackingStrategy = 'MaxHeight'; $view.TextWrapping = 'Wrap'
  $view.Text = $rendered
  $width = $shell.ActualWidth - 28; $height = $shell.ActualHeight - 28
  $view.Measure((New-Object Windows.Size($width, [double]::PositiveInfinity)))
  if ($view.DesiredSize.Height -gt $height) {
    throw "Clipped page [$start,$end): text height $($view.DesiredSize.Height), available $height"
  }
}
$cjk = [string][char]0x4e2d + [char]0x6587 + [char]0x3002
$cases = @(
  (($cjk + "`n") * 100),
  (($cjk * 18 + "`n`n") * 14),
  (('abcdefghijklmnopqrstuvwxyz 0123456789 ' + $cjk) * 40),
  (($cjk + "`n`n`n`n") * 50)
)
$pages = 0
foreach ($size in @(@(300,180), @(220,120), @(460,300), @(301.5,181.5))) {
  Set-TestSize $size[0] $size[1]
  foreach ($text in $cases) {
    $script:bookText = $text
    $cursor = 0; $joined = New-Object Text.StringBuilder
    while ($cursor -lt $bookText.Length) {
      $end = Get-PageEnd $cursor
      if ($end -le $cursor -or $end -gt $bookText.Length) { throw 'Invalid page boundary' }
      Assert-Visible $cursor $end
      [void]$joined.Append($bookText.Substring($cursor, $end - $cursor))
      $cursor = $end; $pages++
    }
    if ($joined.ToString() -cne $bookText) { throw 'Pagination lost source text' }
  }
}
Write-Output "PASS: $pages complete pages fit their viewport and reconstruct the original text."

# Test actual navigation, including the short page before an arbitrary saved offset.
Set-TestSize 300 180
$script:bookText = $cases[1]
$script:pageStarts = @(0); $script:pageIndex = 0; $script:previousPageStarts = @{}
Render-Page
for ($turn = 0; $turn -lt 8; $turn++) {
  $expected = $script:renderedPageEnd
  Next-Page
  if ((Current-Offset) -ne $expected) { throw 'Next page skipped the displayed boundary' }
  if ($textBlock.Text -cne (Format-PageText $bookText.Substring((Current-Offset), $script:renderedPageEnd - (Current-Offset)))) { throw 'Displayed text differs from page text' }
  Assert-Visible (Current-Offset) $script:renderedPageEnd
}
Refresh-FromOffset 47
Prev-Page
if ($script:renderedPageEnd -ne 47) { throw 'Previous page overlapped the saved offset' }
Next-Page
if ((Current-Offset) -ne 47) { throw 'Could not return to saved offset' }
Set-TestSize 220 120
Reset-PageLayout
$expected = $script:renderedPageEnd
Next-Page
if ((Current-Offset) -ne $expected) { throw 'Resize skipped text' }
Prev-Page
Assert-Visible (Current-Offset) $script:renderedPageEnd
Write-Output 'PASS: forward/backward, arbitrary saved offset and resized pages remain contiguous.'

# Preserve indentation and trailing spaces rather than trimming source characters.
if ((Format-PageText '  text  ') -cne '  text  ') { throw 'Formatting removed page edge spaces' }
foreach ($font in @(12, 22, 36, 52)) {
  foreach ($spacing in @(1.05, 1.58, 2.5)) {
    $textBlock.FontSize = $font; $textBlock.LineHeight = $font * $spacing
    Set-TestSize 460 300
    foreach ($paragraph in @(0, 1, 4)) {
      $settings.ParagraphSpacing = $paragraph
      $script:bookText = $cases[1]
      $cursor = 0
      for ($sample = 0; $sample -lt 5; $sample++) {
        $end = Get-PageEnd $cursor
        Assert-Visible $cursor $end
        $cursor = $end
      }
    }
  }
}
Write-Output 'PASS: font sizes 12/22/36/52, line spacing 1.05/1.58/2.5 and paragraph spacing 0/1/4.'
