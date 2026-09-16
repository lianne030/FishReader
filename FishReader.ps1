<#
FishReader — portable transparent desktop reader for Windows.
Run through 启动阅读器.cmd. All state is stored beside this script.
#>
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.IO.Compression.FileSystem

$ErrorActionPreference = 'Stop'
$appRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataRoot = Join-Path $appRoot 'FishReaderData'
New-Item -ItemType Directory -Force -Path $dataRoot | Out-Null
$settingsFile = Join-Path $dataRoot 'settings.json'

Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ReaderNative {
  [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
  [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
  [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
  [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
}
'@

function Get-DefaultSettings {
  @{ FontSize = 22; FontFamily = 'Microsoft YaHei UI'; Foreground = '#D3C6AA'; LineSpacing = 1.58; ParagraphSpacing = 1; Progress = @(); Opacity = 1.0; Width = 460; Height = 300; Left = 100; Top = 100; AlwaysOnTop = $true; Recent = @(); LastBook = ''; LastOffset = 0; ClickThrough = $false }
}
function Load-Settings {
  $d = Get-DefaultSettings
  if (Test-Path $settingsFile) {
    try { $saved = Get-Content $settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json; foreach ($p in $saved.PSObject.Properties) { $d[$p.Name] = $p.Value } } catch {}
  }
  return $d
}
$settings = Load-Settings
function Save-Settings {
  $settings | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsFile -Encoding UTF8
}

function Decode-Html([string]$value) {
  $value = [regex]::Replace($value, '(?is)<script.*?</script>|<style.*?</style>', '')
  $value = [regex]::Replace($value, '(?i)</?(p|div|br|h[1-6]|li|tr|section|article)[^>]*>', "`n")
  $value = [regex]::Replace($value, '(?is)<[^>]+>', '')
  $value = [System.Net.WebUtility]::HtmlDecode($value)
  return [regex]::Replace($value, "`r?`n[ `t]*`r?`n[ `t]*`r?`n+", "`n`n").Trim()
}
function Read-PlainText([string]$path) {
  $bytes = [IO.File]::ReadAllBytes($path)
  if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return [Text.Encoding]::UTF8.GetString($bytes,3,$bytes.Length-3) }
  if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($bytes,2,$bytes.Length-2) }
  $utf8 = [Text.Encoding]::UTF8.GetString($bytes)
  if ($utf8.Contains([char]0xfffd)) { return [Text.Encoding]::GetEncoding(936).GetString($bytes) }
  return $utf8
}
function Resolve-ZipPath([string]$basePath, [string]$ref) {
  $parts = New-Object Collections.Generic.List[string]
  foreach($part in (($basePath + '/' + $ref) -replace '\\','/' -split '/')) {
    if ($part -eq '..') { if ($parts.Count) { $parts.RemoveAt($parts.Count-1) } } elseif ($part -and $part -ne '.') { $parts.Add($part) }
  }
  return ($parts -join '/')
}
function Read-Epub([string]$path) {
  $zip = [IO.Compression.ZipFile]::OpenRead($path)
  try {
    $container = $zip.GetEntry('META-INF/container.xml'); if (!$container) { throw '不是有效的 EPUB 文件。' }
    $sr = New-Object IO.StreamReader($container.Open()); [xml]$cx = $sr.ReadToEnd(); $sr.Dispose()
    $opfPath = ([string]$cx.container.rootfiles.rootfile.'full-path') -replace '\\','/'
    $opfEntry = $zip.GetEntry($opfPath); if (!$opfEntry) { throw 'EPUB 目录文件丢失。' }
    $sr = New-Object IO.StreamReader($opfEntry.Open()); [xml]$opf = $sr.ReadToEnd(); $sr.Dispose()
    $opfDir = [IO.Path]::GetDirectoryName($opfPath) -replace '\\','/'
    $manifest = @{}; foreach($item in $opf.package.manifest.item) { $manifest[[string]$item.id] = [string]$item.href }
    $chapters = @(); $all = New-Object Text.StringBuilder
    foreach($itemref in $opf.package.spine.itemref) {
      $href = $manifest[[string]$itemref.idref]; if (!$href) { continue }
      $entry = $zip.GetEntry((Resolve-ZipPath $opfDir $href)); if (!$entry) { continue }
      $reader = New-Object IO.StreamReader($entry.Open()); $html = $reader.ReadToEnd(); $reader.Dispose(); $txt = Decode-Html $html
      if (!$txt) { continue }; $start = $all.Length; [void]$all.AppendLine($txt); [void]$all.AppendLine()
      $title = ($txt -split "`n" | Where-Object { $_.Trim() } | Select-Object -First 1).Trim(); if (!$title) { $title = "章节 $($chapters.Count+1)" }
      $chapters += [pscustomobject]@{ Title=$title; Start=$start }
    }
    return @{ Text=$all.ToString().Trim(); Chapters=$chapters }
  } finally { $zip.Dispose() }
}
function Read-Docx([string]$path) {
  $zip = [IO.Compression.ZipFile]::OpenRead($path)
  try { $e=$zip.GetEntry('word/document.xml'); if(!$e){throw 'DOCX 内容缺失。'}; $r=New-Object IO.StreamReader($e.Open()); [xml]$x=$r.ReadToEnd(); $r.Dispose(); return @{Text=(($x.SelectNodes("//*[local-name()='p']") | ForEach-Object { $_.InnerText }) -join "`n"); Chapters=@()} } finally {$zip.Dispose()}
}
function Get-SavedProgress([string]$path) {
  $entry=@($settings.Progress | Where-Object { $_.Path -eq $path } | Select-Object -Last 1)
  if($entry.Count){return [int]$entry[0].Offset}
  if($settings.LastBook -eq $path){return [int]$settings.LastOffset}
  return 0
}
function Save-CurrentProgress([int]$offset) {
  if(!$bookPath){return}
  $entries=@($settings.Progress | Where-Object { $_.Path -ne $bookPath })
  $entries += [pscustomobject]@{Path=$bookPath;Offset=$offset;Updated=(Get-Date).ToString('o')}
  $settings.Progress=@($entries | Select-Object -Last 200)
}
function Load-Book([string]$path) {
  if (!(Test-Path -LiteralPath $path)) { throw '找不到该文件。' }
  $ext=[IO.Path]::GetExtension($path).ToLowerInvariant()
  switch ($ext) {
    '.epub' { $result=Read-Epub $path }
    '.docx' { $result=Read-Docx $path }
    '.html' { $result=@{Text=(Decode-Html (Read-PlainText $path)); Chapters=@()} }
    '.htm' { $result=@{Text=(Decode-Html (Read-PlainText $path)); Chapters=@()} }
    '.fb2' { [xml]$x=Read-PlainText $path; $result=@{Text=$x.InnerText; Chapters=@()} }
    '.rtf' { $raw=Read-PlainText $path; $result=@{Text=([regex]::Replace($raw,'\\[a-z]+-?\d* ?|[{}]','')); Chapters=@()} }
    '.md' { $result=@{Text=([regex]::Replace((Read-PlainText $path),'(?m)^#{1,6}\s*|\*\*|__|`','')); Chapters=@()} }
    default { $result=@{Text=(Read-PlainText $path); Chapters=@()} }
  }
  if (!$result.Text.Trim()) { throw '没有提取到可阅读的文字。' }
  if (!$result.Chapters.Count) {
    $found = [regex]::Matches($result.Text, '(?m)^\s*(第[一二三四五六七八九十百千0-9]+[章节卷回部篇]|Chapter\s+\d+|序章|楔子|前言|后记|番外).*$')
    $result.Chapters=@($found | ForEach-Object { [pscustomobject]@{Title=$_.Value.Trim();Start=$_.Index} })
  }
  $script:bookPath=$path; $script:bookText=$result.Text -replace "`r`n","`n"; $script:chapters=@($result.Chapters); $savedOffset=Get-SavedProgress $path; $script:pageStarts=@([math]::Max(0,[math]::Min($savedOffset,$script:bookText.Length-1))); $script:pageIndex=0; $script:previousPageStarts=@{}
  $settings.LastBook=$path; $settings.LastOffset=$savedOffset
  $settings.Recent=@($path)+@($settings.Recent | Where-Object { $_ -ne $path -and (Test-Path -LiteralPath $_) } | Select-Object -First 9); Save-Settings
}

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" AllowsTransparency="True" Background="Transparent" WindowStyle="None" ResizeMode="CanResize" ShowInTaskbar="False" Topmost="True">
  <Border x:Name="Shell" Background="#01000000" Padding="14" CornerRadius="3">
    <TextBlock x:Name="ReaderText" TextWrapping="Wrap" FontFamily="Microsoft YaHei UI" Foreground="#D3C6AA" LineStackingStrategy="BlockLineHeight" LineHeight="35"/>
  </Border>
</Window>
'@
$reader = New-Object Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
$textBlock = $window.FindName('ReaderText'); $shell = $window.FindName('Shell')
$window.Width=[double]$settings.Width; $window.Height=[double]$settings.Height; $window.Left=[double]$settings.Left; $window.Top=[double]$settings.Top; $window.Topmost=[bool]$settings.AlwaysOnTop
$textBlock.FontSize=[double]$settings.FontSize; $textBlock.FontFamily=$settings.FontFamily; $textBlock.Foreground=(New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($settings.Foreground))); $textBlock.LineHeight=$textBlock.FontSize*[double]$settings.LineSpacing; $window.Opacity=[double]$settings.Opacity
function Set-Foreground([string]$color) {
  try { $textBlock.Foreground=(New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($color))); $settings.Foreground=$color; Save-Settings } catch { [System.Windows.MessageBox]::Show('颜色格式无效。','FishReader') }
}
function Set-LineSpacing([double]$spacing) {
  $spacing=[math]::Max(1.05,[math]::Min(2.5,$spacing)); $settings.LineSpacing=$spacing; $textBlock.LineHeight=$textBlock.FontSize*$spacing; Reset-PageLayout; Save-Settings
}
function Set-ParagraphSpacing([int]$spacing) {
  $settings.ParagraphSpacing=[math]::Max(0,[math]::Min(4,$spacing)); Reset-PageLayout; Save-Settings
}
function Format-PageText([string]$value) {
  $count=[int]$settings.ParagraphSpacing+1; $break=(1..$count | ForEach-Object { "`n" }) -join ''
  return [regex]::Replace($value, "`n[ `t]*`n+", $break).Trim()
}


function Set-ClickThrough([bool]$enabled) {
  $script:settings.ClickThrough=$enabled
  if (!$script:hwnd) { return }
  $style=[ReaderNative]::GetWindowLong($script:hwnd, -20)
  if ($enabled) { $style=$style -bor 0x20 } else { $style=$style -band (-bnot 0x20) }
  [void][ReaderNative]::SetWindowLong($script:hwnd,-20,$style); Save-Settings
}
function Test-PageFit([int]$start, [int]$length, [double]$maxWidth, [double]$maxHeight) {
  $probe = New-Object Windows.Controls.TextBlock
  $probe.TextWrapping = 'Wrap'; $probe.FontFamily = $textBlock.FontFamily; $probe.FontSize = $textBlock.FontSize
  $probe.LineStackingStrategy = 'BlockLineHeight'; $probe.LineHeight = $textBlock.LineHeight
  $probe.Text = Format-PageText $bookText.Substring($start, $length)
  $probe.Measure((New-Object Windows.Size($maxWidth, [double]::PositiveInfinity)))
  return $probe.DesiredSize.Height -le ($maxHeight + 0.5)
}
function Get-PageEnd([int]$start) {
  $remaining = $bookText.Length - $start; if ($remaining -le 0) { return $bookText.Length }
  $layoutWidth = if ($window.ActualWidth -gt 40) { $window.ActualWidth } else { $window.Width }
  $layoutHeight = if ($window.ActualHeight -gt 40) { $window.ActualHeight } else { $window.Height }
  $maxWidth = [math]::Max(1, $layoutWidth - 28); $maxHeight = [math]::Max(1, $layoutHeight - 28)
  $rough = [int][math]::Max(256, [math]::Ceiling((($maxWidth / $textBlock.FontSize) * ($maxHeight / $textBlock.LineHeight)) * 2))
  $low = 0; $high = [math]::Min($remaining, $rough)
  while ((Test-PageFit $start $high $maxWidth $maxHeight)) {
    $low = $high; if ($high -eq $remaining) { return $bookText.Length }
    $high = [math]::Min($remaining, $high * 2)
  }
  # WPF may report an unavailable layout during startup or monitor/DPI changes.
  # Keep pagination usable with a conservative fallback instead of degrading to one character per page.
  if ($low -eq 0) {
    $fallbackColumns = [math]::Max(8, [math]::Floor($maxWidth / ($textBlock.FontSize * 1.2)))
    $fallbackLines = [math]::Max(1, [math]::Floor($maxHeight / $textBlock.LineHeight) - 1)
    $fallbackLength = [int][math]::Max(20, [math]::Floor($fallbackColumns * $fallbackLines * 0.75))
    return [math]::Min($bookText.Length, $start + $fallbackLength)
  }
  while (($high - $low) -gt 1) {
    $middle = [int](($low + $high) / 2)
    if (Test-PageFit $start $middle $maxWidth $maxHeight) { $low = $middle } else { $high = $middle }
  }
  $end = $start + $low
  # Prefer a natural sentence boundary only when it is already very close to the measured page end.
  # Searching too far backwards leaves the final line visibly half empty.
  $snapRange = [math]::Min(4, $end - $start)
  $cut = $bookText.LastIndexOfAny([char[]]@("`n",'。','！','？','；',' '), $end - 1, $snapRange)
  if ($cut -ge $start -and ($end - ($cut + 1)) -le 4) { return $cut + 1 }
  return $end
}
function Current-Offset { return [int]$pageStarts[$pageIndex] }
function Reset-PageLayout {
  if (!$bookText) { return }
  $current = Current-Offset
  $script:pageStarts = @($current)
  $script:pageIndex = 0
  $script:previousPageStarts = @{}
  Render-Page
}
function Get-PreviousPageStart([int]$current) {
  if ($current -le 0) { return $null }
  $key = [string]$current
  if ($previousPageStarts.ContainsKey($key)) { return [int]$previousPageStarts[$key] }
  # Rebuild the route from the beginning only when arriving here via a chapter jump or app restart.
  $cursor = 0
  while ($cursor -lt $current) {
    $next = Get-PageEnd $cursor
    if ($next -le $cursor) { break }
    $previousPageStarts[[string]$next] = $cursor
    if ($next -ge $current) { $previousPageStarts[$key] = $cursor; return $cursor }
    $cursor = $next
  }
  return $null
}
function Render-Page {
  if (!$bookText) { $textBlock.Text="右键 → 打开小说`n`nCtrl + Alt + H：关闭阅读器`nCtrl + Alt + T：鼠标穿透"; return }
  $start=Current-Offset; $end=Get-PageEnd $start; $textBlock.Text=Format-PageText $bookText.Substring($start,$end-$start)
  $settings.LastOffset=$start; Save-CurrentProgress $start; Save-Settings
}
function Next-Page {
  if(!$bookText){return}
  $current=Current-Offset; $end=Get-PageEnd $current
  if($end -ge $bookText.Length){return}
  $cachedNext = if($pageIndex -lt $pageStarts.Count-1){[int]$pageStarts[$pageIndex+1]}else{-1}
  if($cachedNext -eq $end){
    $script:pageIndex++
  } else {
    # The layout changed: discard every stale forward boundary so no text can be skipped.
    $script:pageStarts=@($pageStarts[0..$pageIndex])
    $previousPageStarts[[string]$end]=$current
    $script:pageStarts+=,$end
    $script:pageIndex++
  }
  Render-Page
}
function Prev-Page {
  if($pageIndex -gt 0){$script:pageIndex--;Render-Page;return}
  $current=Current-Offset; $previous=Get-PreviousPageStart $current
  if($null -ne $previous -and $previous -lt $current){$script:pageStarts=@($previous,$current);$script:pageIndex=0;Render-Page}
}
function Refresh-FromOffset([int]$offset) { $script:pageStarts=@([math]::Max(0,[math]::Min($offset,$bookText.Length-1)));$script:pageIndex=0;$script:previousPageStarts=@{};Render-Page }
function Current-Chapter {
  $c=$chapters | Where-Object {$_.Start -le (Current-Offset)} | Select-Object -Last 1
  if($c){return $c.Title}; return '未识别章节'
}
function Add-Menu([Windows.Controls.ContextMenu]$menu,[string]$label,[scriptblock]$action,[bool]$checked=$false) { $i=New-Object Windows.Controls.MenuItem; $i.Header=$label; $i.IsCheckable=$checked; if($checked){$i.IsChecked=$true}; $i.Add_Click($action); [void]$menu.Items.Add($i); return $i }
function Add-Seperator($menu){ [void]$menu.Items.Add((New-Object Windows.Controls.Separator)) }
function Show-ReaderMenu {
  $menu=New-Object Windows.Controls.ContextMenu
  $head=New-Object Windows.Controls.MenuItem; $head.Header=if($bookPath){"$([IO.Path]::GetFileNameWithoutExtension($bookPath))  ·  $(Current-Chapter)"}else{'FishReader · 尚未打开小说'}; $head.IsEnabled=$false; [void]$menu.Items.Add($head); Add-Seperator $menu
  Add-Menu $menu '打开小说…' { $d=New-Object System.Windows.Forms.OpenFileDialog; $d.Filter='支持的小说文件|*.txt;*.epub;*.html;*.htm;*.md;*.fb2;*.rtf;*.docx|所有文件|*.*'; if($d.ShowDialog() -eq 'OK'){try{Load-Book $d.FileName;Render-Page}catch{[System.Windows.MessageBox]::Show($_.Exception.Message,'FishReader')}} } | Out-Null
  if($bookText){ Add-Menu $menu '上一页' {Prev-Page} | Out-Null; Add-Menu $menu '下一页' {Next-Page} | Out-Null; Add-Menu $menu '上一章' {$c=$chapters | Where-Object {$_.Start -lt (Current-Offset)} | Select-Object -Last 1;if($c){Refresh-FromOffset $c.Start}} | Out-Null; Add-Menu $menu '下一章' {$c=$chapters | Where-Object {$_.Start -gt (Current-Offset)} | Select-Object -First 1;if($c){Refresh-FromOffset $c.Start}} | Out-Null }
  $recent=New-Object Windows.Controls.MenuItem; $recent.Header='最近阅读'; foreach($p in @($settings.Recent)){if(Test-Path -LiteralPath $p){$mi=New-Object Windows.Controls.MenuItem;$mi.Header=[IO.Path]::GetFileName($p);$mi.Tag=$p;$mi.Add_Click({try{Load-Book $this.Tag;Render-Page}catch{}});[void]$recent.Items.Add($mi)}};if(!$recent.Items.Count){$recent.IsEnabled=$false};[void]$menu.Items.Add($recent); Add-Seperator $menu
  $font=New-Object Windows.Controls.MenuItem;$font.Header='字体大小'; foreach($delta in @(-2,2)){ $mi=New-Object Windows.Controls.MenuItem;$mi.Header=if($delta -gt 0){'增大字号'}else{'减小字号'};$mi.Tag=$delta;$mi.Add_Click({$textBlock.FontSize=[math]::Max(12,[math]::Min(52,$textBlock.FontSize+[double]$this.Tag));$textBlock.LineHeight=$textBlock.FontSize*[double]$settings.LineSpacing;$settings.FontSize=$textBlock.FontSize;Reset-PageLayout;Save-Settings});[void]$font.Items.Add($mi)};[void]$menu.Items.Add($font)
  $spacing=New-Object Windows.Controls.MenuItem;$spacing.Header=('行间距（当前 {0:N2}）' -f [double]$settings.LineSpacing); foreach($item in @(@('紧凑  1.25',1.25),@('标准  1.58',1.58),@('宽松  1.90',1.90),@('更紧凑',-0.1),@('更宽松',0.1))){$mi=New-Object Windows.Controls.MenuItem;$mi.Header=$item[0];$mi.Tag=$item[1];$mi.Add_Click({if([double]$this.Tag -gt 1){Set-LineSpacing ([double]$this.Tag)}else{Set-LineSpacing ([double]$settings.LineSpacing+[double]$this.Tag)}});[void]$spacing.Items.Add($mi)};[void]$menu.Items.Add($spacing)
  $paragraph=New-Object Windows.Controls.MenuItem;$paragraph.Header=('段间距（当前 {0}）' -f [int]$settings.ParagraphSpacing); foreach($item in @(@('紧凑  无额外空行',0),@('标准  一行空白',1),@('宽松  两行空白',2),@('更紧凑',-1),@('更宽松',1))){$mi=New-Object Windows.Controls.MenuItem;$mi.Header=$item[0];$mi.Tag=$item[1];$mi.Add_Click({if($this.Header -like '更*'){Set-ParagraphSpacing ([int]$settings.ParagraphSpacing+[int]$this.Tag)}else{Set-ParagraphSpacing ([int]$this.Tag)}});[void]$paragraph.Items.Add($mi)};[void]$menu.Items.Add($paragraph)
  $colors=New-Object Windows.Controls.MenuItem;$colors.Header='字体颜色'; foreach($item in @(@('暖白','#FFF3E7'),@('纯白','#FFFFFF'),@('护眼绿','#C7F9CC'),@('浅蓝','#CDE8FF'),@('淡黄','#FFF1B8'),@('柔紫','#E9D5FF'))){$mi=New-Object Windows.Controls.MenuItem;$mi.Header=$item[0];$mi.Tag=$item[1];$mi.Add_Click({Set-Foreground ([string]$this.Tag)});[void]$colors.Items.Add($mi)};$custom=New-Object Windows.Controls.MenuItem;$custom.Header='自定义颜色…';$custom.Add_Click({$d=New-Object System.Windows.Forms.ColorDialog;$d.FullOpen=$true;if($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){$c=$d.Color;Set-Foreground ('#{0:X2}{1:X2}{2:X2}' -f $c.R,$c.G,$c.B)}});[void]$colors.Items.Add($custom);[void]$menu.Items.Add($colors)
  $size=New-Object Windows.Controls.MenuItem;$size.Header='显示区域'; foreach($pair in @(@(80,0),@(-80,0),@(0,60),@(0,-60))){$mi=New-Object Windows.Controls.MenuItem;$mi.Header=if($pair[0]-gt 0){'加宽'}elseif($pair[0]-lt 0){'变窄'}elseif($pair[1]-gt 0){'加高'}else{'变矮'};$mi.Tag=$pair;$mi.Add_Click({$window.Width=[math]::Max(220,$window.Width+[double]$this.Tag[0]);$window.Height=[math]::Max(120,$window.Height+[double]$this.Tag[1]);Reset-PageLayout});[void]$size.Items.Add($mi)};[void]$menu.Items.Add($size)
  Add-Menu $menu '窗口置顶' {$window.Topmost=-not $window.Topmost;$settings.AlwaysOnTop=$window.Topmost;Save-Settings} ([bool]$window.Topmost) | Out-Null
  Add-Menu $menu '锁定/解锁鼠标穿透 (Ctrl+Alt+T)' {Set-ClickThrough (-not [bool]$settings.ClickThrough)} ([bool]$settings.ClickThrough) | Out-Null
  Add-Menu $menu '关闭阅读器 (Ctrl+Alt+H)' {$window.Close()} | Out-Null
  Add-Seperator $menu; Add-Menu $menu '退出' {$window.Close()} | Out-Null
  $window.ContextMenu=$menu; $menu.IsOpen=$true
}

$window.Add_MouseLeftButtonDown({if(!$settings.ClickThrough){if($_.ClickCount -ge 2){$window.Close();$_.Handled=$true}else{$window.DragMove()}}})
$window.Add_MouseRightButtonUp({if(!$settings.ClickThrough){Show-ReaderMenu}})
$window.Add_MouseWheel({if($_.Delta -lt 0){Next-Page}else{Prev-Page};$_.Handled=$true})
$window.Add_KeyDown({if($_.Key -in @('PageDown','Space','Right','Down')){Next-Page}elseif($_.Key -in @('PageUp','Left','Up')){Prev-Page}elseif($_.Key -eq 'Escape'){$window.Close()}})
$window.Add_SizeChanged({if($bookText){Reset-PageLayout}})
$window.Add_Loaded({if($bookText){Render-Page}})
$window.Add_SourceInitialized({
  $script:hwnd=(New-Object Windows.Interop.WindowInteropHelper($window)).Handle
  [void][ReaderNative]::RegisterHotKey($script:hwnd,1,3,0x48); [void][ReaderNative]::RegisterHotKey($script:hwnd,2,3,0x54)
  $source=[Windows.Interop.HwndSource]::FromHwnd($script:hwnd); $source.AddHook([Windows.Interop.HwndSourceHook]{param($hwnd,$msg,$w,$l,[ref]$handled) if($msg -eq 0x0312){if($w.ToInt32() -eq 1){$window.Close()}elseif($w.ToInt32() -eq 2){Set-ClickThrough (-not [bool]$settings.ClickThrough)};$handled.Value=$true};return [IntPtr]::Zero})
  if($settings.ClickThrough){Set-ClickThrough $true}
})
$window.Add_Closing({$settings.Width=$window.Width;$settings.Height=$window.Height;$settings.Left=$window.Left;$settings.Top=$window.Top;$settings.FontSize=$textBlock.FontSize;$settings.LineSpacing=$textBlock.LineHeight/$textBlock.FontSize;$settings.AlwaysOnTop=$window.Topmost;Save-Settings;if($script:hwnd){[void][ReaderNative]::UnregisterHotKey($script:hwnd,1);[void][ReaderNative]::UnregisterHotKey($script:hwnd,2)}})

if($settings.LastBook -and (Test-Path -LiteralPath $settings.LastBook)){try{Load-Book $settings.LastBook;Refresh-FromOffset ([int]$settings.LastOffset)}catch{Render-Page}}else{Render-Page}
[void]$window.ShowDialog()
