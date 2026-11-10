# 秦巴药都数据平台 - 本地静态服务（PowerShell 内置实现，无需安装 Python / Node）
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File server.ps1 -Port 8000
#       powershell -NoProfile -ExecutionPolicy Bypass -File server.ps1 -SyncOnly   # 只同步数据快照，不启动服务
#       powershell -NoProfile -ExecutionPolicy Bypass -File server.ps1 -NoSync     # 跳过启动时的快照同步
#       powershell -NoProfile -ExecutionPolicy Bypass -File server.ps1 -SyncOnly -Force  # 强制重建全部快照
param(
  [int]$Port = 8000,
  [string]$Root = "",
  [switch]$SyncOnly,
  [switch]$NoSync,
  [switch]$Force
)

$ErrorActionPreference = 'Continue'
if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Split-Path -Parent $MyInvocation.MyCommand.Path }
$Root = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\')
$script:LogPath = Join-Path $env:TEMP 'qbd-server.log'

# 用于把 Excel 窗口句柄换算成进程号，Quit 后若进程赖着不走就精确收掉自己启动的那一个
if (-not ('Native.Win32' -as [type])) {
  Add-Type -Namespace Native -Name Win32 -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
'@
}

$mime = @{
  '.html' = 'text/html; charset=utf-8'
  '.htm'  = 'text/html; charset=utf-8'
  '.js'   = 'application/javascript; charset=utf-8'
  '.mjs'  = 'application/javascript; charset=utf-8'
  '.css'  = 'text/css; charset=utf-8'
  '.json' = 'application/json; charset=utf-8'
  '.txt'  = 'text/plain; charset=utf-8'
  '.md'   = 'text/markdown; charset=utf-8'
  '.csv'  = 'text/csv; charset=utf-8'
  '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  '.xls'  = 'application/vnd.ms-excel'
  '.png'  = 'image/png'
  '.jpg'  = 'image/jpeg'
  '.jpeg' = 'image/jpeg'
  '.gif'  = 'image/gif'
  '.webp' = 'image/webp'
  '.svg'  = 'image/svg+xml'
  '.ico'  = 'image/x-icon'
  '.pdf'  = 'application/pdf'
  '.woff' = 'font/woff'
  '.woff2'= 'font/woff2'
  '.ttf'  = 'font/ttf'
}

function Write-Log([string]$msg) {
  try { Add-Content -LiteralPath $script:LogPath -Value ("[{0}] {1}" -f (Get-Date -Format 'MM-dd HH:mm:ss'), $msg) -Encoding UTF8 } catch { }
}

function Read-Shared([string]$path) {
  $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
  try {
    $ms = New-Object System.IO.MemoryStream
    $fs.CopyTo($ms)
    return $ms.ToArray()
  } finally {
    $fs.Dispose()
  }
}

function Send-Raw($stream, [byte[]]$head, [byte[]]$body) {
  $stream.Write($head, 0, $head.Length)
  if ($body -and $body.Length -gt 0) { $stream.Write($body, 0, $body.Length) }
  $stream.Flush()
}

function Send-Text($stream, [int]$code, [string]$text) {
  $status = switch ($code) { 200 { 'OK' } 403 { 'Forbidden' } 404 { 'Not Found' } default { 'Internal Server Error' } }
  $body = [System.Text.Encoding]::UTF8.GetBytes($text)
  $head = "HTTP/1.1 $code $status`r`nContent-Type: text/plain; charset=utf-8`r`nContent-Length: $($body.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n`r`n"
  Send-Raw $stream ([System.Text.Encoding]::ASCII.GetBytes($head)) $body
}

function Send-Json($stream, [int]$code, [string]$json) {
  $status = switch ($code) { 200 { 'OK' } 400 { 'Bad Request' } 403 { 'Forbidden' } 404 { 'Not Found' } 500 { 'Internal Server Error' } default { 'OK' } }
  $body = [System.Text.Encoding]::UTF8.GetBytes($json)
  $head = "HTTP/1.1 $code $status`r`nContent-Type: application/json; charset=utf-8`r`nContent-Length: $($body.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n`r`n"
  Send-Raw $stream ([System.Text.Encoding]::ASCII.GetBytes($head)) $body
}

# 按 Content-Length 读取请求体（前端保证 body 为纯 ASCII 的 JSON）
function Read-AsciiBody($reader, [int]$len) {
  if ($len -le 0) { return '' }
  $sb = New-Object System.Text.StringBuilder
  $buf = New-Object 'char[]' 8192
  $left = $len
  while ($left -gt 0) {
    $want = [Math]::Min($buf.Length, $left)
    $n = $reader.Read($buf, 0, $want)
    if ($n -le 0) { break }
    [void]$sb.Append($buf, 0, $n)
    $left -= $n
  }
  return $sb.ToString()
}

# 覆盖写入：$backup 为真时先备份一份 .bak（仅保留最新一份）
# 先写同目录临时文件再原子替换，避免写一半失败把原表写坏；目标被 Excel/WPS 占用时抛错
function Save-BytesFile([string]$target, [byte[]]$bytes, [bool]$backup) {
  $dir = Split-Path -Parent $target
  if (-not (Test-Path -LiteralPath $dir -PathType Container)) { throw '目标目录不存在' }
  if ($backup -and (Test-Path -LiteralPath $target -PathType Leaf)) {
    try {
      $old = Read-Shared $target
      [System.IO.File]::WriteAllBytes(($target + '.bak'), $old)
    } catch { }
  }
  $tmp = $target + '.writing'
  try {
    [System.IO.File]::WriteAllBytes($tmp, $bytes)
  } catch {
    try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch { }
    throw '写入临时文件失败，请检查磁盘空间或目录权限'
  }
  try {
    if (Test-Path -LiteralPath $target -PathType Leaf) {
      try {
        [System.IO.File]::Replace($tmp, $target, $null)
      } catch {
        # 部分分区不支持 Replace：回退为覆盖复制
        [System.IO.File]::Copy($tmp, $target, $true)
        [System.IO.File]::Delete($tmp)
      }
    } else {
      [System.IO.File]::Move($tmp, $target)
    }
  } catch {
    try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch { }
    throw '文件被其他程序占用（可能正在 Excel / WPS 中打开），请关闭后重试'
  }
}

# 文件是否被 Excel / WPS 独占打开（能独占打开说明没被占用）
function Test-FileLocked([string]$path) {
  try {
    $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $fs.Dispose()
    return $false
  } catch {
    return $true
  }
}

# 备份目录：data\backup，按时间戳保留最近 20 份，出问题随时可回滚
function New-Backup([string]$target) {
  try {
    if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { return $null }
    $parent = Split-Path -Parent $target
    $bdir = Join-Path $parent 'backup'
    if (-not (Test-Path -LiteralPath $bdir -PathType Container)) { New-Item -ItemType Directory -Path $bdir -Force | Out-Null }
    $leaf = [System.IO.Path]::GetFileNameWithoutExtension($target)
    $ext = [System.IO.Path]::GetExtension($target)
    $dst = Join-Path $bdir ("{0}_{1}{2}" -f $leaf, (Get-Date -Format 'yyyyMMdd_HHmmss'), $ext)
    [System.IO.File]::WriteAllBytes($dst, (Read-Shared $target))
    Get-ChildItem -LiteralPath $bdir -Filter ("{0}_*{1}" -f $leaf, $ext) -ErrorAction SilentlyContinue |
      Sort-Object LastWriteTime -Descending | Select-Object -Skip 20 |
      ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force } catch { } }
    return $dst
  } catch { return $null }
}

# ===================== xlsx -> js 快照（零依赖：纯 .NET 解包 + XML 解析） =====================
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue

# 单元格引用 A1 / AB12 -> 列号（1 起）
function Get-ColIndexFromRef([string]$ref) {
  if ([string]::IsNullOrEmpty($ref)) { return 0 }
  $m = [regex]::Match($ref, '^[A-Za-z]+')
  if (-not $m.Success) { return 0 }
  $n = 0
  foreach ($ch in $m.Value.ToUpperInvariant().ToCharArray()) {
    $n = $n * 26 + (([int][char]$ch) - 64)
  }
  return $n
}

# 整行是否全空（与页面 isBlankRow 同义）
function Test-BlankRow($row) {
  if ($null -eq $row) { return $true }
  foreach ($v in $row) {
    if ($null -eq $v) { continue }
    if (([string]$v).Trim() -ne '') { return $false }
  }
  return $true
}

# JSON 字符串字面量（与 JSON.stringify 的转义一致，中文原样输出不转 \u）
function ConvertTo-JsonStr([string]$s) {
  if ($null -eq $s) { return '""' }
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    $code = [int]$ch
    if ($ch -eq [char]'"') { [void]$sb.Append('\"') }
    elseif ($ch -eq [char]'\') { [void]$sb.Append('\\') }
    elseif ($code -eq 8) { [void]$sb.Append('\b') }
    elseif ($code -eq 12) { [void]$sb.Append('\f') }
    elseif ($code -eq 10) { [void]$sb.Append('\n') }
    elseif ($code -eq 13) { [void]$sb.Append('\r') }
    elseif ($code -eq 9) { [void]$sb.Append('\t') }
    elseif ($code -lt 32) { [void]$sb.Append(('\u{0:x4}' -f $code)) }
    else { [void]$sb.Append($ch) }
  }
  [void]$sb.Append('"')
  return $sb.ToString()
}

# 数字转文本：整数不带小数点，其余用最短往返表示
function Format-CellNum([double]$d) {
  if ([double]::IsNaN($d)) { return '' }
  if ([double]::IsInfinity($d)) { return '' }
  if ($d -eq [Math]::Truncate($d) -and [Math]::Abs($d) -lt 1e15) {
    return ([long]$d).ToString([System.Globalization.CultureInfo]::InvariantCulture)
  }
  $s = $d.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)
  if ($s -match 'E') { $s = $d.ToString('0.###############', [System.Globalization.CultureInfo]::InvariantCulture) }
  return $s
}

# 数字格式判定：date / time / datetime / ''（普通数字）
function Get-FmtKind([int]$numFmtId, [string]$code) {
  if ($numFmtId -eq 0) { return '' }
  if ($numFmtId -ge 14 -and $numFmtId -le 17) { return 'date' }
  if ($numFmtId -eq 22) { return 'datetime' }
  if ($numFmtId -ge 18 -and $numFmtId -le 21) { return 'time' }
  if ($numFmtId -ge 27 -and $numFmtId -le 36) { return 'date' }
  if ($numFmtId -ge 45 -and $numFmtId -le 47) { return 'time' }
  if ($numFmtId -ge 50 -and $numFmtId -le 58) { return 'date' }
  if ([string]::IsNullOrWhiteSpace($code)) { return '' }
  $c = [regex]::Replace($code, '"[^"]*"', '')
  $c = [regex]::Replace($c, '\[[^\]]*\]', '')
  $c = [regex]::Replace($c, '\\.', '')
  $hasDate = ($c -match '[yY]') -or ($c -match '[dD]')
  $hasTime = ($c -match '[hH]') -or ($c -match '[sS]')
  if ($hasDate) { if ($hasTime) { return 'datetime' } else { return 'date' } }
  if ($hasTime) { return 'time' }
  return ''
}

# Excel 序列号 -> 文本（日期统一 yyyy-MM-dd，与页面 cellText 约定一致）
function Format-OADate([double]$d, [string]$kind, [bool]$date1904) {
  $v = $d
  if ($date1904) { $v = $d + 1462 }
  try { $dt = [DateTime]::FromOADate($v) } catch { return (Format-CellNum $d) }
  if ($kind -eq 'time') { return $dt.ToString('HH:mm:ss') }
  if ($v -lt 1) { return $dt.ToString('HH:mm:ss') }
  if ($dt.TimeOfDay.Ticks -eq 0) { return $dt.ToString('yyyy-MM-dd') }
  return $dt.ToString('yyyy-MM-dd HH:mm:ss')
}

# 解析一张工作表 -> @{ name; aoa; merges }
function Read-XlsxSheetGrid($entry, [string]$sheetName, $shared, $xfFmt, $customFmt, [bool]$date1904) {
  $doc = New-Object System.Xml.XmlDocument
  $doc.Load($entry.Open())

  $cells = @{}
  $maxR = 0
  $maxC = 0
  $rIdx = 0
  foreach ($rowEl in $doc.GetElementsByTagName('row')) {
    $rAttr = [string]$rowEl.GetAttribute('r')
    if ([string]::IsNullOrEmpty($rAttr)) { $rIdx++ } else { $rIdx = [int]$rAttr }
    $cIdx = 0
    foreach ($cel in $rowEl.ChildNodes) {
      if ($cel.LocalName -ne 'c') { continue }
      $ref = [string]$cel.GetAttribute('r')
      if ([string]::IsNullOrEmpty($ref)) { $cIdx++ } else { $cIdx = Get-ColIndexFromRef $ref }
      if ($cIdx -le 0) { $cIdx = 1 }
      $t = [string]$cel.GetAttribute('t')
      $sAttr = [string]$cel.GetAttribute('s')
      $raw = $null
      $isText = $null
      foreach ($cn in $cel.ChildNodes) {
        if ($cn.LocalName -eq 'v') { $raw = $cn.InnerText }
        elseif ($cn.LocalName -eq 'is') { $isText = $cn.InnerText }
      }
      $text = ''
      if ($t -eq 's') {
        if (-not [string]::IsNullOrEmpty($raw)) {
          $si = 0
          if ([int]::TryParse($raw, [ref]$si)) {
            if ($si -ge 0 -and $si -lt $shared.Count) { $text = [string]$shared[$si] }
          }
        }
      } elseif ($t -eq 'inlineStr') {
        if ($null -ne $isText) { $text = $isText }
      } elseif ($t -eq 'str' -or $t -eq 'd' -or $t -eq 'e') {
        if ($null -ne $raw) { $text = $raw }
      } elseif ($t -eq 'b') {
        if ($raw -eq '1') { $text = 'TRUE' } elseif ($raw -eq '0') { $text = 'FALSE' }
      } else {
        if (-not [string]::IsNullOrEmpty($raw)) {
          $num = 0.0
          if ([double]::TryParse($raw, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
            $fmtId = 0
            if (-not [string]::IsNullOrEmpty($sAttr)) {
              $sx = 0
              if ([int]::TryParse($sAttr, [ref]$sx)) {
                if ($sx -ge 0 -and $sx -lt $xfFmt.Count) { $fmtId = [int]$xfFmt[$sx] }
              }
            }
            $code = ''
            if ($customFmt.ContainsKey($fmtId)) { $code = [string]$customFmt[$fmtId] }
            $kind = Get-FmtKind $fmtId $code
            if ($kind -eq 'date' -or $kind -eq 'time' -or $kind -eq 'datetime') {
              $text = Format-OADate $num $kind $date1904
            } else {
              $text = Format-CellNum $num
            }
          } else {
            $text = $raw
          }
        }
      }
      if ($text -ne '') {
        $cells[('{0},{1}' -f $rIdx, $cIdx)] = $text
        if ($rIdx -gt $maxR) { $maxR = $rIdx }
        if ($cIdx -gt $maxC) { $maxC = $cIdx }
      }
    }
  }

  $aoa = New-Object System.Collections.Generic.List[object]
  for ($r = 1; $r -le $maxR; $r++) {
    $rowArr = New-Object 'object[]' $maxC
    for ($c = 1; $c -le $maxC; $c++) {
      $k = ('{0},{1}' -f $r, $c)
      if ($cells.ContainsKey($k)) { $rowArr[$c - 1] = $cells[$k] } else { $rowArr[$c - 1] = '' }
    }
    $aoa.Add($rowArr)
  }
  # 去掉尾部整行空白（与页面快照生成逻辑一致）
  while ($aoa.Count -gt 0) {
    if (Test-BlankRow $aoa[$aoa.Count - 1]) { $aoa.RemoveAt($aoa.Count - 1) } else { break }
  }

  $merges = New-Object System.Collections.Generic.List[object]
  foreach ($mc in $doc.GetElementsByTagName('mergeCell')) {
    $mref = [string]$mc.GetAttribute('ref')
    if ([string]::IsNullOrEmpty($mref)) { continue }
    $parts = $mref.Split(':')
    if ($parts.Length -lt 2) { continue }
    $sCol = Get-ColIndexFromRef $parts[0]
    $eCol = Get-ColIndexFromRef $parts[1]
    $sm = [regex]::Match($parts[0], '\d+')
    $em = [regex]::Match($parts[1], '\d+')
    if (-not $sm.Success -or -not $em.Success) { continue }
    $merges.Add([ordered]@{
      s = [ordered]@{ r = ([int]$sm.Value - 1); c = ($sCol - 1) }
      e = [ordered]@{ r = ([int]$em.Value - 1); c = ($eCol - 1) }
    })
  }

  return [ordered]@{ name = $sheetName; aoa = $aoa.ToArray(); merges = $merges.ToArray() }
}

# 解析整个 xlsx -> 工作表列表（按工作簿内顺序）
function Read-XlsxWorkbook([string]$path) {
  $zip = $null
  try { $zip = [System.IO.Compression.ZipFile]::OpenRead($path) }
  catch { throw ('无法打开 Excel 文件（可能正在被 Excel / WPS 写入）：' + $_.Exception.Message) }

  $sheets = New-Object System.Collections.Generic.List[object]
  try {
    $wbEntry = $zip.GetEntry('xl/workbook.xml')
    if (-not $wbEntry) { throw '不是标准 .xlsx 文件（缺少 xl/workbook.xml）' }
    $wbDoc = New-Object System.Xml.XmlDocument
    $wbDoc.Load($wbEntry.Open())

    $date1904 = $false
    $prNodes = $wbDoc.GetElementsByTagName('workbookPr')
    if ($prNodes.Count -gt 0) {
      $dv = [string]$prNodes[0].GetAttribute('date1904')
      if ($dv -eq '1' -or $dv -eq 'true') { $date1904 = $true }
    }

    $relMap = @{}
    $relEntry = $zip.GetEntry('xl/_rels/workbook.xml.rels')
    if ($relEntry) {
      $relDoc = New-Object System.Xml.XmlDocument
      $relDoc.Load($relEntry.Open())
      foreach ($rel in $relDoc.GetElementsByTagName('Relationship')) {
        $rid = [string]$rel.GetAttribute('Id')
        $tg = [string]$rel.GetAttribute('Target')
        if (-not [string]::IsNullOrEmpty($rid) -and -not [string]::IsNullOrEmpty($tg) -and $tg -notmatch '^https?://') {
          $relMap[$rid] = $tg
        }
      }
    }

    $shared = New-Object System.Collections.Generic.List[string]
    $ssEntry = $zip.GetEntry('xl/sharedStrings.xml')
    if ($ssEntry) {
      $ssDoc = New-Object System.Xml.XmlDocument
      $ssDoc.Load($ssEntry.Open())
      foreach ($si in $ssDoc.GetElementsByTagName('si')) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($node in $si.ChildNodes) {
          if ($node.LocalName -eq 't') { [void]$sb.Append($node.InnerText) }
          elseif ($node.LocalName -eq 'r') {
            foreach ($tn in $node.GetElementsByTagName('t')) { [void]$sb.Append($tn.InnerText) }
          }
        }
        [void]$shared.Add($sb.ToString())
      }
    }

    $xfFmt = New-Object System.Collections.Generic.List[int]
    $customFmt = @{}
    $stEntry = $zip.GetEntry('xl/styles.xml')
    if ($stEntry) {
      $stDoc = New-Object System.Xml.XmlDocument
      $stDoc.Load($stEntry.Open())
      foreach ($nf in $stDoc.GetElementsByTagName('numFmt')) {
        $nid = [string]$nf.GetAttribute('numFmtId')
        if (-not [string]::IsNullOrEmpty($nid)) { $customFmt[[int]$nid] = [string]$nf.GetAttribute('formatCode') }
      }
      foreach ($cx in $stDoc.GetElementsByTagName('cellXfs')) {
        foreach ($xf in $cx.GetElementsByTagName('xf')) {
          $nid = [string]$xf.GetAttribute('numFmtId')
          if ([string]::IsNullOrEmpty($nid)) { [void]$xfFmt.Add(0) } else { [void]$xfFmt.Add([int]$nid) }
        }
      }
    }

    foreach ($shEl in $wbDoc.GetElementsByTagName('sheet')) {
      $sname = [string]$shEl.GetAttribute('name')
      $rid = $shEl.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
      $tgt = $null
      if (-not [string]::IsNullOrEmpty($rid) -and $relMap.ContainsKey($rid)) { $tgt = [string]$relMap[$rid] }
      if ([string]::IsNullOrEmpty($tgt)) { continue }
      $tgt = $tgt.Replace('\', '/')
      if ($tgt.StartsWith('/')) { $tgt = $tgt.TrimStart('/') } else { $tgt = 'xl/' + $tgt }
      $tgt = $tgt.Replace('/./', '/')
      $entry = $zip.GetEntry($tgt)
      if (-not $entry) { continue }
      $sheets.Add((Read-XlsxSheetGrid $entry $sname $shared $xfFmt $customFmt $date1904))
    }
  } finally {
    $zip.Dispose()
  }
  return ,$sheets
}

# 通用 JSON 序列化：$step=0 紧凑，$step=2 缩进（等价 JSON.stringify 的两种输出）
function ConvertTo-JsonDoc($v, [int]$level, [int]$step) {
  $pad = ''
  $padIn = ''
  if ($step -gt 0) {
    $pad = ' ' * ($level * $step)
    $padIn = ' ' * (($level + 1) * $step)
  }
  if ($null -eq $v) { return 'null' }
  if ($v -is [string]) { return (ConvertTo-JsonStr $v) }
  if ($v -is [datetime]) { return (ConvertTo-JsonStr ($v.ToString('yyyy-MM-dd HH:mm'))) }
  if ($v -is [bool]) { if ($v) { return 'true' } else { return 'false' } }
  if ($v -is [System.Collections.IDictionary]) {
    $keys = @($v.Keys)
    if ($keys.Count -eq 0) { return '{}' }
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($k in $keys) {
      $sv = ConvertTo-JsonDoc $v[$k] ($level + 1) $step
      if ($step -gt 0) { $items.Add($padIn + (ConvertTo-JsonStr ([string]$k)) + ': ' + $sv) }
      else { $items.Add((ConvertTo-JsonStr ([string]$k)) + ':' + $sv) }
    }
    if ($step -le 0) { return ('{' + ($items -join ',') + '}') }
    return ('{' + "`n" + ($items -join ",`n") + "`n" + $pad + '}')
  }
  if ($v -is [System.Collections.IEnumerable]) {
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($it in $v) {
      $sv = ConvertTo-JsonDoc $it ($level + 1) $step
      if ($step -gt 0) { $items.Add($padIn + $sv) } else { $items.Add($sv) }
    }
    if ($items.Count -eq 0) { return '[]' }
    if ($step -le 0) { return ('[' + ($items -join ',') + ']') }
    return ('[' + "`n" + ($items -join ",`n") + "`n" + $pad + ']')
  }
  if ($v -is [ValueType]) { return (Format-CellNum ([double]$v)) }
  return (ConvertTo-JsonStr ([string]$v))
}

# 通用表格快照文本：window.__HOME_DATA__ = { generatedAt, sheets:[{name, aoa, merges}] }
function New-SheetsSnapshotText($sheets, [string]$varName, [string]$xlsxName) {
  $arr = New-Object System.Collections.Generic.List[object]
  foreach ($s in $sheets) {
    $arr.Add([ordered]@{ name = [string]$s.name; aoa = $s.aoa; merges = $s.merges })
  }
  $root = [ordered]@{ generatedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm'); sheets = $arr.ToArray() }
  $json = ConvertTo-JsonDoc $root 0 0
  $comment = '/* 自动生成：' + $xlsxName + ' 数据快照。仅在 file:// 双击打开时作为兜底；改完 Excel 后双击 start-login.bat 会自动重建本文件。 */'
  return ($comment + "`n" + 'window.' + $varName + ' = ' + $json + ';' + "`n")
}

# 权限开关取值：1 / 0
function Get-PermBit($row, $idx, [string]$key) {
  if (-not $idx.ContainsKey($key)) { return 0 }
  $i = [int]$idx[$key]
  if ($i -lt 0 -or $i -ge @($row).Count) { return 0 }
  $v = [string]$row[$i]
  $d = 0.0
  if ([double]::TryParse($v, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
    if ($d -eq 1) { return 1 }
  }
  return 0
}

# 用户快照文本：window.__USERS_DATA__ = { generatedAt, note, users:[...], perm:{sheetName, tables} }
function New-UsersSnapshotText($sheets, [string]$varName, [string]$xlsxName) {
  $userSheet = $null
  foreach ($s in $sheets) {
    if (([string]$s.name -replace '\s', '') -eq '用户') { $userSheet = $s; break }
  }
  if (-not $userSheet -and $sheets.Count -gt 0) { $userSheet = $sheets[0] }

  $headers = @()
  $users = New-Object System.Collections.Generic.List[object]
  if ($userSheet) {
    $aoa = $userSheet.aoa
    $hIdx = -1
    for ($i = 0; $i -lt $aoa.Count; $i++) {
      if (-not (Test-BlankRow $aoa[$i])) { $hIdx = $i; break }
    }
    if ($hIdx -ge 0) { $headers = @($aoa[$hIdx] | ForEach-Object { ([string]$_).Trim() }) }
    for ($r = $hIdx + 1; $r -lt $aoa.Count; $r++) {
      $row = @($aoa[$r])
      if (Test-BlankRow $row) { continue }
      $first = ''
      if ($row.Count -gt 0) { $first = [string]$row[0] }
      if ($first.Trim() -eq '') { continue }
      $obj = [ordered]@{}
      for ($c = 0; $c -lt $headers.Count; $c++) {
        $h = [string]$headers[$c]
        if ([string]::IsNullOrEmpty($h)) { $h = '列' + ($c + 1) }
        $v = ''
        if ($c -lt $row.Count) { $v = [string]$row[$c] }
        $obj[$h] = $v
      }
      $users.Add($obj)
    }
  }

  $permName = '用户权限'
  $tables = New-Object System.Collections.Generic.List[object]
  $permSheet = $null
  foreach ($s in $sheets) {
    if (([string]$s.name -replace '\s', '') -eq '用户权限') { $permSheet = $s; break }
  }
  if ($permSheet) {
    $permName = [string]$permSheet.name
    $paoa = $permSheet.aoa
    $ph = -1
    for ($i = 0; $i -lt $paoa.Count; $i++) {
      if (-not (Test-BlankRow $paoa[$i])) { $ph = $i; break }
    }
    if ($ph -ge 0) {
      $head = @($paoa[$ph] | ForEach-Object { ([string]$_).Trim() })
      $idx = @{}
      for ($i = 0; $i -lt $head.Count; $i++) {
        $k = $head[$i] -replace '\s', ''
        if (-not $idx.ContainsKey($k)) { $idx[$k] = $i }
      }
      $iName = 0
      if ($idx.ContainsKey('表名')) { $iName = [int]$idx['表名'] }
      for ($p = $ph + 1; $p -lt $paoa.Count; $p++) {
        $pr = @($paoa[$p])
        if (Test-BlankRow $pr) { continue }
        $tname = ''
        if ($iName -lt $pr.Count) { $tname = ([string]$pr[$iName]).Trim() }
        if ($tname -eq '') { continue }
        $tables.Add([ordered]@{
          name = $tname
          cols = [ordered]@{
            '一般用户' = (Get-PermBit $pr $idx '一般用户')
            '管理用户' = (Get-PermBit $pr $idx '管理用户')
            '超级管理用户' = (Get-PermBit $pr $idx '超级管理用户')
          }
          deletable = (Get-PermBit $pr $idx '可删除')
        })
      }
    }
  }

  $root = [ordered]@{
    generatedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    note = ($xlsxName + ' 的内置快照，供双击(file://)打开时兜底；修改 xlsx 后双击 start-login.bat 会自动重建。')
    users = $users.ToArray()
    perm = [ordered]@{ sheetName = $permName; tables = $tables.ToArray() }
  }
  $json = ConvertTo-JsonDoc $root 0 2
  $comment = '/* 自动生成，源数据：data/' + $xlsxName + '，请勿手工编辑 */'
  return ($comment + "`n" + 'window.' + $varName + ' = ' + $json + ';' + "`n")
}

# 扫描 data 目录下全部 js 快照：与同名 xlsx 的当前内容逐格比对，不一致才解析重建 js（重建前备份旧文件）
# 返回 @{ total; updated; skipped; failed; lines }
# 快照文本解析：取出 window.__XXX__ = { ... }; 里的变量名与 JSON 对象
function Get-SnapshotDoc([string]$text) {
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  $m = [regex]::Match($text, 'window\.(__[A-Za-z0-9_]+__)\s*=\s*([\s\S]*?);\s*$')
  if (-not $m.Success) { return $null }
  $doc = $null
  try { $doc = $m.Groups[2].Value | ConvertFrom-Json } catch { return $null }
  if ($null -eq $doc) { return $null }
  return [pscustomobject]@{ var = ([string]$m.Groups[1].Value); doc = $doc }
}

# 数值/文本统一成可比较的字符串（null -> 空串；数字去掉多余小数位）
function ConvertTo-CompareText($v) {
  if ($null -eq $v) { return '' }
  if ($v -is [bool]) { if ($v) { return 'TRUE' } else { return 'FALSE' } }
  if ($v -is [double] -or $v -is [single] -or $v -is [decimal]) {
    return (([double]$v).ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)).Trim()
  }
  return ([string]$v).Trim()
}

# 表格类快照（window.__HOME_DATA__）内容等价判断：表名 / 逐格文本 / 合并单元格
function Test-SheetsSnapshotEqual($a, $b) {
  $sa = @($a.sheets); $sb = @($b.sheets)
  if ($sa.Count -eq 0) { return $false }
  if ($sa.Count -ne $sb.Count) { return $false }
  for ($i = 0; $i -lt $sa.Count; $i++) {
    if (([string]$sa[$i].name) -ne ([string]$sb[$i].name)) { return $false }
    $ra = @($sa[$i].aoa); $rb = @($sb[$i].aoa)
    if ($ra.Count -ne $rb.Count) { return $false }
    for ($r = 0; $r -lt $ra.Count; $r++) {
      $ca = @($ra[$r]); $cb = @($rb[$r])
      if ($ca.Count -ne $cb.Count) { return $false }
      for ($c = 0; $c -lt $ca.Count; $c++) {
        if ((ConvertTo-CompareText $ca[$c]) -ne (ConvertTo-CompareText $cb[$c])) { return $false }
      }
    }
    $ma = (@($sa[$i].merges) | ForEach-Object { [string]$_ }) -join '|'
    $mb = (@($sb[$i].merges) | ForEach-Object { [string]$_ }) -join '|'
    if ($ma -ne $mb) { return $false }
  }
  return $true
}

# 用户类快照（window.__USERS_DATA__）内容等价判断：用户字段 / 权限表逐行开关
function Test-UsersSnapshotEqual($a, $b) {
  $ua = @($a.users); $ub = @($b.users)
  if ($ua.Count -ne $ub.Count) { return $false }
  for ($i = 0; $i -lt $ua.Count; $i++) {
    $na = (@($ua[$i].PSObject.Properties) | ForEach-Object { $_.Name }) -join '|'
    $nb = (@($ub[$i].PSObject.Properties) | ForEach-Object { $_.Name }) -join '|'
    if ($na -ne $nb) { return $false }
    foreach ($prop in @($ua[$i].PSObject.Properties)) {
      $pb = $ub[$i].PSObject.Properties[$prop.Name]
      $valB = $null
      if ($pb) { $valB = $pb.Value }
      if ((ConvertTo-CompareText $prop.Value) -ne (ConvertTo-CompareText $valB)) { return $false }
    }
  }
  if (([string]$a.perm.sheetName) -ne ([string]$b.perm.sheetName)) { return $false }
  $ta = @($a.perm.tables); $tb = @($b.perm.tables)
  if ($ta.Count -ne $tb.Count) { return $false }
  for ($i = 0; $i -lt $ta.Count; $i++) {
    if (([string]$ta[$i].name) -ne ([string]$tb[$i].name)) { return $false }
    if ((ConvertTo-CompareText $ta[$i].deletable) -ne (ConvertTo-CompareText $tb[$i].deletable)) { return $false }
    foreach ($k in @('一般用户', '管理用户', '超级管理用户')) {
      if ((ConvertTo-CompareText $ta[$i].cols.$k) -ne (ConvertTo-CompareText $tb[$i].cols.$k)) { return $false }
    }
  }
  return $true
}

# 扫描 data 目录下全部 js 快照：与同名 xlsx 的当前内容逐格比对，内容不一致才重建（重建前备份旧文件）
# 判据是"内容是否等价"而不是"修改时间谁更新"，避免 js 时间戳比 xlsx 新、内容却已过期的情况被漏掉
# 返回 @{ total; updated; skipped; failed; lines }
function Sync-SnapshotFiles([string]$dataDir, [switch]$Force) {
  $res = @{ total = 0; updated = 0; skipped = 0; failed = 0; lines = @() }
  if (-not (Test-Path -LiteralPath $dataDir -PathType Container)) { return $res }
  $jsFiles = @(Get-ChildItem -LiteralPath $dataDir -Filter '*.js' -File -ErrorAction SilentlyContinue | Sort-Object Name)
  foreach ($js in $jsFiles) {
    $res.total = $res.total + 1
    $base = [System.IO.Path]::GetFileNameWithoutExtension($js.Name)
    $xlsxPath = Join-Path $dataDir ($base + '.xlsx')
    if (-not (Test-Path -LiteralPath $xlsxPath -PathType Leaf)) {
      $res.skipped = $res.skipped + 1
      $res.lines += [ordered]@{ file = $js.Name; state = 'skip'; msg = '没有同名 xlsx，跳过' }
      continue
    }
    try {
      $oldText = [System.IO.File]::ReadAllText($js.FullName, [System.Text.Encoding]::UTF8)
      $head = $oldText
      if ($head.Length -gt 40000) { $head = $head.Substring(0, 40000) }
      $m = [regex]::Match($head, 'window\.(__[A-Za-z0-9_]+__)\s*=')
      if (-not $m.Success) {
        $res.skipped = $res.skipped + 1
        $res.lines += [ordered]@{ file = $js.Name; state = 'skip'; msg = '不是数据快照（未找到 window.__XXX__ 赋值），已跳过' }
        continue
      }
      $varName = $m.Groups[1].Value
      $sheets = Read-XlsxWorkbook $xlsxPath
      if (-not $sheets -or $sheets.Count -eq 0) { throw 'Excel 里没有可读取的工作表' }
      if ($varName -match 'USER') {
        $text = New-UsersSnapshotText $sheets $varName ([System.IO.Path]::GetFileName($xlsxPath))
      } else {
        $text = New-SheetsSnapshotText $sheets $varName ([System.IO.Path]::GetFileName($xlsxPath))
      }
      if (-not $Force) {
        $oa = Get-SnapshotDoc $oldText
        $ob = Get-SnapshotDoc $text
        $same = $false
        if ($oa -and $ob -and (([string]$oa.var) -eq ([string]$ob.var))) {
          if ($varName -match 'USER') { $same = Test-UsersSnapshotEqual $oa.doc $ob.doc }
          else { $same = Test-SheetsSnapshotEqual $oa.doc $ob.doc }
        }
        if ($same) {
          $res.skipped = $res.skipped + 1
          $res.lines += [ordered]@{ file = $js.Name; state = 'fresh'; msg = '内容与 xlsx 一致，无需重建' }
          continue
        }
      }
      $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($text)
      $bakPath = New-Backup $js.FullName
      Save-BytesFile $js.FullName $bytes $false
      $res.updated = $res.updated + 1
      $res.lines += [ordered]@{ file = $js.Name; state = 'update'; msg = ('已按 ' + $base + '.xlsx 重建（' + $sheets.Count + ' 张表）'); size = $bytes.Length; backup = $bakPath }
    } catch {
      $res.failed = $res.failed + 1
      $msg = $_.Exception.Message
      if ([string]::IsNullOrWhiteSpace($msg)) { $msg = '未知错误' }
      $res.lines += [ordered]@{ file = $js.Name; state = 'fail'; msg = $msg }
    }
  }
  return $res
}

# 取一张工作表的"格式指纹"：列宽 / 行高 / 前几行字体（名称+字号+加粗）
# 用途：保存后自检，确认这次写入只动了数据，没有把整本工作簿的样式重写
# 注意：固定读 A1 起 30 行 × 20 列，不用 UsedRange —— UsedRange 在 Excel 里是惰性计算的，
# 保存前后同一张表可能报出不同尺寸，会造成"别的表被改了"的假警报
function Get-SheetFingerprint($ws) {
  $rowsN = 30
  $colsN = 20
  $cols = New-Object System.Text.StringBuilder
  for ($c = 1; $c -le $colsN; $c++) {
    try { [void]$cols.Append([math]::Round([double]$ws.Columns($c).Width, 2)).Append(',') }
    catch { [void]$cols.Append('x,') }
  }
  $rows = New-Object System.Text.StringBuilder
  for ($r = 1; $r -le $rowsN; $r++) {
    try { [void]$rows.Append([math]::Round([double]$ws.Rows($r).RowHeight, 2)).Append(',') }
    catch { [void]$rows.Append('x,') }
  }
  $fonts = New-Object System.Text.StringBuilder
  for ($r = 1; $r -le 3; $r++) {
    for ($c = 1; $c -le $colsN; $c++) {
      try {
        $cell = $ws.Cells($r, $c)
        [void]$fonts.Append($cell.Font.Name).Append('/').Append($cell.Font.Size).Append('/').Append([int]$cell.Font.Bold).Append(';')
      } catch { [void]$fonts.Append('x;') }
    }
  }
  return @{ cols = $cols.ToString(); rows = $rows.ToString(); fonts = $fonts.ToString() }
}

# 除目标表外，其余各表整体指纹
function Get-WbFingerprint($wb, [string]$skipSheet) {
  $fp = @{}
  foreach ($s in $wb.Worksheets) {
    $n = [string]$s.Name
    if ($n -eq $skipSheet) { continue }
    try { $fp[$n] = Get-SheetFingerprint $s } catch { }
  }
  return $fp
}

# 对比两份指纹，返回发生变化的表名列表
function Compare-Fingerprint($before, $after) {
  $diff = @()
  foreach ($k in $before.Keys) {
    if (-not $after.ContainsKey($k)) { $diff += $k; continue }
    $a = $before[$k]; $b = $after[$k]
    if ($a.cols -ne $b.cols -or $a.rows -ne $b.rows -or $a.fonts -ne $b.fonts) { $diff += $k }
  }
  return $diff
}

# 用本机 Excel 精确改写：只写数据区单元格的值，按需在末尾增删行，原有字体/边框/行高/合并/打印设置一律不动
# 保存后做格式自检：一旦发现别的表或本表的列宽/字体被改动，立即用备份回滚，绝不留下被改坏的文件
# 返回 @{ ok = $true; method = 'excel'; lastRow = n }；失败抛异常（异常消息直接回给页面）
function Invoke-ExcelPatch([string]$target, [string]$sheetName, [int]$headerRows, $rows) {
  $rowCount = @($rows).Count
  if ($rowCount -le 0) { throw '没有可写入的数据行' }

  # 写前备份：.bak（最新一份）+ backup 目录按时间戳留 20 份
  try {
    if (Test-Path -LiteralPath $target -PathType Leaf) {
      $old = Read-Shared $target
      [System.IO.File]::WriteAllBytes(($target + '.bak'), $old)
    }
  } catch { }
  # 先探一下文件是不是正被别的程序开着：被占用时直接给出明确提示，
  # 不进 Excel（否则 COM 会抛出 RPC 失败这种看不懂的错误，还可能留下孤儿进程）
  try {
    $probe = [System.IO.File]::Open($target, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $probe.Close(); $probe.Dispose()
  } catch {
    throw [System.Exception]::new('LOCKED::文件正被 Excel 或 WPS 打开（无法独占写入），请先在 WPS / Excel 里关闭该表格，再点「重试保存」')
  }

  $bakPath = New-Backup $target
  $rollbackReason = ''

  $xl = $null
  $wb = $null
  $excelPid = 0
  try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    # 记下自己启动的这个 Excel 进程，收尾时兜底清理
    try {
      $h = [IntPtr]$xl.Hwnd
      $procId = [uint32]0
      [void][Native.Win32]::GetWindowThreadProcessId($h, [ref]$procId)
      $excelPid = [int]$procId
    } catch { $excelPid = 0 }
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    $xl.EnableEvents = $false
    try { $xl.Calculation = -4135 } catch { }   # 手动重算，避免大表拖慢
    $wb = $xl.Workbooks.Open($target, 0, $false)
    # 被 WPS / 另一个 Excel 打开着时，这里只能拿到只读副本：绝不能继续改
    if ($wb.ReadOnly) {
      throw [System.Exception]::new('LOCKED::文件正被 Excel / WPS 占用（只能只读打开），请先关闭该文件再操作')
    }

    $ws = $null
    foreach ($s in $wb.Worksheets) {
      if ([string]$s.Name -eq $sheetName) { $ws = $s; break }
    }
    if ($null -eq $ws) { throw ('工作簿中没有名为「' + $sheetName + '」的工作表') }

    # 改动前先记下全部表的格式指纹，保存后拿来自检
    $fpBefore = Get-WbFingerprint $wb $sheetName
    $selfBefore = Get-SheetFingerprint $ws

    $maxCol = 0
    foreach ($r in $rows) { if (@($r).Count -gt $maxCol) { $maxCol = @($r).Count } }
    if ($maxCol -le 0) { throw '数据行内容为空' }

    $used = $ws.UsedRange
    $lastRow = [int]$used.Row + [int]$used.Rows.Count - 1
    $targetLast = $headerRows + $rowCount

    # 数据行数对齐：不足则从表尾插入（沿用上方格式），多余则删除表尾多余行
    if ($targetLast -gt $lastRow) {
      $addFrom = $lastRow + 1
      $ws.Rows(("{0}:{1}" -f $addFrom, $targetLast)).Insert(-4121, 0) | Out-Null
      # 新行继承最后一条数据行的格式（字体/边框/底纹），否则新行是光秃秃的无框线
      try {
        $ws.Rows(("{0}:{0}" -f $lastRow)).Copy() | Out-Null
        $ws.Rows(("{0}:{1}" -f $addFrom, $targetLast)).PasteSpecial(-4122) | Out-Null
        $xl.CutCopyMode = 0
      } catch {
        try { $xl.CutCopyMode = 0 } catch { }
      }
    } elseif ($targetLast -lt $lastRow) {
      $ws.Rows(("{0}:{1}" -f ($targetLast + 1), $lastRow)).Delete() | Out-Null
    }

    $arr = New-Object 'object[,]' $rowCount, $maxCol
    for ($i = 0; $i -lt $rowCount; $i++) {
      $r = @($rows[$i])
      for ($j = 0; $j -lt $maxCol; $j++) {
        $v = if ($j -lt $r.Count) { $r[$j] } else { $null }
        if ($v -is [string] -and $v.Trim() -eq '') { $v = $null }
        $arr[$i, $j] = $v
      }
    }

    $range = $ws.Range($ws.Cells($headerRows + 1, 1), $ws.Cells($targetLast, $maxCol))
    # 记录写前的"自动换行"设置：Excel 给单元格赋含换行的文本时会擅自打开 WrapText，写完要还原
    $wrapBefore = $null
    try { $wrapBefore = $range.WrapText } catch { $wrapBefore = $null }
    $written = $false
    try {
      $range.Value2 = $arr
      $written = $true
    } catch { $written = $false }

    if (-not $written) {
      # 数据区内含合并单元格：逐格写，跳过合并区里非左上角的格子与公式格
      for ($i = 0; $i -lt $rowCount; $i++) {
        for ($j = 0; $j -lt $maxCol; $j++) {
          $cell = $ws.Cells($headerRows + 1 + $i, 1 + $j)
          try {
            if ([int]$cell.MergeArea.Cells.Count -gt 1) {
              $anchor = $cell.MergeArea.Cells.Item(1, 1)
              if ([string]$anchor.Address() -ne [string]$cell.Address()) { continue }
            }
          } catch { }
          $cell.Value2 = $arr[$i, $j]
        }
      }
    }

    # 还原自动换行设置，保持原表观感
    if ($null -ne $wrapBefore) {
      try { $range.WrapText = $wrapBefore } catch { }
    }

    $wb.Save()

    # ===== 保存后自检：只允许目标表的数据行数变化，其余表以及本表的列宽/字体必须一模一样 =====
    $fpAfter = Get-WbFingerprint $wb $sheetName
    $changed = Compare-Fingerprint $fpBefore $fpAfter
    $selfAfter = Get-SheetFingerprint $ws
    $selfBad = ($selfAfter.cols -ne $selfBefore.cols) -or ($selfAfter.fonts -ne $selfBefore.fonts)
    if ($changed.Count -gt 0 -or $selfBad) {
      $who = if ($changed.Count -gt 0) { '其他工作表：' + ($changed -join '、') } else { '本表的列宽/字体' }
      $rollbackReason = '保存过程改动了不该动的格式（' + $who + '），已自动回滚到保存前的文件，本次修改未写入'
      throw [System.Exception]::new('SELFCHECK::' + $rollbackReason)
    }

    return @{ ok = $true; method = 'excel'; lastRow = $targetLast; rows = $rowCount }
  } finally {
    if ($wb) {
      try { $wb.Close($false) } catch { }
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
    }
    if ($xl) {
      try { $xl.Quit() } catch { }
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
    }
    # 自检不通过：用备份把文件恢复原样，绝不让被改坏的版本留在磁盘上
    if ($rollbackReason -and $bakPath -and (Test-Path -LiteralPath $bakPath -PathType Leaf)) {
      for ($k = 0; $k -lt 8; $k++) {
        try { [System.IO.File]::Copy($bakPath, $target, $true); break }
        catch { Start-Sleep -Milliseconds 500 }
      }
      Write-Log ('已回滚 ' + $target + ' <- ' + $bakPath)
    }
    $wb = $null; $ws = $null; $range = $null; $used = $null; $xl = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    # 兜底：Quit 后进程若还赖着（COM 引用未完全释放时常见），只收掉自己启动的那一个
    if ($excelPid -gt 0) {
      try {
        Start-Sleep -Milliseconds 400
        $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($p -and [string]::IsNullOrEmpty($p.MainWindowTitle)) { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue }
      } catch { }
    }
  }
}

# ============ 启动前：把 data 目录里所有 js 快照与同名 xlsx 对齐 ============
if (-not $NoSync) {
  Write-Host ""
  Write-Host "  正在检查数据快照（xlsx -> js）..." -ForegroundColor DarkGray
  $syncRes = $null
  try { $syncRes = Sync-SnapshotFiles (Join-Path $Root 'data') -Force:$Force }
  catch {
    $syncRes = @{ total = 0; updated = 0; skipped = 0; failed = 1; lines = @([ordered]@{ file = 'data'; state = 'fail'; msg = $_.Exception.Message }) }
  }
  foreach ($it in $syncRes.lines) {
    if ($it.state -eq 'update') {
      Write-Host ("  [已更新] " + $it.file + " <- " + $it.msg) -ForegroundColor Green
      Write-Log ("快照重建 " + $it.file + ": " + $it.msg + "; 旧文件备份 " + $it.backup)
    } elseif ($it.state -eq 'fail') {
      Write-Host ("  [失败]   " + $it.file + ": " + $it.msg) -ForegroundColor Red
      Write-Log ("快照重建失败 " + $it.file + ": " + $it.msg)
    } elseif ($it.state -eq 'skip') {
      Write-Host ("  [跳过]   " + $it.file + ": " + $it.msg) -ForegroundColor DarkYellow
    }
  }
  if ($syncRes.updated -eq 0 -and $syncRes.failed -eq 0) {
    Write-Host ("  数据快照已与 Excel 一致（共 {0} 个，无需重建）" -f $syncRes.total) -ForegroundColor DarkGray
  }
  if ($syncRes.failed -gt 0) {
    Write-Host "  提示：快照更新失败不影响服务启动，http 模式下页面仍会直接读 Excel。" -ForegroundColor DarkYellow
  }
}

if ($SyncOnly) {
  Write-Host ""
  Write-Host "  快照同步完成（本次不启动服务）。" -ForegroundColor DarkGray
  Write-Log "快照同步完成（-SyncOnly）"
  exit 0
}

$listener = $null
try {
  $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $Port)
  $listener.Start()
} catch {
  Write-Host ("[错误] 端口 {0} 无法使用：{1}" -f $Port, $_.Exception.Message) -ForegroundColor Red
  Write-Log ("启动失败: " + $_.Exception.Message)
  exit 1
}

Write-Host ""
Write-Host "  秦巴药都数据平台 · 本地服务已启动" -ForegroundColor Green
Write-Host ("  服务根目录：{0}" -f $Root)
Write-Host ("  登录地址  ：http://127.0.0.1:{0}/html/login.html" -f $Port)
Write-Host ("  数据目录  ：{0}\data" -f $Root)
Write-Host "  关闭本窗口或按 Ctrl+C 即停止服务。" -ForegroundColor DarkGray
Write-Host ""
Write-Log ("服务启动 port=" + $Port)

while ($true) {
  $client = $null
  try {
    $client = $listener.AcceptTcpClient()
    $client.ReceiveTimeout = 8000
    $stream = $client.GetStream()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII, $false, 8192, $true)
    $requestLine = $reader.ReadLine()
    if ([string]::IsNullOrWhiteSpace($requestLine)) { continue }
    $contentLength = 0
    while ($true) {
      $h = $reader.ReadLine()
      if ($null -eq $h -or $h -eq '') { break }
      if ($h -match '^(?i)Content-Length:\s*(\d+)') { $contentLength = [int]$matches[1] }
    }

    $parts = $requestLine.Split(' ')
    $method = $parts[0].ToUpperInvariant()
    $target = if ($parts.Length -gt 1) { $parts[1] } else { '/' }
    $path = ($target -split '\?')[0]
    try { $path = [System.Uri]::UnescapeDataString($path) } catch { }
    $path = $path.Replace('\', '/')
    if ($path -eq '/' -or $path -eq '') { $path = '/html/index.html' }

    # ===== 精确改写接口：POST /__patch —— 页面只发数据区的值，由本机 Excel 改单元格，格式原样保留 =====
    if ($method -eq 'POST' -and $path -eq '/__patch') {
      if ($contentLength -le 0) { Send-Json $stream 400 '{"ok":false,"error":"请求体为空"}'; continue }
      if ($contentLength -gt 20971520) { Send-Json $stream 400 '{"ok":false,"error":"数据超过 20MB，拒绝写入"}'; continue }
      $bodyText = Read-AsciiBody $reader $contentLength
      $target = $null
      $name = ''
      try {
        $req = $bodyText | ConvertFrom-Json
        $rel = ([string]$req.path).Replace('\', '/').TrimStart('/')
        $sheetName = [string]$req.sheet
        $headerRows = [int]$req.headerRows
        $rows = $req.rows
        if ([string]::IsNullOrWhiteSpace($rel)) { throw '缺少 path' }
        if ([string]::IsNullOrWhiteSpace($sheetName)) { throw '缺少工作表名' }
        if ($headerRows -lt 0) { throw 'headerRows 不合法' }
        if ($null -eq $rows -or @($rows).Count -le 0) { throw '缺少数据行' }
        $ext = [System.IO.Path]::GetExtension($rel).ToLowerInvariant()
        if ($ext -ne '.xlsx' -and $ext -ne '.xls') { throw '只允许改写 .xlsx / .xls 数据文件' }
        $target = [System.IO.Path]::GetFullPath((Join-Path $Root ($rel.Replace('/', '\'))))
        if (-not $target.StartsWith($Root + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw '路径超出服务根目录' }
        if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw '数据文件不存在' }
        $name = Split-Path -Leaf $target
        # 只有"精确改写"这一条路：整体重写会把字体/边框/列宽全压平，正是格式全变的根源。
        # 路走不通就明确报错（提示关掉占用它的 Excel/WPS），绝不偷偷降级去整体重写。
        $r = Invoke-ExcelPatch $target $sheetName $headerRows $rows
        $len = (Get-Item -LiteralPath $target).Length
        Write-Host ("  200  PATCH " + $name + " [" + $sheetName + "] rows=" + $r.rows) -ForegroundColor Green
        Write-Log ("精确改写 " + $target + " sheet=" + $sheetName + " rows=" + $r.rows + " " + $len + " bytes")
        Send-Json $stream 200 ('{"ok":true,"method":"excel","name":"' + $name + '","rows":' + $r.rows + ',"bytes":' + $len + '}')
      } catch {
        $msg = $_.Exception.Message
        if ($null -eq $msg) { $msg = '未知错误' }
        $code = 'error'
        if ($msg.StartsWith('LOCKED::')) { $code = 'locked'; $msg = $msg.Substring(8) }
        elseif ($msg.StartsWith('SELFCHECK::')) { $code = 'selfcheck'; $msg = $msg.Substring(11) }
        elseif ($msg -match 'ComObject|CLSID|Class not registered|80040154|80080005|Retrieving the COM class') {
          $code = 'no_excel'; $msg = '本机未安装 Microsoft Excel，无法保真写入（为免损坏表格格式，已禁止整体重写）'
        }
        elseif ($msg -match '800706BE|800706BA|0x800A|RPC|远程过程调用|调用被拒绝') {
          $code = 'locked'; $msg = '写入引擎被占用（表格可能正在 WPS / Excel 中打开），请关闭该文件后点「重试保存」'
        }
        elseif ($msg -match '只读|read-only|ReadOnly') {
          $code = 'locked'; $msg = '文件正被 Excel / WPS 以只读方式占用，请关闭后重试'
        }
        Write-Host ("  500  PATCH 失败(" + $code + "): " + $msg) -ForegroundColor Red
        Write-Log ("精确改写失败[" + $code + "]: " + $msg)
        $safe = $msg -replace '"', '\"' -replace "`r", ' ' -replace "`n", ' '
        Send-Json $stream 500 ('{"ok":false,"code":"' + $code + '","error":"' + $safe + '"}')
      }
      continue
    }

    # ===== 自动保存接口：POST /__save —— 页面把改后的 xlsx 覆盖回原目录中的文件 =====
    if ($method -eq 'POST' -and $path -eq '/__save') {
      if ($contentLength -le 0) {
        Send-Json $stream 400 '{"ok":false,"error":"请求体为空"}'
        continue
      }
      if ($contentLength -gt 20971520) {
        Send-Json $stream 400 '{"ok":false,"error":"数据超过 20MB，拒绝写入"}'
        continue
      }
      $bodyText = Read-AsciiBody $reader $contentLength
      try {
        $req = $bodyText | ConvertFrom-Json
        $rel = [string]$req.path
        $b64 = [string]$req.data
        if ([string]::IsNullOrWhiteSpace($rel) -or [string]::IsNullOrWhiteSpace($b64)) { throw '缺少 path 或 data' }
        $rel = $rel.Replace('\', '/').TrimStart('/')
        $ext = [System.IO.Path]::GetExtension($rel).ToLowerInvariant()
        # 整体重写 xlsx 会压平字体/边框/列宽（这正是"删一行六张表全变"的根源），这里彻底关死
        if ($ext -eq '.xlsx' -or $ext -eq '.xls') {
          throw '为保护原有格式，表格只允许走精确改写接口，已禁止整体重写'
        }
        if ($ext -ne '.js') { throw '只允许写入 data 目录下的 .js 快照' }
        if ($rel -notmatch '^data/(homeData|usersData)\.js$') { throw '只允许写入 data/homeData.js 或 data/usersData.js' }
        $target = [System.IO.Path]::GetFullPath((Join-Path $Root ($rel.Replace('/', '\'))))
        if (-not $target.StartsWith($Root + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw '路径超出服务根目录' }
        $bytes = [Convert]::FromBase64String($b64)
        if ($ext -eq '.js') {
          # 快照文件：必须是这两个全局变量之一的赋值脚本，避免把坏内容写进页面
          $text = [System.Text.Encoding]::UTF8.GetString($bytes)
          if ($text -notmatch 'window\.__(HOME|USERS)_DATA__') { throw '快照内容校验失败，已拒绝写入' }
          Save-BytesFile $target $bytes $false
        } else {
          Save-BytesFile $target $bytes $true
        }
        $name = Split-Path -Leaf $target
        Write-Host ("  200  SAVE " + $name + " (" + $bytes.Length + " bytes)") -ForegroundColor Green
        Write-Log ("保存 " + $target + " " + $bytes.Length + " bytes")
        Send-Json $stream 200 ('{"ok":true,"name":"' + $name + '","bytes":' + $bytes.Length + '}')
      } catch {
        $msg = $_.Exception.Message
        if ($msg -eq $null) { $msg = '未知错误' }
        Write-Host ("  500  SAVE 失败: " + $msg) -ForegroundColor Red
        Write-Log ("保存失败: " + $msg)
        Send-Json $stream 500 ('{"ok":false,"error":"' + ($msg -replace '"', '\"') + '"}')
      }
      continue
    }

    # ===== 还原接口：POST /__restore —— 用 data\backup 里的备份把表格恢复原样 =====
    if ($method -eq 'POST' -and $path -eq '/__restore') {
      if ($contentLength -le 0) { Send-Json $stream 400 '{"ok":false,"error":"请求体为空"}'; continue }
      $bodyText = Read-AsciiBody $reader $contentLength
      try {
        $req = $bodyText | ConvertFrom-Json
        $rel = ([string]$req.path).Replace('\', '/').TrimStart('/')
        $pick = [string]$req.file
        if ([string]::IsNullOrWhiteSpace($rel)) { throw '缺少 path' }
        $target = [System.IO.Path]::GetFullPath((Join-Path $Root ($rel.Replace('/', '\'))))
        if (-not $target.StartsWith($Root + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw '路径超出服务根目录' }
        if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw '目标文件不存在' }
        $bdir = Join-Path (Split-Path -Parent $target) 'backup'
        if (-not (Test-Path -LiteralPath $bdir -PathType Container)) { throw '还没有任何备份，无法还原' }
        $leafName = [System.IO.Path]::GetFileNameWithoutExtension($target)
        $ext = [System.IO.Path]::GetExtension($target)
        if ([string]::IsNullOrWhiteSpace($pick)) {
          $src = Get-ChildItem -LiteralPath $bdir -Filter ("{0}_*{1}" -f $leafName, $ext) -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        } else {
          $src = Get-Item -LiteralPath (Join-Path $bdir (Split-Path -Leaf $pick)) -ErrorAction SilentlyContinue
        }
        if (-not $src) { throw '找不到可用的备份文件' }
        if (Test-FileLocked $target) { throw '文件正在 Excel / WPS 中打开，请先关闭后再还原' }
        New-Backup $target | Out-Null        # 先把当前状态也存一份，还原后仍可反悔
        [System.IO.File]::Copy($src.FullName, $target, $true)
        $name = Split-Path -Leaf $target
        Write-Host ("  200  RESTORE " + $name + " <- " + $src.Name) -ForegroundColor Yellow
        Write-Log ("还原 " + $target + " <- " + $src.FullName)
        Send-Json $stream 200 ('{"ok":true,"from":"' + $src.Name + '","bytes":' + $src.Length + '}')
      } catch {
        $msg = $_.Exception.Message
        if ($null -eq $msg) { $msg = '未知错误' }
        $safe = $msg -replace '"', '\"' -replace "`r", ' ' -replace "`n", ' '
        Send-Json $stream 500 ('{"ok":false,"error":"' + $safe + '"}')
      }
      continue
    }

    $candidate = [System.IO.Path]::GetFullPath((Join-Path $Root $path.TrimStart('/').Replace('/', '\')))
    if (-not $candidate.StartsWith($Root + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
      Send-Text $stream 403 "403 Forbidden"
      continue
    }
    if (Test-Path -LiteralPath $candidate -PathType Container) { $candidate = Join-Path $candidate 'index.html' }

    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
      $bytes = Read-Shared $candidate
      $ext = [System.IO.Path]::GetExtension($candidate).ToLowerInvariant()
      $ctype = if ($mime.ContainsKey($ext)) { $mime[$ext] } else { 'application/octet-stream' }
      $head = "HTTP/1.1 200 OK`r`nContent-Type: $ctype`r`nContent-Length: $($bytes.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n`r`n"
      if ($method -eq 'HEAD') {
        Send-Raw $stream ([System.Text.Encoding]::ASCII.GetBytes($head)) $null
      } else {
        Send-Raw $stream ([System.Text.Encoding]::ASCII.GetBytes($head)) $bytes
      }
      Write-Host ("  200  " + $path) -ForegroundColor DarkGray
    } else {
      Send-Text $stream 404 ("404 Not Found: " + $path)
      Write-Host ("  404  " + $path) -ForegroundColor DarkYellow
    }
  } catch {
    Write-Host ("  [异常] " + $_.Exception.Message) -ForegroundColor Red
    Write-Log ("异常: " + $_.Exception.Message)
    try { Send-Text $stream 500 "500 Internal Server Error" } catch { }
  } finally {
    if ($client) { try { $client.Close() } catch { } }
  }
}
