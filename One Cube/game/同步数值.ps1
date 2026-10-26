# 同步数值.ps1 —— 把「数值清单.md」里的 tune 块回写进 game/OneCube.html
#
# 用法（本机是 Windows PowerShell 5.1，没有 pwsh；默认策略禁止运行脚本，所以要带 Bypass）：
#   powershell -ExecutionPolicy Bypass -File 同步数值.ps1            # 预览：只报告 旧值 → 新值，不改文件
#   powershell -ExecutionPolicy Bypass -File 同步数值.ps1 -Apply     # 实际写入：先备份 → 改 → node --check 校验
#
# 注意：本文件必须保存为「UTF-8 带 BOM」。PowerShell 5.1 读无 BOM 的 .ps1 会按 ANSI 解码，
#       里面的中文会变乱码并把引号配对弄坏导致语法错误。
#
# 约定：数值清单.md 里必须有一个 ```tune 代码块，内容形如
#   PHY.G = 10
#   score.shoot = 800
# 行内以 # 开头的是注释，空行忽略。
#
# 新增可调项时：在下面的 $MAP 里加一条 'key' = '正则（用第一个捕获组圈住数值）' 即可。

[CmdletBinding()]
param(
  [switch]$Apply,
  [string]$MdPath,
  [string]$HtmlPath
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $MdPath)   { $MdPath   = Join-Path $here '数值清单.md' }
if (-not $HtmlPath) {
  # 兼容两种布局：脚本与 OneCube.html 同目录（都在 game\ 里），或脚本在上一层（OneCube\ 里）
  $same = Join-Path $here 'OneCube.html'
  if (Test-Path $same) { $HtmlPath = $same } else { $HtmlPath = Join-Path $here 'game\OneCube.html' }
}

if (-not (Test-Path $MdPath))   { throw "找不到数值清单: $MdPath" }
if (-not (Test-Path $HtmlPath)) { throw "找不到游戏文件: $HtmlPath" }

# key -> 精确定位用的正则（第一个捕获组必须是数值本身，其余部分原样保留）
$MAP = [ordered]@{
  # 1. 逻辑与物理
  'PHY.G'          = 'var PHY = \{ G:(-?[\d.]+)'
  'PHY.MAXSPEED'   = 'MAXSPEED:(-?[\d.]+)'
  'PHY.ACC'        = 'ACC:(-?[\d.]+)'
  'PHY.JUMP_H'     = 'JUMP_H:(-?[\d.]+)'
  'PHY.HALF'       = 'HALF:(-?[\d.]+)'
  'CROSS_GATE'     = 'var CROSS_GATE = (-?[\d.]+);'
  # 2. 相机
  'cam.yaw'        = 'var cam=\{ yaw:(-?[\d.]+)'
  'cam.pitch'      = 'var cam=\{ yaw:-?[\d.]+, pitch:(-?[\d.]+)'
  'cam.dist'       = 'var cam=\{ yaw:-?[\d.]+, pitch:-?[\d.]+, dist:(-?[\d.]+)'
  'orbit.sens'     = 'setFromAxisAngle\(axis, ulen\*([\d.]+)\)'
  # 3. 计分
  'score.init'     = 'newScore\(\)\{ return \{ score:(-?\d+)'
  'score.shoot'    = 'score\.score-=(-?\d+); updateHUD\(\);'
  'score.warn'     = 'score\.score<(-?\d+)\) sfx\.tick\(\);'
  'score.bonusRed' = 'if \(o\.type===OBS\.RED\) return (-?\d+);'
  'score.bonusBlue'= 'if \(o\.type===OBS\.BLUE\) return (-?\d+);'
  'score.bonusBlk' = 'OBS\.BLUE\) return \d+;\s*return (-?\d+);'   # 黑障奖励是裸 return（不是 if 分支）
  'score.portal'   = 'score\.score\+=(-?\d+); updateHUD\(\);'
  # 4. 子弹 / 残影 / 粒子
  'bullet.speed'   = 'var speed=(-?[\d.]+), STEP='
  'bullet.life'    = 'dv:aimDir\[1\], life:(-?[\d.]+)'
  'trail.life'     = 'var speed=-?[\d.]+, STEP=-?[\d.]+, TRAIL_LIFE=(-?[\d.]+);'
  'trail.step'     = 'var speed=-?[\d.]+, STEP=(-?[\d.]+), TRAIL_LIFE='
  'trail.px'       = 'var trails=\[\], TRAIL_PX=(\d+);'
  'fx.life'        = 'var FX_LIFE=(-?[\d.]+), FX_PX='
  'fx.px'          = 'var FX_LIFE=-?[\d.]+, FX_PX=(\d+),'
  'fx.n'           = 'FX_PX=-?[\d.]+, FX_N=(\d+),'
  'fx.spd'         = 'FX_N=-?[\d.]+, FX_SPD=(\d+),'
  'fx.drag'        = 'FX_SPD=-?[\d.]+, FX_DRAG=(-?[\d.]+),'
  'fx.g'           = 'FX_DRAG=-?[\d.]+, FX_G=(\d+);'
  # 6. UI
  'slice.px'       = 'var PX=(\d+),SH='
  'slice.sh'       = 'var PX=-?\d+,SH=(\d+),SL='
  'ui.btnW'        = '#btn-editor, #btn-import\{ width:(\d+)px'          # 主按钮 #btn-start 是 168，见《数值清单》附注
  'ui.cardSize'    = '#level-grid\{ display:grid; grid-template-columns:repeat\(5,(\d+)px\)'
  'ui.perPage'     = 'var LEVEL_PAGE=0, LVL_PER_PAGE=(\d+);'
}

# ---- 解析 MD 里的 tune 块 ----
$md = Get-Content -Raw -Encoding UTF8 $MdPath
$m = [regex]::Match($md, '(?s)```tune\r?\n(.*?)```')
if (-not $m.Success) { throw "数值清单.md 里没有找到 ```tune 代码块" }
$want = [ordered]@{}
foreach ($line in ($m.Groups[1].Value -split "`r?`n")) {
  $s = ($line -replace '#.*$','').Trim()
  if ($s -eq '' -or $s.StartsWith('#')) { continue }
  $kv = [regex]::Match($s, '^([\w\.]+)\s*=\s*(\S+)\s*$')
  if (-not $kv.Success) { Write-Warning "跳过无法解析的行: $s"; continue }
  $want[$kv.Groups[1].Value] = $kv.Groups[2].Value
}
Write-Host ("清单里的数值项: {0} 个" -f $want.Count) -ForegroundColor Cyan

# ---- 应用到 HTML ----
$html = Get-Content -Raw -Encoding UTF8 $HtmlPath
$changed = 0; $same = 0; $missing = @()
foreach ($key in $want.Keys) {
  if (-not $MAP.Contains($key)) { $missing += "$key（脚本里没有对应正则）"; continue }
  $pat = $MAP[$key]; $val = $want[$key]
  $rx = [regex]::new($pat)
  $mm = $rx.Matches($html)
  if ($mm.Count -eq 0) { $missing += "$key（正则没匹配到）"; continue }
  if ($mm.Count -gt 1)  { $missing += "$key（匹配到 $($mm.Count) 处，需收紧正则）"; continue }
  $old = $mm[0].Groups[1].Value
  if ($old -eq $val) { $same++; Write-Host ("  = {0,-16} {1}" -f $key, $val) -ForegroundColor DarkGray; continue }
  $changed++
  Write-Host ("  * {0,-16} {1}  ->  {2}" -f $key, $old, $val) -ForegroundColor Yellow
  $html = $rx.Replace($html, { param($x) $x.Value.Replace($x.Groups[1].Value, $val) }, 1)
}
if ($missing.Count) {
  Write-Host "`n以下项未应用：" -ForegroundColor Red
  $missing | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
}
Write-Host ("`n汇总：需修改 {0} 项，已一致 {1} 项，未应用 {2} 项" -f $changed, $same, $missing.Count)

if (-not $Apply) {
  Write-Host "`n（预览模式，未写入。加 -Apply 才会真正修改并校验）" -ForegroundColor Cyan
  exit 0
}
if ($changed -eq 0) { Write-Host "无需写入。" -ForegroundColor Cyan; exit 0 }

# ---- 备份 → 写入 → 语法校验 ----
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$bak = "$HtmlPath.bak_$stamp"
Copy-Item $HtmlPath $bak -Force
[System.IO.File]::WriteAllText($HtmlPath, $html, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "`n已写入。备份: $bak" -ForegroundColor Green

# 抽取最后一个 <script> 做 node --check
$lines = [System.IO.File]::ReadAllLines($HtmlPath)
$s1 = @(); for ($i=0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match '^\s*<script(\s|>)') { $s1 += $i } }
$s2 = @(); for ($i=0; $i -lt $lines.Count; $i++) { if ($lines[$i] -match '^\s*</script>')        { $s2 += $i } }
$tmp = Join-Path $env:TEMP 'oc_tune_check.js'
[System.IO.File]::WriteAllText($tmp, ($lines[($s1[-1]+1)..($s2[-1]-1)] -join "`n"), (New-Object System.Text.UTF8Encoding($false)))
$node = 'C:\Program Files\nodejs\node.exe'
if (Test-Path $node) {
  & $node --check $tmp 2>&1 | Select-Object -First 5
  if ($LASTEXITCODE -eq 0) { Write-Host "语法校验通过 ✅" -ForegroundColor Green }
  else { Write-Host "语法校验失败 ✗ —— 请用备份回滚: $bak" -ForegroundColor Red }
} else { Write-Host "未找到 node.exe，跳过语法校验" -ForegroundColor Yellow }
Remove-Item $tmp -Force -ErrorAction SilentlyContinue
