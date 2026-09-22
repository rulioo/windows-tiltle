# pack.ps1 —— 打发布包: 把 DeskTiler.exe / 收款码 / README / 安装说明 放进 install/,
# 再压成一个带顶层文件夹的 zip。重复运行会覆盖上一次的产物。
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\pack.ps1
#
# 打包前请先 build.cmd 出新的 exe —— 本脚本会检查 exe 是不是比 uMain.pas 旧(旧了只警告, 不拦)。
# install/ 已列入 .gitignore(二进制不进仓库)。

param([string]$Root = "")

$ErrorActionPreference = 'Stop'
if ($Root -eq "") { $Root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..")) }

$exe = Join-Path $Root 'DeskTiler.exe'
$jpg = Join-Path $Root '300收款码.jpg'
$rmd = Join-Path $Root 'README.md'
$txt = Join-Path $Root 'tools\install-readme.txt'
$out = Join-Path $Root 'install'

foreach($f in @($exe,$jpg,$rmd,$txt)){
  if(-not (Test-Path $f)){ throw "缺文件: $f" }
}

# ---- 版本: 以 exe 里嵌的版本资源为准(它是 build.cmd 编译时写进去的) ----
# 取不到就退回 build.number —— 编译失败也会让它 +1, 所以 exe 里的才是准的。
$vi  = (Get-Item $exe).VersionInfo
$m   = [regex]::Match([string]$vi.FileVersion, '^\d+\.\d+\.\d+')
if($m.Success){
  $ver = $m.Value
} else {
  $ver = '1.0.' + ((Get-Content (Join-Path $Root 'build.number') -Raw).Trim())
  Write-Warning "exe 的版本资源读不出干净版本号(读到的是 '$($vi.FileVersion)'), 改用 build.number"
}

# ---- 陈旧检查: uMain.pas 比 exe 新说明还没重新编译 ----
$srcNewer = (Get-Item (Join-Path $Root 'uMain.pas')).LastWriteTime -gt (Get-Item $exe).LastWriteTime
if($srcNewer){ Write-Warning "uMain.pas 比 DeskTiler.exe 新 —— 先跑 build.cmd 再打包, 否则包里是旧程序" }

# ---- 摆平目录(清掉上一版的 zip, 免得新旧混在一起) ----
if(-not (Test-Path $out)){ New-Item -ItemType Directory -Path $out | Out-Null }
Get-ChildItem $out -Filter 'DeskTiler-v*-win32.zip' -ErrorAction SilentlyContinue | Remove-Item -Force

Copy-Item $exe (Join-Path $out 'DeskTiler.exe') -Force
Copy-Item $jpg (Join-Path $out '300收款码.jpg') -Force
Copy-Item $rmd (Join-Path $out 'README.md') -Force

# 安装说明: 正文是仓库里的 tools/install-readme.txt(@VER@ 占位), 这里替换版本号,
# 并用 UTF-8 **带 BOM** 落盘 —— 记事本等工具才不会把中文读成乱码。
$body = [IO.File]::ReadAllText($txt, [Text.Encoding]::UTF8).Replace('@VER@', "v$ver")
[IO.File]::WriteAllText((Join-Path $out '安装说明.txt'), $body, (New-Object Text.UTF8Encoding($true)))

# ---- 压包: 顶层带一个同名文件夹, 解压出来不会散一地 ----
$topName = "DeskTiler-v$ver"
$stage   = Join-Path $env:TEMP ("ds_pack_" + [Guid]::NewGuid().ToString('N'))
$inner   = Join-Path $stage $topName
New-Item -ItemType Directory -Path $inner -Force | Out-Null
try {
  Get-ChildItem $out -File | Where-Object { $_.Extension -ne '.zip' } | ForEach-Object {
    Copy-Item $_.FullName $inner -Force
  }
  $zip = Join-Path $out ($topName + '-win32.zip')
  Compress-Archive -Path $inner -DestinationPath $zip -Force
} finally {
  Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output "DeskTiler 发布包 v$ver"
Get-ChildItem $out | Sort-Object Name | ForEach-Object {
  Write-Output ("  {0,-32} {1,10:N0} 字节" -f $_.Name, $_.Length)
}
Write-Output "  -> $out"
