# makeicon.ps1 - 生成 DeskTiler.ico (32x32, 32bpp 带透明)
# 图案: 透明底上 2x2 四个“蓝顶窗口”方块, 呼应“窗口平铺”主题。

param([string]$Out = "")

if ($Out -eq "") {
    $Out = Join-Path $PSScriptRoot "..\DeskTiler.ico"
}
$Out = [System.IO.Path]::GetFullPath($Out)

$size = 32
$xordata = New-Object byte[] ($size * $size * 4)  # 初始全透明

# 四个格子的布局
$m = 3     # 外边距
$g = 2     # 格子间距
$cw = ($size - 2 * $m - $g) / 2   # 每个格子边长 12
$cols = 2

function Set-Px([int]$cx, [int]$cy, [int]$r, [int]$gg, [int]$b) {
    if ($cx -lt 0 -or $cx -ge 32 -or $cy -lt 0 -or $cy -ge 32) { return }
    $row = 31 - $cy            # 图标 DIB 为自下而上
    $off = ($row * 32 + $cx) * 4
    $xordata[$off]     = $b
    $xordata[$off + 1] = $gg
    $xordata[$off + 2] = $r
    $xordata[$off + 3] = 255
}

$blueR, $blueG, $blueB = 24, 88, 238     # #1858EE
$bodyR, $bodyG, $bodyB = 236, 244, 255   # #ECF4FF

for ($ci = 0; $ci -lt 2; $ci++) {
    for ($ri = 0; $ri -lt 2; $ri++) {
        $x0 = $m + $ci * ($cw + $g)
        $y0 = $m + $ri * ($cw + $g)
        for ($ly = 0; $ly -lt $cw; $ly++) {
            for ($lx = 0; $lx -lt $cw; $lx++) {
                $isTitle = ($ly -le 2)          # 顶部标题栏(含顶边)
                $isEdge  = ($lx -eq 0 -or $lx -eq ($cw - 1) -or $ly -eq ($cw - 1))
                if ($isTitle -or $isEdge) {
                    Set-Px ($x0 + $lx) ($y0 + $ly) $blueR $blueG $blueB
                } else {
                    Set-Px ($x0 + $lx) ($y0 + $ly) $bodyR $bodyG $bodyB
                }
            }
        }
    }
}

# AND mask: 全部 0(以 alpha 表示透明)
$andmask = New-Object byte[] 128

$imgLen = 40 + $xordata.Length + $andmask.Length
$entryOffset = 6 + 16  # ICONDIR(6) + ICONDIRENTRY(16)

$fs = [System.IO.File]::Create($Out)
try {
    $bw = New-Object System.IO.BinaryWriter($fs)

    # ICONDIR
    $bw.Write([uint16]0)          # reserved
    $bw.Write([uint16]1)          # type = icon
    $bw.Write([uint16]1)          # count

    # ICONDIRENTRY
    $bw.Write([byte]$size)        # width  (0 => 256)
    $bw.Write([byte]$size)        # height
    $bw.Write([byte]0)            # colors
    $bw.Write([byte]0)            # reserved
    $bw.Write([uint16]1)          # planes
    $bw.Write([uint16]32)         # bitcount
    $bw.Write([uint32]$imgLen)    # bytes in resource
    $bw.Write([uint32]$entryOffset)

    # BITMAPINFOHEADER
    $bw.Write([uint32]40)                       # biSize
    $bw.Write([int32]$size)                     # biWidth
    $bw.Write([int32]($size * 2))               # biHeight (XOR+AND)
    $bw.Write([uint16]1)                        # biPlanes
    $bw.Write([uint16]32)                       # biBitCount
    $bw.Write([uint32]0)                        # biCompression = BI_RGB
    $bw.Write([uint32]$imgLen)                  # biSizeImage
    $bw.Write([int32]0)                         # biXPelsPerMeter
    $bw.Write([int32]0)                         # biYPelsPerMeter
    $bw.Write([uint32]0)                        # biClrUsed
    $bw.Write([uint32]0)                        # biClrImportant

    $bw.Write($xordata)
    $bw.Write($andmask)
    $bw.Flush()
    $bw.Close()
} finally {
    $fs.Dispose()
}

Write-Output "icon written: $Out"
