# verify.ps1 - DeskTiler 功能自检(含: 两列布局, 目标显示器下拉, 全选复选框, 平铺后激活)
# 全部在同一个 PowerShell 进程/桌面内完成, 保证窗口可被枚举。
#   1) 界面元素: 一键全排/Cmd快排/PowerShell快排/目录快排/应用快排 按钮(含从左到右的顺序)
#      + 顶部 全选/选定窗口置顶/本窗口置顶 复选框 + 计数
#   2) 计数联动: 顶部“全选”复选框 打勾/取消 联动 “已选择”(BM_GETCHECK 校验状态)
#   2b) 两列布局 & 显示器下拉: 表头列数==2(无 PID); TComboBox 下拉 >=2 且含 “自动选择”
#   2c) 顶栏 5 个左侧控件顺序/不重叠(含 选定窗口置顶 / 本窗口置顶)
#   2d) 置顶行为: 选定窗口置顶 -> 目标窗口 WS_EX_TOPMOST 置位/清除; 本窗口置顶 -> 主窗体自身置位/清除
#   2e) 点选列表一行: 「应用快排」按钮改名为「<应用名>快排」并按文字加宽(文字量得下 + 底行重排不错位)
#   2f) 列表右键菜单: 右键某行弹出「结束进程/转到应用」两项;
#       结束进程 -> 原生确认框点名应用与 PID, 点「否」不杀(进程仍在), 点「是」真杀(进程退出);
#       转到应用 -> 该行窗口被激活到前台。这一段自己起一个新实例当靶子, 只动自己造的进程。
#   2g) 「转到应用」+「本窗口置顶」: 重叠时工具窗口先最小化让开, 靶子中心真的露出来且成为前台;
#       错开摆放(不重叠)时不让开。这一段也自己起一个靶子实例。
#   2h) 右键列表某一行: 选中的正好是那一行(单选, 不累加), 「应用快排」跟着那一行走 ——
#       行号断言在“列表冻着”时做(开着自动刷新的话表每 2.5 秒按 Z 序重排, 行号不是身份);
#       之后再临时打开自动刷新跨过一轮整表重建, 量“选中的还是同一个窗口、仍旧只有一行”。
#       期望值不写死窗口标题 —— 都现取(投递一次左键问程序自己), 机器上的真实窗口随时会变。
#   3) 表头排序: -sorttest 在程序内按 PID/应用 升序/降序各排一次(真实 SortList 路径)
#   4) -tileapp DeskTiler 平铺两个目标 -> 校验为均分网格(等大、不相交、相邻)
#   说明: 本机单显示器, 手动指定屏的“跨屏平铺”无法自动化; 覆盖控件存在 + 自动兜底路径
#   说明: (2f) 用真光标 + 真鼠标事件驱动原生菜单(它的模态循环不认 PostMessage 造的鼠标消息),
#         落点前用 WindowFromPoint 确认底下就是被测程序自己的列表/菜单, 不是就不点; 结束后还原光标。
#         落点用**客户区原点**算(窗口矩形含 2px 边框, 直接拿它当客户区原点会偏 2px),
#         并且挑行的中间而不是上边缘 —— 行顶就是行边界, 偏一点点就落到上一行去了。
$ErrorActionPreference = 'Continue'
$exe = 'E:\cc\windows-tiltle\DeskTiler.exe'
$log = 'E:\cc\windows-tiltle\verify.log'
$sortFile = 'E:\cc\windows-tiltle\sorttest.txt'
Remove-Item $log -ErrorAction SilentlyContinue
Remove-Item $sortFile -ErrorAction SilentlyContinue
$script:fail = 0
function Log($m){ Add-Content -Path $log -Value $m -Encoding UTF8; Write-Output $m }
function Chk($ok,$msg){ if($ok){ Log ("OK   " + $msg) } else { Log ("FAIL " + $msg); $script:fail++ } }

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public struct VERECT { public int L,T,R,B; }
public struct VEPT { public int X,Y; }
public class VW {
  public delegate bool ChildEnum(IntPtr h, IntPtr l);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, ChildEnum cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  // 右键菜单是原生弹出菜单(#32768), 它的模态循环按“真实光标位置”跟踪,
  // PostMessage 造的鼠标消息它不认 —— 只能喂真光标 + 真鼠标事件。
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern bool GetCursorPos(out VEPT p);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, int dx, int dy, uint d, IntPtr e);
  [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(VEPT p);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr h, int id);
  [DllImport("user32.dll")] public static extern int GetMenuItemCount(IntPtr m);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out VERECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out VERECT r);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextLengthW(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool PostMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowLongW(IntPtr h, int idx);   // idx=-20 -> GWL_EXSTYLE
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint f);
  [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint f);    // f=2 -> GA_ROOT
  [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint f, IntPtr extra);
  [DllImport("kernel32.dll", CharSet=CharSet.Unicode)] public static extern int CompareStringW(int locale, int dwCmpFlags, string a, int ca, string b, int cb);
}
"@

# ---- 子窗口收集 ----
function Get-ChildInfo($root){
  $res = New-Object System.Collections.Generic.List[object]
  $cb = { param($h,$l)
    $len=[VW]::GetWindowTextLengthW($h)
    $cap=''
    if($len -gt 0){ $sb=New-Object System.Text.StringBuilder ($len+1); [void][VW]::GetWindowTextW($h,$sb,$sb.Capacity); $cap=$sb.ToString() }
    $sb2=New-Object System.Text.StringBuilder 128; [void][VW]::GetClassNameW($h,$sb2,128)
    $res.Add([pscustomobject]@{ Hwnd=$h; Cls=$sb2.ToString(); Cap=$cap })
    return $true }
  $script:g = [Runtime.InteropServices.GCHandle]::Alloc($cb)
  try { [void][VW]::EnumChildWindows($root,$cb,[IntPtr]::Zero) } finally { $script:g.Free() }
  return $res
}
function Get-Cap($h){
  $len=[VW]::GetWindowTextLengthW($h)
  if($len -le 0){ return '' }
  $sb=New-Object System.Text.StringBuilder ($len+1)
  [void][VW]::GetWindowTextW($h,$sb,$sb.Capacity)
  return $sb.ToString()
}
function Get-RectOf($h){
  $r = New-Object VERECT
  [void][VW]::GetWindowRect($h,[ref]$r)
  return $r
}
# 客户区左上角的屏幕坐标。窗口矩形**不等于**客户区: 列表有 2px 边框, 用窗口矩形去算
# 屏幕落点, 真点击落进客户区时会比验证过的 y 少 2 —— 踩着行边界就翻到上一行去了。
# 边框按对称算((窗口宽 - 客户宽)/2), 标准边框都成立。
function Get-ClientOrigin($h){
  $r = New-Object VERECT; [void][VW]::GetWindowRect($h,[ref]$r)
  $c = New-Object VERECT; [void][VW]::GetClientRect($h,[ref]$c)
  $bx = [int](($r.R - $r.L - $c.R) / 2)
  $by = [int](($r.B - $r.T - $c.B) / 2)
  return ,@(($r.L + $bx),($r.T + $by))
}
# 两个矩形是否相交(边贴边不算)
function Rects-Overlap($a,$b){
  return ($a.L -lt $b.R -and $b.L -lt $a.R -and $a.T -lt $b.B -and $b.T -lt $a.B)
}
function EX-TopMost($h){ return (([VW]::GetWindowLongW($h,-20)) -band 0x00000008) -ne 0 }   # WS_EX_TOPMOST
# WindowFromPoint 给的是最底下的子窗口(列表/按钮…), 要判断“这一片屏幕上盖着的是谁”得回溯到顶层
function Root-Of($h){ if($h -eq [IntPtr]::Zero){ return [IntPtr]::Zero }; return [VW]::GetAncestor($h,2) }
# 文字在“窗体那套字体”(Microsoft YaHei UI 9pt)下的像素宽 —— 走的也是 GDI 度量,
# 与程序内 Self.Canvas.TextWidth 同一套算法, 所以“按钮够不够宽装下这行字”可以直接断言。
# (本机 100% 缩放: 顶栏按钮实测宽度与代码里写的 80/86/92/126/108 一模一样, 9pt 即 12px。)
function Measure-Text($s,$bold = $false){
  Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
  Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
  $f = $null; $bmp = $null; $gfx = $null
  try {
    $style = if($bold){ [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $f = New-Object System.Drawing.Font('Microsoft YaHei UI',9,$style)
    $bmp = New-Object System.Drawing.Bitmap 1,1
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $flags = [System.Windows.Forms.TextFormatFlags]::NoPadding
    return [System.Windows.Forms.TextRenderer]::MeasureText($s,$f,(New-Object System.Drawing.Size 2000,100),$flags).Width
  } catch { return -1 }
  finally { if($gfx){$gfx.Dispose()}; if($bmp){$bmp.Dispose()}; if($f){$f.Dispose()} }
}
function Parse-Status($s){
  if($s -match '窗口总数:\s*(\d+).*已选择:\s*(\d+)'){ return @([int]$matches[1],[int]$matches[2]) }
  return $null
}
function Same-Seq($a,$b){
  if($a.Count -ne $b.Count){ return $false }
  for($k=0;$k -lt $a.Count;$k++){ if($a[$k] -ne $b[$k]){ return $false } }
  return $true
}
function CmpStr($x,$y){
  # 与应用内一致的区域比较: LOCALE_USER_DEFAULT($400) + NORM_IGNORECASE(1)
  # 返回 1=小于 2=等于 3=大于
  return [VW]::CompareStringW(0x0400,0x1,$x,$x.Length,$y,$y.Length)
}
function Start-Run($argsStr){
  # 同桌面启动并等待退出(用于 -tileapp/-sorttest 等自测 CLI)
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $exe
  $psi.Arguments = $argsStr
  $q = [System.Diagnostics.Process]::Start($psi)
  $q.WaitForExit()
  return $q.ExitCode
}

# ---- 顶层窗口查找(按类名; 原生弹出菜单 #32768 / 消息框 #32770 都是顶层窗口) ----
function Find-TopWindow($cls){
  $script:foundH = [IntPtr]::Zero
  $cb = { param($h,$l)
    if($script:foundH -ne [IntPtr]::Zero){ return $true }
    if([VW]::IsWindowVisible($h)){
      $sb = New-Object System.Text.StringBuilder 128
      [void][VW]::GetClassNameW($h,$sb,128)
      if($sb.ToString() -eq $cls){ $script:foundH = $h }
    }
    return $true }
  $script:gTop = [Runtime.InteropServices.GCHandle]::Alloc($cb)
  try { [void][VW]::EnumWindows($cb,[IntPtr]::Zero) } finally { $script:gTop.Free() }
  return $script:foundH
}
function To-Pt($x,$y){ $p = New-Object VEPT; $p.X=[int]$x; $p.Y=[int]$y; return $p }
# 真光标 + 真鼠标事件(不是 PostMessage): 原生菜单只认真实输入
function Real-Click($x,$y,$right){
  [void][VW]::SetCursorPos([int]$x,[int]$y)
  Start-Sleep -Milliseconds 200
  if($right){
    [VW]::mouse_event(0x0008,0,0,0,[IntPtr]::Zero); Start-Sleep -Milliseconds 80
    [VW]::mouse_event(0x0010,0,0,0,[IntPtr]::Zero)
  } else {
    [VW]::mouse_event(0x0002,0,0,0,[IntPtr]::Zero); Start-Sleep -Milliseconds 80
    [VW]::mouse_event(0x0004,0,0,0,[IntPtr]::Zero)
  }
}
# 在列表某行上右键唤出菜单, 返回菜单窗口句柄(0 = 没弹出来)。
# 参数是**客户端** y(行号换算出来的, 窗口怎么挪都不变); 屏幕坐标必须用客户区原点现算 ——
# 踩过两个坑: 一是缓存了旧的 rect, 二是拿窗口矩形当客户区原点(差 2px 边框),
# 两次都让真右键按到了上一行(菜单里点名的成了 Weixin 而不是靶子)。
function Open-RowMenu($clientY){
  $o = Get-ClientOrigin $script:hLv
  $rowX = $o[0] + 60
  $rowY = $o[1] + $clientY
  # 安全闸: 落点必须真的在被测程序自己的列表上。被别的窗口盖住就宁可不动手 ——
  # 真鼠标一点下去是会打到盖在上面那个窗口上的。
  $hit = [VW]::WindowFromPoint((To-Pt $rowX $rowY))
  if($hit -ne $script:hLv){
    $null = Log ("  row point ({0},{1}) is covered by hwnd={2} (listview={3}) -> abort menu" -f $rowX,$rowY,([int64]$hit),([int64]$script:hLv))
    return [IntPtr]::Zero
  }
  # 退到前台这一步在别的进程抢焦点时可能失败(Win 的前台锁), 失败时那一下右键就只够把窗口激活、
  # 不会弹菜单 —— 所以按“再点一次”处理: 最多试三轮, 每轮都把当次的前台窗口记下来。
  for($try=1; $try -le 3; $try++){
    [void][VW]::SetForegroundWindow($script:hMain)
    Start-Sleep -Milliseconds 350
    if($try -eq 1){ $null = Log ("  foreground before right-click: {0} (driver {1})" -f ([int64][VW]::GetForegroundWindow()),([int64]$script:hMain)) }
    Real-Click $rowX $rowY $true
    for($k=0; $k -lt 8; $k++){
      Start-Sleep -Milliseconds 200
      $m = Find-TopWindow '#32768'
      if($m -ne [IntPtr]::Zero){ return $m }
    }
    $null = Log ("  right-click attempt {0}: no menu (fg={1})" -f $try,([int64][VW]::GetForegroundWindow()))
    Start-Sleep -Milliseconds 250
  }
  return [IntPtr]::Zero
}
# 点菜单里的第 frac 段(两项时: 0.25=第一项 0.75=第二项)
function Click-MenuItem($m,$frac){
  $mr = Get-RectOf $m
  $x = [int]($mr.L + ($mr.R-$mr.L)/2)
  $y = [int]($mr.T + ($mr.B-$mr.T)*$frac)
  $hit = [VW]::WindowFromPoint((To-Pt $x $y))
  if($hit -ne $m){
    $null = Log ("  menu item point ({0},{1}) hit hwnd={2}, not the menu {3} -> abort click" -f $x,$y,([int64]$hit),([int64]$m))
    return $false
  }
  Real-Click $x $y $false
  Start-Sleep -Milliseconds 900
  return $true
}
# 读原生消息框的正文(它的正文是有句柄的 Static, 跨进程读得到; VCL 的 MessageDlg 就读不到)
function Get-DialogText($dlg){
  $st = @(Get-ChildInfo $dlg) | Where-Object { $_.Cls -eq 'Static' -and $_.Cap -ne '' } | Select-Object -First 1
  if($st){ return $st.Cap }
  return ''
}

# ============ 清理历史 DeskTiler 实例(只保留两个不可杀死的无窗僵尸) ============
Log '== cleanup stale =='
Get-Process DeskTiler -ErrorAction SilentlyContinue | ForEach-Object {
  try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; Log ("killed stale pid {0}" -f $_.Id) }
  catch { Log ("cannot kill stale pid {0} (windowless, ok)" -f $_.Id) }
}
Start-Sleep -Milliseconds 900

Log '== verify start =='
# ============ 启动两个目标实例 ============
$p1 = [System.Diagnostics.Process]::Start($exe)
$p2 = [System.Diagnostics.Process]::Start($exe)
function Wait-Hwnd($p){
  for($k=0;$k -lt 15;$k++){
    Start-Sleep -Milliseconds 700
    $p.Refresh()
    if($p.MainWindowHandle -ne [IntPtr]::Zero){ return $p.MainWindowHandle }
  }
  return [IntPtr]::Zero
}
$h1 = Wait-Hwnd $p1
$h2 = Wait-Hwnd $p2
Log ("target pids: {0},{1} hwnd: {2},{3}" -f $p1.Id,$p2.Id,$h1.ToInt64(),$h2.ToInt64())
Chk ($h1 -ne [IntPtr]::Zero -and $h2 -ne [IntPtr]::Zero) '两个目标窗口都创建成功'

if($h1 -ne [IntPtr]::Zero -and $h2 -ne [IntPtr]::Zero){
  $kids = Get-ChildInfo $h1

  # ---- 关闭自动刷新, 让测试期间列表/计数稳定 ----
  $auto = $kids | Where-Object { $_.Cap -eq '自动刷新' } | Select-Object -First 1
  if($auto){
    [void][VW]::SendMessageW($auto.Hwnd, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)  # BM_CLICK -> 取消勾选, 停 Timer
    Start-Sleep -Milliseconds 600
    $kids = Get-ChildInfo $h1
    Log 'auto-refresh disabled for test'
  } else { Log 'WARN: 自动刷新复选框未找到' }

  # ---- 强制重新枚举一次: 列表是启动时那一帧的快照, 关掉自动刷新后就冻住了,
  #      而第二个目标窗口可能是在第一个目标建表之后才出现的 -> 不刷新就漏掉它。
  #      (踩过: “批勾的窗口被置顶”断言因此假失败, 因为 h2 压根不在表里。) ----
  $ref = $kids | Where-Object { $_.Cap -eq '刷新(&R)' } | Select-Object -First 1
  if($ref){
    [void][VW]::SendMessageW($ref.Hwnd, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)   # BM_CLICK
    Start-Sleep -Milliseconds 1200
    $kids = Get-ChildInfo $h1
    Log 'forced re-enumeration (刷新) so both targets are in the list'
  } else { Log 'WARN: 刷新按钮未找到' }
  $p2.Refresh(); Log ("[diag] 第二个实例存活 @刷新后 = " + (-not $p2.HasExited))

  # ---- (1) 快捷按钮 & 顶部复选框 ----
  $caps = @($kids | ForEach-Object { $_.Cap })
  Chk (($caps -contains '一键全排') -and ($caps -contains 'Cmd快排')) '按钮: 一键全排 / Cmd快排 均存在'
  # PowerShell快排: 平铺全部 Windows Terminal / PowerShell 窗口(2026-09 按用户要求加回)
  Chk ($caps -contains 'PowerShell快排') 'PowerShell快排 按钮存在(平铺全部 Windows Terminal 窗口)'
  # 目录快排/应用快排: 2026-09-17 按用户要求新增的第 4、5 个快捷按钮
  Chk ($caps -contains '目录快排') '目录快排 按钮存在(平铺全部 explorer 窗口)'
  Chk ($caps -contains '应用快排') '应用快排 按钮存在(平铺列表中选中那行所属应用的全部窗口)'
  # 主按钮 2026-09-17 由「平铺排列(&T)」改名为「平铺所选应用(&T)」(动作不变: 平铺勾选的那些窗口)
  Chk ($caps -contains '平铺所选应用(&T)') '主按钮: 平铺所选应用(&T) 存在(由“平铺排列”改名而来)'
  $btnTile = $kids | Where-Object { $_.Cap -eq '平铺所选应用(&T)' } | Select-Object -First 1
  if($btnTile){
    # 名字长了两个字(还带粗体的 "(&T)"), 宽度必须跟着加宽, 否则文字会被截掉 —— 按同一字体量一遍
    $rt = Get-RectOf $btnTile.Hwnd
    $wt = $rt.R - $rt.L
    $needT = Measure-Text '平铺所选应用(&T)' $true
    Log ("  main btn W={0} (L={1} R={2}); bold text needs {3}px" -f $wt,$rt.L,$rt.R,$needT)
    Chk ($needT -gt 0 -and $wt -ge $needT + 6) ('主按钮够宽: 粗体文字 {0}px + 留白 <= 按钮 {1}px' -f $needT,$wt)
    # 加宽后不能压到左边那排选择框(列数/间距/显示器 都在同一行, 左停靠, 主按钮右停靠)
    $leftMax = 0
    foreach($c in ($kids | Where-Object { $_.Cls -eq 'TComboBox' })){
      $rc = Get-RectOf $c.Hwnd
      if($rc.R -gt $leftMax){ $leftMax = $rc.R }
    }
    Log ("  main btn L={0} vs left-most right edge (combos) = {1}" -f $rt.L,$leftMax)
    Chk ($rt.L -gt $leftMax) ('主按钮加宽后没有压住左边的选择框 (按钮左缘 {0} > 下拉右缘 {1})' -f $rt.L,$leftMax)
  } else { Chk $false '主按钮「平铺所选应用(&T)」未找到' }
  Chk (($caps -contains '全选(&A)') -and ($caps -contains '自动刷新')) '顶部复选框: 全选 + 自动刷新 均存在'
  # 置顶控制(2026-09-17 新增): 选定窗口置顶(管勾选的目标窗口) + 本窗口置顶(管 DeskTiler 自己)
  Chk (($caps -contains '选定窗口置顶') -and ($caps -contains '本窗口置顶')) '顶部复选框: 选定窗口置顶 + 本窗口置顶 均存在'

  # 底部快捷按钮行顺序(2026-09-17 调整: “一键全排”放到最右端, “应用快排”紧挨在它左边)
  $actOrder = @('目录快排','PowerShell快排','Cmd快排','应用快排','一键全排')
  $actRect = @{}
  foreach($n in $actOrder){
    $c = $kids | Where-Object { $_.Cap -eq $n } | Select-Object -First 1
    if($c){ $actRect[$n] = Get-RectOf $c.Hwnd }
  }
  $afound = @($actOrder | Where-Object { $actRect.ContainsKey($_) })
  Chk ($afound.Count -eq $actOrder.Count) ('底部 5 个快捷按钮齐全 ({0}/5)' -f $afound.Count)
  if($afound.Count -eq $actOrder.Count){
    foreach($n in $actOrder){ $rr=$actRect[$n]; Log ("  actbtn '{0}': L={1} R={2}" -f $n,$rr.L,$rr.R) }
    $aord = $true
    for($k=1;$k -lt $actOrder.Count;$k++){
      if($actRect[$actOrder[$k]].L -le $actRect[$actOrder[$k-1]].L){ $aord = $false }
    }
    Chk $aord '快捷按钮从左到右 = 目录快排/PowerShell快排/Cmd快排/应用快排/一键全排(一键全排在最右)'
    Chk ($actRect['一键全排'].R -gt $actRect['应用快排'].R) '“一键全排”确实排在“应用快排”右侧(最右端)'
    $aNoOv = $true
    for($k=0;$k -lt $actOrder.Count;$k++){
      for($j=$k+1;$j -lt $actOrder.Count;$j++){
        if(Rects-Overlap $actRect[$actOrder[$k]] $actRect[$actOrder[$j]]){ $aNoOv = $false }
      }
    }
    Chk $aNoOv '5 个快捷按钮两两不重叠'
  }

  # ---- (2) 计数标签 & 联动 ----
  $stat = $kids | Where-Object { $_.Cap -like '窗口总数*' } | Select-Object -First 1
  Chk ($stat -ne $null) '计数标签存在(窗口总数/已选择)'
  $lv = $null
  foreach($k in $kids){ if($k.Cls -eq 'TListView'){ $lv = $k.Hwnd; break } }  # VCL 注册类名为 TListView
  Chk ($lv -ne $null) ('列表控件存在 hwnd=' + $(if($lv){$lv.ToInt64()}else{'0'}))

  if($stat -and $lv){
    $st = Parse-Status (Get-Cap $stat.Hwnd)
    $cnt = [VW]::SendMessageW($lv, 0x1004, [IntPtr]::Zero, [IntPtr]::Zero).ToInt64()  # LVM_GETITEMCOUNT
    Log ("initial status: {0} | listview rows: {1}" -f (Get-Cap $stat.Hwnd), $cnt)
    Chk ($st -ne $null) '计数文本可解析'
    if($st){
      Chk ($st[0] -eq $cnt) ('“窗口总数”=列表行数 ({0}=={1})' -f $st[0],$cnt)
      $total = $st[0]

      $chkAll = $kids | Where-Object { $_.Cap -eq '全选(&A)' } | Select-Object -First 1
      if($chkAll){
        $chk0 = [VW]::SendMessageW($chkAll.Hwnd, 0x00F0, [IntPtr]::Zero, [IntPtr]::Zero).ToInt64()  # BM_GETCHECK
        Chk ($chk0 -eq 0) '初始“全选”复选框未勾选(与已选择=0一致)'

        # 打勾 -> 全部选中, 复选框保持勾选态
        [void][VW]::SendMessageW($chkAll.Hwnd, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)  # BM_CLICK
        Start-Sleep -Milliseconds 500
        $st1 = Parse-Status (Get-Cap $stat.Hwnd)
        $chk1 = [VW]::SendMessageW($chkAll.Hwnd, 0x00F0, [IntPtr]::Zero, [IntPtr]::Zero).ToInt64()
        Log ("after 全选打勾: {0} | BM_GETCHECK={1}" -f (Get-Cap $stat.Hwnd), $chk1)
        Chk ($st1 -and $st1[1] -eq $total -and $chk1 -eq 1) ('全选打勾后 “已选择”=总数且为勾选态 ({0}=={1})' -f $(if($st1){$st1[1]}else{'?'}), $total)

        # 再点 -> 全部取消
        [void][VW]::SendMessageW($chkAll.Hwnd, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 500
        $st2 = Parse-Status (Get-Cap $stat.Hwnd)
        $chk2 = [VW]::SendMessageW($chkAll.Hwnd, 0x00F0, [IntPtr]::Zero, [IntPtr]::Zero).ToInt64()
        Log ("after 全选取消: {0} | BM_GETCHECK={1}" -f (Get-Cap $stat.Hwnd), $chk2)
        Chk ($st2 -and $st2[1] -eq 0 -and $chk2 -eq 0) '再次点击后 “已选择”=0 且复选框未勾选'
      } else { Log 'WARN: 全选复选框未找到' }
    }
  }

  # ---- (2b) 两列布局 & 显示器下拉 ----
  if($lv){
    # 表头列数必须为 2(应用程序+窗口标题, PID 列已去掉): LVM_GETHEADER -> HDM_GETITEMCOUNT
    $hdr = [VW]::SendMessageW($lv, 0x101F, [IntPtr]::Zero, [IntPtr]::Zero)
    $cols = 0
    if($hdr -ne [IntPtr]::Zero){ $cols = [VW]::SendMessageW($hdr, 0x1200, [IntPtr]::Zero, [IntPtr]::Zero).ToInt64() }
    Log ("listview header cols = {0}" -f $cols)
    Chk ($cols -eq 2) '表头列为 2(应用程序+窗口标题, PID 列已去掉)'

    # 下拉框: 列数 + 显示器; 显示器下拉须至少含 “自动选择”(CB_GETCOUNT>=1)
    $combos = @($kids | Where-Object { $_.Cls -eq 'TComboBox' })
    Log ("TComboBox count = {0}" -f $combos.Count)
    Chk ($combos.Count -ge 2) '存在 列数+显示器 两个下拉(TComboBox)'
    $anyItems = $false
    foreach($cb in $combos){
      $n = [VW]::SendMessageW($cb.Hwnd, 0x0146, [IntPtr]::Zero, [IntPtr]::Zero).ToInt64()  # CB_GETCOUNT
      if($n -ge 1){ $anyItems = $true }
    }
    Chk $anyItems '显示器下拉至少含 “自动选择”(CB_GETCOUNT>=1)'
  }

  # ---- (2c) 顶栏左侧控件: 顺序 + 互不重叠(上次的 关于链接 被盖住就是这么暴露的) ----
  $topOrder = @('全选(&A)','刷新(&R)','自动刷新','选定窗口置顶','本窗口置顶')
  $topRect = @{}
  foreach($n in $topOrder){
    $c = $kids | Where-Object { $_.Cap -eq $n } | Select-Object -First 1
    if($c){ $topRect[$n] = Get-RectOf $c.Hwnd }
  }
  $found = @($topOrder | Where-Object { $topRect.ContainsKey($_) })
  Chk ($found.Count -eq $topOrder.Count) ('顶栏左侧 5 个控件齐全 ({0}/5)' -f $found.Count)
  foreach($n in $found){ $rr=$topRect[$n]; Log ("  topbar '{0}': L={1} R={2} T={3} B={4}" -f $n,$rr.L,$rr.R,$rr.T,$rr.B) }
  if($found.Count -eq $topOrder.Count){
    $ordered = $true
    for($k=1;$k -lt $topOrder.Count;$k++){
      if($topRect[$topOrder[$k]].L -le $topRect[$topOrder[$k-1]].L){ $ordered = $false }
    }
    Chk $ordered '顶栏左侧控件从左到右 = 全选/刷新/自动刷新/选定窗口置顶/本窗口置顶(顺序正确)'
    $noOverlap = $true
    for($k=0;$k -lt $topOrder.Count;$k++){
      for($j=$k+1;$j -lt $topOrder.Count;$j++){
        if(Rects-Overlap $topRect[$topOrder[$k]] $topRect[$topOrder[$j]]){ $noOverlap = $false }
      }
    }
    Chk $noOverlap '顶栏左侧控件两两不重叠'
    if($stat){
      $rs = Get-RectOf $stat.Hwnd
      $clash = $false
      foreach($n in $topOrder){ if(Rects-Overlap $topRect[$n] $rs){ $clash = $true } }
      Chk (-not $clash) '计数标签未被任何顶栏左侧控件压住'
      Chk ($rs.R - $rs.L -gt 120) ('计数标签宽度足够显示计数 ({0}px)' -f ($rs.R-$rs.L))
    }
  }

  # ---- (2d) 置顶行为 ----
  $chkTop  = $kids | Where-Object { $_.Cap -eq '选定窗口置顶' } | Select-Object -First 1
  $chkSelf = $kids | Where-Object { $_.Cap -eq '本窗口置顶' } | Select-Object -First 1
  $msg = $kids | Where-Object { $_.Cls -eq 'TStaticText' -and $_.Cap -notlike '窗口总数*' } | Select-Object -First 1
  Log ("prereq: chkTop={0} chkSelf={1} msg={2}" -f $(if($chkTop){'y'}else{'n'}),$(if($chkSelf){'y'}else{'n'}),$(if($msg){'y'}else{'n'}))
  $p2.Refresh(); Log ("[diag] 第二个实例存活 @2d前 = " + (-not $p2.HasExited))
  if($chkTop -and $chkSelf){
    $BM_CLICK   = 0x00F5
    $BM_GETCHK  = 0x00F0
    $zero = [IntPtr]::Zero
    function Click-Chk($h){ [void][VW]::SendMessageW($h,$BM_CLICK,$zero,$zero); Start-Sleep -Milliseconds 450 }
    function Chk-State($h){ return [VW]::SendMessageW($h,$BM_GETCHK,$zero,$zero).ToInt64() }

    Chk ((Chk-State $chkTop.Hwnd) -eq 0 -and (Chk-State $chkSelf.Hwnd) -eq 0) '两个置顶复选框初始均未勾选'
    Chk (-not (EX-TopMost $h1)) '初始: 目标窗口非置顶'

    # (i) 一个都没勾就打开“选定窗口置顶” -> 只提示, 不置顶
    Click-Chk $chkTop.Hwnd
    $m1 = if($msg){ Get-Cap $msg.Hwnd } else { '' }
    Log ("after 窗口置顶 (0 checked): chk={0} msg='{1}' topmost={2}" -f (Chk-State $chkTop.Hwnd), $m1, (EX-TopMost $h1))
    Chk ((Chk-State $chkTop.Hwnd) -eq 1 -and -not (EX-TopMost $h1) -and $m1.Contains('勾选')) '无勾选时打开“选定窗口置顶”: 仅提示, 不误置顶(复选框仍为勾选态)'

    # (ii) “选定窗口置顶”开着时点全选 -> 批勾的窗口也要跟着置顶(OnChkAllClick 补的那一次)
    Click-Chk $chkAll.Hwnd
    $st3 = Parse-Status (Get-Cap $stat.Hwnd)
    Log ("after 全选(ChkTop on): {0} | target topmost = {1}" -f (Get-Cap $stat.Hwnd), (EX-TopMost $h1))
    Chk ($st3 -and $st3[1] -eq $total) '开启“选定窗口置顶”后点全选: 计数仍正确'
    # h1 是被我们点按钮的那个实例, 它**不会把自己列进自己的表**; h2 一定会出现 -> 用 h2 断言
    Chk ((EX-TopMost $h2) -and -not (EX-TopMost $h1)) '批勾的窗口立刻被置顶(h2 置位; h1 是自己, 不在自己的表里)'
    if($msg){ Chk ((Get-Cap $msg.Hwnd).StartsWith('已把')) ('提示文本: ' + (Get-Cap $msg.Hwnd)) }

    # (iii) 关掉“选定窗口置顶” -> 消失
    Click-Chk $chkTop.Hwnd
    Log ("after 取消窗口置顶: h2 topmost = {0}" -f (EX-TopMost $h2))
    Chk (-not (EX-TopMost $h2)) '取消“选定窗口置顶”后 WS_EX_TOPMOST 被清除'

    Click-Chk $chkAll.Hwnd   # 全不选, 恢复原状
    $st4 = Parse-Status (Get-Cap $stat.Hwnd)
    Chk ($st4 -and $st4[1] -eq 0) '收尾: 取消全选后 已选择=0'

    # (iv) “本窗口置顶”只影响 DeskTiler 自己
    Click-Chk $chkSelf.Hwnd
    Log ("after 本窗口置顶: self topmost = {0}" -f (EX-TopMost $h1))
    Chk (EX-TopMost $h1) '打开“本窗口置顶”: 主窗体自身 WS_EX_TOPMOST 置位'
    Click-Chk $chkSelf.Hwnd
    Log ("after 取消本窗口置顶: self topmost = {0}" -f (EX-TopMost $h1))
    Chk (-not (EX-TopMost $h1)) '取消“本窗口置顶”: 主窗体恢复普通层级'
  }

  $p2.Refresh(); Log ("[diag] 第二个实例存活 @2d后 = " + (-not $p2.HasExited))
  # ---- (2e) 在列表里点选一行 -> 「应用快排」按钮改名为「应用名+快排」并按文字加宽(2026-09-17 用户要求) ----
  if($lv){
    $btnApp = $kids | Where-Object { $_.Cap -eq '应用快排' } | Select-Object -First 1
    if($btnApp){
      $bh  = $btnApp.Hwnd
      $rb0 = Get-RectOf $bh
      $wb0 = $rb0.R - $rb0.L
      Log ("appbtn before: cap='{0}' W={1} L={2}" -f (Get-Cap $bh), $wb0, $rb0.L)

      # 往列表第 1 行上真按一下: 客户端坐标 x=60 落在第 0 列(应用程序, 宽 130)的文字区,
      # 避开最左边那几个像素的勾选框 —— 点文字只选中、不会顺带改勾选状态。
      # y 不写死(行高随 DPI 变): 从表头下缘起逐个候选值试, 试到按钮改名就停。
      $clicked = $false
      foreach($y in 30,26,34,22,38,42,46,50){
        $lp = [IntPtr]((([int]$y) -shl 16) -bor 60)
        [void][VW]::PostMessageW($lv, 0x0201, [IntPtr]1, $lp)          # WM_LBUTTONDOWN (MK_LBUTTON)
        [void][VW]::PostMessageW($lv, 0x0202, [IntPtr]::Zero, $lp)     # WM_LBUTTONUP
        Start-Sleep -Milliseconds 350
        if((Get-Cap $bh) -ne '应用快排'){ $clicked = $true; break }
      }

      $capA = Get-Cap $bh
      $rb1  = Get-RectOf $bh
      $wb1  = $rb1.R - $rb1.L
      Log ("appbtn after : cap='{0}' W={1} R={2} (row1 click hit = {3})" -f $capA, $wb1, $rb1.R, $clicked)
      Chk $clicked ('点选列表一行后「应用快排」改名成「应用名+快排」: ' + $capA)

      if($clicked){
        # 名字必须真是那一行的应用名, 不能是随便串上去的 —— 拿系统里正在跑的进程名对一遍
        $app = $capA -replace '快排$',''
        $known = @(Get-Process -ErrorAction SilentlyContinue | ForEach-Object { $_.ProcessName })
        Chk ($known -contains $app) ('按钮上的应用名确实来自运行中的进程: ' + $app)
        Chk ($wb1 -ge $wb0) ('按钮没有因为改名变窄 ({0}px -> {1}px)' -f $wb0, $wb1)
        # 真正的需求是“文字显示得完整”: 按窗体同一字体量一遍文字像素宽, 按钮必须比它宽
        $need = Measure-Text $capA
        Log ("  text '{0}' needs {1}px, button is {2}px" -f $capA, $need, $wb1)
        Chk ($need -gt 0 -and $wb1 -ge $need) ('按钮够宽, 文字能完整显示 ({0}px 文字 <= {1}px 按钮)' -f $need, $wb1)

        # 宽度一变, 底行就得重排(五个按钮是从右往左贴着摆的): 顺序与不重叠都要重新成立。
        # 注意 $kids 是改名**之前**抓的快照, 里面那个控件的 Cap 还写着「应用快排」——
        # 按老名字再查一次会把同一个按钮当成第二个控件收进来, 所以这个位置直接用改名后的 rect。
        $ar = @{}
        foreach($n in $actOrder){
          if($n -eq '应用快排'){ $ar[$capA] = $rb1; continue }
          $c = $kids | Where-Object { $_.Cap -eq $n } | Select-Object -First 1
          if($c){ $ar[$n] = Get-RectOf $c.Hwnd }
        }
        $names = @($actOrder | ForEach-Object { if($_ -eq '应用快排'){ $capA } else { $_ } })
        $ord2 = $true
        for($k=1;$k -lt $names.Count;$k++){
          if($ar[$names[$k]].L -le $ar[$names[$k-1]].L){ $ord2 = $false }
        }
        $noOv2 = $true
        for($k=0;$k -lt $names.Count;$k++){
          for($j=$k+1;$j -lt $names.Count;$j++){
            if(Rects-Overlap $ar[$names[$k]] $ar[$names[$j]]){ $noOv2 = $false }
          }
        }
        foreach($n in $names){ $rr=$ar[$n]; Log ("  actbtn(after rename) '{0}': L={1} R={2}" -f $n,$rr.L,$rr.R) }
        Log ("  reorder check: order={0} noOverlap={1} buttons={2}/5" -f $ord2,$noOv2,$ar.Count)
        Chk ($ord2 -and $noOv2 -and $ar.Count -eq 5) '加宽后底行仍是同样顺序且两两不重叠(重排生效)'
      }
    } else { Chk $false '「应用快排」按钮未找到(无法验证改名)' }
  }
}

# ============ (3) 表头排序自检(SortList 升降序, 程序内执行, 无跨进程读内存) ============
$p2.Refresh(); Log ("[diag] 第二个实例存活 @2e后 = " + (-not $p2.HasExited))
Log '== run -sorttest =='
Start-Run "-sorttest `"$sortFile`""
if(Test-Path $sortFile){
  $lines = Get-Content $sortFile
  function Get-Block($name){
    $idx = -1
    for($j=0;$j -lt $lines.Count;$j++){ if($lines[$j] -eq $name){ $idx=$j; break } }
    if($idx -lt 0){ return $null }
    $out = New-Object System.Collections.Generic.List[string]
    for($j=$idx+1;$j -lt $lines.Count;$j++){
      if($lines[$j].StartsWith('--')){ break }
      $out.Add($lines[$j])
    }
    return $out
  }
  $z  = Get-Block '-- z-order --'
  $pa = Get-Block '-- pid asc --'
  $pd = Get-Block '-- pid desc --'
  $aa = Get-Block '-- app asc --'
  $ad = Get-Block '-- app desc --'
  Log ("sorttest blocks: z={0} pa={1} pd={2} aa={3} ad={4}" -f $(if($z){$z.Count}else{0}),$(if($pa){$pa.Count}else{0}),$(if($pd){$pd.Count}else{0}),$(if($aa){$aa.Count}else{0}),$(if($ad){$ad.Count}else{0}))
  Chk (($z -and $z.Count -ge 2 -and $pa -and $pd -and $aa -and $ad) -and ($z.Count -eq $pa.Count)) '枚举到多窗口且各区块行数一致'
  if($pa -and $pd){
    $pidsA = @($pa | ForEach-Object { [int](($_ -split '\|')[0]) })
    $pidsD = @($pd | ForEach-Object { [int](($_ -split '\|')[0]) })
    $revA = @($pidsA); [array]::Reverse($revA)
    $okNum = $true
    for($k=1;$k -lt $pidsA.Count;$k++){ if($pidsA[$k] -lt $pidsA[$k-1]){ $okNum=$false; break } }
    Chk $okNum 'PID 升序: 非递减'
    Chk (Same-Seq $pidsD $revA) 'PID 再次点击: 降序 = 升序的反序'
  }
  if($aa -and $ad -and $aa.Count -ge 2){
    # 应用名列允许同名并列(TList.Sort 不稳定), 故用“按区域设置单调”而非“精确反序”校验
    $appA = @($aa | ForEach-Object { ($_ -split '\|')[1] })
    $appD = @($ad | ForEach-Object { ($_ -split '\|')[1] })
    $zapp = @($z  | ForEach-Object { ($_ -split '\|')[1] })
    $okAsc=$true
    for($k=1;$k -lt $appA.Count;$k++){ if((CmpStr $appA[$k-1] $appA[$k]) -eq 3){ $okAsc=$false; break } }
    Chk $okAsc '应用名 升序: 按区域设置非递减(文本排序生效)'
    $okDesc=$true
    for($k=1;$k -lt $appD.Count;$k++){ if((CmpStr $appD[$k-1] $appD[$k]) -eq 1){ $okDesc=$false; break } }
    Chk $okDesc '应用名 降序: 按区域设置非递增(再点一次切换方向)'
    Chk (-not (Same-Seq $appA $zapp)) '应用名升序相对 Z 序已变化(确实发生了排序)'
  }
} else { Chk $false 'sorttest 未生成文件' }

# ============ (4) -tileapp 平铺两个 DeskTiler 目标 -> 均分网格 ============
$p2.Refresh(); Log ("[diag] 第二个实例存活 @sorttest后 = " + (-not $p2.HasExited))
Log '== run -tileapp DeskTiler =='
Start-Run "-tileapp DeskTiler"
Start-Sleep -Milliseconds 600
$r1 = New-Object VERECT; [void][VW]::GetWindowRect($h1,[ref]$r1)
$r2 = New-Object VERECT; [void][VW]::GetWindowRect($h2,[ref]$r2)
$w1=$r1.R-$r1.L; $hgt1=$r1.B-$r1.T
$w2=$r2.R-$r2.L; $hgt2=$r2.B-$r2.T
Log ("tiled A: ({0},{1}) {2}x{3}   B: ({4},{5}) {6}x{7}" -f $r1.L,$r1.T,$w1,$hgt1,$r2.L,$r2.T,$w2,$hgt2)

$sameSize = ($w1 -eq $w2 -and $hgt1 -eq $hgt2)
$overlap  = ($r1.L -lt $r2.R -and $r2.L -lt $r1.R -and $r1.T -lt $r2.B -and $r2.T -lt $r1.B)
Chk ($sameSize -and -not $overlap) '两窗口等大且不重叠(成功平铺为均分网格)'
Chk ($w1 -gt 100 -and $hgt1 -gt 100) '平铺尺寸正常(非 0/极小)'
# 相邻: 同行左右相邻 或 同列上下相邻
$ra = [pscustomobject]@{L=$r1.L;T=$r1.T;R=$r1.R;B=$r1.B}
$rb = [pscustomobject]@{L=$r2.L;T=$r2.T;R=$r2.R;B=$r2.B}
$r0,$r1b = if($ra.L -le $rb.L){ $ra,$rb } else { $rb,$ra }
$adj = $false
if($r0.T -eq $r1b.T -and $r0.B -eq $r1b.B){ $gap=$r1b.L-$r0.R; if($gap -ge 0 -and $gap -le 24){ $adj=$true }; Log ("gap horizontal = {0}" -f $gap) }
elseif($r0.L -eq $r1b.L -and $r0.R -eq $r1b.R){ $gap=$r1b.T-$r0.B; if($gap -ge 0 -and $gap -le 24){ $adj=$true }; Log ("gap vertical = {0}" -f $gap) }
else { Log 'orientation neither pure-row nor pure-col' }
Chk $adj '两窗口相邻成 1×2 或 2×1 网格'

# ============ (2f) 列表右键菜单: 「结束进程」/「转到应用」 ============
# 放在 (4) 之后: 这一段会把第二个实例真的杀掉, 不能影响前面还要用到它的断言。
# 被杀的都是本脚本自己起的实例($p2), 不碰机器上任何别人的进程。
Log '== (2f) list context menu =='
$script:hMain = $h1
$script:hLv   = $lv

# 逐行点选, 看「应用快排」按钮是否变成「DeskTiler快排」—— 那就是第二个实例所在的行。
# (列表里读不到文字, 但按钮改名是可以跨进程读的; 驱动实例不把自己列进自己的表, 所以只有一行是它。)
function Find-VictimRow($btnTile,$lv){
  $nItems = [int64][VW]::SendMessageW($lv,0x1004,[IntPtr]::Zero,[IntPtr]::Zero)
  # 扫**全部**行: 原来是写死的 12 行, 机器上窗口多一条(13 行)时靶子正好落在最后一行就找不到
  # (投递点击不需要那一行可见, 行距 20px 对任何行号都成立)
  for($k=0; $k -lt [Math]::Min($nItems,40); $k++){
    $rowY = -1
    foreach($cand in @(($k*20+32),($k*20+27),($k*20+37))){
      $lp = [IntPtr]((([int]$cand) -shl 16) -bor 60)
      [void][VW]::PostMessageW($lv,0x0201,[IntPtr]1,$lp)          # WM_LBUTTONDOWN
      [void][VW]::PostMessageW($lv,0x0202,[IntPtr]::Zero,$lp)     # WM_LBUTTONUP
      Start-Sleep -Milliseconds 220
      if(([int64][VW]::SendMessageW($lv,0x100C,[IntPtr](-1),[IntPtr]2)) -eq $k){ $rowY=$cand; break }  # LVM_GETNEXTITEM/LVNI_SELECTED
    }
    if($rowY -lt 0){ continue }
    $rc = Get-Cap $btnTile.Hwnd
    $null = Log ("  row {0} (client y={1}) -> '{2}'" -f $k,$rowY,$rc)
    if($rc -eq 'DeskTiler快排'){ return $rowY }
  }
  return -1
}

# 投递一下左键, 反查应用眼里的这一行是谁 —— 「应用快排」按钮会改名为「<应用名>快排」。
# 投递的左键走的是客户端坐标, 窗口挪到哪儿都不影响。
function Get-RowNameAt($btnTile,$lv,$clientY){
  $lp = [IntPtr]((([int]$clientY) -shl 16) -bor 60)
  [void][VW]::PostMessageW($lv,0x0201,[IntPtr]1,$lp)
  [void][VW]::PostMessageW($lv,0x0202,[IntPtr]::Zero,$lp)
  Start-Sleep -Milliseconds 220
  return (Get-Cap $btnTile.Hwnd)
}
function Test-VictimRow($btnTile,$lv,$clientY){
  return ((Get-RowNameAt $btnTile $lv $clientY) -eq 'DeskTiler快排')
}

# 找行找到的是行的**上边缘**, 而那就是行边界: 真点击只要偏下半个像素就落到上一行。
# 所以再往行中间挪, 每挪一个位置都投递点击复核一次, 挑出第一个仍然属于靶子行的 y。
function Pick-VictimPoint($btnTile,$lv,$rowTop){
  foreach($off in @(10,8,6,12,0)){
    $cy = $rowTop + $off
    if((Get-RowNameAt $btnTile $lv $cy) -eq 'DeskTiler快排'){
      $null = Log ("  落点 client y={0} (行顶 {1} + {2}) 复核通过" -f $cy,$rowTop,$off)
      return $cy
    }
  }
  return -1
}

# 找行 -> 挑落点复核 -> 右键。三段都成了才返回菜单句柄(0 = 没成)
function Open-VictimMenu($btnTile,$lv){
  for($attempt=1; $attempt -le 2; $attempt++){
    $rowTop = Find-VictimRow $btnTile $lv
    if($rowTop -lt 0){ $null = Log ("  第 {0} 次找行: 列表里没看到靶子那一行" -f $attempt); continue }
    $cy = Pick-VictimPoint $btnTile $lv $rowTop
    if($cy -lt 0){ $null = Log ("  第 {0} 次找行: 行顶 {1} 附近复核都不是靶子行(列表动过), 重来" -f $attempt,$rowTop); continue }
    $m = Open-RowMenu $cy
    if($m -ne [IntPtr]::Zero){ return $m }
    $null = Log ("  第 {0} 次找行: 行是对的, 但菜单没弹出来" -f $attempt)
  }
  return [IntPtr]::Zero
}

# 这一段自己起靶子, 不用前面那个已经跑了近百秒的实例 —— 否则本机别的风吹草动(实例被关掉)
# 会让右键菜单的结论跟着一起垮。先收掉旧的, 保证列表里只剩一个新的 DeskTiler 行, 找行才不含糊。
#
# 找行是按「应用名 = DeskTiler」认的(跨进程读不到 PID 列, 那一列早没了), 所以列表里**同名的行
# 只能有一条**。踩过: 机器上另有 DeskTiler 实例时, 这段右键点到了人家那一行, 确认框点名的是
# PID 13168 而不是靶子 26916, 后面三次「转到应用」也一路把前台切到那个别人的窗口上。
# 驱动自己以外的 DeskTiler 一律收掉 —— 这工具没有任何未保存状态, 开头那一段也是这么干的。
Get-Process DeskTiler -ErrorAction SilentlyContinue | Where-Object { $_.Id -ne $p1.Id } | ForEach-Object {
  $null = Log ("  收掉多余的 DeskTiler 实例 pid={0}" -f $_.Id)
  try { Stop-Process -Id $_.Id -Force -ErrorAction Stop } catch { $null = Log ("    收不掉: " + $_.Exception.Message) }
}
Start-Sleep -Milliseconds 900
$pv = [System.Diagnostics.Process]::Start($exe)
$hv = Wait-Hwnd $pv
Chk ($hv -ne [IntPtr]::Zero) '右键菜单: 为这一段新起的目标实例已就绪'
Log ("  fresh target pid={0} hwnd={1}" -f $pv.Id,([int64]$hv))

# 列表是“关掉自动刷新”之后冻住的快照, 得点一次「刷新」才收得进新实例
$ref2 = $kids | Where-Object { $_.Cap -eq '刷新(&R)' } | Select-Object -First 1
if($ref2){
  [void][VW]::SendMessageW($ref2.Hwnd,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)   # BM_CLICK
  Start-Sleep -Milliseconds 1200
}

$pt0 = New-Object VEPT
$cursorSaved = [VW]::GetCursorPos([ref]$pt0)
$btnTile = $kids | Where-Object { $_.Cls -eq 'TButton' -and $_.Cap -like '*快排' -and (@('目录快排','PowerShell快排','Cmd快排') -notcontains $_.Cap) } | Select-Object -First 1
$lvr = Get-RectOf $lv
Log ("  lv rect={0},{1},{2},{3}; target pid={4} hwnd={5}" -f $lvr.L,$lvr.T,$lvr.R,$lvr.B,$pv.Id,([int64]$hv))

$victimY = Find-VictimRow $btnTile $lv
Chk ($victimY -ge 0) ('在列表里定位到目标实例所在的行(client y=' + $victimY + ')')

if($victimY -ge 0){

  # (a) 右键那一行 -> 弹出原生菜单
  $m = Open-VictimMenu $btnTile $lv
  Chk ($m -ne [IntPtr]::Zero) '右键列表某行: 弹出原生右键菜单(#32768)'
  if($m -ne [IntPtr]::Zero){
    $mr = Get-RectOf $m
    $hm = [VW]::SendMessageW($m,0x01E1,[IntPtr]::Zero,[IntPtr]::Zero)          # MN_GETHMENU
    $mi = -1
    if($hm -ne [IntPtr]::Zero){ $mi = [VW]::GetMenuItemCount($hm) }
    Log ("  menu: {0}x{1}px at ({2},{3}) items={4}" -f ($mr.R-$mr.L),($mr.B-$mr.T),$mr.L,$mr.T,$mi)
    Chk ($mi -eq 2) '右键菜单正好两项(结束进程 / 转到应用)'
    # 菜单已弹出、还没点任何一项: 记一下此刻列表的矩形与行数, 万一以后又不一致, 好拿它对表。
    # (注意: 这里只能读, 不能往列表投递点击 —— 那种消息会把菜单顶掉, 后面那一下点击就落空了。)
    $lvn = Get-ClientOrigin $lv
    $null = Log ("  [diag] menu open: lv client origin={0},{1}; items={2}" -f `
      $lvn[0],$lvn[1],([int64][VW]::SendMessageW($lv,0x1004,[IntPtr]::Zero,[IntPtr]::Zero)))

    # (b) 第一项 = 结束进程 -> 原生确认框(正文有句柄, 跨进程读得到)
    [void](Click-MenuItem $m 0.25)
    Start-Sleep -Milliseconds 300
    $dlg = Find-TopWindow '#32770'
    $dtit = ''; $dtxt = ''
    if($dlg -ne [IntPtr]::Zero){ $dtit = Get-Cap $dlg; $dtxt = Get-DialogText $dlg }
    Log ("  dialog: title='{0}' text='{1}'" -f $dtit,($dtxt -replace "`r`n",' // '))
    Chk ($dlg -ne [IntPtr]::Zero -and $dtit -eq '结束进程') '「结束进程」弹出确认框(标题=结束进程)'
    Chk ($dtxt.Contains('DeskTiler') -and $dtxt.Contains('PID ' + $pv.Id)) '确认框点名了要结束的应用与 PID'

    # (c) 点「否」 -> 对话框关掉, 进程一个都不能少
    if($dlg -ne [IntPtr]::Zero){
      [void][VW]::SendMessageW([VW]::GetDlgItem($dlg,7),0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)   # IDNO / BM_CLICK
      Start-Sleep -Milliseconds 800
      $pv.Refresh()
      $gone = ((Find-TopWindow '#32770') -eq [IntPtr]::Zero)
      Log ("  after NO: dialog closed={0} target alive={1}" -f $gone,(-not $pv.HasExited))
      Chk ($gone -and -not $pv.HasExited) '点「否」: 对话框关闭且目标进程仍存活(没有误杀)'
    }

    # (d) 第二项 = 转到应用 -> 该行窗口被激活到前台
    # 这一下是真光标 + 真鼠标事件, 而且要把前台抢过来。机器上如果**有人同时在操作**,
    # Windows 的前台锁会把这次 SetForegroundWindow 顶掉(实测偶发, 同一份程序有时红有时绿)。
    # 所以动作重试几次, 判定仍然只做一次 —— 真是程序没激活的话, 重试多少次都过不了。
    $fg = [IntPtr]::Zero
    $tried = 0
    for($try = 1; $try -le 3; $try++){
      $m2 = Open-VictimMenu $btnTile $lv
      if($m2 -eq [IntPtr]::Zero){ $null = Log ("  第 {0} 次: 右键没能弹出菜单" -f $try); continue }
      $tried++
      [void](Click-MenuItem $m2 0.75)
      Start-Sleep -Milliseconds 700
      $fg = [VW]::GetForegroundWindow()
      $null = Log ("  第 {0} 次「转到应用」: fg={1} '{2}' (target {3} 最小化={4}; 工具最小化={5})" -f `
        $try,([int64]$fg),(Get-Cap $fg),([int64]$hv),[VW]::IsIconic($hv),[VW]::IsIconic($script:hMain))
      if([int64]$fg -eq [int64]$hv){ break }
    }
    if($tried -gt 0){ Chk ([int64]$fg -eq [int64]$hv) '「转到应用」把目标窗口激活到前台' }
    else { Chk $false '再次右键未能弹出菜单(无法验证「转到应用」)' }

    # (e) 再走一遍结束进程, 这次点「是」 -> 进程真的没了
    $m3 = Open-VictimMenu $btnTile $lv
    if($m3 -ne [IntPtr]::Zero){
      [void](Click-MenuItem $m3 0.25)
      Start-Sleep -Milliseconds 300
      $dlg2 = Find-TopWindow '#32770'
      $t2 = ''
      if($dlg2 -ne [IntPtr]::Zero){ $t2 = Get-DialogText $dlg2 }
      Log ("  kill dialog: '" + ($t2 -replace "`r`n",' // ') + "'")
      # 安全闸: 只有确认框点名的就是我们自己起的那个进程, 才真去按「是」
      if($dlg2 -ne [IntPtr]::Zero -and $t2.Contains('PID ' + $pv.Id)){
        [void][VW]::SendMessageW([VW]::GetDlgItem($dlg2,6),0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)  # IDYES
        $ended = $pv.WaitForExit(5000)
        Log ("  after YES: target exited={0}" -f $ended)
        Chk $ended '点「是」: 目标进程确实被结束'
      } else {
        if($dlg2 -ne [IntPtr]::Zero){ [void][VW]::SendMessageW([VW]::GetDlgItem($dlg2,7),0x00F5,[IntPtr]::Zero,[IntPtr]::Zero) }
        Chk $false ('确认框点名的不是目标进程, 已按「否」放弃: ' + $t2)
      }
    } else { Chk $false '第三次右键未能弹出菜单(无法验证真正结束进程)' }
  }
}
if($cursorSaved){ [void][VW]::SetCursorPos($pt0.X,$pt0.Y) }   # 真光标挪过, 还回去

# ============ (2g) 勾着「本窗口置顶」时「转到应用」必须真的看得见 ============
# 用户报的问题: 勾上「本窗口置顶」后点「转到应用」什么也看不到 —— 置顶窗口永远浮在普通窗口
# 之上, 目标窗口就算被激活也只能躲在工具窗口后面(实测连前台都没换来)。现在的做法: 两者矩形
# 相交时先把工具窗口最小化让开再激活目标; 错开摆放(不重叠)时不让开, 不打扰“让工具一直浮着”。
# 仍然只动本脚本自己起的实例。
Log '== (2g) 转到应用 @ 本窗口置顶 =='
$SWP_NA = 0x0010 -bor 0x0040            # SWP_NOACTIVATE | SWP_SHOWWINDOW
$pw = [System.Diagnostics.Process]::Start($exe)
$hw = Wait-Hwnd $pw
Chk ($hw -ne [IntPtr]::Zero) '(2g) 为置顶场景新起的目标实例已就绪'

if($hw -ne [IntPtr]::Zero -and $ref2 -and $chkSelf){
  $null = Log ("  pid={0} hwnd={1}" -f $pw.Id,([int64]$hw))
  # 列表是冻住的快照, 点一次「刷新」把新实例收进来
  [void][VW]::SendMessageW($ref2.Hwnd,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)
  Start-Sleep -Milliseconds 1200

  $dr0 = Get-RectOf $script:hMain
  # 靶子摆里侧、工具窗口摆外侧: 让置顶的工具窗口正好压住靶子中心
  [void][VW]::SetWindowPos($hw,[IntPtr]::Zero,200,200,700,600,$SWP_NA)
  Start-Sleep -Milliseconds 400
  [void][VW]::SetWindowPos($script:hMain,[IntPtr]::Zero,60,60,1000,900,$SWP_NA)
  Start-Sleep -Milliseconds 500
  $nc = Get-RectOf $hw
  $tcx = [int](($nc.L+$nc.R)/2); $tcy = [int](($nc.T+$nc.B)/2)

  [void][VW]::SendMessageW($chkSelf.Hwnd,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)   # 勾上「本窗口置顶」
  Start-Sleep -Milliseconds 600
  Chk (EX-TopMost $script:hMain) '(2g) 「本窗口置顶」已生效(工具窗口 WS_EX_TOPMOST)'

  # 前置核验: 量“置顶的工具窗口是不是真的压住了靶子中心”。
  # 有两件事会让它不成立, 但都不是被测程序的错: ①工具窗口被谁(用户)最小化了 —— 上面那次
  # SetWindowPos 带的是 SWP_SHOWWINDOW, 它**不会**把已最小化的窗口还原, 窗口就一直缩在任务栏上;
  # ②有人同时在点别的窗口, 前台/Z 序当场被搅动。所以先还原再重摆, 最多试 3 次, 判定只做一次。
  $coverBefore = [IntPtr]::Zero
  for($try = 1; $try -le 3; $try++){
    if([VW]::IsIconic($script:hMain)){ [void][VW]::ShowWindow($script:hMain,9); Start-Sleep -Milliseconds 400 }
    [void][VW]::SetWindowPos($script:hMain,[IntPtr]::Zero,60,60,1000,900,$SWP_NA)
    Start-Sleep -Milliseconds 500
    $coverBefore = Root-Of ([VW]::WindowFromPoint((To-Pt $tcx $tcy)))
    $rcDr = Get-RectOf $script:hMain
    $null = Log ("  第 {0} 次前置: 靶子中心 ({1},{2}) 上盖着 {3} '{4}' (工具={5} 靶子={6}); 工具矩形={7},{8},{9},{10} 最小化={11} 置顶={12} 靶子最小化={13}" `
      -f $try,$tcx,$tcy,([int64]$coverBefore),(Get-Cap $coverBefore),([int64]$script:hMain),([int64]$hw),`
      $rcDr.L,$rcDr.T,$rcDr.R,$rcDr.B,[VW]::IsIconic($script:hMain),(EX-TopMost $script:hMain),[VW]::IsIconic($hw))
    if($coverBefore -eq $script:hMain){ break }
  }
  Chk ($coverBefore -eq $script:hMain) '(2g) 前置: 置顶的工具窗口确实压住了靶子中心(否则这条用例没意义)'

  $mg = Open-VictimMenu $btnTile $lv
  if($mg -ne [IntPtr]::Zero){
    [void](Click-MenuItem $mg 0.75)                                             # 第二项 = 转到应用
    Start-Sleep -Milliseconds 900
    $fg = [VW]::GetForegroundWindow()
    $coverAfter = Root-Of ([VW]::WindowFromPoint((To-Pt $tcx $tcy)))
    $null = Log ("  点「转到应用」后: fg={0} 工具最小化={1} 靶子最小化={2} 靶子中心上盖着={3}" -f `
      ([int64]$fg),[VW]::IsIconic($script:hMain),[VW]::IsIconic($hw),([int64]$coverAfter))
    Chk ([VW]::IsIconic($script:hMain)) '(2g) 重叠时工具窗口最小化让开'
    Chk ($coverAfter -eq $hw) '(2g) 靶子中心真的露出来了(不再被工具窗口压着)'
    Chk ([int64]$fg -eq [int64]$hw) '(2g) 靶子窗口成为前台窗口'
  } else { Chk $false '(2g) 右键没能弹出菜单, 无法验证“重叠 + 置顶”场景' }

  # 场景二: 置顶但两者错开摆放 -> 不该让开(用户勾置顶就是想让它一直浮着)
  [void][VW]::ShowWindow($script:hMain,9)                                       # SW_RESTORE
  Start-Sleep -Milliseconds 500
  [void][VW]::SetWindowPos($script:hMain,[IntPtr]::Zero,700,450,900,600,$SWP_NA)
  [void][VW]::SetWindowPos($hw,[IntPtr]::Zero,40,40,420,300,$SWP_NA)
  Start-Sleep -Milliseconds 600
  $nc2 = Get-RectOf $hw
  $tcx2 = [int](($nc2.L+$nc2.R)/2); $tcy2 = [int](($nc2.T+$nc2.B)/2)
  $cover2 = Root-Of ([VW]::WindowFromPoint((To-Pt $tcx2 $tcy2)))
  $null = Log ("  场景二前置: 靶子中心 ({0},{1}) 上盖着 {2} (工具={3})" -f $tcx2,$tcy2,([int64]$cover2),([int64]$script:hMain))
  Chk ($cover2 -eq $hw) '(2g) 前置: 错开摆放后工具窗口确实不压靶子'

  # 同样重试几次再判定: 抢前台会被别人的操作顶掉, 理由同 (2f)(d)
  $fg2 = [IntPtr]::Zero
  $tried2 = 0
  for($try = 1; $try -le 3; $try++){
    $mg2 = Open-VictimMenu $btnTile $lv
    if($mg2 -eq [IntPtr]::Zero){ $null = Log ("  第 {0} 次: 右键没能弹出菜单" -f $try); continue }
    $tried2++
    [void](Click-MenuItem $mg2 0.75)
    Start-Sleep -Milliseconds 900
    $fg2 = [VW]::GetForegroundWindow()
    $null = Log ("  第 {0} 次「转到应用」(不重叠): fg={1} '{2}' (target {3}) 工具最小化={4}" -f `
      $try,([int64]$fg2),(Get-Cap $fg2),([int64]$hw),[VW]::IsIconic($script:hMain))
    if([int64]$fg2 -eq [int64]$hw){ break }
  }
  if($tried2 -gt 0){
    Chk (-not ([VW]::IsIconic($script:hMain))) '(2g) 不重叠时不让开(工具窗口没被最小化)'
    Chk ([int64]$fg2 -eq [int64]$hw) '(2g) 不重叠时靶子窗口同样成为前台'
  } else { Chk $false '(2g) 右键没能弹出菜单, 无法验证“不重叠 + 置顶”场景' }

  # 还原: 取消置顶 + 工具窗口回到原来的位置尺寸
  [void][VW]::SendMessageW($chkSelf.Hwnd,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)
  Start-Sleep -Milliseconds 400
  [void][VW]::SetWindowPos($script:hMain,[IntPtr]::Zero,$dr0.L,$dr0.T,($dr0.R-$dr0.L),($dr0.B-$dr0.T),$SWP_NA)
  Start-Sleep -Milliseconds 300
  Log ("  还原: 置顶={0} 工具最小化={1}" -f (EX-TopMost $script:hMain),[VW]::IsIconic($script:hMain))
}
if($pw){ try { Stop-Process -Id $pw.Id -Force -ErrorAction SilentlyContinue } catch {} }

# ============ (2h) 右键某一行: 选中的必须正好是那一行 ============
# 用户报的问题: “当右键点击列表项某一行, 总是莫名其妙其他的行被选中”。两个原因叠在一起:
#   1) 坐标慢一拍 —— 列表视图收到 WM_RBUTTONDOWN 时会先同步发出 NM_RCLICK, VCL 当场把菜单
#      弹起来并进入菜单自己的模态循环, 等菜单关掉才轮到 Lv.OnMouseDown; 于是 OnPopup 读到的
#      永远是**上一次**右键的坐标, 菜单老作用在上一行。现在坐标改在消息派发前接住
#      (Application.OnMessage), 实测 4 次右键 4 次对得上。
#   2) 多选累加 —— Lv.MultiSelect=True 时 it.Selected:=True 是加选不是改选, 右键一圈下来好几行
#      都亮着; 已改成单选(平铺选哪些窗口本来就是靠勾选框的)。
# 两把尺子一起量: 「选中的行集合正好是右键那一行」, 以及「应用快排」按钮改成的那行应用名
# 就是被右键那一行(那是程序自己认定的目标行)。这两条都在**列表冻着**(自动刷新关)时做,
# 行号才算数; 之后再打开自动刷新跨过一轮整表重建, 量“选中的还是不是同一个窗口”。
Log '== (2h) 右键某一行 =='
function As-Idx($v){ $x = $v.ToInt64(); if($x -gt 0x7FFFFFFF){ return $x - 0x100000000 }; return $x }
# 跨进程 SendMessage 的返回值是零扩展的 32 位: “没有” -1 回来是 4294967295, 得折回去
function Sel-Set($lv){
  $res = @()
  # 句柄没了的话 SendMessage 一律返回 0, 下面这个循环会一路加到上限, 打出来是 {0,0,0,...} ——
  # 看着像“选中了几十行”, 其实是窗口没了。这种时候直接给个不可能当行号的哨兵值, 好认。
  if(-not [VW]::IsWindow($lv)){ return ,@(-99) }
  $i = As-Idx ([VW]::SendMessageW($lv,0x100C,[IntPtr](-1),[IntPtr]2))     # LVM_GETNEXTITEM / LVNI_SELECTED
  while($i -ge 0 -and $res.Count -lt 40){
    $res += [int]$i
    $i = As-Idx ([VW]::SendMessageW($lv,0x100C,[IntPtr]$i,[IntPtr]2))
  }
  return ,$res
}
function Set-Text($set){ if($set.Count -eq 0){ return '{}' }; return '{' + ($set -join ',') + '}' }

$pt0 = New-Object VEPT
$cursorSaved2 = [VW]::GetCursorPos([ref]$pt0)
$dr1 = Get-RectOf $script:hMain
# 把窗口拉高一点, 保证要点的几行都在可视区里(窗口怎么挪都不影响投递点击 —— 那是客户区坐标)
[void][VW]::SetWindowPos($script:hMain,[IntPtr]::Zero,$dr1.L,$dr1.T,($dr1.R-$dr1.L),[Math]::Max(($dr1.B-$dr1.T),900),$SWP_NA)
Start-Sleep -Milliseconds 700

# 这一节要连着真点好几行, 而用户可能正开着别的窗口压在工具窗口上(实测有一次整屏被 Chrome 盖住,
# 于是哪一行都点不到)。临时把自己抬到最顶层只为“点得到行” —— 量的本来就不是 Z 序;
# 前面 (2g) 才是量置顶的那一节, 所以这里完事要还原。
$drTop = EX-TopMost $script:hMain
$SWP_NMNA = 0x0001 -bor 0x0002 -bor 0x0010                       # NOSIZE | NOMOVE | NOACTIVATE
[void][VW]::SetWindowPos($script:hMain,[IntPtr](-1),0,0,0,0,$SWP_NMNA)     # HWND_TOPMOST
Start-Sleep -Milliseconds 400
Log ("  工具窗口临时抬到最顶层(原来 topmost={0}), 免得行被别的窗口盖着点不到" -f $drTop)

# 先点一次「刷新」把**陈旧行**清掉: 前面 (2f)/(2g) 杀掉的靶子实例在列表里还留着行,
# 等自动刷新把它剔掉时, 它下面的行会整体上移一位 —— 那是列表内容真的变了, 别当成 bug 量进去。
if($ref2){
  [void][VW]::SendMessageW($ref2.Hwnd,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)   # BM_CLICK
  Start-Sleep -Milliseconds 1200
}

# 自动刷新这一节全程都是关着的(整套自检从开头就把它关了), 阶段二才临时打开
$autoWas = $false
if($auto){
  $autoWas = ([int64][VW]::SendMessageW($auto.Hwnd,0x00F0,[IntPtr]::Zero,[IntPtr]::Zero)) -ne 0   # BM_GETCHECK
  if($autoWas){
    [void][VW]::SendMessageW($auto.Hwnd,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)   # 关掉它, 让列表冻住
    Start-Sleep -Milliseconds 500
  }
  Log ("  auto-refresh was={0} (行号断言阶段要它关着)" -f $autoWas)
}

$nItems2 = [int][VW]::SendMessageW($lv,0x1004,[IntPtr]::Zero,[IntPtr]::Zero)     # LVM_GETITEMCOUNT
Log ("  items={0}" -f $nItems2)

# 行 k 的落点(客户区 y): 行顶 = 32 + 20k, 取行中间
function Row-Y($k){ return (32 + $k*20 + 10) }
function Row-Clickable($clientY){
  $o = Get-ClientOrigin $lv
  return ([VW]::WindowFromPoint((To-Pt ($o[0]+60) ($o[1]+$clientY))) -eq $lv)
}
function Close-Menu-Esc{
  [VW]::keybd_event(0x1B,0,0,[IntPtr]::Zero); Start-Sleep -Milliseconds 60
  [VW]::keybd_event(0x1B,0,2,[IntPtr]::Zero); Start-Sleep -Milliseconds 700
  return ((Find-TopWindow '#32768') -eq [IntPtr]::Zero)
}

# 挑两行都能点的(下面那张表是机器上真实的窗口, 行数够不够、有没有被盖住都不由我们说了算)
$pairs = @()
$coverLogged = $false
foreach($p in @(@(2,5),@(4,7),@(3,6),@(1,4))){
  if($pairs.Count -ge 2){ break }
  if($p[1] -ge $nItems2){ continue }
  if(-not (Row-Clickable (Row-Y $p[0])) -or -not (Row-Clickable (Row-Y $p[1]))){
    # 点不到就说清楚是谁盖在上面 —— 真鼠标一点下去会打到它身上, 宁可不动手
    if(-not $coverLogged){
      $coverLogged = $true
      $o = Get-ClientOrigin $lv
      $hit = [VW]::WindowFromPoint((To-Pt ($o[0]+60) ($o[1]+(Row-Y $p[0]))))
      $sb = New-Object System.Text.StringBuilder 128
      [void][VW]::GetClassNameW($hit,$sb,128)
      $dr = Get-RectOf $script:hMain
      Log ("  第 {0} 行那一点点不到列表: 上面盖着 hwnd={1} cls='{2}' cap='{3}'; 工具窗口={4},{5},{6},{7} 列表={8}" -f `
        $p[0],([int64]$hit),$sb.ToString(),(Get-Cap $hit),$dr.L,$dr.T,$dr.R,$dr.B,([int64]$lv))
    }
    continue
  }
  $pairs += ,$p
}
Chk ($pairs.Count -ge 2) ('(2h) 挑到 {0} 组可点的行(每组: 先选中前一行, 再右键后一行)' -f $pairs.Count)

# 这一段的“行号”断言必须在**列表冻着**的时候做: 列表按 Z 序排, 一开自动刷新就要每 2.5 秒
# 重排一次, 从“读到某行的名字”到“真的把光标点上去”之间那一行可能已经换人了(实测有一轮
# 整块挪了 3 行, 于是右键盘点到隔壁窗口上) —— 那是表在动, 不是选中行飘。重建的影响放到
# 下面阶段二单独量, 并且只按“哪个窗口”量、不按行号。
$lastRow = -1; $lastName = ''
foreach($p in $pairs){
  $prevRow = $p[0]; $row = $p[1]
  $cyPrev = Row-Y $prevRow; $cyRow = Row-Y $row
  $nNow = [int][VW]::SendMessageW($lv,0x1004,[IntPtr]::Zero,[IntPtr]::Zero)
  if($row -ge $nNow){
    Log ("  条目只剩 {0} 行, 跳过第 {1} 行" -f $nNow,$row)
    continue
  }

  # 被右键那一行的应用名: 投递一次左键问程序自己(按钮会改名为「<应用名>快排」)
  $nameRow = Get-RowNameAt $btnTile $lv $cyRow
  # 再把选中行挪到**别的**那一行 —— 这正是用户的操作: 列表里选着 A, 却去右键 B
  $null = Get-RowNameAt $btnTile $lv $cyPrev
  $selBefore = Sel-Set $lv
  Log ("  前置: 选中={0} (要右键的是第 {1} 行 '{2}')" -f (Set-Text $selBefore),$row,$nameRow)

  # 这一下是真鼠标。落点是“客户区原点 + 行号换算的 y”, 所以这 600 毫秒里工具窗口**不能动** ——
  # 机器上有人正在拖窗口/最大化时, 窗口一挪同样的落点就落到别的行上去了(实测: 工具窗口被最大化后
  # 客户区原点变了, 算好第 5 行的落点, 右键按中的是第 2 行 —— 代码没错, 是窗口动了)。
  # 所以每次点击前后各量一次列表矩形: 动过就作废重来, 只有“窗口没动、选中行还是不对”才算程序的账。
  $m = [IntPtr]::Zero
  $drifted = $false
  for($try = 1; $try -le 3; $try++){
    $rcB = Get-RectOf $lv
    $m = Open-RowMenu $cyRow
    $rcA = Get-RectOf $lv
    $drifted = ($rcB.L -ne $rcA.L) -or ($rcB.T -ne $rcA.T)
    if($drifted){
      $null = Log ("  第 {0} 次右键期间列表动过({1},{2} -> {3},{4}), 这次不算数, 重来" -f $try,$rcB.L,$rcB.T,$rcA.L,$rcA.T)
      if($m -ne [IntPtr]::Zero){ [void](Close-Menu-Esc) }
      continue
    }
    if($m -ne [IntPtr]::Zero){ break }
  }
  if($drifted){
    Log ("  WARN: 连着三次都没等到窗口不动(机器上有人在动窗口), 第 {0} 行这组跳过, 不记失败" -f $row)
    continue
  }
  if($m -eq [IntPtr]::Zero){
    Chk $false ('(2h) 右键第 {0} 行没能弹出菜单(被别的窗口盖住?)' -f $row)
    continue
  }
  $selAfter = Sel-Set $lv
  $btnCap = Get-Cap $btnTile.Hwnd
  Log ("  右键第 {0} 行后: 选中={1} 按钮='{2}'" -f $row,(Set-Text $selAfter),$btnCap)
  Chk ($selAfter.Count -eq 1 -and $selAfter[0] -eq $row) ('(2h) 右键第 {0} 行 -> 选中的就是这一行(不是别的行)' -f $row)
  Chk ($btnCap -eq $nameRow) ('(2h) 「应用快排」跟着右键那一行走(期望 ' + $nameRow + ', 实际 ' + $btnCap + ')')

  $closed = Close-Menu-Esc
  Log ("  ESC 关菜单 = {0}" -f $closed)
  $lastRow = $row; $lastName = $nameRow
}

# 阶段二: 打开自动刷新, 跨过一轮整表重建 —— 重建只按窗口句柄还原选中行, 所以能保证的是
# “选中的还是同一个窗口、且仍旧只有一行”; 行号不保证(表按 Z 序排, 期间有窗口动过就会挪位)。
if($auto -and $lastRow -ge 0){
  $p1.Refresh()
  Log ("  阶段二开始: driver pid={0} alive={1} title='{2}' 条目={3}" -f `
    $p1.Id,(-not $p1.HasExited),(Get-Cap $script:hMain),([int][VW]::SendMessageW($lv,0x1004,[IntPtr]::Zero,[IntPtr]::Zero)))
  [void][VW]::SendMessageW($auto.Hwnd,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero)
  Start-Sleep -Milliseconds 3300
  $selKeep = Sel-Set $lv
  $capKeep = Get-Cap $btnTile.Hwnd
  $p1.Refresh()
  $errDlg = Find-TopWindow '#32770'
  Log ("  开自动刷新 + 跨过一轮重建后: 选中={0} 按钮='{1}' (重建前是第 {2} 行 '{3}')" -f `
    (Set-Text $selKeep),$capKeep,$lastRow,$lastName)
  Log ("    driver alive={0} title='{1}' 条目={2} 弹窗={3}" -f `
    (-not $p1.HasExited),(Get-Cap $script:hMain),([int][VW]::SendMessageW($lv,0x1004,[IntPtr]::Zero,[IntPtr]::Zero)),([int64]$errDlg))
  if($errDlg -ne [IntPtr]::Zero){ Log ("    弹窗正文: " + (Get-DialogText $errDlg)) }
  Chk ($selKeep.Count -eq 1) '(2h) 整表重建后仍是单选(没有几行一起亮)'
  Chk ($capKeep -eq $lastName) ('(2h) 整表重建后选中的还是同一个窗口(应用名 ' + $lastName + ' 不变)')
  if($selKeep.Count -eq 1 -and $selKeep[0] -ne $lastRow){
    Log ("  行号 {0} -> {1}: 表按 Z 序排, 这一会儿有窗口动过, 行号整体挪位不奇怪" -f $lastRow,$selKeep[0])
  }
} else { Log '  WARN: 没找到自动刷新复选框, 跳过“整表重建”那一半' }

# 还原: 自动刷新回到原样 + 窗口尺寸与置顶状态回到原样
if($auto -and -not $autoWas){ [void][VW]::SendMessageW($auto.Hwnd,0x00F5,[IntPtr]::Zero,[IntPtr]::Zero) }
if(-not $drTop){
  [void][VW]::SetWindowPos($script:hMain,[IntPtr](-2),0,0,0,0,$SWP_NMNA)   # HWND_NOTOPMOST
}
[void][VW]::SetWindowPos($script:hMain,[IntPtr]::Zero,$dr1.L,$dr1.T,($dr1.R-$dr1.L),($dr1.B-$dr1.T),$SWP_NA)
Start-Sleep -Milliseconds 400
Log ("  还原: 置顶={0} 工具窗口={1},{2},{3},{4}" -f (EX-TopMost $script:hMain),$dr1.L,$dr1.T,$dr1.R,$dr1.B)
if($cursorSaved2){ [void][VW]::SetCursorPos($pt0.X,$pt0.Y) }

# ============ 清理 ============
try { Stop-Process -Id $p1.Id -Force -ErrorAction SilentlyContinue } catch {}
if($pv){ try { Stop-Process -Id $pv.Id -Force -ErrorAction SilentlyContinue } catch {} }
try { Stop-Process -Id $p2.Id -Force -ErrorAction SilentlyContinue } catch {}
Log '== verify done =='
Log ("RESULT: {0} failure(s)" -f $script:fail)
exit $script:fail
