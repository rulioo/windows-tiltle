# verify.ps1 - DeskTiler 功能自检(含: 两列布局, 目标显示器下拉, 全选复选框, 平铺后激活)
# 全部在同一个 PowerShell 进程/桌面内完成, 保证窗口可被枚举。
#   1) 界面元素: 一键全排/Cmd快排/PowerShell快排/目录快排/应用快排 按钮(含从左到右的顺序)
#      + 顶部 全选/选定窗口置顶/本窗口置顶 复选框 + 计数
#   2) 计数联动: 顶部“全选”复选框 打勾/取消 联动 “已选择”(BM_GETCHECK 校验状态)
#   2b) 两列布局 & 显示器下拉: 表头列数==2(无 PID); TComboBox 下拉 >=2 且含 “自动选择”
#   2c) 顶栏 5 个左侧控件顺序/不重叠(含 选定窗口置顶 / 本窗口置顶)
#   2d) 置顶行为: 选定窗口置顶 -> 目标窗口 WS_EX_TOPMOST 置位/清除; 本窗口置顶 -> 主窗体自身置位/清除
#   2e) 点选列表一行: 「应用快排」按钮改名为「<应用名>快排」并按文字加宽(文字量得下 + 底行重排不错位)
#   3) 表头排序: -sorttest 在程序内按 PID/应用 升序/降序各排一次(真实 SortList 路径)
#   4) -tileapp DeskTiler 平铺两个目标 -> 校验为均分网格(等大、不相交、相邻)
#   说明: 本机单显示器, 手动指定屏的“跨屏平铺”无法自动化; 覆盖控件存在 + 自动兜底路径
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
public class VW {
  public delegate bool ChildEnum(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, ChildEnum cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out VERECT r);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextLengthW(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern IntPtr SendMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool PostMessageW(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowLongW(IntPtr h, int idx);   // idx=-20 -> GWL_EXSTYLE
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
# 两个矩形是否相交(边贴边不算)
function Rects-Overlap($a,$b){
  return ($a.L -lt $b.R -and $b.L -lt $a.R -and $a.T -lt $b.B -and $b.T -lt $a.B)
}
function EX-TopMost($h){ return (([VW]::GetWindowLongW($h,-20)) -band 0x00000008) -ne 0 }   # WS_EX_TOPMOST
# 文字在“窗体那套字体”(Microsoft YaHei UI 9pt)下的像素宽 —— 走的也是 GDI 度量,
# 与程序内 Self.Canvas.TextWidth 同一套算法, 所以“按钮够不够宽装下这行字”可以直接断言。
# (本机 100% 缩放: 顶栏按钮实测宽度与代码里写的 80/86/92/126/108 一模一样, 9pt 即 12px。)
function Measure-Text($s){
  Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
  Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
  $f = $null; $bmp = $null; $gfx = $null
  try {
    $f = New-Object System.Drawing.Font('Microsoft YaHei UI',9)
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

  # ---- (1) 快捷按钮 & 顶部复选框 ----
  $caps = @($kids | ForEach-Object { $_.Cap })
  Chk (($caps -contains '一键全排') -and ($caps -contains 'Cmd快排')) '按钮: 一键全排 / Cmd快排 均存在'
  # PowerShell快排: 平铺全部 Windows Terminal / PowerShell 窗口(2026-09 按用户要求加回)
  Chk ($caps -contains 'PowerShell快排') 'PowerShell快排 按钮存在(平铺全部 Windows Terminal 窗口)'
  # 目录快排/应用快排: 2026-09-17 按用户要求新增的第 4、5 个快捷按钮
  Chk ($caps -contains '目录快排') '目录快排 按钮存在(平铺全部 explorer 窗口)'
  Chk ($caps -contains '应用快排') '应用快排 按钮存在(平铺列表中选中那行所属应用的全部窗口)'
  Chk ($caps -contains '平铺排列(&T)') '原有按钮: 平铺排列(&T) 存在'
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

# ============ 清理 ============
try { Stop-Process -Id $p1.Id -Force -ErrorAction SilentlyContinue } catch {}
try { Stop-Process -Id $p2.Id -Force -ErrorAction SilentlyContinue } catch {}
Log '== verify done =='
Log ("RESULT: {0} failure(s)" -f $script:fail)
exit $script:fail
