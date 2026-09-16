unit uMain;

{ DeskTiler 主窗体
  - 枚举当前所有可见的顶层应用窗口(主窗口)
  - 勾选要整理的窗口后, 按网格均匀平铺到所在显示器的工作区 }

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  System.Generics.Defaults, System.Types,
  Winapi.Windows, Winapi.Messages,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.Graphics, Vcl.Samples.Spin, Vcl.Imaging.jpeg;

type
  HMONITOR = NativeUInt;

  TWinInfo = record
    Handle: HWND;
    PID: DWORD;
    Title: string;
    AppName: string;
  end;

  TMainForm = class(TForm)
  private
    FEnumList: TList<TWinInfo>;
    FProcNames: TDictionary<DWORD, string>;

    // 表头排序状态(FSortCol 即表头列号: 0=应用程序 1=窗口标题 2=PID, 2 仅供 -sorttest 内部用)
    FSortCol: Integer;   // -1 = 未排序(保持 Z 序)
    FSortAsc: Boolean;

    // UI 控件
    Lv: TListView;
    BtnRefresh: TButton;
    ChkAll, ChkAuto: TCheckBox;   // ChkAll=顶部“全选”复选框; ChkAuto=自动刷新
    BtnTile, BtnAllTile, BtnCmdTile, BtnPsTile: TButton;
    BtnDirTile, BtnAppTile: TButton;   // 目录快排(explorer) / 应用快排(列表中鼠标选中的那行所属应用)
    pnlActs: TPanel;         // 底部快捷按钮行(含 PowerShell快排/Cmd快排/一键全排); 字段可见以便重排
    CmbCols: TComboBox;
    SpinGap: TSpinEdit;
    CmbMon: TComboBox;       // 目标显示器下拉(0=自动选择); Items[1..n] 与 FMonitors 平行
    FMonitors: TArray<HMONITOR>;
    LblStatus, LblMsg: TStaticText; // LblStatus=窗口计数; LblMsg=提示/操作结果 (TStaticText=有句柄, 便于自动化校验)
    LblAbout: TLabel;           // 顶部“关于”链接(悬停显示收款码图片)
    FAboutPopup: TForm;         // 悬停弹出的收款码小窗(无边框、置顶、不抢焦点)
    FAboutShown: Boolean;       // 收款码小窗当前是否可见(用于跳过鼠标进入自身引发的重入)
    FAboutMiss: Integer;        // 鼠标连续几次巡检都不在链接/小窗内(达到阈值才关闭)
    FAboutHideRect: TRect;      // 上次关闭时小窗的位置(屏幕坐标)
    FAboutHideGuard: Boolean;   // True 时: 鼠标仍停在小窗原位就不重新弹出(防"点一下又弹回来")
    AboutTimer: TTimer;         // 移出“关于”/图片后延时关闭, 顺带处理移向图片途中的过渡
    Timer: TTimer;
    FUpdatingAll: Boolean;      // 全选复选框批量勾选期间为 True, 抑制逐行 OnChange 刷新
    FSelectedApp: string;       // 列表里最近被鼠标点选那行的“应用程序”名(供“应用快排”用):
                                // 不能用 Lv.Selected 现取 —— 列表每 2.5 秒自动刷新时会 Clear 重建,
                                // 选中行会被抹掉; 而且点完行还要把鼠标移到按钮上, 中间早就刷新过了。

    procedure BuildUI;
    procedure RefreshApps;
    procedure UpdateStatus;
    procedure PopulateList;      // 用 FEnumList 重建列表, 尽量保留勾选与选中行
    procedure SortList;          // 按 FSortCol / FSortAsc 排序 FEnumList
    procedure UpdateHeaderArrows; // 在表头画上/下箭头指示排序方向
    procedure OnRefreshClick(Sender: TObject);
    procedure OnChkAllClick(Sender: TObject);  // 顶部“全选”复选框: 打勾=全选, 取消=全不选
    procedure OnTileClick(Sender: TObject);
    procedure OnAllTileClick(Sender: TObject);
    procedure OnCmdTileClick(Sender: TObject);
    procedure OnPsTileClick(Sender: TObject);   // 平铺全部 Windows Terminal / PowerShell 窗口
    procedure OnDirTileClick(Sender: TObject);  // 平铺全部资源管理器(explorer)窗口
    procedure OnAppTileClick(Sender: TObject);  // 平铺列表中鼠标选中那行所属应用的全部窗口
    procedure OnColumnClick(Sender: TObject; Column: TListColumn);
    procedure OnAutoToggle(Sender: TObject);
    procedure OnTimerTick(Sender: TObject);
    procedure OnListChange(Sender: TObject; Item: TListItem; Change: TItemChange);
    procedure OnListClicked(Sender: TObject);
    procedure OnListSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure OnFormResize(Sender: TObject);
    procedure OnMonDropDown(Sender: TObject);   // 展开显示器下拉时刷新监视器列表(热插拔)

    procedure OnAboutMouseEnter(Sender: TObject);  // “关于”悬停: 弹出收款码图片
    procedure OnAboutMouseLeave(Sender: TObject);  // 移出“关于”: 延时关闭收款码
    procedure OnAboutClick(Sender: TObject);
    procedure OnAboutPopupEnter(Sender: TObject);  // 鼠标进入收款码小窗: 保持显示
    procedure OnAboutPopupLeave(Sender: TObject);  // 移出收款码小窗: 延时关闭
    procedure OnAboutPopupClick(Sender: TObject);  // 点击收款码图片 → 直接关闭
    procedure OnAboutPopupMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);          // 点击兜底: 图片收不到点击时也能关
    procedure OnAboutTimer(Sender: TObject);       // 巡检计时器: 鼠标不在链接/小窗内则关闭
    procedure ShowAboutPic;                        // 创建小窗并显示收款码
    procedure HideAboutPic;                        // 关闭收款码小窗

    // 底部快捷按钮组靠右摆放(PowerShell快排 / Cmd快排 / 一键全排), 窗口宽度变化时重算
    procedure LayoutActButtons;

    function GetProcessName(APID: DWORD): string;

    // 平铺并返回实际移动的窗口数(格子列/行由 out 返回)
    function TileWindowsFromList(const L: TList<HWND>; out ACols, ARows: Integer): Integer;
    function PickTargetWorkArea(const L: TList<HWND>; out Work: TRect): Boolean;
    function SelectedMonitor: HMONITOR;          // 0 = 自动(多数窗口所在屏)
    procedure RefreshMonitors;                   // 用 EnumDisplayMonitors 重建下拉
    function ResolveTargetArea(const L: TList<HWND>; out Work: TRect): Boolean;
    procedure ActivateWindow(hw: HWND);  // 还原并置顶激活某窗口, 避免被遮挡
    procedure ChooseGrid(const N: Integer; const AreaW, AreaH: Integer;
      out Cols, Rows: Integer);
    procedure TileAll;                       // 一键重排全部窗口
    function IsCmdWin(const W: TWinInfo): Boolean;
    function IsPsWin(const W: TWinInfo): Boolean;
    function IsDirWin(const W: TWinInfo): Boolean;

    class function EnumWndProc(h: HWND; lParam: LPARAM): BOOL; stdcall; static;
    class function EnumMonProc(hMonitor: HMONITOR; hdcMonitor: HDC;
      lprcMonitor: PRect; dwData: LPARAM): BOOL; stdcall; static;
  public
    // 平铺某一可执行程序的所有窗口(命令行 -tileapp, 亦供测试)
    procedure TileByApp(const AName: string);
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    // 勾选全部已勾选窗口并平铺
    procedure TileSelected;
    // 自动勾选列表前 ACount 个窗口并平铺(命令行 -tilefirst)
    procedure TileFirstN(ACount: Integer);
    // 把当前窗口列表写到文件(命令行 -list)
    procedure DumpList(const AFileName: string);
    // 自检: 按 PID/应用 升/降序各排一次并写入文件(命令行 -sorttest)
    procedure SortSelfTest(const AFileName: string);
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

uses
  Winapi.Dwmapi;

{$I uVersion.inc}   // 版本常量: 由 build.cmd / tools\bumpversion.ps1 每次构建自动生成

const
  ID_Cloaked = 14; { DWMWA_CLOAKED }
  PROCESS_QUERY_LIMITED_INFORMATION = $1000;
  MONITOR_DEFAULTTONEAREST = 2;
  MONITOR_DEFAULTTOPRIMARY = 1;
  MONITORINFOF_PRIMARY = 1;

type
  PMonitorInfo = ^TMonitorInfo;
  TMonitorInfo = record
    cbSize: DWORD;
    rcMonitor: TRect;
    rcWork: TRect;
    dwFlags: DWORD;
    szDevice: array[0..31] of WideChar;   { MONITORINFOEX 的 CCHDEVICENAME=32 }
  end;
  TMonitorEnumProc = function(hMonitor: HMONITOR; hdcMonitor: HDC;
    lprcMonitor: PRect; dwData: LPARAM): BOOL; stdcall;

{ 以下 API 不在 Winapi.Windows 中, 自行声明 }
function MonitorFromWindow(hWnd: HWND; dwFlags: DWORD): HMONITOR; stdcall;
  external 'user32.dll' name 'MonitorFromWindow';
function GetMonitorInfo(hMonitor: HMONITOR; lpmi: PMonitorInfo): BOOL; stdcall;
  external 'user32.dll' name 'GetMonitorInfoW';
function EnumDisplayMonitors(hdc: HDC; lprcClip: PRect;
  lpfnEnumProc: TMonitorEnumProc; dwData: LPARAM): BOOL; stdcall;
  external 'user32.dll' name 'EnumDisplayMonitors';
function QueryFullProcessImageNameW(hProcess: THandle; dwFlags: DWORD;
  lpExeName: PWideChar; lpdwSize: PDWORD): BOOL; stdcall;
  external 'kernel32.dll' name 'QueryFullProcessImageNameW';
function CompareStringW(Locale: LCID; dwCmpFlags: DWORD; lpString1: PWideChar;
  cchCount1: Integer; lpString2: PWideChar; cchCount2: Integer): Integer; stdcall;
  external 'kernel32.dll' name 'CompareStringW';

const
  LOCALE_USER_DEFAULT = $400;
  NORM_IGNORECASE = 1;
  CSTR_LESS_THAN = 1;
  CSTR_EQUAL = 2;
  CSTR_GREATER_THAN = 3;

{ 按当前用户区域设置的排序规则比较两串(中文按拼音/笔顺排序), 忽略大小写 }
function LocaleCompareText(const A, B: string): Integer;
var
  n: Integer;
begin
  n := CompareStringW(LOCALE_USER_DEFAULT, NORM_IGNORECASE,
    PWideChar(A), Length(A), PWideChar(B), Length(B));
  case n of
    CSTR_LESS_THAN: Result := -1;
    CSTR_GREATER_THAN: Result := 1;
  else
    Result := 0;
  end;
end;

{ 计算一个进程的可执行文件名(仅用于显示), 结果按进程缓存 }
function TMainForm.GetProcessName(APID: DWORD): string;
var
  s: string;
  hProc: THandle;
  buf: array[0..2047] of WideChar;
  size: DWORD;
begin
  if FProcNames.TryGetValue(APID, s) then
    Exit(s);

  Result := '';
  hProc := OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, APID);
  if hProc = 0 then
    hProc := OpenProcess(PROCESS_QUERY_INFORMATION, False, APID);
  if hProc <> 0 then
  try
    size := 2048;
    if QueryFullProcessImageNameW(hProc, 0, PWideChar(@buf), @size) then
    begin
      SetString(s, PWideChar(@buf), size);
      Result := ChangeFileExt(ExtractFileName(s), '');
    end;
  finally
    CloseHandle(hProc);
  end;

  if Result = '' then
    Result := 'PID' + IntToStr(APID);
  FProcNames.Add(APID, Result);
end;

{ EnumWindows 回调: 收集“可作为应用打开”的顶层主窗口 }
class function TMainForm.EnumWndProc(h: HWND; lParam: LPARAM): BOOL;
var
  SelfPtr: TMainForm;
  pid: DWORD;
  owner: HWND;
  n: Integer;
  tmp: array[0..1023] of WideChar;
  cloaked: LongBool;
  rc: TRect;
  wi: TWinInfo;
begin
  Result := True; // 继续枚举
  SelfPtr := TMainForm(Pointer(lParam));
  if SelfPtr = nil then
    Exit;

  pid := 0;
  GetWindowThreadProcessId(h, @pid);
  if (pid = 0) or (pid = GetCurrentProcessId) then
    Exit; // 排除本进程(含本程序窗体与隐藏主窗)

  // 只保留无 owner 的主窗口, 排除对话框/下拉/弹出等附属窗口
  owner := GetWindow(h, GW_OWNER);
  if owner <> 0 then
    Exit;

  if not IsWindowVisible(h) then
    Exit;

  if (GetWindowLong(h, GWL_EXSTYLE) and WS_EX_TOOLWINDOW) <> 0 then
    Exit;

  // 排除 DWM 折叠(cloaked)窗口, 如其它虚拟桌面 / 后台 UWP
  cloaked := False;
  if DwmGetWindowAttribute(h, ID_Cloaked, @cloaked, SizeOf(cloaked)) = S_OK then
    if cloaked then
      Exit;

  // 窗口必须有标题
  n := GetWindowTextW(h, tmp, Length(tmp));
  if n <= 0 then
    Exit;

  // 无可见面积的窗口跳过
  if GetWindowRect(h, rc) then
    if (rc.Right <= rc.Left) or (rc.Bottom <= rc.Top) then
      Exit;

  wi.Handle := h;
  wi.PID := pid;
  SetString(wi.Title, PWideChar(@tmp), n);
  wi.AppName := SelfPtr.GetProcessName(pid);
  SelfPtr.FEnumList.Add(wi);
end;

{ 重新枚举并刷新列表, 尽量保留此前的勾选; 若已设置表头排序则按排序显示 }
procedure TMainForm.RefreshApps;
begin
  if (FEnumList = nil) or (FProcNames = nil) or (Lv = nil) then
    Exit;

  FEnumList.Clear;
  FProcNames.Clear;
  EnumWindows(@EnumWndProc, LPARAM(Self));

  if FSortCol >= 0 then
    SortList;   // 用户选过表头排序, 刷新后仍保持
  PopulateList;
  UpdateStatus;
end;

{ 用 FEnumList 重建列表视图, 只保留仍存在的窗口的勾选 }
procedure TMainForm.PopulateList;
var
  checkedOld: TDictionary<NativeUInt, Boolean>;
  selOld: NativeUInt;
  i: Integer;
  item: TListItem;
begin
  if (FEnumList = nil) or (Lv = nil) then
    Exit;

  // 重建前记住当前选中行是哪个窗口, 重建后按句柄选回去; 否则每 2.5 秒一次的自动刷新
  // 会把用户刚点选的那一行抹掉, “应用快排”就要用户跟刷新抢时间了。
  selOld := 0;
  if Lv.Selected <> nil then
    selOld := NativeUInt(Lv.Selected.Data);

  checkedOld := TDictionary<NativeUInt, Boolean>.Create;
  try
    for i := 0 to Lv.Items.Count - 1 do
      if Lv.Items[i].Checked then
        checkedOld.Add(NativeUInt(Lv.Items[i].Data), True);

    Lv.Items.BeginUpdate;
    try
      Lv.Items.Clear;
      for i := 0 to FEnumList.Count - 1 do
      begin
        item := Lv.Items.Add;
        item.Caption := FEnumList[i].AppName;   // 第 0 列 = 应用程序(窄)
        item.SubItems.Add(FEnumList[i].Title);  // 第 1 列 = 窗口标题(撑满)
        item.Data := Pointer(NativeUInt(FEnumList[i].Handle));
        if checkedOld.ContainsKey(NativeUInt(FEnumList[i].Handle)) then
          item.Checked := True;
        if (selOld <> 0) and (NativeUInt(FEnumList[i].Handle) = selOld) then
          item.Selected := True;   // 选中行原位保留
      end;
    finally
      Lv.Items.EndUpdate;
    end;
  finally
    checkedOld.Free;
  end;
end;

{ 按 FSortCol / FSortAsc 对 FEnumList 排序 }
procedure TMainForm.SortList;
var
  cmp: TComparison<TWinInfo>;
  col: Integer;
  asc: Boolean;
begin
  if (FEnumList = nil) or (FEnumList.Count < 2) or (FSortCol < 0) then
    Exit;
  col := FSortCol;
  asc := FSortAsc;
  cmp := function(const A, B: TWinInfo): Integer
    begin
      case col of
        0: Result := LocaleCompareText(A.AppName, B.AppName);  // 第 0 列 应用程序
        1: Result := LocaleCompareText(A.Title, B.Title);      // 第 1 列 窗口标题
        2: if A.PID < B.PID then                                // 仅供 -sorttest 内部使用
            Result := -1
          else if A.PID > B.PID then
            Result := 1
          else
            Result := 0;
      else
        Result := LocaleCompareText(A.Title, B.Title);
      end;
      if not asc then
        Result := -Result;
    end;
  FEnumList.Sort(TComparer<TWinInfo>.Construct(cmp));
end;

{ 在表头画 ▲/▼, 指示当前排序列与方向 }
procedure TMainForm.UpdateHeaderArrows;
const
  LVM_FIRST = $1000;
  LVM_GETHEADER = LVM_FIRST + 31;
  HDM_FIRST = $1200;
  HDM_GETITEMW = HDM_FIRST + 11;
  HDM_SETITEMW = HDM_FIRST + 12;
  HDI_FORMAT = 4;
  HDF_SORTUP = $400;
  HDF_SORTDOWN = $200;
type
  THDItemW = record
    mask: UINT;
    cxy: Integer;
    pszText: PWideChar;
    hbm: HBITMAP;
    cchTextMax: Integer;
    fmt: Integer;
    lParam: LPARAM;
    iImage: Integer;
    iOrder: Integer;
  end;
var
  hdr: HWND;
  item: THDItemW;
  i: Integer;
  f: Integer;
begin
  if (Lv = nil) or (Lv.Handle = 0) or (Lv.Columns.Count = 0) then
    Exit;
  hdr := SendMessage(Lv.Handle, LVM_GETHEADER, 0, 0);
  if hdr = 0 then
    Exit;

  for i := 0 to Lv.Columns.Count - 1 do
  begin
    FillChar(item, SizeOf(item), 0);
    item.mask := HDI_FORMAT;
    SendMessage(hdr, HDM_GETITEMW, i, LPARAM(@item));
    f := item.fmt and not (HDF_SORTUP or HDF_SORTDOWN);
    if i = FSortCol then
      if FSortAsc then
        f := f or HDF_SORTUP
      else
        f := f or HDF_SORTDOWN;
    item.mask := HDI_FORMAT;
    item.fmt := f;
    SendMessage(hdr, HDM_SETITEMW, i, LPARAM(@item));
  end;
end;

procedure TMainForm.UpdateStatus;
var
  sel, i: Integer;
begin
  if Lv = nil then
    Exit;
  sel := 0;
  for i := 0 to Lv.Items.Count - 1 do
    if Lv.Items[i].Checked then
      Inc(sel);
  LblStatus.Caption := Format('窗口总数: %d   已选择: %d', [FEnumList.Count, sel]);
  if ChkAll <> nil then   // 顶部“全选”同步: 列表全部勾选才显示打勾
    ChkAll.Checked := (FEnumList.Count > 0) and (sel = FEnumList.Count);
end;

{ 在多数所选窗口所在的显示器上, 选一个工作区 }
function TMainForm.PickTargetWorkArea(const L: TList<HWND>; out Work: TRect): Boolean;
type
  TMonCnt = record
    Mon: HMONITOR;
    Cnt: Integer;
  end;
var
  arr: array of TMonCnt;
  i, k, best, bestIdx: Integer;
  mon: HMONITOR;
  mi: TMonitorInfo;
  hw: HWND;
begin
  SetLength(arr, 0);
  for i := 0 to L.Count - 1 do
  begin
    hw := L[i];
    if not IsWindow(hw) then
      Continue;
    mon := MonitorFromWindow(hw, MONITOR_DEFAULTTONEAREST);
    k := -1;
    for bestIdx := 0 to Length(arr) - 1 do
      if arr[bestIdx].Mon = mon then
      begin
        k := bestIdx;
        Break;
      end;
    if k < 0 then
    begin
      SetLength(arr, Length(arr) + 1);
      k := Length(arr) - 1;
      arr[k].Mon := mon;
      arr[k].Cnt := 0;
    end;
    Inc(arr[k].Cnt);
  end;

  best := -1;
  bestIdx := -1;
  for i := 0 to Length(arr) - 1 do
    if arr[i].Cnt > best then
    begin
      best := arr[i].Cnt;
      bestIdx := i;
    end;

  if bestIdx >= 0 then
    mon := arr[bestIdx].Mon
  else
    mon := MonitorFromWindow(0, MONITOR_DEFAULTTOPRIMARY);

  FillChar(mi, SizeOf(mi), 0);
  mi.cbSize := SizeOf(mi);
  if GetMonitorInfo(mon, @mi) then
    Work := mi.rcWork
  else
  begin
    // 兜底: 主屏工作区(排除任务栏)
    Work := Rect(0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN));
    SystemParametersInfo(SPI_GETWORKAREA, 0, @Work, 0);
  end;
  Result := True;
end;

{ 当前“显示器”下拉选中的屏句柄; 0 = 自动(多数所选窗口所在屏) }
function TMainForm.SelectedMonitor: HMONITOR;
begin
  Result := 0;
  if (CmbMon = nil) or (CmbMon.ItemIndex <= 0) then
    Exit;
  if CmbMon.ItemIndex - 1 < Length(FMonitors) then
    Result := FMonitors[CmbMon.ItemIndex - 1];
end;

{ 用 EnumDisplayMonitors 重建下拉: “自动选择” + 每台显示器一项 }
procedure TMainForm.RefreshMonitors;
var
  want: HMONITOR;
  i, idx: Integer;
begin
  if CmbMon = nil then
    Exit;
  want := SelectedMonitor;          // 记住当前选择, 刷新后按句柄还原
  CmbMon.Items.BeginUpdate;
  try
    CmbMon.Items.Clear;
    FMonitors := nil;
    CmbMon.Items.Add('自动选择');
    EnumDisplayMonitors(0, nil, @EnumMonProc, LPARAM(Self));
    idx := 0;
    if want <> 0 then
      for i := 0 to Length(FMonitors) - 1 do
        if FMonitors[i] = want then
        begin
          idx := i + 1;
          Break;
        end;
    CmbMon.ItemIndex := idx;        // 始终有效, 不会 -1
  finally
    CmbMon.Items.EndUpdate;
  end;
end;

procedure TMainForm.OnMonDropDown(Sender: TObject);
begin
  RefreshMonitors;   // 展开下拉时刷新监视器列表(热插拔)
end;

{ EnumDisplayMonitors 回调: 给每台显示器追加一项并记录句柄 }
class function TMainForm.EnumMonProc(hMonitor: HMONITOR; hdcMonitor: HDC;
  lprcMonitor: PRect; dwData: LPARAM): BOOL; stdcall;
var
  S: TMainForm;
  mi: TMonitorInfo;
  w, h: Integer;
  dev: string;
begin
  Result := True;   // 继续枚举
  S := TMainForm(Pointer(dwData));
  if S = nil then
    Exit;
  w := 0;
  h := 0;
  if lprcMonitor <> nil then
  begin
    w := lprcMonitor.Right - lprcMonitor.Left;
    h := lprcMonitor.Bottom - lprcMonitor.Top;
  end;
  dev := '';
  FillChar(mi, SizeOf(mi), 0);
  mi.cbSize := SizeOf(mi);
  if GetMonitorInfo(hMonitor, @mi) then
  begin
    dev := PChar(@mi.szDevice);   // 如 “\\.\DISPLAY1”
    if LastDelimiter('\', dev) > 0 then
      dev := Copy(dev, LastDelimiter('\', dev) + 1, MaxInt);
    if (mi.dwFlags and MONITORINFOF_PRIMARY) <> 0 then
      dev := dev + ' (主)';
  end;
  S.CmbMon.Items.Add(Format('显示器 %d   %dx%d  %s',
    [Length(S.FMonitors) + 1, w, h, dev]));
  S.FMonitors := S.FMonitors + [hMonitor];
end;

{ 解析本次平铺的目标工作区: 手动指定屏 -> 其 rcWork; 否则回落自动选择 }
function TMainForm.ResolveTargetArea(const L: TList<HWND>; out Work: TRect): Boolean;
var
  mon: HMONITOR;
  mi: TMonitorInfo;
begin
  mon := SelectedMonitor;
  if mon <> 0 then
  begin
    FillChar(mi, SizeOf(mi), 0);
    mi.cbSize := SizeOf(mi);
    if GetMonitorInfo(mon, @mi) then
    begin
      Work := mi.rcWork;
      Exit(True);
    end;
    // 该屏已被拔出 -> 回落到自动选择
  end;
  Result := PickTargetWorkArea(L, Work);
end;

{ 按“窗口数与屏幕形状尽量匹配 + 空位尽量少”选择网格 }
procedure TMainForm.ChooseGrid(const N: Integer; const AreaW, AreaH: Integer;
  out Cols, Rows: Integer);
const
  DesiredAspect = 1.6;  // 期望窗口接近 16:10
  WasteWeight = 0.35;   // 空位惩罚系数
var
  c, r, waste: Integer;
  cellA, cost, best: Double;
begin
  Cols := 1;
  Rows := N;
  if N <= 1 then
  begin
    Rows := 1;
    Exit;
  end;
  best := 1.0E300;
  for c := 1 to N do
  begin
    r := (N + c - 1) div c;
    waste := c * r - N;
    if AreaH > 0 then
      cellA := (AreaW / c) / (AreaH / r)
    else
      cellA := 1;
    cost := Abs(Ln(cellA / DesiredAspect)) + WasteWeight * waste;
    if cost < best then
    begin
      best := cost;
      Cols := c;
      Rows := r;
    end;
  end;
end;

{ 把某窗口还原并置顶激活(绕过 Windows 前台锁) }
procedure TMainForm.ActivateWindow(hw: HWND);
var
  ct, fg: DWORD;
begin
  if not IsWindow(hw) then
    Exit;
  if IsIconic(hw) then
    ShowWindow(hw, SW_RESTORE);
  ct := GetCurrentThreadId;
  fg := GetWindowThreadProcessId(GetForegroundWindow, nil);
  // 本线程不是前台线程时, 先挂接到前台线程再置前, 避免 SetForegroundWindow 被系统拒绝
  if (fg <> 0) and (fg <> ct) then
  begin
    AttachThreadInput(ct, fg, True);
    try
      SetForegroundWindow(hw);
      BringWindowToTop(hw);
    finally
      AttachThreadInput(ct, fg, False);
    end;
  end
  else
  begin
    SetForegroundWindow(hw);
    BringWindowToTop(hw);
  end;
end;

{ 核心: 把给定的一批窗口均匀平铺; 返回实际移动的窗口数, 网格列/行由 out 返回 }
function TMainForm.TileWindowsFromList(const L: TList<HWND>;
  out ACols, ARows: Integer): Integer;
var
  N, Cols, Rows, i, c, r, gap, cellW, cellH, x, y, WorkW, WorkH: Integer;
  rc: TRect;
  hw: HWND;
begin
  ACols := 0;
  ARows := 0;
  Result := 0;
  N := L.Count;
  if N = 0 then
    Exit;
  if not ResolveTargetArea(L, rc) then
    Exit;

  WorkW := rc.Right - rc.Left;
  WorkH := rc.Bottom - rc.Top;
  if (WorkW <= 0) or (WorkH <= 0) then
    Exit;

  if CmbCols.ItemIndex > 0 then // 用户指定了列数
    Cols := CmbCols.ItemIndex
  else
    ChooseGrid(N, WorkW, WorkH, Cols, Rows);
  Rows := (N + Cols - 1) div Cols;

  gap := SpinGap.Value;
  cellW := (WorkW - gap * (Cols - 1)) div Cols;
  cellH := (WorkH - gap * (Rows - 1)) div Rows;
  if cellW < 1 then cellW := 1;
  if cellH < 1 then cellH := 1;

  for i := 0 to N - 1 do
  begin
    hw := L[i];
    if not IsWindow(hw) then
      Continue;

    // 先还原最小化 / 最大化, 才能重新定尺寸
    if IsIconic(hw) then
      ShowWindow(hw, SW_RESTORE);
    if IsZoomed(hw) then
      ShowWindow(hw, SW_RESTORE);

    r := i div Cols;
    c := i mod Cols;
    x := rc.Left + c * (cellW + gap);
    y := rc.Top + r * (cellH + gap);
    MoveWindow(hw, x, y, cellW, cellH, True);
    Inc(Result);
  end;

  // 全部移完后, 逐个激活/置顶这些窗口, 避免被其它活动窗口遮挡
  for i := 0 to N - 1 do
    if IsWindow(L[i]) then
      ActivateWindow(L[i]);

  ACols := Cols;
  ARows := Rows;
end;

procedure TMainForm.TileSelected;
var
  L: TList<HWND>;
  i, cols, rows, n: Integer;
begin
  L := TList<HWND>.Create;
  try
    for i := 0 to Lv.Items.Count - 1 do
      if Lv.Items[i].Checked then
        L.Add(HWND(NativeUInt(Lv.Items[i].Data)));

    if L.Count = 0 then
    begin
      Application.MessageBox('请先在列表中勾选要平铺的窗口。', 'DeskTiler',
        MB_OK + MB_ICONINFORMATION);
      Exit;
    end;

    n := TileWindowsFromList(L, cols, rows);
    if n > 0 then
      LblMsg.Caption := Format('已将 %d 个勾选窗口平铺为 %d 列 × %d 行', [n, cols, rows]);
    UpdateStatus;
  finally
    L.Free;
  end;
end;

procedure TMainForm.TileFirstN(ACount: Integer);
var
  L: TList<HWND>;
  i, cols, rows: Integer;
begin
  L := TList<HWND>.Create;
  try
    if ACount < 1 then
      ACount := 1;
    for i := 0 to Lv.Items.Count - 1 do
    begin
      if i >= ACount then
        Break;
      Lv.Items[i].Checked := True;
      L.Add(HWND(NativeUInt(Lv.Items[i].Data)));
    end;
    TileWindowsFromList(L, cols, rows);
  finally
    L.Free;
  end;
end;

{ 一键: 把所有窗口按当前设置平铺 }
procedure TMainForm.TileAll;
var
  L: TList<HWND>;
  i, cols, rows, n: Integer;
begin
  L := TList<HWND>.Create;
  try
    for i := 0 to FEnumList.Count - 1 do
      if IsWindow(FEnumList[i].Handle) then
        L.Add(FEnumList[i].Handle);

    if L.Count = 0 then
    begin
      LblMsg.Caption := '当前没有可平铺的窗口';
      Exit;
    end;

    n := TileWindowsFromList(L, cols, rows);
    if n > 0 then
      LblMsg.Caption := Format('已一键平铺全部 %d 个窗口（%d 列 × %d 行）', [n, cols, rows]);
  finally
    L.Free;
  end;
end;

{ 按可执行文件名(去掉扩展名)过滤并平铺, 用于 -tileapp 与自动化测试 }
procedure TMainForm.TileByApp(const AName: string);
var
  L: TList<HWND>;
  i, cols, rows, n: Integer;
begin
  L := TList<HWND>.Create;
  try
    for i := 0 to FEnumList.Count - 1 do
      if SameText(FEnumList[i].AppName, AName) and IsWindow(FEnumList[i].Handle) then
        L.Add(FEnumList[i].Handle);

    if L.Count = 0 then
    begin
      LblMsg.Caption := Format('未检测到 %s 窗口', [AName]);
      Exit;
    end;

    n := TileWindowsFromList(L, cols, rows);
    if n > 0 then
      LblMsg.Caption := Format('已平铺 %d 个 %s 窗口（%d 列 × %d 行）', [n, AName, cols, rows]);
  finally
    L.Free;
  end;
end;

{ cmd 窗口判定: 归属进程为 cmd/command; 或 Windows 终端里标题含 cmd.exe / 命令提示符 }
function TMainForm.IsCmdWin(const W: TWinInfo): Boolean;
var
  a, t: string;
begin
  a := LowerCase(W.AppName);
  t := LowerCase(W.Title);
  Result := (a = 'cmd') or (a = 'command')
    or ((a = 'windowsterminal') and
      ((Pos('cmd.exe', t) > 0) or (Pos('命令提示符', t) > 0)
      or (Pos('command prompt', t) > 0)));
end;

{ PowerShell/终端窗口判定: Windows Terminal(WindowsTerminal.exe / wt.exe) 的所有窗口,
  外加老式控制台里直接跑的 powershell.exe / pwsh.exe 窗口。 }
function TMainForm.IsPsWin(const W: TWinInfo): Boolean;
var
  a: string;
begin
  a := LowerCase(W.AppName);
  Result := (a = 'windowsterminal') or (a = 'wt')
    or (a = 'powershell') or (a = 'pwsh');
end;

{ 资源管理器窗口判定: 进程名 explorer —— 即每个“文件夹”窗口。
  桌面/任务栏那些 shell 窗口不是“可见+有标题+无 owner”的顶层主窗口, 本来就不会进枚举结果。 }
function TMainForm.IsDirWin(const W: TWinInfo): Boolean;
begin
  Result := SameText(W.AppName, 'explorer');
end;

procedure TMainForm.DumpList(const AFileName: string);
var
  sl: TStringList;
  i: Integer;
begin
  sl := TStringList.Create;
  try
    sl.Add('hwnd=pid=app=title');
    for i := 0 to FEnumList.Count - 1 do
      sl.Add(Format('%d=%d=%s=%s', [Integer(FEnumList[i].Handle), FEnumList[i].PID,
        FEnumList[i].AppName, FEnumList[i].Title]));
    sl.SaveToFile(AFileName, TEncoding.UTF8);
  finally
    sl.Free;
  end;
end;

{ 自检排序: 复现“点表头”走的排序路径(SortList), 把各次结果写入文件供自动化比对 }
procedure TMainForm.SortSelfTest(const AFileName: string);
var
  sl: TStringList;
  procedure Emit(const ATag: string);
  var
    k: Integer;
  begin
    sl.Add(ATag);
    for k := 0 to FEnumList.Count - 1 do
      sl.Add(Format('%d|%s|%s', [Integer(FEnumList[k].PID), FEnumList[k].AppName,
        FEnumList[k].Title]));
  end;
begin
  sl := TStringList.Create;
  try
    RefreshApps;                 // 填充 FEnumList(Lv 创建于构造阶段, 与 -list/-tileapp 同前提)
    Emit('-- z-order --');

    FSortCol := 2; FSortAsc := True;  SortList;
    Emit('-- pid asc --');
    FSortAsc := False;                SortList;
    Emit('-- pid desc --');

    FSortCol := 0; FSortAsc := True;  SortList;
    Emit('-- app asc --');
    FSortAsc := False;                SortList;
    Emit('-- app desc --');

    FSortCol := 1; FSortAsc := True;  SortList;
    Emit('-- title asc --');

    sl.SaveToFile(AFileName, TEncoding.UTF8);
  finally
    sl.Free;
  end;
end;

{ ---------- 事件处理器 ---------- }

procedure TMainForm.OnRefreshClick(Sender: TObject);
begin
  RefreshApps;
end;

procedure TMainForm.OnChkAllClick(Sender: TObject);
var
  i: Integer;
  doCheck: Boolean;
begin
  if Lv = nil then
    Exit;
  doCheck := ChkAll.Checked;   // 用户点击后的新状态
  FUpdatingAll := True;        // 批量设置期间跳过逐行 OnChange -> UpdateStatus
  try
    Lv.Items.BeginUpdate;
    try
      for i := 0 to Lv.Items.Count - 1 do
        Lv.Items[i].Checked := doCheck;
    finally
      Lv.Items.EndUpdate;
    end;
  finally
    FUpdatingAll := False;
  end;
  UpdateStatus;                // 一次性刷新计数 + 复选框状态
end;

procedure TMainForm.OnTileClick(Sender: TObject);
begin
  TileSelected;
end;

procedure TMainForm.OnAllTileClick(Sender: TObject);
begin
  TileAll;
end;

procedure TMainForm.OnCmdTileClick(Sender: TObject);
var
  L: TList<HWND>;
  i, cols, rows, n: Integer;
begin
  L := TList<HWND>.Create;
  try
    for i := 0 to FEnumList.Count - 1 do
      if IsCmdWin(FEnumList[i]) and IsWindow(FEnumList[i].Handle) then
        L.Add(FEnumList[i].Handle);

    if L.Count = 0 then
    begin
      LblMsg.Caption := '未检测到 cmd 窗口';
      Exit;
    end;

    n := TileWindowsFromList(L, cols, rows);
    if n > 0 then
      LblMsg.Caption := Format('已快速平铺 %d 个 cmd 窗口（%d 列 × %d 行）', [n, cols, rows]);
  finally
    L.Free;
  end;
end;

procedure TMainForm.OnPsTileClick(Sender: TObject);
var
  L: TList<HWND>;
  i, cols, rows, n: Integer;
begin
  L := TList<HWND>.Create;
  try
    for i := 0 to FEnumList.Count - 1 do
      if IsPsWin(FEnumList[i]) and IsWindow(FEnumList[i].Handle) then
        L.Add(FEnumList[i].Handle);

    if L.Count = 0 then
    begin
      LblMsg.Caption := '未检测到 Windows Terminal / PowerShell 窗口';
      Exit;
    end;

    n := TileWindowsFromList(L, cols, rows);
    if n > 0 then
      LblMsg.Caption := Format('已快速平铺 %d 个终端窗口（%d 列 × %d 行）', [n, cols, rows]);
  finally
    L.Free;
  end;
end;

{ 目录快排: 平铺全部资源管理器(explorer)窗口 —— 判定只看进程名, 与窗口里打开的是哪个文件夹无关 }
procedure TMainForm.OnDirTileClick(Sender: TObject);
var
  L: TList<HWND>;
  i, cols, rows, n: Integer;
begin
  L := TList<HWND>.Create;
  try
    for i := 0 to FEnumList.Count - 1 do
      if IsDirWin(FEnumList[i]) and IsWindow(FEnumList[i].Handle) then
        L.Add(FEnumList[i].Handle);

    if L.Count = 0 then
    begin
      LblMsg.Caption := '未检测到资源管理器(explorer)窗口';
      Exit;
    end;

    n := TileWindowsFromList(L, cols, rows);
    if n > 0 then
      LblMsg.Caption := Format('已快速平铺 %d 个资源管理器窗口（%d 列 × %d 行）', [n, cols, rows]);
  finally
    L.Free;
  end;
end;

{ 应用快排: 目标是“鼠标当前选中的应用程序” —— 即列表里被点选(高亮)的那一行,
  取该行的「应用程序」列(可执行文件名)后, 平铺同名的全部窗口。
  目标应用名记在 FSelectedApp(见 OnListSelectItem), 不现取 Lv.Selected —— 列表每 2.5 秒
  自动重建一次, 等鼠标移到按钮上时选中行早没了。
  没有选中过任何行时给出提示而不是猜: 猜错会把毫不相干的窗口铺满屏幕, 比不做事更糟。 }
procedure TMainForm.OnAppTileClick(Sender: TObject);
begin
  if FSelectedApp = '' then
  begin
    LblMsg.Caption := '请先在列表里点选一行（该行的「应用程序」即目标）, 再点“应用快排”';
    Exit;
  end;
  TileByApp(FSelectedApp);   // 内部会写好“已平铺 N 个 xxx 窗口”的提示
end;

procedure TMainForm.OnColumnClick(Sender: TObject; Column: TListColumn);
begin
  // 点同一列在升序/降序间切换; 点新列则从升序开始
  if Column.Index = FSortCol then
    FSortAsc := not FSortAsc
  else
  begin
    FSortCol := Column.Index;
    FSortAsc := True;
  end;
  SortList;
  PopulateList;
  UpdateHeaderArrows;
end;

procedure TMainForm.OnAutoToggle(Sender: TObject);
begin
  if Timer <> nil then
    Timer.Enabled := ChkAuto.Checked;
end;

procedure TMainForm.OnTimerTick(Sender: TObject);
begin
  RefreshApps;
end;

procedure TMainForm.OnListChange(Sender: TObject; Item: TListItem; Change: TItemChange);
begin
  if FUpdatingAll then
    Exit;   // 全选复选框批量勾选中, 末尾统一刷新一次
  UpdateStatus;
end;

procedure TMainForm.OnListClicked(Sender: TObject);
begin
  UpdateStatus;
end;

{ 记住“鼠标当前选中的应用程序”: 列表行被点选/被键盘移动到就记下来。
   这里记的是应用名而不是行对象 —— 行每次自动刷新都会被重建, 应用名不会失效。 }
procedure TMainForm.OnListSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
begin
  if Selected and (Item <> nil) and (Item.Caption <> '') then
    FSelectedApp := Item.Caption;
end;

procedure TMainForm.OnFormResize(Sender: TObject);
begin
  if (Lv <> nil) and (Lv.Columns.Count >= 2) then
    Lv.Columns[1].Width := Lv.ClientWidth - Lv.Columns[0].Width - 4;  // 标题列撑满, 应用列固定 130
  LayoutActButtons;   // 快捷按钮组始终贴着底行右端
end;

{ 底部快捷按钮组: 从右往左依次摆 一键全排 / Cmd快排 / PowerShell快排 / 目录快排 / 应用快排,
  所以从左往右看就是 应用快排 → 目录快排 → PowerShell快排 → Cmd快排 → 一键全排
  (“一键全排”是兜底的那个, 始终排在最右端)。 }
procedure TMainForm.LayoutActButtons;
const
  GAP = 6;      // 按钮之间的横向间隙
  RMARGIN = 12; // 整组距面板右缘留白
var
  x, h, y: Integer;

  { 把某个按钮贴着 x 的右端摆好, 返回它左缘的横坐标（即下一个按钮的右端） }
  function Place(const B: TButton; var CurX: Integer; const AY, AH: Integer): Integer;
  begin
    CurX := CurX - B.Width;
    B.SetBounds(CurX, AY, B.Width, AH);
    Result := CurX;
  end;

begin
  if (pnlActs = nil) or (BtnPsTile = nil) or (BtnCmdTile = nil) or (BtnAllTile = nil)
    or (BtnDirTile = nil) or (BtnAppTile = nil) then Exit;
  h := 26;
  y := (pnlActs.ClientHeight - h) div 2;
  if y < 0 then y := 0;
  x := pnlActs.ClientWidth - RMARGIN;
  Place(BtnAllTile, x, y, h);   Dec(x, GAP);
  Place(BtnCmdTile, x, y, h);   Dec(x, GAP);
  Place(BtnPsTile,  x, y, h);   Dec(x, GAP);
  Place(BtnDirTile, x, y, h);   Dec(x, GAP);
  Place(BtnAppTile, x, y, h);
end;

{ ---------- 构造 UI ---------- }

procedure TMainForm.BuildUI;
var
  pnlTop, pnlBottom, pnlOpts: TPanel;   // pnlActs 是字段(供 LayoutActButtons 使用)
  col: TListColumn;
begin
  Self.Caption := Format('DeskTiler v%s — 桌面窗口均匀平铺', [APP_VER_STR]);
  Self.Position := poScreenCenter;
  Self.ClientWidth := 680;
  Self.ClientHeight := 580;
  Self.Font.Name := 'Microsoft YaHei UI';
  Self.Font.Size := 9;
  Self.DoubleBuffered := True;
  Self.OnResize := OnFormResize;

  // ---- 顶部工具条 ----
  pnlTop := TPanel.Create(Self);
  pnlTop.Parent := Self;
  pnlTop.Align := alTop;
  pnlTop.Height := 42;
  pnlTop.BevelOuter := bvNone;
  pnlTop.Caption := '';

  // 全选复选框(顶栏最左): 打勾=全选, 取消=全不选; 列表勾选变动时自动反映
  ChkAll := TCheckBox.Create(pnlTop);
  ChkAll.Parent := pnlTop;
  ChkAll.Align := alLeft;
  ChkAll.Width := 80;
  ChkAll.Caption := '全选(&A)';
  ChkAll.Checked := False;
  ChkAll.OnClick := OnChkAllClick;
  ChkAll.AlignWithMargins := True;
  ChkAll.Margins.SetBounds(8, 10, 4, 10);

  BtnRefresh := TButton.Create(pnlTop);
  BtnRefresh.Parent := pnlTop;
  BtnRefresh.Align := alLeft;
  BtnRefresh.Width := 86;
  BtnRefresh.Caption := '刷新(&R)';
  BtnRefresh.OnClick := OnRefreshClick;
  BtnRefresh.AlignWithMargins := True;
  BtnRefresh.Margins.SetBounds(0, 6, 4, 6);

  ChkAuto := TCheckBox.Create(pnlTop);
  ChkAuto.Parent := pnlTop;
  ChkAuto.Align := alLeft;
  ChkAuto.Width := 92;
  ChkAuto.Caption := '自动刷新';
  ChkAuto.Checked := True;
  ChkAuto.OnClick := OnAutoToggle;
  ChkAuto.AlignWithMargins := True;
  ChkAuto.Margins.SetBounds(8, 6, 4, 6);

  // “关于”链接: 悬停显示收款码图片(位于顶栏最右侧贴边)
  LblAbout := TLabel.Create(pnlTop);
  LblAbout.Parent := pnlTop;
  LblAbout.Align := alRight;
  LblAbout.Caption := '关于';
  LblAbout.AutoSize := False;
  LblAbout.Width := 44;
  LblAbout.Layout := tlCenter;
  LblAbout.Font.Color := clBlue;
  LblAbout.Font.Style := [fsUnderline];
  LblAbout.Cursor := crHandPoint;
  LblAbout.OnMouseEnter := OnAboutMouseEnter;
  LblAbout.OnMouseLeave := OnAboutMouseLeave;
  LblAbout.OnClick := OnAboutClick;
  LblAbout.AlignWithMargins := True;
  LblAbout.Margins.SetBounds(6, 12, 12, 12);

  LblStatus := TStaticText.Create(pnlTop);
  LblStatus.Parent := pnlTop;
  // 必须用 alClient(占满左侧按钮与右侧“关于”链接之间的剩余宽度), 不能也用 alRight:
  // LblStatus 是**有窗口句柄**的控件, 若与无句柄的“关于”TLabel 抢同一块右侧空间,
  // 它会盖住链接 —— 鼠标事件全被它的窗口吃掉, 链接的 OnMouseEnter/OnClick 永远不触发。
  LblStatus.Align := alClient;
  LblStatus.AutoSize := False;
  LblStatus.Alignment := taRightJustify;
  LblStatus.Color := clBtnFace;
  LblStatus.AlignWithMargins := True;
  LblStatus.Margins.SetBounds(0, 0, 12, 0);

  // ---- 列表 ----
  Lv := TListView.Create(Self);
  Lv.Parent := Self;
  Lv.Align := alClient;
  Lv.ViewStyle := vsReport;
  Lv.Checkboxes := True;
  Lv.ReadOnly := True;
  Lv.RowSelect := True;
  Lv.MultiSelect := True;
  Lv.HideSelection := False;
  Lv.ColumnClick := True;   // 点击表头排序
  Lv.DoubleBuffered := True;
  Lv.OnChange := OnListChange;
  Lv.OnClick := OnListClicked;
  Lv.OnSelectItem := OnListSelectItem;   // 记录选中的应用程序(供“应用快排”)
  Lv.OnColumnClick := OnColumnClick;

  col := Lv.Columns.Add;
  col.Caption := '应用程序';
  col.Width := 130;                      // 固定窄列
  col := Lv.Columns.Add;
  col.Caption := '窗口标题';
  col.Width := Self.ClientWidth - 130 - 4; // 占位; OnFormResize 会按宽度撑满

  // ---- 底部: 操作结果/提示条(最先创建 => 位于最底部) ----
  LblMsg := TStaticText.Create(Self);
  LblMsg.Parent := Self;
  LblMsg.Align := alBottom;
  LblMsg.Height := 22;
  LblMsg.AutoSize := False;
  LblMsg.Color := clBtnFace;
  LblMsg.Caption := '提示: 点表头可排序; 勾选窗口后点“平铺排列”, 或点“一键全排”重排全部窗口。';
  LblMsg.AlignWithMargins := True;
  LblMsg.Margins.SetBounds(12, 0, 8, 0);

  // ---- 底部操作区: 两行(上行: 列数/间距/显示器; 下行: 快捷按钮) ----
  pnlBottom := TPanel.Create(Self);
  pnlBottom.Parent := Self;
  pnlBottom.Align := alBottom;
  pnlBottom.Height := 80;
  pnlBottom.BevelOuter := bvNone;
  pnlBottom.Caption := '';

  // 下行: 快捷按钮组(靠右)
  pnlActs := TPanel.Create(pnlBottom);
  pnlActs.Parent := pnlBottom;
  pnlActs.Align := alBottom;
  pnlActs.Height := 40;
  pnlActs.BevelOuter := bvNone;
  pnlActs.Caption := '';

  // 快捷按钮组: 从左到右 = 应用快排 / 目录快排 / PowerShell快排 / Cmd快排 / 一键全排。
  // 位置由 LayoutActButtons 显式摆放, 不用 alRight 停靠 —— 同一个父容器里多个 alRight
  // 兄弟控件的停靠次序并不等于创建次序(实测对不上, 也正是“关于”链接被盖住那类坑的来源),
  // 显式算坐标才不会摆错。(创建次序无所谓, 摆放次序只看 LayoutActButtons。)
  BtnAppTile := TButton.Create(pnlActs);    // 平铺“列表中选中那行”所属应用的全部窗口
  BtnAppTile.Parent := pnlActs;
  BtnAppTile.Align := alNone;
  BtnAppTile.Width := 92;
  BtnAppTile.Caption := '应用快排';
  BtnAppTile.OnClick := OnAppTileClick;
  BtnAppTile.ShowHint := True;   // 悬停给出用法: 得先在列表里点一行
  BtnAppTile.Hint := '平铺列表中选中那一行所属应用程序的全部窗口（先在列表里点选一行）';

  BtnDirTile := TButton.Create(pnlActs);    // 平铺全部资源管理器(explorer)窗口
  BtnDirTile.Parent := pnlActs;
  BtnDirTile.Align := alNone;
  BtnDirTile.Width := 92;
  BtnDirTile.Caption := '目录快排';
  BtnDirTile.OnClick := OnDirTileClick;
  BtnDirTile.ShowHint := True;
  BtnDirTile.Hint := '平铺全部资源管理器（文件夹）窗口';

  BtnPsTile := TButton.Create(pnlActs);     // 平铺全部 Windows Terminal 窗口
  BtnPsTile.Parent := pnlActs;
  BtnPsTile.Align := alNone;
  BtnPsTile.Width := 108;
  BtnPsTile.Caption := 'PowerShell快排';
  BtnPsTile.OnClick := OnPsTileClick;

  BtnCmdTile := TButton.Create(pnlActs);
  BtnCmdTile.Parent := pnlActs;
  BtnCmdTile.Align := alNone;
  BtnCmdTile.Width := 92;
  BtnCmdTile.Caption := 'Cmd快排';
  BtnCmdTile.OnClick := OnCmdTileClick;

  BtnAllTile := TButton.Create(pnlActs);
  BtnAllTile.Parent := pnlActs;
  BtnAllTile.Align := alNone;
  BtnAllTile.Width := 92;
  BtnAllTile.Caption := '一键全排';
  BtnAllTile.OnClick := OnAllTileClick;

  // 上行: 三个选项组; 每个说明 label 都在其控件左方
  pnlOpts := TPanel.Create(pnlBottom);
  pnlOpts.Parent := pnlBottom;
  pnlOpts.Align := alClient;
  pnlOpts.BevelOuter := bvNone;
  pnlOpts.Caption := '';

  // 列数(label → 下拉框, 约 2 汉字宽)
  with TLabel.Create(pnlOpts) do
  begin
    Parent := pnlOpts;
    Align := alLeft;
    Width := 40;
    Caption := '列数';
    AutoSize := False;
    Layout := tlCenter;
    AlignWithMargins := True;
    Margins.SetBounds(12, 8, 0, 8);
  end;

  CmbCols := TComboBox.Create(pnlOpts);
  CmbCols.Parent := pnlOpts;
  CmbCols.Align := alLeft;
  CmbCols.Width := 58;
  CmbCols.Style := csDropDownList;
  CmbCols.Items.Add('自动');
  CmbCols.Items.Add('1列');
  CmbCols.Items.Add('2列');
  CmbCols.Items.Add('3列');
  CmbCols.Items.Add('4列');
  CmbCols.Items.Add('5列');
  CmbCols.Items.Add('6列');
  CmbCols.Items.Add('7列');
  CmbCols.Items.Add('8列');
  CmbCols.ItemIndex := 0;
  CmbCols.AlignWithMargins := True;
  CmbCols.Margins.SetBounds(0, 8, 6, 8);

  // 间距(label → 数值框, 约 4 位数字宽)
  with TLabel.Create(pnlOpts) do
  begin
    Parent := pnlOpts;
    Align := alLeft;
    Width := 40;
    Caption := '间距';
    AutoSize := False;
    Layout := tlCenter;
    AlignWithMargins := True;
    Margins.SetBounds(6, 8, 0, 8);
  end;

  SpinGap := TSpinEdit.Create(pnlOpts);
  SpinGap.Parent := pnlOpts;
  SpinGap.Align := alLeft;
  SpinGap.Width := 60;
  SpinGap.MinValue := 0;
  SpinGap.MaxValue := 40;
  SpinGap.Value := 4;
  SpinGap.AlignWithMargins := True;
  SpinGap.Margins.SetBounds(0, 8, 6, 8);

  // 目标显示器(label → 下拉框): “自动”=多数所选窗口所在屏; 或手动指定一台
  with TLabel.Create(pnlOpts) do
  begin
    Parent := pnlOpts;
    Align := alLeft;
    Width := 48;
    Caption := '显示器';
    AutoSize := False;
    Layout := tlCenter;
    AlignWithMargins := True;
    Margins.SetBounds(6, 8, 0, 8);
  end;

  CmbMon := TComboBox.Create(pnlOpts);
  CmbMon.Parent := pnlOpts;
  CmbMon.Align := alLeft;
  CmbMon.Width := 150;
  CmbMon.Style := csDropDownList;
  CmbMon.DropDownCount := 12;
  CmbMon.OnDropDown := OnMonDropDown;
  CmbMon.AlignWithMargins := True;
  CmbMon.Margins.SetBounds(0, 8, 4, 8);
  RefreshMonitors;   // 初始填入 “自动” + 当前每台显示器

  // 平铺排列(&T): 与参数(列数/间距/显示器)同一行, 靠右放 —— 主按钮, 醒目且不占快捷按钮行
  BtnTile := TButton.Create(pnlOpts);
  BtnTile.Parent := pnlOpts;
  BtnTile.Align := alRight;
  BtnTile.Width := 118;
  BtnTile.Caption := '平铺排列(&T)';
  BtnTile.Default := True;
  BtnTile.OnClick := OnTileClick;
  BtnTile.Font.Style := [fsBold];
  BtnTile.AlignWithMargins := True;
  BtnTile.Margins.SetBounds(2, 8, 12, 8);

  LayoutActButtons;   // 摆好底行五个快捷按钮(显式定位)
end;

{ ---- “关于”链接: 悬停显示收款码图片 ---- }

procedure TMainForm.ShowAboutPic;
var
  path: string;
  P: TPoint;
  mon: TMonitor;
  wa: TRect;
  w, h, capH: Integer;
  img: TImage;
  pnl: TPanel;
begin
  if FAboutShown then Exit;                 // 已显示则不再处理
  // 刚点掉小窗时鼠标还停在小窗原来的位置上, 此时 VCL 可能又补一个 MouseEnter 过来。
  // 不挡住的话会出现"点一下关掉、立刻又弹回来", 看上去就像点击没用。
  // 只在鼠标还停在小窗原位时挡; 鼠标一移开(或重新移回“关于”)就恢复正常。
  if FAboutHideGuard and PtInRect(FAboutHideRect, Mouse.CursorPos) then Exit;
  if FAboutPopup = nil then
  begin
    path := IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName))
          + '300收款码.jpg';
    if not FileExists(path) then
    begin
      LblMsg.Caption := '未找到收款码图片: ' + path;
      Exit;
    end;
    FAboutPopup := TForm.CreateNew(Self);   // 无 DFM 的顶层小窗
    FAboutPopup.BorderStyle := bsNone;
    FAboutPopup.FormStyle := fsStayOnTop;
    FAboutPopup.Position := poDesigned;
    FAboutPopup.Visible := False;
    FAboutPopup.DoubleBuffered := True;
    FAboutPopup.Color := clWhite;
    FAboutPopup.OnMouseEnter := OnAboutPopupEnter;
    FAboutPopup.OnMouseLeave := OnAboutPopupLeave;
    FAboutPopup.OnClick := OnAboutPopupClick;
    FAboutPopup.OnMouseDown := OnAboutPopupMouseDown;   // 点击兜底

    // 底部版本条: 构建号每次自增, 一眼看出是否已更新
    capH := 22;
    pnl := TPanel.Create(FAboutPopup);
    pnl.Parent := FAboutPopup;
    pnl.Align := alBottom;          // 有窗口句柄的控件, 鼠标进出事件可靠
    pnl.Height := capH;
    pnl.BevelOuter := bvNone;
    pnl.Color := clWhite;
    pnl.Font.Color := clGray;
    pnl.Alignment := taCenter;
    pnl.Caption := Format('DeskTiler v%s   ·   build %d   ·   %s',
      [APP_VER_STR, APP_VER_BUILD, APP_BUILD_STAMP]);
    pnl.Cursor := crHandPoint;
    pnl.OnMouseEnter := OnAboutPopupEnter;
    pnl.OnMouseLeave := OnAboutPopupLeave;
    pnl.OnClick := OnAboutPopupClick;

    img := TImage.Create(FAboutPopup);
    img.Parent := FAboutPopup;
    img.Align := alClient;
    img.Stretch := True;
    img.Cursor := crHandPoint;      // 提示可点击
    img.Picture.LoadFromFile(path);
    img.OnMouseEnter := OnAboutPopupEnter;
    img.OnMouseLeave := OnAboutPopupLeave;
    img.OnClick := OnAboutPopupClick;   // 点击图片即关闭

    // 等比缩放到合理大小, 避免图太大顶出屏幕
    w := img.Picture.Width;
    h := img.Picture.Height;
    if (w > 360) or (h > 360) then
      if w >= h then
      begin
        h := MulDiv(h, 360, w);
        w := 360;
      end
      else
      begin
        w := MulDiv(w, 360, h);
        h := 360;
      end;
    if w < 260 then w := 260;   // 至少留出版本条文字的宽度
    FAboutPopup.ClientWidth := w;
    FAboutPopup.ClientHeight := h + capH;
  end;

  // 定位: 放在“关于”链接的右下方(贴顶栏下缘), 移出屏幕就收回
  P := LblAbout.ClientToScreen(Point(0, 0));
  mon := Screen.MonitorFromPoint(P);
  if mon <> nil then wa := mon.WorkareaRect else wa := Screen.WorkAreaRect;

  // 右缘与“关于”链接右缘对齐, 且上缘紧贴链接下缘 —— 中间不留空隙。
  // 这一点很重要: 鼠标从链接移向图片时, 中间只要有几像素的空档, 就会在
  // “既不在链接内也不在小窗内”的瞬间被巡检判为移出而把图片关掉, 用户根本
  // 来不及移过去点击。贴着放, 鼠标垂直下移始终落在两者之一里面。
  FAboutPopup.Left := P.X + LblAbout.Width - FAboutPopup.ClientWidth;
  if FAboutPopup.Left < wa.Left then FAboutPopup.Left := wa.Left + 4;
  if FAboutPopup.Left + FAboutPopup.ClientWidth > wa.Right then
    FAboutPopup.Left := wa.Right - FAboutPopup.ClientWidth - 4;
  if FAboutPopup.Left < wa.Left then FAboutPopup.Left := wa.Left;

  FAboutPopup.Top := P.Y + LblAbout.Height;          // 紧贴链接下缘
  if FAboutPopup.Top + FAboutPopup.ClientHeight > wa.Bottom then
    FAboutPopup.Top := P.Y - FAboutPopup.ClientHeight; // 下方放不下 → 贴链接上缘
  if FAboutPopup.Top < wa.Top then FAboutPopup.Top := wa.Top;

  // 不抢焦点地显示
  SetWindowLong(FAboutPopup.Handle, GWL_EXSTYLE,
    GetWindowLong(FAboutPopup.Handle, GWL_EXSTYLE) or WS_EX_NOACTIVATE or WS_EX_TOOLWINDOW);
  SetWindowPos(FAboutPopup.Handle, HWND_TOPMOST,
    FAboutPopup.Left, FAboutPopup.Top, FAboutPopup.Width, FAboutPopup.Height,
    SWP_SHOWWINDOW or SWP_NOACTIVATE);
  // 关键: 用 SetWindowPos 显示**不会**更新 VCL 自己的 Visible 标志。不补这一句,
  // 之后调用 Hide 时 VCL 认为"本来就不可见"而直接返回, 小窗永远关不掉。
  // 窗口带 WS_EX_NOACTIVATE, 所以这里的显示不会抢焦点。
  FAboutPopup.Visible := True;
  FAboutShown := True;
  FAboutMiss := 0;
  AboutTimer.Enabled := True;   // 启动巡检, 由它判断何时关闭
end;

procedure TMainForm.HideAboutPic;
begin
  AboutTimer.Enabled := False;
  if FAboutShown and (FAboutPopup <> nil) then
  begin
    FAboutHideRect := FAboutPopup.BoundsRect;     // 记住原位, 防止鼠标没动就立刻重新弹出
    FAboutHideGuard := True;
    FAboutPopup.Visible := False;                 // 同步 VCL 状态
    if IsWindowVisible(FAboutPopup.Handle) then   // 兜底: 确保窗口真的隐藏
      ShowWindow(FAboutPopup.Handle, SW_HIDE);
    FAboutShown := False;
  end;
end;

procedure TMainForm.OnAboutMouseEnter(Sender: TObject);
begin
  ShowAboutPic;   // 显示收款码, 同时启动巡检计时器
end;

procedure TMainForm.OnAboutMouseLeave(Sender: TObject);
begin
  AboutTimer.Enabled := True;   // 交给巡检判断: 鼠标确实不在链接/小窗内才关
end;

procedure TMainForm.OnAboutClick(Sender: TObject);
begin
  ShowAboutPic;
end;

procedure TMainForm.OnAboutPopupEnter(Sender: TObject);
begin
  AboutTimer.Enabled := True;   // 鼠标在小窗内 → 巡检会继续保持显示
end;

procedure TMainForm.OnAboutPopupLeave(Sender: TObject);
begin
  AboutTimer.Enabled := True;
end;

procedure TMainForm.OnAboutPopupClick(Sender: TObject);
begin
  HideAboutPic;   // 点击图片(或小窗) → 立即关闭
end;

procedure TMainForm.OnAboutPopupMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbLeft then HideAboutPic;   // 点击兜底: 图片没收到点击时也能关
end;

{ 巡检计时器: 鼠标只要还在「关于」链接或收款码小窗内就继续保持显示, 否则关闭。
  这里按鼠标位置判断, 而不是只靠 MouseEnter/MouseLeave 事件 —— 无窗口句柄的
  子控件(图片)在跨窗口切换时可能收不到 Leave, 只靠事件会出现"关不掉"的情况。 }
procedure TMainForm.OnAboutTimer(Sender: TObject);
var
  pt: TPoint;
  link: TRect;
begin
  AboutTimer.Enabled := False;
  if not (FAboutShown and (FAboutPopup <> nil)) then Exit;

  pt := Mouse.CursorPos;   // 屏幕坐标
  link := Rect(LblAbout.ClientOrigin.X, LblAbout.ClientOrigin.Y,
               LblAbout.ClientOrigin.X + LblAbout.Width,
               LblAbout.ClientOrigin.Y + LblAbout.Height);
  if PtInRect(link, pt) or PtInRect(FAboutPopup.BoundsRect, pt) then
  begin
    FAboutMiss := 0;
    AboutTimer.Enabled := True;   // 鼠标还在, 继续巡检
    Exit;
  end;
  Inc(FAboutMiss);
  // 连续 2 次(约 600ms)都不在链接/小窗内才关闭: 留出从链接移到图片上点击的时间
  if FAboutMiss < 2 then
  begin
    AboutTimer.Enabled := True;
    Exit;
  end;
  HideAboutPic;
end;

constructor TMainForm.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FEnumList := TList<TWinInfo>.Create;
  FProcNames := TDictionary<DWORD, string>.Create;
  FSortCol := -1;   // 默认按 Z 序
  FSortAsc := True;

  BuildUI;
  RefreshApps;

  Timer := TTimer.Create(Self);
  Timer.Interval := 2500;
  Timer.Enabled := ChkAuto.Checked;
  Timer.OnTimer := OnTimerTick;

  // “关于”收款码: 移出“关于”/图片后延时关闭(避免中途消失或残留)
  AboutTimer := TTimer.Create(Self);
  AboutTimer.Interval := 300;
  AboutTimer.Enabled := False;
  AboutTimer.OnTimer := OnAboutTimer;
end;

destructor TMainForm.Destroy;
begin
  Timer.Free;
  AboutTimer.Free;
  FAboutPopup.Free;
  FProcNames.Free;
  FEnumList.Free;
  inherited Destroy;
end;

end.
