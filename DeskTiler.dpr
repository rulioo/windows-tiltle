program DeskTiler;

{ DeskTiler - 桌面窗口平铺小工具
  * 列出当前所有打开的顶层应用窗口
  * 勾选需要整理的窗口后, 均匀平铺到桌面(所在显示器工作区)
  由 build.cmd 用命令行 dcc32 编译。 }

uses
  System.SysUtils,
  Vcl.Forms,
  uMain in 'uMain.pas';

{$R *.res}

var
  N: Integer;

begin
  Application.Initialize;
  Application.Title := 'DeskTiler';
  Application.MainFormOnTaskBar := True;
  Application.CreateForm(TMainForm, MainForm);

  // 命令行自测模式:
  //   DeskTiler.exe -list            把枚举到的窗口列表写到 winlist.txt 后退出
  //   DeskTiler.exe -tilefirst N     自动勾选列表前 N 个窗口并平铺后退出(用于自动化验证)
  //   DeskTiler.exe -tileapp exe     平铺“可执行文件名=exe”的全部窗口后退出(测试用)
  //   DeskTiler.exe -sorttest file   按 PID/应用 升降序自检排序, 结果写 file 后退出
  if (ParamCount >= 2) and SameText(ParamStr(1), '-sorttest') then
  begin
    MainForm.SortSelfTest(ParamStr(2));
    Exit;
  end;

  if (ParamCount >= 2) and SameText(ParamStr(1), '-tilefirst') then
  begin
    N := StrToIntDef(ParamStr(2), 2);
    MainForm.TileFirstN(N);
    Exit;
  end;

  if (ParamCount >= 2) and SameText(ParamStr(1), '-tileapp') then
  begin
    MainForm.TileByApp(ParamStr(2));
    Exit;
  end;

  if (ParamCount >= 1) and SameText(ParamStr(1), '-list') then
  begin
    MainForm.DumpList(ExtractFilePath(ParamStr(0)) + 'winlist.txt');
    Exit;
  end;

  Application.Run;
end.
