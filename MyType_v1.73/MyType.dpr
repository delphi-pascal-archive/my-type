program MyType;

uses
  Forms,
  Windows,
  SysUtils,
  s_frmMain in 's_frmMain.pas' {frmMain};

var
  hwndPrev : HWND;
 {$R *.res}
begin

  Application.Initialize;
  hwndPrev := FindWindow('TfrmMain',' MyType - Редактор записей');

  if hwndPrev >= 0  then //  При последней сборке поменять на <=
begin
  Application.CreateForm(TfrmMain, frmMain);
  Application.Title := 'MyType';
  Application.ShowMainForm:=false;

  Application.Run;
  end else
  begin
  SetForegroundWindow(hwndPrev);
  Application.Terminate;
    end;

   end.




