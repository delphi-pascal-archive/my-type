unit s_FrmMain;

 { При последней сборке в файле MyType.dpr
  if hwndPrev >= 0 следует поменять на <= 0
  (Запрет на запуск второй копии программы) }

 interface

 uses
 Windows, Forms, SysUtils, ADODB, DBGrids, DB, DBCtrls, Messages, Classes,
 Graphics, Controls, Grids, StdCtrls, ExtCtrls, IniFiles, TrayIcon, Menus,
 ComCtrls, Registry, XPMan;

 type

 TfrmMain  = class(TForm)
 tbl:TADOTable;DtSrc:TDataSource;dbs:TDBGrid;dbm:TDBMemo;pnlDbGrid:TPanel;
 Panel1:TPanel;Panel2:TPanel;StBar:TStatusBar;TrPop:TPopupMenu;
 MainMenu1:TMainMenu;N1:TMenuItem;N2:TMenuItem;N3:TMenuItem;N4:TMenuItem;
 N5:TMenuItem;N7:TMenuItem;N8:TMenuItem;N9:TMenuItem;N10:TMenuItem;
 N11:TMenuItem;N12:TMenuItem;N13:TMenuItem;N14:TMenuItem;N15:TMenuItem;
 N16:TMenuItem;N17:TMenuItem;N18:TMenuItem;N19:TMenuItem;N20:TMenuItem;
 N22:TMenuItem;N23:TMenuItem;N25:TMenuItem;Splitter1:TSplitter;
 XPManif:TXPManifest;

 procedure Resz;
 procedure SetIco;
 procedure RegInit;
 function  Psk(ds: string): boolean;
 procedure Snt(ds: WideString; vs: integer);
 procedure FormCreate(Sender: TObject);
 procedure FormDestroy(Sender: TObject);
 procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
 procedure FormShow(Sender: TObject);
 procedure FormCanResize(Sender: TObject; var NewWidth, NewHeight: Integer;
           var Resize: Boolean);
 procedure dbsKeyPress(Sender: TObject; var Key: Char);
 procedure dbsMouseUp(Sender: TObject; Button: TMouseButton;
           Shift: TShiftState; X, Y: Integer);
 procedure tblAfterDelete(DataSet: TDataSet);
 procedure ShowHint(Sender: TObject);
 procedure N1Click(Sender: TObject);
 procedure N2Click(Sender: TObject);
 procedure N3Click(Sender: TObject);
 procedure N4Click(Sender: TObject);
 procedure N7Click(Sender: TObject);
 procedure N10Click(Sender: TObject);
 procedure N13Click(Sender: TObject);
 procedure N16Click(Sender: TObject);
 procedure N17Click(Sender: TObject);
 procedure N18Click(Sender: TObject);
 procedure N20Click(Sender: TObject);
 procedure N23Click(Sender: TObject);
 procedure N25Click(Sender: TObject);
 procedure TrPopChange(Sender: TObject; Source: TMenuItem; Rebuild: Boolean);
 procedure TrayIconOnDblClick(Sender: TObject);

  private

 procedure WMQueryEndSession(var Msg: TWMQueryEndSession);
  message WM_QUERYENDSESSION;

    { Private declarations }

 procedure DllMessage(var Msg: TMessage);
  message WM_USER + 2008;

  public

    { Public declarations }

      end;

  type

  TMyTypeHook = procedure(AppHandle: HWND; State: boolean);

{ ID раскладок клавиатуры}
 const
 KlEn = 67699721; KlRu = 68748313;

 var
 frmMain: TfrmMain; HandleDLL: THandle; TrayIcon: TTrayIcon;
 MsgLParam, YetEvent: Integer; CnClose: Boolean; kwEn,lwEn,kwRu,lwRu: string;

 implementation

 {$R *.dfm}

 function gkl: HKL;
begin
Result:= GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow, nil));
end;

 procedure CtrlV;
 begin
 keybd_event(VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY, 0);
 keybd_event(86, 0, KEYEVENTF_EXTENDEDKEY, 0);
 keybd_event(VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY or KEYEVENTF_KEYUP, 0);
 keybd_event(86, 0, KEYEVENTF_EXTENDEDKEY or KEYEVENTF_KEYUP, 0);Inc(YetEvent);
 end;

 procedure psl(ds: Word; vs: Boolean);
var b: array [0..1] of TInput;
begin
 FillChar(b,SizeOf(b),0);b[0].Itype:=INPUT_KEYBOARD;b[1].Itype:=INPUT_KEYBOARD;
 if vs then begin b[0].ki.wScan:=ds;b[0].ki.dwFlags:=4;b[1].ki.wScan:=ds;
 b[1].ki.dwFlags:=4 or 2; end else begin b[0].ki.wVk:=ds;b[1].ki.wVk:=ds;
 b[1].ki.dwFlags:=2; end; SendInput(Length(b),b[0],SizeOf(TInput));
end;

 function znc(ds: Byte; var vs, ks: Char): Boolean;
 const scn: array [0..43] of Byte =
 (192, 49, 50, 51, 52, 53, 54, 55, 56, 57, 48, 81, 87, 69, 82, 84, 89, 85, 73,
   79, 80,219,221, 65, 83, 68, 70, 71, 72, 74, 75, 76,186,222, 90, 88, 67, 86,
   66, 78, 77,188,190, 32);
 zn: array [0..1] of array[0..43] of Char =
 // En
 (('-','1','2','3','4','5','6','7','8','9','0','q','w','e','r','t','y','u','i',
  'o','p','-','-','a','s','d','f','g','h','j','k','l','-','-','z','x','c','v',
  'b','n','m','-','-',' '),
 // Ru
 ('ё','1','2','3','4','5','6','7','8','9','0','й','ц','у','к','е','н','г','ш',
  'щ','з','х','ъ','ф','ы','в','а','п','р','о','л','д','ж','э','я','ч','с','м',
  'и','т','ь','б','ю',' '));
 var  i: integer;
 begin
Result:=true;for i:=0 to 44 do begin if i=44 then begin Result:=False;Exit;end;
if ds=scn[i]then begin vs:=zn[0,i];ks:=zn[1,i];Exit;end;end;
 end;

 procedure TfrmMain.RegInit;
 var  h: TRegistry;
 begin
 h:= TRegistry.Create; with h do begin RootKey:= HKEY_LOCAL_MACHINE;
  OpenKey('\Software\Microsoft\Windows\CurrentVersion\Run', True);
  if n23.Checked then WriteString('MyType', Application.ExeName)
  else if ValueExists('MyType') then DeleteValue('MyType');CloseKey;Free;
 end;//with
 end;

 procedure TfrmMain.FormCanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
 begin
Resz;
 end;

 procedure TfrmMain.dbsMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
 begin
Resz;
 end;

 procedure TfrmMain.Resz;
 begin
dbs.Columns[0].Width:= dbs.ClientWidth - 1;
 end;

 procedure TfrmMain.WMQueryEndSession(var Msg: TWMQueryEndSession);
 begin
CnClose:= true;Msg.Result:= 1;
 end;

 procedure TfrmMain.FormCloseQuery(Sender:TObject;var CanClose:Boolean);
 begin
CanClose:= CnClose;Visible:= false;
 end;

 procedure TfrmMain.FormShow(Sender:TObject);
 begin
dbs.SetFocus;stbar.Panels[0].Text:='Всего: '+IntToStr(tbl.RecordCount)+' зап.';
 end;

 procedure TfrmMain.dbsKeyPress(Sender: TObject; var Key: Char);
 begin
if dbs.SelectedField.FieldName = 'Abridged' then case Key of
'a'..'z','а'..'я','0'..'9','ё',chr(VK_BACK),chr(VK_RETURN):Exit; else Key:= #0;
end; //case
 end;

function TfrmMain.Psk(ds: string): boolean;
begin
if Length(ds) > 0 then tbl.Locate('Abridged', ds, [loPartialKey]);
if ds=tbl.FieldValues['Abridged'] then Result:=true else Result:=false;
end;

 procedure TfrmMain.DllMessage(var Msg: TMessage);
 var b, r: char;
 begin
 if (Active) or (n12.Caption = 'Включить') then Exit;
if YetEvent>0 then begin Dec(YetEvent);Exit;end;if Msg.wParam=VK_BACK then begin
 Delete(kwEn,length(kwEn),length(kwEn));Delete(kwRu,length(kwRu),length(kwRu));
end else begin if znc(Msg.wParam,b,r)then if MsgLParam=msg.LParam then begin
if b=' ' then begin lwEn:= kwEn; lwRu:= kwRu; kwEn:= ''; kwRu:= ''; end;
if b='-' then kwEn:= '' else kwEn:= kwEn + b; kwRu:= kwRu + r; end else
  begin kwEn:= b;kwRu:= r; end else begin kwEn:= ''; kwRu:= '';end; end;
 if kwEn=' ' then begin kwEn:= '';kwRu:= '';end;MsgLParam:= msg.LParam;
if n4.Checked then begin if Msg.wParam=VK_SPACE then if n25.Checked then
if gkl = KLEn then begin if psk(lwEn) then snt(dbm.Text,length(lwEn));
end else begin if psk(lwRu) then snt(dbm.Text, length(lwRu)); end else
begin if psk(lwEn) then snt(dbm.Text, length(lwEn));if psk(lwRu) then
snt(dbm.Text,length(lwRu));end;end else if n25.Checked then if gkl=KLEn then
begin if psk(kwEn) then snt(dbm.Text, length(kwEn)); end else begin
if psk(kwRu) then snt(dbm.Text, length(kwRu)); end else begin
if psk(kwEn) then snt(dbm.Text, length(kwEn));
if psk(kwRu) then snt(dbm.Text, length(kwRu));end;
 end;

 procedure TfrmMain.Snt(ds:WideString;vs:integer);
 var i: integer;
 begin
if n4.Checked then Inc(vs); // Нужно убрать разделитель
for i:= 1 to vs do Psl(VK_BACK, false); // Стираем кейворд
if n7.Checked then//Используем клипборд.Недостатки - теряется содержимое буфера
begin
dbm.Text:= ds;dbm.SelectAll;dbm.CopyToClipboard;CtrlV; // Эмуляция
end else for i:= 1 to length(ds) do Psl(Word(ds[i]), true);
if n20.Checked then Psl(VK_SPACE, false); // Добавляем пробел
kwEn:= ''; kwRu:= '';
end;

 procedure TfrmMain.N1Click(Sender:TObject);
 begin
if n1.Caption= 'Редактор записей' then begin
Visible:= true; frmmain.SetFocus;end else Visible:= false;
 end;

procedure TfrmMain.N2Click(Sender: TObject);
begin
if n2.Checked then begin
n2.Checked:= false;dbm.WordWrap:= false;dbm.ScrollBars:= ssBoth;
end else begin n2.Checked:=true;dbm.ScrollBars:=ssVertical;dbm.WordWrap:=true;
end;
end;

 procedure TfrmMain.N3Click(Sender: TObject);
 begin
 CnClose:= true; Close;
 end;

 procedure TfrmMain.N4Click(Sender: TObject);
 begin
 n4.Checked:= not n4.Checked;
  end;

 procedure TfrmMain.N7Click(Sender: TObject);
begin
 n7.Checked:= not n7.Checked;
end;

procedure TfrmMain.N10Click(Sender:TObject);
 begin
 if n10.Caption= 'Отключить'then
    n10.Caption:= 'Включить' else n10.Caption:= 'Отключить'; SetIco;
end;

  procedure TfrmMain.N13Click(Sender:TObject);
 begin
visible:= false;
 end;

 procedure TfrmMain.N16Click(Sender:TObject);
 begin
dbs.EditorMode:= true;
 end;

 procedure TfrmMain.N17Click(Sender:TObject);
 begin
tbl.Insert;
 end;

 procedure TfrmMain.N18Click(Sender:TObject);
 begin
tbl.Delete;
 end;

 procedure TfrmMain.N20Click(Sender:TObject);
 begin
 n20.Checked:= not n20.Checked;
 end;

 procedure TfrmMain.N23Click(Sender:TObject);
 begin
 n23.Checked:= not n23.Checked; RegInit;
 end;

 procedure TfrmMain.N25Click(Sender: TObject);
begin
n25.Checked:= not n25.Checked;
end;

procedure TfrmMain.TrPopChange(Sender: TObject; Source: TMenuItem;
Rebuild: Boolean);
begin
if visible then n1.Caption:='Закрыть редактор'
else n1.Caption:='Редактор записей';
end;

 {$R ico.res}
 procedure TfrmMain.SetIco;
 var d: TIcon;
 begin
d := TIcon.Create; try if n10.Caption = 'Отключить' then
begin n12.Caption:= 'Отключить'; d.Handle:= Application.Icon.Handle;end else
begin n12.Caption:= 'Включить'; d.Handle:= LoadIcon(hInstance, 'ZZZICON');end;
TrayIcon.Icon:= d;finally d.Free;end;
 end;

 procedure TfrmMain.ShowHint(Sender: TObject);
 begin
if Length(Application.Hint) > 0 then begin StBar.SimplePanel:= true;
StBar.SimpleText:= Application.Hint;end else StBar.SimplePanel:= false;
 end;

 procedure TfrmMain.tblAfterDelete(DataSet:TDataSet);
 begin
stbar.Panels[0].Text:= 'Всего: ' + inttostr(tbl.RecordCount) + ' зап.';
 end;

 procedure TfrmMain.FormCreate(Sender: TObject);
 var  Text: TIniFile; MyTypeHook: TMyTypeHook; WS: string;
 begin
TrayIcon:= TTrayIcon.Create(TrayIcon);
with TrayIcon do begin
ToolTip:= 'MyType';Icon:= Application.Icon;Active:= True;PopupMenu:= TrPop;
OnDblClick:= TrayIconOnDblClick; end;
 HandleDLL:= LoadLibrary('MyHook.dll');Application.OnHint:=ShowHint;
 @MyTypeHook:=GetProcAddress(HandleDLL,'MyTypeHook');
 if @MyTypeHook=nil then Exit;MyTypeHook(Handle,true);kwEn:= ''; kwRu:= '';
 Text:= TIniFile.Create(extractFileDir(ParamSTR(0)) + '\mti.ini'); try
 with Text do  begin
Top         := ReadInteger('frmMain', 'Top', 450);
Left        := ReadInteger('frmMain', 'Left', 275);
Height      := ReadInteger('frmMain', 'Height', 250);
Width       := ReadInteger('frmMain', 'Width', 490);
n20.Checked := ReadBool   ('frmMain', 'n20ch', true);
n2.Checked  := ReadBool   ('frmMain', 'n2ch', true);

if n2.Checked then
begin dbm.ScrollBars:= ssVertical;dbm.WordWrap:= true;end else
begin dbm.WordWrap:= false;dbm.ScrollBars:= ssBoth;end;

n4.Checked   := ReadBool    ('frmMain', 'n4ch',false);
n25.Checked  := ReadBool    ('frmMain', 'n25ch',false);
n7.Checked   := ReadBool    ('frmMain', 'n7ch',false);
n23.Checked  := ReadBool    ('frmMain', 'n23ch',false);
n10.Caption  := ReadString  ('frmMain', 'dis','Отключить'); Setico;
Panel1.Width := ReadInteger ('frmMain', 'CollWidth',140);
WS           := ReadString  ('frmMain', 'WindowState','wsNormal');
end; //with
if WS = 'wsMaximized' then WindowState:= wsMaximized
else WindowState:= wsNormal;

tbl.ConnectionString:= 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='
+ extractFileDir(ParamSTR(0))+'\KeyWord.mdb;'+
'Mode=ReadWrite;Persist Security Info=False';
tbl.Active:= true;Resz;CnClose:= false;YetEvent:= 0; finally Text.Free;end;
 end;

 procedure TfrmMain.FormDestroy(Sender:TObject);
 var  Text: TIniFile; WS: string; MyTypeHook: TMyTypeHook;
 begin
   TrayIcon.Destroy;  @MyTypeHook:= GetProcAddress(HandleDll, 'MyTypeHook');
if @MyTypeHook = nil then Exit; MyTypeHook(Handle, false);
 FreeLibrary(HandleDLL);
 { For some reason in Win XP you need to call FreeLibrary twice
  maybe because you get two functions from the DLL }
 FreeLibrary(HandleDLL);
 Text:= TIniFile.Create(extractFileDir(ParamSTR(0)) + '\mti.ini'); try
 with Text do begin  if WindowState = wsNormal then begin
WriteInteger('frmMain', 'Top',         Top);
WriteInteger('frmMain', 'Left',        Left);
WriteInteger('frmMain', 'Height',      Height);
WriteInteger('frmMain', 'Width',       Width);
WriteInteger('frmMain', 'CollWidth',   Panel1.Width);
WriteBool   ('frmMain', 'n20ch',       n20.Checked);
WriteBool   ('frmMain', 'n7ch',        n7.Checked);
WriteBool   ('frmMain', 'n4ch',        n4.Checked);
WriteBool   ('frmMain', 'n2ch',        n2.Checked);
WriteBool   ('frmMain', 'n25ch',       n25.Checked);
WriteBool   ('frmMain', 'n23ch',       n23.Checked);RegInit;
WriteString ('frmMain', 'dis',         n10.Caption);

WS:= 'wsNormal'; end else
WS:= 'wsMaximized'; WriteString ('frmMain', 'WindowState', WS); end; //with
tbl.UpdateBatch; finally Text.Free;end;
 end;

procedure TfrmMain.TrayIconOnDblClick(Sender: TObject);
begin
Visible:= not Visible;
end;

end.
