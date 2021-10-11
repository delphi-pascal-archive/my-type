{ Library Code for a Key Hook DLL

This code post's the Char from the DLL in way,
as full a Char message or a DLLMessage WM_USER

Взято с сайта: http://www.swissdelphicenter.ch}

library MyHook;

uses Windows,Messages;

type PHookRec = ^THookRec;
     THookRec = record
     AppHnd   : integer;
     end;

var  Hooked:Boolean;
     hK, hM, hA: HWND;
     Hr: PHookRec;

function KeyHookFunc(Code, VirtualKey, KeyStroke: integer): LRESULT; stdcall;
//var  Kv: TKeyBoardState;  b: array[0..1] of char;
begin
Result:=0; if Code=HC_NOREMOVE then Exit;
Result:=CallNextHookEx(hK,Code,VirtualKey,KeyStroke);

 {I moved the CallNextHookEx up here but if you want to block
 or change any keys then move it back down}

 if Code<0 then Exit;if Code=HC_ACTION then
begin
if((KeyStroke and (1 shl 30))<>0) then if not IsWindow(hA) then
begin

 {I moved the OpenFileMapping up here so it would not be opened
 unless the app the DLL is attatched to gets some Key messages}

 hM:=OpenFileMapping(FILE_MAP_WRITE,False,'MyHookMap');
Hr:=MapViewOfFile(hM,FILE_MAP_WRITE,0,0,0);
if Hr<>nil then hA:=Hr.AppHnd;
end;
if ((KeyStroke and (1 shl 30))<>0) then
//begin GetKeyboardState(Kv);
//if ToAscii(VirtualKey,KeyStroke,Kv,b,0)=1 then

{I included way to get the Char, as post a WM_USER message to the program}

 SendMessage(hA, WM_USER+2008, VirtualKey, GetFocus);
// end;
 end;
end;

procedure MyTypeHook(AppHandle: HWND; State: boolean); export;
begin
if State then begin if Hooked then Exit;

 {You need to use a mapped file, because this DLL attatches to every app
that gets windows messages when it's hooked, and you can't get info except
through a Globally avaiable Mapped file :( }

hK:=SetWindowsHookEx(WH_KEYBOARD,KeyHookFunc,hInstance,0);
if hK>0 then begin
hM:=CreateFileMapping($FFFFFFFF, // $FFFFFFFF gets a page memory file
nil,                             // no security attributes
PAGE_READWRITE,                  // read/write access
0,                               // size: high 32-bits
SizeOf(THookRec),                // size: low 32-bits
'MyHookMap');                    // name of map object

Hooked:=true;
Hr:=MapViewOfFile(hM,FILE_MAP_WRITE,0,0,0);hA:=AppHandle;

{Set the App handles to the mapped file}

Hr.AppHnd:=AppHandle; end; end else begin
 if Hr<>nil then begin UnmapViewOfFile(Hr); CloseHandle(hM); Hr:=nil; end;
 if Hooked then UnhookWindowsHookEx(hK); Hooked:=false; end;
 end;

procedure EntryProc(dwReason: DWORD);
begin
if (dwReason=Dll_Process_Detach) then begin
if Hr<>nil then begin UnmapViewOfFile(Hr); CloseHandle(hM); end;
UnhookWindowsHookEx(hK); end;
end;

exports MyTypeHook;

begin
Hr:=nil;Hooked:=False;hK:=0;DLLProc:=@EntryProc;
EntryProc(Dll_Process_Attach);
end.
