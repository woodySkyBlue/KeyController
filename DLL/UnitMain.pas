unit UnitMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.Classes, System.SysUtils,
  plCommMem, KeyUnit, UnitUtils;

//外部のアプリケーションから使用する関数
// 260_グローバルフック ( キー・マウス )
//   http://mrxray.on.coocan.jp/Delphi/plSamples/260_HookKeyMouseEvent.htm
// 共有メモリを使用して異なるプロセス間でデータを共有
//   http://mrxray.on.coocan.jp/Delphi/Others/CommonMemory.htm#top
function StartKeyHook(Wnd: HWND): Boolean; stdcall;
procedure StopKeyHook; stdcall;
procedure AddKeyControlInfo(aStream: TStream); stdcall;
procedure ClearKeyControlInfo; stdcall;
procedure ReadLogStream(const aList: TStringList); stdcall

implementation

type
  //TkeyApplicationListクラスサイズの共有メモリ構造体
  TMapBlock = record
    Size: Cardinal;
  end;

  //共有メモリの内容の構造体
  PHookInfo  = ^THookInfo;
  THookInfo  = record
    HookKeyHandle   : HHOOK;
    HookMouseHandle : HHOOK;
    HostWnd         : HWND;
  end;

var
  //メモリマップドファイルのハンドル
  hMapFile : THandle;
  // Ctrl/shift/無変換キー状態
  FModifierState: TModifierState;
  FKeyApplicationList: TkeyApplicationList;
  FList: TStringList;

const
  BLOCK_SIZE: Cardinal = SizeOf(TMapBlock);
  //メモリマップドファイル名
  MAP_KEY_NAME = 'plMouseKeyHookDLL';
  MAP_LIST_NAME = 'KeyListHookDLL';
  MAP_SIZE = 4096 * 10000;


// AValueパラメータ値を4ビット区切りの2進数文字列に変換
function ProcIntToBinaryGrouped(AValue: NativeInt): string;
begin
  Result := '';
  for var Cnt := 31 downto 0 do begin
    Result := Result + Char(Ord((AValue shr Cnt) and 1) + Ord('0'));
    if (Cnt mod 4 = 0) and (Cnt <> 0) then
      Result := Result + ' ';
  end;
end;

procedure WriteDebugLogStream(S: string);
begin
  FList.Add(Format('%s %s', [FormatDateTime('nn:ss.zzz', Now), S]));
end;

procedure ReadLogStream(const aList: TStringList);
begin
  //aList.Assign(FList);
  aList.Clear;
  for var Cnt := 0 to FList.Count-1 do
    aList.Add(FList.Strings[Cnt]);
  FList.Clear;
end;

// 共有メモリアクセス準備
function MapFileMemory(var hMap: THandle; var pMap: pointer; aName: LPCWSTR): Integer;
//  成功すると0を返す.失敗すると負数を返す.
//  負の値で処理を分岐できるようになっているが,このコードでは未使用.
begin
  //メモリマップドファイルを開く
  hMap := OpenFileMapping(FILE_MAP_ALL_ACCESS, False, aName);
  if hMap = 0 then begin
    Result := -1;
    exit;
  end;
  //メモリマップドファイルの割り当て
  pMap := MapViewOfFile(hMap, FILE_MAP_ALL_ACCESS, 0, 0, 0);
  if pMap = nil then begin
    Result := -2;
    CloseHandle(hMap);
    exit;
  end;
  Result := 0;
end;

// 共有メモリアクセス終了
procedure UnMapFileMemory(hMap: THandle; pMap: Pointer);
begin
  if pMap <> nil then UnmapViewOfFile(pMap);
  if hMap <> 0 then CloseHandle(hMap);
end;

// KeyApplicationListクラスの内容をメモリマップドファイルに保存する
procedure ProcSaveKeyAppListToMemMappedFile;
var
  FMapBlock: TMapBlock;
begin
  var FStream := TMemoryStream.Create;
  try
    FKeyApplicationList.SaveToStream(FStream);
    FMapBlock.Size := FStream.Size;
    TplCommMem.Write(MAP_LIST_NAME, @FMapBlock, 0, BLOCK_SIZE);
    TplCommMem.Write(MAP_LIST_NAME, FStream.Memory, BLOCK_SIZE, FStream.Size);
  finally
    FreeAndNil(FStream);
  end;
end;

// メモリマップドファイルの内容をKeyApplicationListクラスに読み込む
procedure ProcLoadKeyAppListFromMemMappedFile;
var
  FMapBlock: TMapBlock;
begin
  var FStream := TMemoryStream.Create;
  try
    TplCommMem.Read(MAP_LIST_NAME, @FMapBlock, 0, BLOCK_SIZE);
    FStream.Size := FMapBlock.Size;
    TplCommMem.Read(MAP_LIST_NAME, FStream.Memory, BLOCK_SIZE, FStream.Size);
    FStream.Position := 0;
    FKeyApplicationList.Clear;
    FKeyApplicationList.LoadFromStream(FStream);
  finally
    FreeAndNil(FStream);
  end;
end;

// フック本体
// KeyHookProcは、StartKeyHook関数内のSetWindowsHookEx関数を使ってWH_KEYBOARDフックを設定すると、
// Windowsがキーボードイベント（キーDown,Up）を検知すると、コールバックとして自動的に呼び出される
function KeyHookProc(nCode:integer; wPar: WPARAM; lPar: LPARAM): LRESULT; stdcall;
var
  LpMap: Pointer;
  LMapWnd: THANDLE;
begin
  Result := 0;
  if MapFileMemory(LMapWnd, LpMap, MAP_KEY_NAME) = 0 then begin
    ProcLoadKeyAppListFromMemMappedFile;
    var FBinary := ProcIntToBinaryGrouped(lPar);
    //WriteDebugLogFile('C:\temp\log.txt', Format('%s %3d %11d %s', [FormatDateTime('hh:ss.zzz', Now), wPar, lPar, FBinary]));
    WriteDebugLogStream(Format('%3d %11d %s', [wPar, lPar, FBinary]));
    Result := CallNextHookEx(pHookInfo(LpMap)^.HookKeyHandle, nCode, wPar, lPar);
    if FKeyApplicationList.KeyConvert(nCode, wPar, lPar) then begin
      // 押されたキーがコントロールキーとして登録されている時はキーの表示を禁止する
      Result := -1;
    end;
    ProcSaveKeyAppListToMemMappedFile;
    UnMapFileMemory(LMapWnd, LpMap);
  end;
end;

// フック関数の登録
function StartKeyHook(Wnd: HWND): Boolean; stdcall;
var
  LpMap   : Pointer;
  LMapWnd : THandle;
begin
  Result := False;
  //メモリマップドファイル使用準備
  MapFileMemory(LMapWnd, LpMap, MAP_KEY_NAME);
  if LpMap <> nil then begin
    //フック情報構造体初期化とフック関数の登録
    pHookInfo(LpMap)^.HostWnd := Wnd;
    pHookInfo(LpMap)^.HookKeyHandle := SetWindowsHookEx(WH_KEYBOARD, Addr(KeyHookProc), hInstance, 0);
    if pHookInfo(LpMap)^.HookKeyHandle > 0 then begin
      Result := True;
    end;
    //メモリマップドファイル使用終了処理
    UnMapFileMemory(LMapWnd, LpMap);
  end;
end;

// フックの解除
procedure StopKeyHook; stdcall;
var
  LpMap   : Pointer;
  LMapWnd : THandle;
begin
  //メモリマップドファイル使用準備
  MapFileMemory(LMapWnd, LpMap, MAP_KEY_NAME);
  if LpMap = nil then begin
    LMapWnd := 0;
    exit;
  end;
  //フック解除
  if pHookInfo(LpMap)^.HookMouseHandle > 0 then begin
    UnhookWindowsHookEx(pHookInfo(LpMap)^.HookMouseHandle);
  end;
  if pHookInfo(LpMap)^.HookKeyHandle > 0 then begin
    UnhookWindowsHookEx(pHookInfo(LpMap)^.HookKeyHandle);
  end;
  //メモリマップドファイル使用終了処理
  UnMapFileMemory(LMapWnd, LpMap);
end;

procedure AddKeyControlInfo(aStream: TStream); stdcall;
begin
  aStream.Position := 0;
  FKeyApplicationList.LoadFromStream(aStream);
  // KeyApplicationListクラスの内容をメモリマップドファイルに保存する
  ProcSaveKeyAppListToMemMappedFile;
end;

procedure ClearKeyControlInfo; stdcall;
begin
  FKeyApplicationList.Clear;
end;

initialization
begin
  ReportMemoryLeaksOnShutdown := True;
  FModifierState := [];
  hMapFile := CreateFileMapping(High(NativeUInt), nil, PAGE_READWRITE, 0, SizeOf(THookInfo), MAP_KEY_NAME);
  TplCommMem.Create(MAP_SIZE, MAP_LIST_NAME);
  FKeyApplicationList := TkeyApplicationList.Create;
  FList := TStringList.Create;
end;

finalization
begin
  TplCommMem.Free(MAP_LIST_NAME);
  if hMapFile <> 0 then CloseHandle(hMapFile);
  if Assigned(FKeyApplicationList) then FKeyApplicationList.Free;
  if Assigned(FList) then FList.Free;
end;

end.
