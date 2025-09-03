unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IOUtils,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  KeyUnit;

type
  TFormMain = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    MemoMessage: TMemo;
    Button4: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    function ProcStrToState(S: string): TModifierState;
    procedure ProcSetKeyController(S: string; aController: TKeyController);
    procedure ProcSetKeyControllerList(aList: TStringList; aApplicationList: TkeyApplicationList);
    procedure ProcSetKeyApplicationList(aFileName: string; aApplicationList: TkeyApplicationList);
    procedure ProcLoadKeyControlInfo;
  public
    { Public 宣言 }
  end;

var
  FormMain: TFormMain;

implementation

{$R *.dfm}

uses UnitUtils;

function StartKeyHook(Wnd: HWND): HHOOK; stdcall; external 'KEYHOOK.DLL';
procedure StopKeyHook; stdcall; external 'KEYHOOK.DLL';
procedure AddKeyControlInfo(aStream: TStream); stdcall; external 'KEYHOOK.DLL';
procedure ClearKeyControlInfo; stdcall; external 'KEYHOOK.DLL';
procedure ReadLogStream(const aList: TStringList); stdcall; external 'KEYHOOK.DLL';


function TFormMain.ProcStrToState(S: string): TModifierState;
begin
  Result := [];
  var FValue := StrToIntDef(S, 0);
  if (FValue and 1) <> 0 then Include(Result, tksShift);
  if (FValue and 2) <> 0 then Include(Result, tksCtrl);
  if (FValue and 4) <> 0 then Include(Result, tksNonConvert);
end;

procedure TFormMain.ProcSetKeyController(S: string; aController: TKeyController);
var
  FkeyComb: TKeyCombination;
begin
  //     0           ,1              ,2                  ,3              ,4               ,5
  // S : [ControlKey],[ModifierState],[OperationKeyCount],[OperationKey1],[ModifierState1],[],[]....
  // S の先頭文字が "//" の時はコメント行
  var FList := TStringList.Create;
  try
    FList.CommaText := S;
    FkeyComb.Key := StrToInt(FList[0]);
    FkeyComb.State := ProcStrToState(FList[1]);
    aController.ControlKey := FkeyComb;
    var FSize := StrToIntDef(FList[2], 0);
    if (FSize > 0) and (FList.Count >= 2*FSize+3) then begin
      for var Cnt := 0 to FSize-1 do begin
        FkeyComb.Key := StrToIntDef(FList[Cnt * 2 + 3], VK_ESCAPE);
        FkeyComb.State := ProcStrToState(FList[Cnt * 2 + 4]);
        aController.AddOperationKey(FkeyComb);
      end;
    end;
  finally
    FreeAndNil(FList);
  end;
end;

procedure TFormMain.ProcSetKeyControllerList(aList: TStringList; aApplicationList: TkeyApplicationList);
begin
  // aList[0] : 対象アプリケーションクラス名（文字列が "//" で始まってもコメント行ではない）
  //            0           ,1              ,2                  ,3              ,4               ,5
  // aList[1] : [ControlKey],[ModifierState],[OperationKeyCount],[OperationKey1],[ModifierState1],[],[]....
  // aList[1以上] の先頭文字が "//" の時はコメント行
  if aList.Count >= 2 then begin
    // 読み込んだファイルに対象アプリケーションクラス名行とコントロールキー行があるとき
    var FKeyCotrollerList := TKeyControllerList.Create;
    for var Cnt := 1 to aList.Count-1 do begin
      // コントロールキー行の読み込み
      var S := aList[Cnt];
      if (Copy(S, 1, 2) <> '//') and (SplitCount(S) >= 5) then begin
        // コメント行でなく、コントロール行の設定が最低の5項目以上あるとき
        var FKeyController := TKeyController.Create;
        // コントロールキーの読み込み
        ProcSetKeyController(aList[Cnt], FKeyController);
        FKeyCotrollerList.Add(FKeyController);
      end;
    end;
    // アプリケーションクラス名の読み込み
    aApplicationList.Add(aList[0], FKeyCotrollerList);
  end;
end;

procedure TFormMain.ProcLoadKeyControlInfo;
var
  FName: string;
begin
  var FKeyApplicationList := TkeyApplicationList.Create;
  try
    var FFileList := TDirectory.GetFiles(ExtractFilePath(Application.ExeName));
    for FName in FFileList do begin
      if TPath.GetExtension(FName) = '.txt' then
        // アプリケーションパス内のTEXTファイルからキー変換情報を読み込む
        ProcSetKeyApplicationList(FName, FKeyApplicationList);
    end;
    var FStream := TMemoryStream.Create;
    try
      FKeyApplicationList.SaveToStream(FStream);
      AddKeyControlInfo(FStream);
    finally
      FreeAndNil(FStream);
    end;
  finally
    FreeAndNil(FKeyApplicationList);
  end;
end;

procedure TFormMain.ProcSetKeyApplicationList(aFileName: string; aApplicationList: TkeyApplicationList);
begin
  var FList := TStringList.Create;
  try
    FList.LoadFromFile(aFileName);
    ProcSetKeyControllerList(FList, aApplicationList);
  finally
    FreeAndNil(FList);
  end;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  ReportMemoryLeaksOnShutdown := True;
  // フック関数の登録
  StartKeyHook(Handle);
  // キーコントロール情報の読み込み
  ProcLoadKeyControlInfo;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  StopKeyHook;
end;

procedure TFormMain.Button1Click(Sender: TObject);
begin
  // フック関数の登録
  StartKeyHook(Handle);
end;

procedure TFormMain.Button2Click(Sender: TObject);
begin
  StopKeyHook;
end;

procedure TFormMain.Button3Click(Sender: TObject);
begin
  // キーコントロール情報のクリア
  ClearKeyControlInfo;
  // キーコントロール情報の読み込み
  ProcLoadKeyControlInfo;
end;

procedure TFormMain.Button4Click(Sender: TObject);
begin
  MemoMessage.Clear;
  var FList := TStringList.Create;
  try
    ReadLogStream(FList);
    for var Cnt := 0 to FList.Count-1 do
      MemoMessage.Lines.Add(FList.Strings[Cnt]);
  finally
    FList.Free;
  end;
end;

end.
