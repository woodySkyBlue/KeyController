//*********************************************************************************************************************
//
// SIN-LABO/A.NAKASHIMA Utils Library
//
// 2012.01.08 Ver3.00.00  Delphi2010に対応
// 2013.02.27 Ver3.00.01  CreateFile関数にエラーチェックを追加
// 2013.07.25 Ver3.00.02  Streamデータ文字列読み書きルーチンをUnicodeに対応
// 2014.02.05 Ver3.00.03  SplitCount関数のパラメータに空白文字列を指定したときの
//                        戻り値を"1"→"0"に変更
// 2014.03.12 Ver3.01.00  TUnitDataList高度レコード型を追加
// 2014.04.23 Ver3.02.00  CheckFileNameメソッドの廃止
// 2014.11.27 Ver5.00.00  DelphiXE7に対応
// 2015.04.09 Ver5.00.01  PopnStdDevExメソッドのパラメータにSingle,Extended型も対応できるように拡張
// 2015.06.30 Ver5.01.00  子フォームの中央にダイアログボックスを表示するCenterChildDlg関数追加
// 2016.04.26 Ver5.02.00  プロセス起動中確認メソッドの追加
// 2016.09.15 Ver5.03.00  Split, SplitCount関数の拡張
// 2016.10.04 Ver5.03.01  ファイルフォーマット変更：ANSI→UTF8
// 2017.01.27 Ver5.04.00  ReadStreamAnsiStringメソッドの追加
// 2017.04.04 Ver6.00.00  Delphi10対応
//                        TUnitDataListレコード型にLineFeedTextメソッド追加
// 2017.06.28 Ver6.01.00  GetActiveProcessId, WinExecAndWait32V2, WinExit関数追加
// 2017.07.25 Ver6.02.00  TUnitDataList.Addメソッド拡張
// 2017.08.10 Ver6.03.00  TUnitDataListレコード型にTextメソッド追加
//                        TUnitDataList.Addメソッド拡張
// 2017.09.20 Ver6.03.01  CODE_PAGE_SHIFTJIS定数追加
// 2017.10.10 Ver6.03.02  PopnStdDevExメソッド処理不具合修正（配列数=1のとき計算結果NG）
// 2017.11.02 Ver6.04.00  実行ファイルのバージョン番号の取得関数(GetStrVersion)追加
// 2017.11.15 Ver6.05.00  SplitDelete関数追加
// 2017.12.08 Ver6.06.00  TUnitDataList.TabTextメソッド追加
// 2018.05.25 Ver6.07.00  TUnitDataList.Addメソッド拡張
// 2018.06.20 Ver6.08.00  CheckBeforeSaveFile関数追加
// 2018.07.03 Ver6.09.00  TUnitDataList.addメソッド機能追加（Digitパラメータ）
// 2018.10.30 Ver6.10.00  CheckBeforeSaveFolder関数追加（将来CheckBeforeSaveFile関数を修正すること）
// 2018.12.12 Ver6.11.00  マップ補間レコード(TLerpMapManager)型追加
// 2019.04.11 Ver6.12.00  RetryFileExist, FormatEx関数追加
// 2019.11.05 Ver7.00.00  Delphi10.3対応
// 2019.12.09 Ver7.01.00  IncludeTrailingSlashDelimiter,ExcludeTrailingSlashDelimiter関数追加
//                        TUnitDataListレコードにSpaceTextメソッド追加
// 2020.03.26 Ver7.01.01  CheckBeforeSaveFile,CheckBeforeSaveFolder関数でフォルダ作成エラー発生時に
//                        エラーメッセージダイアログを表示しないように修正
// 2020.04.20 Ver7.02.00  フォームの表示位置設定
//                        CheckDirectoryExists関数追加
// 2020.06.05 Ver7.03.00  コンポーネントを交換する関数（ChangeComponent）追加
//
//*********************************************************************************************************************

unit UnitUtils;

interface

// "暗黙的文字列キャスト(W1057)"をエラーに昇格
{$WARN IMPLICIT_STRING_CAST ERROR}
// "暗黙的文字列キャストによるデータ喪失の可能性(W1058)"をエラーに昇格
{$WARN IMPLICIT_STRING_CAST_LOSS ERROR}
// "set 式で WideChar がバイト文字に変換されました(W1050)"をエラーに昇格
{$WARN WIDECHAR_REDUCED ERROR}

uses
  Winapi.Windows, Winapi.Messages, Winapi.TlHelp32,
  System.Classes, System.SysUtils, System.Types, System.Math, System.Ioutils, System.IniFiles, System.TypInfo,
  Vcl.Forms, Vcl.Dialogs, Vcl.Controls, Vcl.Graphics, Vcl.StdCtrls, Privilege;

type
  //ネットワークフォルダが存在するか確認するためのスレッドクラス
  TDirectoryExistsThread = class(TThread)
  private
    FDirectoryName: string;
    FResult: Integer;
  protected
    procedure Execute; override;
  public
    property DirectoryName: string read FDirectoryName write FDirectoryName;
    property Result: Integer read FResult write FResult;
  end;

  TMessageLanguage = (mlJapanese, mlEnglish);

  TVerResourceKey = (
        vrComments,         // コメント
        vrCompanyName,      // 会社名
        vrFileDescription,  // 説明
        vrFileVersion,      // ファイルバージョン
        vrInternalName,     // 内部名
        vrLegalCopyright,   // 著作権
        vrLegalTrademarks,  // 商標
        vrOriginalFilename, // 正式ファイル名
        vrPrivateBuild,     // プライベートビルド情報
        vrProductName,      // 製品名
        vrProductVersion,   // 製品バージョン
        vrSpecialBuild);    // スペシャルビルド情報

  TUnitDataList = record
  private
    FItems: array of string;
    function GetCount: Integer; inline;
    function GetItem(Index: Integer): string;
    function GetStrings(Index: Integer): string;
    function ProcAddFloat(const Value: Extended; Digits: Integer; const S1, S2: string): Integer;
    procedure SetStrings(Index: Integer; const Value: string);
  public
    procedure Clear;
    function Add(const Value: string): Integer; overload;
    function Add(const Value: Single; Digits: Integer): Integer; overload;
    function Add(const S1: string; const Value: Single; const Digits: Integer; const S2: string): Integer; overload;
    function Add(const Value: Double; Digits: Integer): Integer; overload;
    function Add(const S1: string; const Value: Double; const Digits: Integer; const S2: string): Integer; overload;
    function Add(const Value: Extended; Digits: Integer): Integer; overload;
    function Add(const S1: string; const Value: Extended; const Digits: Integer; const S2: string): Integer; overload;
    function Add(const Value: Word): Integer; overload;
    function Add(const S1: string; const Value: Word; const S2: string): Integer; overload;
    function Add(const Value: Integer): Integer; overload;
    function Add(const S1: string; const Value: Integer; const S2: string): Integer; overload;
    function Add(const FLG: Boolean): Integer; overload;
    function Add(const S1: string; const FLG: Boolean; const S2: string): Integer; overload;
    function CommaText: string;
    function TabText: string;
    function SpaceText: string;
    function LineFeedText: string;
    function Text: string;
    property Count: Integer read GetCount;
    property Strings[Index: Integer]: string read GetStrings write SetStrings; default;
  end;

  TMapData = record
    X: Integer;
    Y: Integer;
  end;

  TTableData = record
    Axis: Extended;
    Value: Extended;
  end;
  TTableDatas = array of TTableData;

  TLerpMapManager = record
    //マップサイズが(1×1)でもOK (1×0)はNG
  private
    FXAxisArray: array of Extended;
    FYAxisArray: array of Extended;
    FMapArray: array of array of Extended;
    function GetAxisSize: TMapData;
    function GetXAxis(Index: Integer): Extended;
    function GetYAxis(Index: Integer): Extended;
    procedure SetAxisSize(const Value: TMapData);
    procedure SetXAxis(Index: Integer; const Value: Extended);
    procedure SetYAxis(Index: Integer; const Value: Extended);
    function _Lerp(AValue, X, X1, Y, Y1: Extended): Extended;
    function _Index(AValue: Extended; ATables: TTableDatas): Integer;
    function _Linear(AAxisValue: Extended; ATables: TTableDatas): Extended;
    function GetSize: TMapData;
    function GetMap(XIndex, YIndex: Integer): Extended;
    procedure SetMap(XIndex, YIndex: Integer; const Value: Extended);
  public
    function GetMapValue(XValue, YValue: Extended): Extended;
    property Size: TMapData read GetSize;
    property XAxis[Index: Integer]: Extended read GetXAxis write SetXAxis;
    property YAxis[Index: Integer]: Extended read GetYAxis write SetYAxis;
    property AxisSize: TMapData read GetAxisSize write SetAxisSize;
    property Map[XIndex, YIndex: Integer]: Extended read GetMap write SetMap;
  end;

const
  KeyWordStr: array [TVerResourceKey] of String = (
        'Comments',
        'CompanyName',
        'FileDescription',
        'FileVersion',
        'InternalName',
        'LegalCopyright',
        'LegalTrademarks',
        'OriginalFilename',
        'PrivateBuild',
        'ProductName',
        'ProductVersion',
        'SpecialBuild');

  //Split,SplitCountメソッド用セパレート定数
  C_COMMA = ','; //カンマ
  C_SPACE = ' '; //スペース
  C_DOT = '.';   //ドット
  C_SLASH = '/'; //スラッシュ
  CODE_PAGE_SHIFTJIS = 932; //Shift-jis送受信用コードページ定数


//ミリ秒単位の時刻の取得 FDELPHI#16-255
function Now_ms: TDateTime;
//Delimiter文字で区切られた文字列からゼロベース配列の指定番号文字を求める
function Split(S: string; Pos: Integer): string; overload;
function Split(S: string; Delimiter: char; Pos: Integer): string; overload;
function Split(List: TStrings; Index, Pos: Integer): string; overload;
function Split(List: TStrings; Delimiter: char; Index, Pos: Integer): string; overload;
//Delimiter文字で区切られた文字列の分割数を求める
function SplitCount(S: string; Delimiter: char = C_COMMA): Integer; overload;
function SplitCount(List: TStrings; Index: Integer; Delimiter: char = C_COMMA): Integer; overload;
//CommaText文字列の先頭文字列を削除する関数
function CommaDelete(S: string): string;
//Delimiter文字で区切られた文字列の先頭文字列を削除する関数
function SplitDelete(S: string; Delimiter: char = C_COMMA): string; overload;
//Delimiter文字で区切られた文字列のPos分を削除する関数
function SplitDelete(S: string; Pos: Integer; Delimiter: char = C_COMMA): string; overload;
//MessageDlg,ShowMessageBoxを指定フォームの中央に表示
function CenterDlg(Form: TForm; AMessage, TitleCaption: String;
                   DlgType: TMsgDlgType; Buttons: TMsgDlgButtons): TModalResult;
//MessageDlg,ShowMessageBoxを指定子フォームの中央に表示
function CenterChildDlg(MainForm, ChildForm: TForm; AMessage, TitleCaption: String;
                        DlgType: TMsgDlgType; Buttons: TMsgDlgButtons): TModalResult;
//パラメータ名の空ファイルを作成する関数（エラーチェックを追加 2013/02/27）
function CreateFile(FileName: string): Boolean;
//ファイル名(パスは含まない)として使用できるか確認する関数→IoUtils.TPath.HasValidFileNameCharsメソッドに置き換える
//function CheckFileName(const FileName: string): Boolean;
//ファイル名(パスを含む)として使用できるか確認する関数
function CheckPathFileName(const FileName: string): Boolean;
//実行ファイルのバージョン番号の取得
function GetStrVersion: string;
//すべてのバージョン情報を取得する関数
function GetVersionInfo(KeyWord: TVerResourceKey): string;
//デバッグ用ログファイル保存共通ルーチン
procedure WriteDebugLogFile(FileName: TFileName; SData: string);
//FolderNameに指定されたネットワークフォルダが存在するか確認する
function NetDirectoryExists(const FolderName: String; TimeOut: Cardinal = 3000): Boolean;
//Mathユニット[PopnStdDev]拡張関数 Ver1.10
function PopnStdDevEx(const Data: array of Single): Extended; overload;
function PopnStdDevEx(const Data: array of Double): Extended; overload;
function PopnStdDevEx(const Data: array of Extended): Extended; overload;
//表示位置を設定出来るInputQuery
function InputQueryEx(const ACaption, APrompt: string; x, y: Integer;  var Value: string): Boolean;
// テキスト回転描画出来るようにする文字描画拡張ルーチン
procedure TextOutEx(Canvas: Tcanvas; X, Y, Z: Integer; S: String);
//Streamデータから文字列を読み出す共通ルーチン
function ReadStreamString(Stream: TStream): string;
//Streamデータから文字列(Shift-JIS)を読み出す共通ルーチン
function ReadStreamAnsiString(Stream: TStream): string;
//Streamデータに文字列Sを書き込む共通ルーチン
procedure WriteStreamString(Stream: TStream; S: string);
//文字列中の改行コードをスペース文字列に変換
function ReplaceLineFeedToSpace(S: string): string;
//指定ファイルが起動中のプロセスか確認する関数
function CheckActiveProcess(aExeName:string): Boolean;
//指定ファイルのプロセスIDを取得する関数
function GetActiveProcessId(aExeName: string): Cardinal;
//指定ファイルを実行する関数（アプリの起動完了まで待つ）
function WinExecAndWait32V2(aExeName: string; Visibility: Integer): DWORD;
//指定ファイルを終了する関数（強制終了処理あり）
procedure WinExit(aExeName: string);
//ファイル保存前チェック関数（AFileName=パス付ファイル名 FollowPath=True：フォルダが存在しないときは自動作成）
function CheckBeforeSaveFile(AFileName: string; FollowPath: Boolean = False): Boolean;
//ファイル保存前チェック関数（AFileName=パス付ファイル名 FollowPath=True：フォルダが存在しないときは自動作成）
function CheckBeforeSaveFolder(AFileName: string; FollowPath: Boolean = False): Boolean;
//指定フォルダの存在確認（AFileName=パス付ファイル名 FollowPath=True：フォルダが存在しないときは自動作成）
function CheckDirectoryExists(const ADirectory: string; AFollowPath: Boolean = False): Boolean;
//ファイルの存在確認
function RetryFileExists(AFileName: string; ARetry: Integer = 5): Boolean;
//フォーマット拡張ルーチン
function FormatEx(const Format: string; const Args: array of const; LanguageCode: Integer = 0): string;
//IsPathDelimiterの"/"バージョン
function IsSlashDelimiter(const S: string; Index: Integer): Boolean;
//IncludeTrailingPathDelimiterの"/"バージョン
function IncludeTrailingSlashDelimiter(const S: string): string;
//ExcludeTrailingPathDelimiterの"/"バージョン
function ExcludeTrailingSlashDelimiter(const S: string): string;
//フォームの初期サイズ設定
procedure InitialFormSize(AForm: TForm; AIniFile: TMemIniFile; ASection: string; IsDaialog: Boolean); overload;
procedure InitialFormSize(AForm: TForm; AIniFile, ASection: string; IsDialog: Boolean); overload;
//コンポーネントを交換する関数
function ChangeComponent(Original: TComponent; NewClass: TComponentClass): TComponent;

function PathGetCharTypeW(ch: PAnsiChar):Integer; stdcall; external 'shlwapi.dll';
function QueryFullProcessImageNameW(Process: THandle; Flags: DWORD; Buffer: PChar;
                                    Size: PDWORD): Boolean; stdcall; external 'kernel32.dll';


implementation

{ TDirectoryExistsThread }

procedure TDirectoryExistsThread.Execute;
begin
  Result := -1;   //検索タイムアウト
  if DirectoryExists(DirectoryName) then
    Result := 0   //指定フォルダが見つかった
  else
    Result := 1;  //指定フォルダが見つからない
  //Terminateプロパティをtrueに設定しスレッド終了を通知する
  Terminate;
end;

{ UnitUtils }

var
  PCount: TLargeInteger;
  Tick: TDateTime;

//ミリ秒単位の時刻の取得 FDELPHI#16-255
function Now_ms: TDateTime;
var qq,yy: TLargeInteger;
    q: comp absolute qq;
    y: comp absolute yy;
    p: comp absolute pcount;
begin
   QueryPerformanceFrequency(qq);
   QueryPerformanceCounter(yy);
   Result := Tick + (y - p) / (24 * 3600 * q);
end;

//Delimiter文字で区切られた文字列からゼロベース配列の指定番号文字を求める
function Split(S: string; Pos: Integer): string;
var
  List: TStrings;
begin
  Result := '';
  if (S <> '') and (Pos >= 0) then begin
    List := TStringList.Create;
    try
      List.Delimiter := C_COMMA;
      List.QuoteChar := '"'; //指定しなくても'"'が指定される
      List.StrictDelimiter := True;
      List.DelimitedText := S;
      if List.Count > Pos then
        Result := List.Strings[Pos];
    finally
      FreeAndNil(List);
    end;
  end;
end;

//Delimiter文字で区切られた文字列からゼロベース配列の指定番号文字を求める
function Split(S: string; Delimiter: char; Pos: Integer): string;
var
  List: TStrings;
begin
  Result := '';
  if (S <> '') and (Pos >= 0) then begin
    List := TStringList.Create;
    try
      List.Delimiter := Delimiter;
      List.QuoteChar := '"'; //指定しなくても'"'が指定される
      List.StrictDelimiter := True;
      List.DelimitedText := S;
      if List.Count > Pos then
        Result := List.Strings[Pos];
    finally
      FreeAndNil(List);
    end;
  end;
end;

function Split(List: TStrings; Index, Pos: Integer): string;
begin
  Result := '';
  if List.Count > Index then begin
    Result := Split(List.Strings[Index], Pos);
  end;
end;

function Split(List: TStrings; Delimiter: char; Index, Pos: Integer): string;
begin
  Result := '';
  if List.Count > Index then begin
    Result := Split(List.Strings[Index], Delimiter, Pos);
  end;
end;

//Delimiter文字で区切られた文字列の分割数を求める
function SplitCount(S: string; Delimiter: char = C_COMMA): Integer;
var
  List: TStrings;
begin
  Result := 0;
  if S <> '' then begin
    List := TStringList.Create;
    try
      List.Delimiter := Delimiter;
      List.QuoteChar := '"'; //指定しなくても'"'が指定される
      List.StrictDelimiter := True;
      List.DelimitedText := S;
      Result := List.Count;
    finally
      FreeAndNil(List);
    end;
  end;
end;

function SplitCount(List: TStrings; Index: Integer; Delimiter: char = C_COMMA): Integer;
begin
  Result := 0;
  if List.Count > Index then begin
    Result := SplitCount(List.Strings[Index], Delimiter);
  end;
end;

//CommaText文字列の先頭文字列を削除する関数
// 例）"10,20,30" → "20,30"
function CommaDelete(S: String): String;
var
  SData: String;
begin
  SData := S;
  Delete(SData, 1, Pos(C_COMMA, S));
  Result := SData;
end;

function SplitDelete(S: string; Delimiter: char = C_COMMA): string;
var
  SData: string;
begin
  SData := S;
  Delete(SData, 1, Pos(Delimiter, S));
  Result := SData;
end;

function SplitDelete(S: string; Pos: Integer; Delimiter: char = C_COMMA): string;
var
  Cnt: Integer;
begin
  Result := S;
  if Pos > 0 then begin
    if Pos >= SplitCount(S, Delimiter) then begin
      Result := '';
    end
    else begin
      for Cnt := 0 to Pos-1 do begin
        Result := SplitDelete(Result, Delimiter);
      end;
    end;
  end;
end;

//MessageDlg,ShowMessageBoxを指定フォームの中央に表示
//FDelph Mes16#580参照
//2010.06.08 マルチモニタに対応した？かも？
function CenterDlg(Form:TForm; AMessage, TitleCaption:
  String; DlgType: TMsgDlgType; Buttons: TMsgDlgButtons): TModalResult;
var
  mon: TMonitor;
  FLeft, FTop: Integer;
begin
  with CreateMessageDialog(AMessage, DlgType, Buttons) do begin
    try
      Position := poDesigned;
      FTop := Form.Top + Form.Height div 2;
      FLeft := Form.Left + Form.Width div 2;
      mon := Screen.MonitorFromPoint(Point(FLeft, FTop), mdNearest);
      Top := Form.Top + (Form.Height - Height) div 2;
      if (Top + Height) > mon.WorkareaRect.Bottom then Top := mon.WorkareaRect.Bottom - Height;
      if Top < mon.WorkareaRect.Top then Top := mon.WorkareaRect.Top;
      //Formの場合Position制御が自動で行われているため横位置はこのように制御させる
      Left := Form.Left + (Form.Width - Width) div 2;
      if (Left + Width) > Screen.DesktopWidth then Left := Screen.DesktopWidth - Width;
      if Left < 0 then Left := 0;
      Caption := TitleCaption;
      Result := ShowModal;
    finally
      Free;
    end;// from try..finally...
  end;// from with...
end;

function CenterChildDlg(MainForm, ChildForm: TForm; AMessage, TitleCaption: String;
                        DlgType: TMsgDlgType; Buttons: TMsgDlgButtons): TModalResult;
var
  FDlgForm: TForm;
  mon: TMonitor;
  FLeft, FTop: Integer;
begin
  FDlgForm := CreateMessageDialog(AMessage, DlgType, Buttons);
  try
    FDlgForm.Position := poDesigned;
    FTop := MainForm.Top + ChildForm.Top + ChildForm.Height div 2;
    FLeft := MainForm.Left + ChildForm.Left + ChildForm.Width div 2;
    mon := Screen.MonitorFromPoint(Point(FLeft, FTop), mdNearest);
    FDlgForm.Top := MainForm.Top + ChildForm.Top + (ChildForm.Height - FDlgForm.Height) div 2;
    if (FDlgForm.Top + FDlgForm.Height) > mon.WorkareaRect.Bottom then begin
      FDlgForm.Top := mon.WorkareaRect.Bottom - FDlgForm.Height;
    end;
    if FDlgForm.Top < mon.WorkareaRect.Top then FDlgForm.Top := mon.WorkareaRect.Top;
    FDlgForm.Left := MainForm.Left + ChildForm.Left + (ChildForm.Width - FDlgForm.Width) div 2;
    if (FDlgForm.Left + FDlgForm.Width) > Screen.DesktopWidth then begin
      FDlgForm.Left := Screen.DesktopWidth - FDlgForm.Width;
    end;
    if FDlgForm.Left < 0 then FDlgForm.Left := 0;
    FDlgForm.Caption := TitleCaption;
    Result := FDlgForm.ShowModal;
  finally
    FreeAndNil(FDlgForm);
  end;
end;

//パラメータ名の空ファイルを作成する関数
//エラーチェックを追加（2013/02/27）
function CreateFile(FileName: string): Boolean;
var
  F: TextFile;
begin
  try
    if FileName <> '' then begin
      AssignFile(F, FileName);
      Rewrite(F);
      CloseFile(F);
    end;
    Result := FileExists(FileName);
  except
    Result := False;
  end;
end;

//ファイル名(パスを含む)として使用できるか確認する関数
//戻り値 True:ファイル名として使用可能
function CheckPathFileName(const FileName: string): Boolean;
var
  FileHandle: Integer;
begin
  FileHandle := FileCreate(FileName);
  if FileHandle = -1 then
    Result := False
  else begin
    Result := True;
    FileClose(FileHandle);
    System.SysUtils.DeleteFile(FileName);
  end;
end;

//実行ファイルのバージョン番号の取得
function GetStrVersion: string;
var
  VerInfoSize  : DWORD;
  VerInfo      : Pointer;
  VerValueSize : DWORD;
  VerValue     : PVSFixedFileInfo;
  Dummy        : DWORD;
begin
  VerInfoSize := GetFileVersionInfoSize(PChar(Application.ExeName), Dummy);
  GetMem(VerInfo, VerInfoSize);
  try
    GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo );
    VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
    with VerValue^ do begin
      Result := Format('Ver%d.%.2d.%.2d' , [(dwFileVersionMS shr 16)
                                          , (dwFileVersionMS and $FFFF)
                                          , (dwFileVersionLS shr 16)])
    end;
  finally
    FreeMem(VerInfo, VerInfoSize);
  end;
end;

//すべてのバージョン情報を取得する関数                                         //
//http://www2.big.or.jp/~osamu/Delphi/tips.cgi?index=0226.txt                  //
function GetVersionInfo(KeyWord: TVerResourceKey): string;
const
  Translation = '\VarFileInfo\Translation';
  FileInfo = '\StringFileInfo\%0.4s%0.4s\';
var
  BufSize, HWnd: DWORD;
  VerInfoBuf: Pointer;
  VerData: Pointer;
  VerDataLen: Longword;
  PathLocale: String;
begin
  // 必要なバッファのサイズを取得
  BufSize := GetFileVersionInfoSize(PChar(Application.ExeName), HWnd);
  if BufSize <> 0 then
  begin
    // メモリを確保
    GetMem(VerInfoBuf, BufSize);
    try
      GetFileVersionInfo(PChar(Application.ExeName), 0, BufSize, VerInfoBuf);
      // 変数情報ブロック内の変換テーブルを指定
      VerQueryValue(VerInfoBuf, PChar(Translation), VerData, VerDataLen);
      if not (VerDataLen > 0) then
        raise Exception.Create('情報の取得に失敗しました');
      // 8桁の１６進数に変換
      // →'\StringFileInfo\027382\FileDescription'
      PathLocale := Format(FileInfo + KeyWordStr[KeyWord],
        [IntToHex(Integer(VerData^) and $FFFF, 4),
         IntToHex((Integer(VerData^) shr 16) and $FFFF, 4)]);
      VerQueryValue(VerInfoBuf, PChar(PathLocale), VerData, VerDataLen);
      if VerDataLen > 0 then
      begin
        // VerDataはゼロで終わる文字列ではないことに注意
        result := '';
        SetLength(result, VerDataLen);
        StrLCopy(PChar(result), VerData, VerDataLen);
      end;
    finally
      // 解放
      FreeMem(VerInfoBuf);
    end;
  end;
end;

//デバッグ用ログファイル保存共通ルーチン
procedure WriteDebugLogFile(FileName: TFileName; SData: string);
var
  FS: TextFile;
begin
  AssignFile(FS, FileName);
  Append(FS);
  //Writeln(FS, FormatDateTime('yyyy/mm/dd_hh:nn:ss:zzz', Now_ms) + ' ' + SData);
  Writeln(FS, SData);
  CloseFile(FS);
end;

//FolderNameに指定されたネットワークフォルダが存在するか確認する
//http://homepage1.nifty.com/MADIA/delphi/delphi_bbs/200802/200802_08020008.html
// 入力 FolderName=検索するフォルダ名
//      TimeOut=検索処理を停止する時間(ms)
// 出力 True=フォルダが存在する False=フォルダが存在しないかTimeOut時間内に見つからない
//http://homepage1.nifty.com/MADIA/delphi/delphi_bbs/200802/200802_08020008.html参照
//ネットワーク上のフォルダをDirectoryExistsメソッドで処理するとフォルダが見つからない場合
//4分ほど処理が戻ってこないため、スレッドでDirectoryExistsメソッド処理を行う
//TimeOut時間内にフォルダが見つからない場合は、DirectoryExistTimeOutメソッドは終了するが
//＜注意＞
//DirectoryExistsのスレッドはDirectoryExistsプロセスが終了するまでは終わらない
function NetDirectoryExists(const FolderName: String; TimeOut: Cardinal = 3000): Boolean;
var
  FDirectoryExistsThread: TDirectoryExistsThread;
  FStartTime: Cardinal;
begin
  Result := False;
  //Createパラメータをtrueに設定しResumeメソッド呼び出し後スレッドを開始させる
  FDirectoryExistsThread := TDirectoryExistsThread.Create(True);
  //スレッド終了後自動的にThreadオブジェクトを破棄する
  FDirectoryExistsThread.FreeOnTerminate := True;
  FDirectoryExistsThread.DirectoryName := FolderName;
  //スレッド処理開始
  FDirectoryExistsThread.Start;
  //FDirectoryExistsThread.Resume;
  FStartTime := GetTickCount;
  while (FDirectoryExistsThread.Terminated = False) and (GetTickCount-FStartTime < TimeOut) do begin
    ;//スレッド終了またはTimeOut時間以上の時Whileルーチンを抜ける
  end;
  if FDirectoryExistsThread.Result = 0 then Result := True;
end;

//Mathユニット[PopnStdDev]拡張関数 Ver1.10
function PopnStdDevEx(const Data: array of Single): Extended;
var
  PopnData: Extended;
begin
  PopnData := PopnVariance(Data);
  if (PopnData < 0) or (Length(Data) <= 1) then PopnData := 0;
  Result := Sqrt(PopnData);
end;

function PopnStdDevEx(const Data: array of Double): Extended;
var
  PopnData: Extended;
begin
  PopnData := PopnVariance(Data);
  if (PopnData < 0) or (Length(Data) <= 1) then PopnData := 0;
  Result := Sqrt(PopnData);
end;

function PopnStdDevEx(const Data: array of Extended): Extended;
var
  PopnData: Extended;
begin
  PopnData := PopnVariance(Data);
  if (PopnData < 0) or (Length(Data) <= 1) then PopnData := 0;
  Result := Sqrt(PopnData);
end;

//表示位置を設定出来るInputQuery
//Delphi DialogsユニットのInputQuery関数丸写し...なんだかな～
function InputQueryEx(const ACaption, APrompt: string; x, y: Integer; var Value: string): Boolean;
  function GetAveCharSize(Canvas: TCanvas): TPoint;
  var
    I: Integer;
    Buffer: array[0..51] of Char;
  begin
    for I := 0 to 25 do Buffer[I] := Chr(I + Ord('A'));
    for I := 0 to 25 do Buffer[I + 26] := Chr(I + Ord('a'));
    GetTextExtentPoint(Canvas.Handle, Buffer, 52, TSize(Result));
    Result.X := Result.X div 52;
  end;

var
  Form: TForm;
  Prompt: TLabel;
  Edit: TEdit;
  DialogUnits: TPoint;
  ButtonTop, ButtonWidth, ButtonHeight: Integer;
begin
  Result := False;
  Form := TForm.Create(Application);
  with Form do
    try
      Canvas.Font := Font;
      DialogUnits := GetAveCharSize(Canvas);
      BorderStyle := bsDialog;
      Caption := ACaption;
      ClientWidth := MulDiv(180, DialogUnits.X, 4);
      Position := poDesigned;
      Left := x;
      Top := y;
      Prompt := TLabel.Create(Form);
      with Prompt do begin
        Parent := Form;
        Caption := APrompt;
        Left := MulDiv(8, DialogUnits.X, 4);
        Top := MulDiv(8, DialogUnits.Y, 8);
        Constraints.MaxWidth := MulDiv(164, DialogUnits.X, 4);
        WordWrap := True;
      end;
      Edit := TEdit.Create(Form);
      with Edit do begin
        Parent := Form;
        Left := Prompt.Left;
        Top := Prompt.Top + Prompt.Height + 5;
        Width := MulDiv(164, DialogUnits.X, 4);
        MaxLength := 255;
        Text := Value;
        SelectAll;
      end;
      ButtonTop := Edit.Top + Edit.Height + 15;
      ButtonWidth := MulDiv(50, DialogUnits.X, 4);
      ButtonHeight := MulDiv(14, DialogUnits.Y, 8);
      with TButton.Create(Form) do begin
        Parent := Form;
        Caption := 'OK';
        ModalResult := mrOk;
        Default := True;
        SetBounds(MulDiv(38, DialogUnits.X, 4), ButtonTop, ButtonWidth,
          ButtonHeight);
      end;
      with TButton.Create(Form) do begin
        Parent := Form;
        Caption := 'キャンセル';
        ModalResult := mrCancel;
        Cancel := True;
        SetBounds(MulDiv(92, DialogUnits.X, 4), Edit.Top + Edit.Height + 15,
          ButtonWidth, ButtonHeight);
        Form.ClientHeight := Top + Height + 13;          
      end;
      if ShowModal = mrOk then begin
        Value := Edit.Text;
        Result := True;
      end;
    finally
      Form.Free;
    end;
end;

// テキスト回転描画出来るようにする文字描画拡張ルーチン
//   x, y:キャンバス上座標（X, Y）z:文字回転角(反時計回り1度)
procedure TextOutEx(Canvas: TCanvas; X, Y, Z: Integer; S: String);
var
  lfText: TLOGFONT;    // フォント情報（SDK の‘LOGFONT’を参照）
  hfNew, hfOld: HFONT; // フォントハンドル
begin
  GetObject(Canvas.Font.Handle, SizeOf(TLOGFONT), @lfText);
  lfText.lfEscapement := z * 10;
  lfText.lfOrientation := lfText.lfEscapement;
  hfNew := CreateFontIndirect(lfText);
  hfOld := SelectObject(Canvas.Handle, hfNew);
  Canvas.TextOut(x, y, S);
  SelectObject(Canvas.Handle, hfOld);
  DeleteObject(hfNew);
end;

//Streamデータから文字列を読み出す共通ルーチン(Unicode対応)
//http://edn.embarcadero.com/jp/article/38699参照
function ReadStreamString(Stream: TStream): string;
var
  Index: Integer;
  SData: string;
begin
  Stream.Read(Index, SizeOf(Index));
  SetLength(SData, Index div 2);
  Stream.Read(Pointer(SData)^, Index);
  Result := SData;
end;

//StreamデータからShift-JIS文字列を読み出す共通ルーチン(Unicode対応)
function ReadStreamAnsiString(Stream: TStream): string;
var
  Index: Integer;
  SData: AnsiString;
begin
  Stream.Read(Index, SizeOf(Index));
  SetLength(SData, Index);
  Stream.Read(Pointer(SData)^, Index);
  Result := string(SData);
end;

//Streamデータに文字列Sを書き込む共通ルーチン(Unicode対応)
//http://edn.embarcadero.com/jp/article/38699参照
procedure WriteStreamString(Stream: TStream; S: string);
var
  Index: Integer;
begin
  Index := Length(S) * SizeOf(Char);
  Stream.Write(Index, SizeOf(Index));
  Stream.Write(Pointer(S)^, Index);
end;

//文字列中の改行コードをスペース文字列に変換
function ReplaceLineFeedToSpace(S: string): string;
begin
  Result := StringReplace(S, #13#10, ' ', [rfReplaceAll]);
end;

//指定ファイルが起動中のプロセスか確認する関数
function CheckActiveProcess(aExeName:string): Boolean;
var
  hProcesss: Cardinal;
  P32: TProcessEntry32;
begin
  Result := False;
  hProcesss := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if hProcesss > 0 then begin
    try
      P32.dwSize := Sizeof(TProcessEntry32);
      if Process32First(hProcesss, P32) then begin
        repeat
          if aExeName = (P32.szExeFile) then Result := True;
        Until(Process32Next(hProcesss, P32) = False);
      end;
    finally
      CloseHandle(hProcesss) ;
    end;
  end;
end;

//指定ファイルのプロセスIDを取得する関数
//  http://mrxray.on.coocan.jp/Delphi/plSamples/330_AppProcessList.htm#09
function GetActiveProcessId(aExeName: string): Cardinal;
const
  PROCESS_QUERY_LIMITED_INFORMATION = $1000;
  PROCESS_NAME_NATIVE = 1;
var
  ListHandle  : Cardinal;
  ProcEntry   : TProcessEntry32;
  ProcessID   : DWORD;
  hProcHandle : THandle;
  ExePath     : string;
  Buff        : array[0..MAX_PATH-1] of Char;
  STR_SIZE    : DWORD;
begin
  Result := 0;
  //デバッグの特権を有効にする
  Privilege.SetPrivilege(SE_DEBUG_NAME, True);
  //プロセスのスナップショットのハンドルを取得
  ListHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if ListHandle > 0 then begin
    try
      //最初のプロセスに関する情報をTProcessEntry32レコード型に取得
      ProcEntry.dwSize := SizeOf(TProcessEntry32);
      Process32First(ListHandle, ProcEntry);
      repeat
        ExePath := '';
        FillChar(Buff, SizeOf(Buff), #0);
        //プロセスIDを取得
        ProcessID := ProcEntry.th32ProcessID;
        //プロセスID値からプロセスのオープンハンドルを取得
        hProcHandle := OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, ProcessID);
        try
          //オープンハンドルからパス名を取得
          if hProcHandle > 0 then begin
            //オープンハンドルからパス名を取得
            STR_SIZE := Length(Buff);
            if QueryFullProcessImageNameW(hProcHandle, 0, @Buff, @STR_SIZE) then begin
              ExePath := String(Buff);
              //実行ファイル名が同じだったら終了
              if string(Buff) = Trim(aExeName) then begin
                Result := ProcessID;
                break;
              end;
            end;
          end;
        finally
          CloseHandle(hProcHandle);
        end;
      //次のプロセスに関する情報をTProcessEntry32レコード型に取得
      until Process32Next(ListHandle, ProcEntry) = False;
    finally
      CloseHandle(ListHandle);
    end;
  end;
end;

//指定ファイルを実行する関数（アプリの起動完了まで待つ）
//  http://mrxray.on.coocan.jp/Delphi/plSamples/485_OpenAppFile.htm#11
//  Visibilityパラメータ
//    https://msdn.microsoft.com/ja-jp/library/windows/desktop/ms633548(v=vs.85).aspx
//    SW_HIDE=    ウィンドウを非表示にし、他のウィンドウをアクティブにする
//    SW_MINIMIZE=指定されたウィンドウを最小化し、Z順序で次のトップレベルのウィンドウを起動
//    SW_SHOW=    ウィンドウをアクティブにして、現在のサイズと位置に表示
//  戻り値=WAIT_FAILED時起動に失敗
function WinExecAndWait32V2(aExeName: string; Visibility: Integer): DWORD;
var
  LStartupInfo: TStartupInfo;
  LProcessInfo: TProcessInformation;
  LRet: DWORD;
begin
  Result := 0;
  FillChar(LStartupInfo, SizeOf(LStartupInfo), #0);
  LStartupInfo.cb := SizeOf(TStartupInfo);
  LStartupInfo.dwFlags := STARTF_USESHOWWINDOW;
  LStartupInfo.wShowWindow := Visibility;
  SetLength(aExeName, Length(aExeName));
  //参照カウンタ対策
  UniqueString(aExeName);
  if not CreateProcess(nil, PChar(aExeName), nil, nil, False,
           CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil, nil, LStartupInfo, LProcessInfo) then begin
    Result := WAIT_FAILED;
  end
  else begin
    //プロセスが生成されるまで待つ
    repeat
      LRet := WaitForInputIdle(LProcessInfo.hProcess, 50);
      Application.ProcessMessages;
    until LRet <> WAIT_TIMEOUT;
    GetExitCodeProcess(LProcessInfo.hProcess, Result);
    CloseHandle(LProcessInfo.hProcess);
    CloseHandle(LProcessInfo.hThread);
  end;
end;

//-----------------------------------------------------------------------------
//  EnumWindowsのコールバック関数
//  ウィンドウのプロセスIDが引数のlParと同じだったら，そのウインドウを閉じる
//-----------------------------------------------------------------------------
function EnumWndProc(hWindow: HWND; lPar: PCardinal):
  Boolean; Stdcall;
var
  dwProcessID : Cardinal;
begin
  Result := True;

  if IsWindowVisible(hWindow) then begin
    GetWindowThreadProcessId(hWindow, dwProcessID);
    if dwProcessID = lPar^ then begin
      PostMessage(hWindow, WM_CLOSE, 0, 0);
      Result := False;
    end;
  end;
end;

//指定ファイルを終了する関数（強制終了処理あり）
procedure WinExit(aExeName: string);
var
  ProcessID   : Cardinal;
  hProcHandle : Cardinal;
begin
  //ExeFullPathのプロセスIDを取得
  //ExeFullPathのプログラムが起動していないと取得できない
  ProcessID := GetActiveProcessId(aExeName);
  if ProcessID > 0 then begin
    //EnumWindowsのコールバック関数内でアプリの閉じる作業を実行
    //http://mrxray.on.coocan.jp/Delphi/plSamples/330_AppProcessList.htm#09
    EnumWindows(@EnumWndProc, LPARAM(@ProcessID));
    Sleep(500);
    //500ms待機後アプリが終了できたかチェック
    ProcessID := GetActiveProcessId(aExeName);
    if ProcessID > 0 then begin
      //TerminateProcess関数でプロセスを強制的に終了させる
      //http://mrxray.on.coocan.jp/Delphi/plSamples/330_AppProcessList.htm#07
      hProcHandle := OpenProcess(PROCESS_TERMINATE, False, ProcessID);
      try
        TerminateProcess(hProcHandle, 0);
      finally
        CloseHandle(hProcHandle);
      end;
    end;
  end;
end;

//ファイル保存前チェック関数（AFileName=パス付ファイル名 FollowPath=True：フォルダが存在しないときは自動作成）
function CheckBeforeSaveFile(AFileName: string; FollowPath: Boolean = False): Boolean;
var
  SPath, SName: string;
begin
  Result := False;
  if TPath.HasValidPathChars(AFileName, False) then begin
    //使用可能なパス＋ファイル名
    SName := ExtractFileName(AFileName);
    if not SameText(SName, '') and TPath.HasValidFileNameChars(SName, False) then begin
      //正しいファイル名が指定されている→SData=パス抽出
      SPath := ExtractFilePath(AFileName);
      if SameText(SPath, '') then begin
        //パス指定がない
        Result := True;
      end
      else begin
        //パス指定がある
        if DirectoryExists(SPath) then begin
          //パスが存在する
          Result := true;
        end
        else if FollowPath then begin
          //パスが存在しないときは作成するとき
          try
            Result := ForceDirectories(SPath);
          except
          end;
        end;
      end;
    end;
  end;
end;

//ファイル保存前チェック関数（AFileName=パス付ファイル名 FollowPath=True：フォルダが存在しないときは自動作成）
function CheckBeforeSaveFolder(AFileName: string; FollowPath: Boolean = False): Boolean;
var
  SPath, SName: string;
begin
  Result := False;
  if TPath.HasValidPathChars(AFileName, False) then begin
    //使用可能なパス＋ファイル名
    SName := ExtractFileName(AFileName);
    if not SameText(SName, '') and TPath.HasValidFileNameChars(SName, False) then begin
      //正しいファイル名が指定されている→SData=パス抽出
      SPath := ExtractFilePath(AFileName);
      if SameText(SPath, '') then begin
        //パス指定がない
        Result := True;
      end
      else begin
        //パス指定がある
        if DirectoryExists(SPath) then begin
          //パスが存在する
          Result := true;
        end
        else if FollowPath then begin
          //パスが存在しないときは作成するとき
          try
            Result := ForceDirectories(SPath);
          except
          end;
        end;
      end;
    end;
  end;
end;

//ファイルの存在確認
function RetryFileExists(AFileName: string; ARetry: Integer): Boolean;
var
  FCounter: Integer;
begin
  FCounter := 0;
  while not FileExists(AFileName) and (FCounter < ARetry) do begin
    inc(FCounter);
    Sleep(100);
  end;
  Result := (FCounter < ARetry);
end;

//指定フォルダの存在確認（AFileName=パス付ファイル名 FollowPath=True：フォルダが存在しないときは自動作成
function CheckDirectoryExists(const ADirectory: string; AFollowPath: Boolean = False): Boolean;
begin
  Result := False;
  if DirectoryExists(ADirectory) then begin
    //フォルダが存在するとき
    Result := True;
  end
  else if AFollowPath then begin
    //フォルダが存在しないとき かつ 自動フォルダ作成モードのとき
    try
      Result := ForceDirectories(ADirectory);
    except
    end;
  end;
end;

//フォーマット拡張ルーチン
function FormatEx(const Format: string; const Args: array of const; LanguageCode: Integer): string;
begin
  try
    Result := System.SysUtils.Format(Format, Args);
  except
    on E: Exception do begin
      case LanguageCode of
        0: Result := '[Error] Format: ' + E.Message;
        1: Result := '[Error] Format: String format is incorrect or does not match argument type.';
      end;
    end;
  end;
end;

//IsPathDelimiterの"/"バージョン
function IsSlashDelimiter(const S: string; Index: Integer): Boolean;
begin
  Result := IsDelimiter(C_SLASH, S, High(S));
end;

//IncludeTrailingPathDelimiterの"/"バージョン
function IncludeTrailingSlashDelimiter(const S: string): string;
begin
  Result := S;
  if not IsSlashDelimiter(Result, High(Result)) then
    Result := Result + C_SLASH;
end;

//ExcludeTrailingPathDelimiterの"/"バージョン
function ExcludeTrailingSlashDelimiter(const S: string): string;
begin
  Result := S;
  if IsSlashDelimiter(Result, High(Result)) then
    SetLength(Result, Length(Result)-1);
end;

//フォームの初期サイズ設定
procedure InitialFormSize(AForm: TForm; AIniFile: TMemIniFile; ASection: string; IsDaialog: Boolean);
const
  //Windows10オフセット対策
  WINDOWS10_OFFSET = 8;
var
  FRect, FDesktopRect, FWorkAreaRect: TRect;
begin
  FRect.Top := AIniFile.ReadInteger(ASection, 'Top', 0);
  FRect.Left := AIniFile.ReadInteger(ASection, 'Left', 0);
  if not IsDaialog then begin
    if AForm.Constraints.MinWidth > 0 then
      AForm.Width := AIniFile.ReadInteger(ASection, 'Width', AForm.Constraints.MinWidth)
    else
      AForm.Width := AIniFile.ReadInteger(ASection, 'Width', AForm.Width);
    if AForm.Constraints.MinHeight > 0 then
      AForm.Height := AIniFile.ReadInteger(ASection, 'Height', AForm.Constraints.MinHeight)
    else
      AForm.Height := AIniFile.ReadInteger(ASection, 'Height', AForm.Height);
  end;
  FRect.Width := AForm.Width;
  FRect.Height := AForm.Height;
  FDesktopRect := Screen.DesktopRect;
  FWorkAreaRect := Screen.MonitorFromRect(FRect, mdNearest).WorkareaRect;
  if FRect.Top <= FDesktopRect.Top then
    AForm.Top := FDesktopRect.Top
  else if FRect.Top >= FWorkAreaRect.Bottom then
    AForm.Top := FWorkAreaRect.Bottom - (FRect.Height div 2)
  else
    AForm.Top := FRect.Top;
  if FRect.Left <= FDesktopRect.Left - WINDOWS10_OFFSET then
    AForm.Left := FDesktopRect.Left - WINDOWS10_OFFSET
  else if FRect.Left >= FDesktopRect.Right then
    AForm.Left := FDesktopRect.Right - (FRect.Width div 2)
  else
    AForm.Left := FRect.Left;
end;

procedure InitialFormSize(AForm: TForm; AIniFile, ASection: string; IsDialog: Boolean);
var
  FIniFile: TMemIniFile;
begin
  FIniFile := TMemIniFile.Create(AIniFile, TEncoding.Unicode);
  try
    InitialFormSize(AForm, FIniFile, ASection, IsDialog);
  finally
    FIniFile.Free;
  end;
end;

//コンポーネントを交換する関数
function ChangeComponent(Original: TComponent; NewClass: TComponentClass): TComponent;
//=============================================================================
// コンポーネントを交換する関数
// http://mrxray.on.coocan.jp/Delphi/CompoInstall/CompInstallDD.htm
// uses節にSystem.TypInfoを追加する必要がある
//=============================================================================
var
  APropList   : TPropList;
  New         : TComponent;
  Stream      : TStream;
  Methods     : array of TMethod;
  MethodCount : Integer;
  i           : Integer;
begin
  MethodCount := GetPropList(Original.ClassInfo, [tkMethod], @APropList[0]);
  SetLength(Methods, MethodCount);
  for i := 0 to MethodCount - 1 do begin
    Methods[i] := GetMethodProp(Original, APropList[i]);
  end;
  Stream := TMemoryStream.Create;
  try
    Stream.WriteComponent(Original);
    New := NewClass.Create(Original.Owner);
    if New is TControl then TControl(New).Parent := TControl(Original).Parent;
    Original.Free;
    Stream.Position := 0;
    Stream.ReadComponent(New);
  finally
    Stream.free
  end;
  for i := 0 to MethodCount - 1 do begin
    SetMethodProp(New, APropList[i], Methods[i]);
  end;
  Result := New;
end;


//*************************************************************************************************
{$REGION '// TUnitDataList //'}
//*************************************************************************************************

function TUnitDataList.ProcAddFloat(const Value: Extended; Digits: Integer; const S1, S2: string): Integer;
begin
  SetLength(Self.FItems, Self.Count+1);
  if Digits < 0 then
    Self.FItems[Self.Count-1] := S1 + FloatToStr(Value) + S2
  else
    Self.FItems[Self.Count-1] := S1 + FloatToStrF(Value, ffFixed, 15, Digits) + S2;
  Result := Self.Count-1;
end;

function TUnitDataList.Add(const Value: Single; Digits: Integer): Integer;
begin
  Result := ProcAddFloat(Value, Digits, '', '');
end;

function TUnitDataList.Add(const S1: string; const Value: Single; const Digits: Integer; const S2: string): Integer;
begin
  Result := ProcAddFloat(Value, Digits, S1, S2);
end;

function TUnitDataList.Add(const Value: Double; Digits: Integer): Integer;
begin
  Result := ProcAddFloat(Value, Digits, '', '');
end;

function TUnitDataList.Add(const S1: string; const Value: Double; const Digits: Integer; const S2: string): Integer;
begin
  Result := ProcAddFloat(Value, Digits, S1, S2);
end;

function TUnitDataList.Add(const Value: Extended; Digits: Integer): Integer;
begin
  Result := ProcAddFloat(Value, Digits, '', '');
end;

function TUnitDataList.Add(const S1: string; const Value: Extended; const Digits: Integer; const S2: string): Integer;
begin
  Result := ProcAddFloat(Value, Digits, S1, S2);
end;

function TUnitDataList.Add(const Value: string): Integer;
begin
  SetLength(Self.FItems, Self.Count+1);
  Self.FItems[Self.Count-1] := Value;
  Result := Self.Count-1;
end;

function TUnitDataList.Add(const Value: Integer): Integer;
begin
  SetLength(Self.FItems, Self.Count+1);
  Self.FItems[Self.Count-1] := IntToStr(Value);
  Result := Self.Count-1;
end;

function TUnitDataList.Add(const S1: string; const Value: Integer; const S2: string): Integer;
begin
  Result := Self.Add(S1 + IntToStr(Value) + S2);
end;

function TUnitDataList.Add(const Value: Word): Integer;
begin
  SetLength(Self.FItems, Self.Count+1);
  Self.FItems[Self.Count-1] := IntToStr(Value);
  Result := Self.Count-1;
end;

function TUnitDataList.Add(const S1: string; const Value: Word; const S2: string): Integer;
begin
  Result := Self.Add(S1 + IntToStr(Value) + S2);
end;

function TUnitDataList.Add(const FLG: Boolean): Integer;
begin
  SetLength(Self.FItems, Self.Count+1);
  Self.FItems[Self.Count-1] := BoolToStr(FLG, True);
  Result := Self.Count-1;
end;

function TUnitDataList.Add(const S1: string; const FLG: Boolean; const S2: string): Integer;
begin
  Result := Self.Add(S1 + BoolToStr(FLG, True) + S2);
end;

procedure TUnitDataList.Clear;
begin
  SetLength(Self.FItems, 0);
end;

function TUnitDataList.CommaText: string;
var
  Cnt: Integer;
  SData: string;
begin
  Result := '';
  if Self.Count > 0 then begin
    SData := Self.GetItem(0);
    if SplitCount(SData) > 1 then
      Result := '"' + SData + '"'
    else
      Result := SData;
    for Cnt := 1 to Self.Count-1 do begin
      SData := Self.GetItem(Cnt);
      if SplitCount(SData) > 1 then
        Result := Result + ',"' + SData + '"'
      else
        Result := Result + ',' + SData;
    end;
  end;
end;

function TUnitDataList.TabText: string;
const
  C_TAB = #9;
var
  Cnt: Integer;
  SData: string;
begin
  Result := '';
  if Self.Count > 0 then begin
    SData := Self.GetItem(0);
    if SplitCount(SData) > 1 then
      Result := '"' + SData + '"'
    else
      Result := SData;
    for Cnt := 1 to Self.Count-1 do begin
      SData := Self.GetItem(Cnt);
      if SplitCount(SData) > 1 then
        Result := Result + C_TAB + '"' + SData + '"'
      else
        Result := Result + C_TAB + SData;
    end;
  end;
end;

function TUnitDataList.SpaceText: string;
var
  Cnt: Integer;
  SData: string;
begin
  Result := '';
  if Self.Count > 0 then begin
    SData := Self.GetItem(0);
    if SplitCount(SData) > 1 then
      Result := '"' + SData + '"'
    else
      Result := SData;
    for Cnt := 1 to Self.Count-1 do begin
      SData := Self.GetItem(Cnt);
      if SplitCount(SData) > 1 then
        Result := Result + ' "' + SData + '"'
      else
        Result := Result + ' ' + SData;
    end;
  end;
end;

function TUnitDataList.GetCount: Integer;
begin
  Result := Length(Self.FItems);
end;

function TUnitDataList.GetItem(Index: Integer): string;
begin
  Result := '';
  if Cardinal(Index) < Cardinal(Self.Count) then begin
    Result := Self.FItems[Index];
  end;
end;

function TUnitDataList.GetStrings(Index: Integer): string;
begin
  Result := GetItem(Index);
end;

function TUnitDataList.LineFeedText: string;
var
  Cnt: Integer;
begin
  Result := '';
  if Self.Count >= 1 then begin
    Result := Self.GetItem(0);
    for Cnt := 1 to Self.Count-1 do begin
      Result := Result + #13#10 + Self.GetItem(Cnt);
    end;
  end;
end;

procedure TUnitDataList.SetStrings(Index: Integer; const Value: string);
begin
  if Cardinal(Index) < Cardinal(Self.Count) then begin
    Self.FItems[Index] := Value;
  end;
end;

function TUnitDataList.Text: string;
var
  Cnt: Integer;
begin
  Result := '';
  if Self.Count >= 1 then begin
    Result := Self.GetItem(0);
    for Cnt := 1 to Self.Count-1 do begin
      Result := Result + ' ' + Self.GetItem(Cnt);
    end;
  end;
end;

//*************************************************************************************************
{$ENDREGION}
//*************************************************************************************************

//*************************************************************************************************
{$REGION '// TUnitMapManager //'}
//*************************************************************************************************

function TLerpMapManager.GetAxisSize: TMapData;
begin
  Result.X := Length(FXAxisArray);
  Result.Y := Length(FYAxisArray);
end;

function TLerpMapManager.GetMap(XIndex, YIndex: Integer): Extended;
begin
  Result := 0.0;
  if (XIndex <= High(FMapArray)) and (YIndex <= High(FMapArray[0])) then begin
    Result := FMapArray[XIndex, YIndex];
  end;
end;

function TLerpMapManager.GetSize: TMapData;
begin
  Result.X := Length(FXAxisArray);
  Result.Y := Length(FYAxisArray);
end;

function TLerpMapManager.GetMapValue(XValue, YValue: Extended): Extended;
var
  CntX, CntY: Integer;
  FXTables, FYTables: TTableDatas;
begin
  SetLength(FXTables, Length(FXAxisArray));
  SetLength(FYTables, Length(FYAxisArray));
  for CntY := 0 to High(FYTables) do begin
    for CntX := 0 to High(FXTables) do begin
      FXTables[CntX].Axis := FXAxisArray[CntX];
      FXTables[CntX].Value := FMapArray[CntX, CntY];
    end;
    FYTables[CntY].Axis := FYAxisArray[CntY];
    FYTables[CntY].Value := _Linear(XValue, FXTables);
  end;
  Result := _Linear(YValue, FYTables);
end;

function TLerpMapManager.GetXAxis(Index: Integer): Extended;
begin
  Result := 0.0;
  if Index <= High(FXAxisArray) then begin
    Result := FXAxisArray[Index];
  end;
end;

function TLerpMapManager.GetYAxis(Index: Integer): Extended;
begin
  Result := 0.0;
  if Index <= High(FYAxisArray) then begin
    Result := FYAxisArray[Index];
  end;
end;

procedure TLerpMapManager.SetAxisSize(const Value: TMapData);
begin
  SetLength(FXAxisArray, Value.X);
  SetLength(FYAxisArray, Value.Y);
  SetLength(FMapArray, Value.X, Value.Y);
end;

procedure TLerpMapManager.SetMap(XIndex, YIndex: Integer; const Value: Extended);
begin
  if (XIndex <= High(FMapArray)) and (YIndex <= High(FMapArray[0])) then begin
    FMapArray[XIndex, YIndex] := Value;
  end;
end;

procedure TLerpMapManager.SetXAxis(Index: Integer; const Value: Extended);
begin
  if Index <= High(FXAxisArray) then begin
    FXAxisArray[Index] := Value;
  end;
end;

procedure TLerpMapManager.SetYAxis(Index: Integer; const Value: Extended);
begin
  if Index <= High(FYAxisArray) then begin
    FYAxisArray[Index] := Value;
  end;
end;

function TLerpMapManager._Index(AValue: Extended; ATables: TTableDatas): Integer;
begin
  Result := 1;
  if AValue < ATables[1].Axis then
    Result := 0
  else if AValue >= ATables[High(ATables)].Axis then
    Result := High(ATables)
  else
    while AValue >= ATables[Result+1].Axis do inc(Result);
end;

function TLerpMapManager._Lerp(AValue, X, X1, Y, Y1: Extended): Extended;
begin
  Result := Y + (Y1 - Y) * (AValue - X) / (X1 - X);
end;

function TLerpMapManager._Linear(AAxisValue: Extended; ATables: TTableDatas): Extended;
var
  FIndex: Integer;
begin
  if AAxisValue < ATables[0].Axis then begin
    Result := ATables[0].Value;
  end
  else if AAxisValue >= ATables[High(ATables)].Axis then begin
    Result := ATables[High(ATables)].Value;
  end
  else begin
    FIndex := _Index(AAxisValue, ATables);
    Result := _Lerp(AAxisValue, ATables[FIndex].Axis, ATables[FIndex+1].Axis
                    , ATables[FIndex].Value, ATables[FIndex+1].Value);
  end;
end;

//*************************************************************************************************
{$ENDREGION}
//*************************************************************************************************

initialization
  QueryPerformanceCounter(Pcount);
  Tick:=now;

end.
