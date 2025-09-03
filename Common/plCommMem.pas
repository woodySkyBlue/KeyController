unit plCommMem;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes, Vcl.Dialogs;

type
  // データ数もレコード型にしておく 
  TCommMemFileHeader = record
    Count : Integer;  // データ数
  end;

  TCommMemHeader = record
    Offset : Cardinal;              // 書き込んだ位置（先頭からのオフセットバイト）
    Size : Cardinal;                // データのバイト数
    dType : array[0..32] of Char;   // データの形式 (種類)．格納内容は任意
  end;
  PCommMemHeader = ^TCommMemHeader;

  // 上のレコード型の動的配列
  TCommMemHeaderArray = array of TCommMemHeader;
  PCommMemHeaderArray = ^TCommMemHeaderArray;

  PplCommMem = ^TplCommMem;
  TplCommMem = record
  private
    class var
      FhCommMem : THandle;
    class procedure Open(var hMutex: THandle; var hMapObject: THandle;
      var pBaseAddr: Pointer; MapFileName: string); static;
    class procedure Close(hMutex: THandle; hObject: THandle; pBaseAddr: Pointer); static;
  public
    class procedure Create(AllocSize: DWORD; MapFileName: string); static;
    class procedure Free(MapFileName: string); static;
    class function Write(MapFileName: string; Source: Pointer; Offset: DWORD;
      Length: DWORD): Boolean; static;
    class function Read(MapFileName: string; Destination: Pointer; Offset: DWORD;
      Length: DWORD): Boolean; static;
    class function Size(MapFileName: string): NativeUInt; static;
  end;

  // 以下は共有メモリで複数のデータを扱う場合の関数類
  // これらの関数類は TplCommMem では使用していない 
  function WriteCommMemHeader(AMapName: string; var AInfoArray: TCommMemHeaderArray;
    ADataCount: Cardinal): Boolean;
  function WriteCommMemData(AMapName: string; ASrcStream: TMemoryStream;
    var AInfoArray: TCommMemHeaderArray; ADataType: string): Boolean;
  function ReadCommMemHeader(AMapName: string; var AInfoArray: TCommMemHeaderArray;
    var ADataCount: Cardinal): Boolean;
  function ReadCommMemData(AMapName: string; ADstStream: TMemoryStream;
    AInfoArray: TCommMemHeaderArray; AIndex: Integer; var ADataType: string):
    Boolean;
  function StringToMemStream(AText: string; ADstStream: TMemoryStream): Boolean;
  function MemStreamToString(ASrcStream: TMemoryStream): string;

implementation

var
  FErrorText : string;
const
  // 追加した関数類で使用する定数値
  // この値も上の TCommMemXXX レコード型も TplCommMem では使用していない
  FFileHeaderSize : Cardinal = SizeOf(TCommMemFileHeader);
  FInfoHeaderSize : Cardinal = SizeOf(TCommMemHeader);

{ TplCommMem }

//=============================================================================
//  名前付きメモリマップドファイルを使用した共有メモリの作成
//
//  AllocSize   : 確保する共有メモリ (メモリマップドファイル) のサイズ (バイト)
//  MapFileName : 共有メモリ名 (メモリマップドファイル名)
//
//  引数の MapFileName のメモリマップドファイルが存在する場合，新規作成はしない
//  この場合，エラー値は ERROR_ALREADY_EXISTS となる
//
//  メモリマップドファイルのサイズは，ページサイズ (4 KB) 単位で指定する
//  ページサイズの倍数値でない場合は自動調整される
//=============================================================================
class procedure TplCommMem.Create(AllocSize: DWORD; MapFileName: string);
begin
  FhCommMem:= CreateFileMapping(INVALID_HANDLE_VALUE,
                                nil,
                                PAGE_READWRITE,
                                0,
                                AllocSize,
                                PChar(MapFileName));
end;

//=============================================================================
//  共有メモリ破棄
//  共有メモリのメモリマップドファイルの破棄は，作成したアプリでないとできない
//=============================================================================
class procedure TplCommMem.Free(MapFileName: string);
var
  LhMapObj : THandle;
begin
  LhMapObj:= OpenFileMapping(FILE_MAP_ALL_ACCESS, False, PChar(MapFileName));
  if LhMapObj <> 0 then begin
    CloseHandle(LhMapObj);
    CloseHandle(FhCommMem);
    FhCommMem := 0;
  end;
end;

//=============================================================================
//  共有メモリにデータを書き込む
//
//  MapFileName : メモリマップドファイルの名前 ( 共有メモリの名前 )
//  Source      : 共有メモリに書き込むデータのアドレス
//  Offset      : データを書き込み開始する共有メモリの先頭からの位置 (バイト単位)
//  Length      : 共有メモリに書き込むデータのバイトサイズ  
//=============================================================================
class function TplCommMem.Write(MapFileName: string; Source: Pointer;
  Offset: DWORD; Length: DWORD): Boolean;
var
  LhMutex     : THandle;
  LhMapObj    : THandle;
  LpBaseAddr  : Pointer;
  LMemoryInfo : MEMORY_BASIC_INFORMATION;
begin
  Result := False;

  Open(LhMutex, LhMapObj, LpBaseAddr, MapFileName);
  if LhMapObj <> 0 then begin
    VirtualQuery(LpBaseAddr, LMemoryInfo, SizeOf(LMemoryInfo));
    if (Offset + Length) < LMemoryInfo.RegionSize then begin
     // ミューテックスオブジェクトの状態をチェック
     // 処理中の時は待つ
      WaitForSingleObject(LhMutex, INFINITE);
      // データを書き込む
      CopyMemory(Pointer(NativeUInt(LpBaseAddr) + Offset), Source, Length);
      // フラッシュして確実に更新する
      FlushViewOfFile(Pointer(NativeUInt(LpBaseAddr) + Offset), Length);
      Result := True;
    end else begin
      // 書き込みサイズオーバー
    end;
  end;
  Close(LhMutex, LhMapObj, LpBaseAddr);
end;

//=============================================================================
//  共有メモリからデータを読み出す
//
//  MapFileName : メモリマップドファイルの名前 ( 共有メモリの名前 )
//  Destination : 共有メモリから読み込んだデータの格納アドレス
//  Offset      : データを読み出す共有メモリの先頭からの位置 (バイト単位)
//  Length      : 共有メモリから読み込むデータのバイトサイズ
//=============================================================================
class function TplCommMem.Read(MapFileName: string; Destination: Pointer;
  Offset: DWORD; Length: DWORD): Boolean;
var
  LhMutex     : THandle;
  LhMapObj    : THandle;
  LpBaseAddr  : Pointer;
  LMemoryInfo : MEMORY_BASIC_INFORMATION;
begin
  Result := False;

  Open(LhMutex, LhMapObj, LpBaseAddr, MapFileName);
  if LhMapObj <> 0 then begin
    VirtualQuery(LpBaseAddr, LMemoryInfo, SizeOf(LMemoryInfo));
    if (Offset + Length) < LMemoryInfo.RegionSize then begin
     // ミューテックスオブジェクトの状態をチェック
     // 処理中の時は待つ
      WaitForSingleObject(LhMutex, INFINITE);
      // データを読み出す
      CopyMemory(Destination, Pointer(NativeUInt(LpBaseAddr) + Offset), Length);
      Result := True;
    end else begin
      // 読み出しサイズオーバー
    end;
  end else begin
    ZeroMemory(Destination, Length);
  end;
  Close(LhMutex, LhMapObj, LpBaseAddr);
end;

//=============================================================================
//  共有メモリ (メモリマッブドファイル) のバイトサイトを取得
//=============================================================================
class function TplCommMem.Size(MapFileName: string): NativeUInt;
var
  LhMutex     : THandle;
  LhMapObj    : THandle;
  LpBaseAddr  : Pointer;
  LMemoryInfo : MEMORY_BASIC_INFORMATION;
begin
  Result := 0;

  Open(LhMutex, LhMapObj, LpBaseAddr, MapFileName);
  if LhMapObj <> 0 then begin
    VirtualQuery(LpBaseAddr, LMemoryInfo, SizeOf(LMemoryInfo));
    Result := LMemoryInfo.RegionSize;
  end;
  Close(LhMutex, LhMapObj, LpBaseAddr);
end;

//=============================================================================
//  ミューテックスオブジェクトを生成
//  メモリマッブドファイルを開く
//=============================================================================
class procedure TplCommMem.Open(var hMutex: THandle; var hMapObject: THandle;
  var pBaseAddr: Pointer; MapFileName: string);
var
  LMutexName : string;
begin
  hMapObject := 0;
  // ミューテックスオブジェクトを生成
  LMutexName := MapFileName + '_Mutex';
  hMutex := CreateMutex(nil, False, PChar(LMutexName));
  if hMutex = 0 then Exit;

  // メモリマッブドファイルのハンドルを取得
  hMapObject := OpenFileMapping(FILE_MAP_ALL_ACCESS, False, PChar(MapFileName));
  if hMapObject <> 0 then begin
    // 共有メモリの先頭アドレスのポインタを取得
    pBaseAddr := MapViewOfFile(hMapObject, FILE_MAP_ALL_ACCESS, 0, 0, 0);
  end;
end;

//=============================================================================
//  メモリマッブドファイルを閉じる
//  ミューテックスオブジェクトの所有権の破棄とハンドルの開放
//=============================================================================
class procedure TplCommMem.Close(hMutex: THandle; hObject: THandle; pBaseAddr: Pointer);
begin
  if hObject <> 0 then begin
   UnmapViewOfFile(pBaseAddr);
    CloseHandle(hObject);
  end;

  // ミューテックスオブジェクトの所有権の破棄とハンドルの開放
  if hMutex <> 0 then begin
    ReleaseMutex(hMutex);
    CloseHandle(hMutex);
  end;
end;

//-----------------------------------------------------------------------------
//  データの情報を取得してレコード型のメンバの値を調べる
//  WriteCommMemHeader が実行されていれば初期化さて 0 になっている
//-----------------------------------------------------------------------------
function CheckInfoHeader(AMapName: string; AIndex: Integer): Boolean;
var
  LInfoHeader : TCommMemHeader;
  LInfoOffset : Cardinal;
begin
  // データ情報を格納しているのレコード型の値を読み出す
  LInfoOffset := FFileHeaderSize + Cardinal(AIndex) * FInfoHeaderSize;
  TplCommMem.Read(AMapName,
                 @LInfoHeader,
                 LInfoOffset,
                 FInfoHeaderSize);
  Result := (LInfoHeader.Offset > 0) and (LInfoHeader.Offset > 0);
end;

//-----------------------------------------------------------------------------
//  TplCommMem を利用して複数データを共有メモリに書き込むための関数
//  共有メモリのヘッダ部の初期処理
//  WriteCommMemDat 関数でデータを書き込む前にこの関数を実行する
//
//  AMapName        : 共有メモリ名
//  var AInfoArray  : データの情報を格納する TCommMemHeader;レコード型の動的配列
//  ADataCount      : 配列の要素数．この値がデータ数
//-----------------------------------------------------------------------------
function WriteCommMemHeader(AMapName: string; var AInfoArray: TCommMemHeaderArray;
  ADataCount: Cardinal): Boolean;
var
  LFileHeader : TCommMemFileHeader;
  LMemStream  : TMemoryStream;
  LMemSize    : Cardinal;
begin
  Result := False;

  if (ADataCount <= 0) then begin
    FErrorText := '配列の要素数は 1 以上必要です．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;

  // 配列の要素数を設定
  SetLength(AInfoArray, ADataCount);

  // ヘッダ部の利用域を初期化
  LMemStream := TMemoryStream.Create;
  try
    LMemSize := FFileHeaderSize + Cardinal(Length(AInfoArray)) * FInfoHeaderSize;
    if (LMemSize > TplCommMem.Size(AMapName)) then begin
      FErrorText := '書き込むサイズが共有メモリのサイズを越えています．';
      MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
      Exit;
    end;

    LMemStream.Size := LMemSize;
    ZeroMemory(LMemStream.Memory, LMemSize);
    Result := Tplcommmem.Write(AMapName, LMemStream.Memory, 0, LMemSize);
  finally
    FreeAndNil(LMemStream);
  end;
  if not Result then begin
    FErrorText := '共有メモリにアクセスできません．' + sLineBreak
                +  '指定名の共有メモリが作成されていない可能性があります．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;

  // 共有メモリの先頭にデータ数をレコード型として書き込む
  LFileHeader.Count := ADataCount;
  Result := TplCommMem.Write(AMapName, @LFileHeader, 0, FFileHeaderSize);
end;

//-----------------------------------------------------------------------------
//  TplCommMem を利用して複数データを共有メモリに書き込むための関数
//  データを書き込む関数
//  実際に書き込むのは引数の TMemoryStream のデータ 
//  書き込むデータは TMemoryStream に格納する必要がある
//
//  AMapName        : 共有メモリ名
//  ASrcStream      : 共有メモリに書き込むデータが格納されているメモリストリーム
//  var AInfoArray  : データの情報を格納する TCommMemHeader;レコード型の動的配列
//                    このメンバに値をセットして共有メモリにデータを書き込む
//  ADataType       : データの形式．内容は任意
//
//  共有メモリのヘッダ部のすぐ後にデータを書き込む
//  この関数を実行する度に順番にシーケンシャルにデータを書き込んでいく
//-----------------------------------------------------------------------------
function WriteCommMemData(AMapName: string; ASrcStream: TMemoryStream;
  var AInfoArray: TCommMemHeaderArray; ADataType: string): Boolean;
{$WRITEABLECONST ON}
const
  LIndex      : Integer = 0;
  LDataCount  : Integer = 0;
  LDataOffset : Cardinal = 0;
{$WRITEABLECONST OFF}
var
  LInfoOffset : Cardinal;
begin
  Result := False;

  if Length(AInfoArray) <= 0 then begin
    FErrorText := '配列の要素数が設定されていません．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;

  // 最初のデータ情報を格納しているのレコード型の値をチェック
  if not CheckInfoHeader(AMapName, 0) then LIndex := 0;
  if LIndex = 0 then begin
    // 最初のデータの書き込みの時の処理
    // ヘッダ部からデータ数を読み出す
    Result := TplcommMem.Read(AMapName, @LDataCount, 0, FFileHeaderSize);
    // 最初のデータを書き込む共有メモリ上の位置
    LDataOffset := FFileHeaderSize + cardinal(LDataCount) * FInfoHeaderSize;
  end;
  // テータの情報を格納するレコード型のデータの書き込み位置
  LInfoOffset := FFileHeaderSize + Cardinal(LIndex) * FInfoHeaderSize;

  if LDataCount <> Length(AInfoArray) then begin
    FErrorText := 'データ数は途中変更できません．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;
  if LIndex > High(AInfoArray) then begin
    FErrorText := '書き込むデータ数が配列の要素数を超えています．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;
  if ((LInfoOffset + FInfoHeaderSize) > TplCommMem.Size(AMapName))
      or ((LDataOffset + ASrcStream.Size) > TplCommMem.Size(AMapName)) then begin
    FErrorText := '書き込むサイズが共有メモリのサイズを越えています．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;

  // テータの情報を各メンバに設定してレコード型のデータを書き込む
  AInfoArray[LIndex].Offset := LDataOffset;
  AInfoArray[LIndex].Size   := ASrcStream.Size;
  // 静的文字配列の要素を全て Null 文字にして String 型の文字列を代入
  FillChar(AInfoArray[LIndex].dType[0], SizeOF(AInfoArray[LIndex].dType), #0);
  StrPCopy(AInfoArray[LIndex].dType, ADataType);
  Result := TplCommMem.Write(AMapName,
                             @AInfoArray[LIndex],
                             LInfoOffset,
                             FInfoHeaderSize);
  if Result then begin
    // TMemoryStream のデータを共有メモリに書き込む
    Result := TplCommMem.Write(AMapName,
                               ASrcStream.Memory,
                               LDataOffset,
                               ASrcStream.Size);
    if Result then begin
      // 次のデータの書き込み位置
      LDataOffset := LDataOffset + ASrcStream.Size;
      // 次の動的配列の要素番号
      LIndex := LIndex + 1;
    end;
  end;
end;

//-----------------------------------------------------------------------------
//  共有メモリからヘッダ部のデータを取得する関数
//  ヘッダ部には共有メモリに書き込んだデータ数
//  各データの情報を格納したレコード型のデータがデータ数分ある
//  実際のデータはこのヘッダ部の後に書き込まれている
//
//  AMapName      　: 共有メモリ名
//  var AInfoArray : データの情報を格納した TCommMemHeader;レコード型の動的配列
//                   この関数はこのデータと次の引数の値を取得する
//  ar ADataCount　: データ数
//-----------------------------------------------------------------------------
function ReadCommMemHeader(AMapName: string; var AInfoArray: TCommMemHeaderArray;
  var ADataCount: Cardinal): Boolean;
var
  LFileHeader : TCommMemFileHeader;
  LInfoSize   : Integer;
begin
  ADataCount := 0;

  // 共有メモリの先頭からデータ数の値をレコード型として読み出す
  Result := TplCommMem.Read(AMapName, @LFileHeader, 0, FFileHeaderSize);

  if not Result then begin
    FErrorText := '共有メモリにアクセスできません．' + sLineBreak
                +  '指定名の共有メモリが作成されていない可能性があります．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;

  ADataCount := LFileHeader.Count;
  // 動的配列の要素数を設定
  SetLength(AInfoArray, ADataCount);

  // データの情報を格納しているレコード型のデータを全て読み込む
  // LInfoSize はテータの情報を格納しているレコード型の全バイト数
  // 動的配列は，1 次元であれば最初の要素のアドレス指定で読み込める
  // 動的配列の最初の要素番号は常に 0
  LInfoSize := ADataCount * FInfoHeaderSize;
  Result := TplCommMem.Read(AMapName, @AInfoArray[0], FFileHeaderSize, LInfoSize);

  // 最初のデータ情報を格納しているのレコード型の値をチェック
  if not CheckInfoHeader(AMapName, 0) then begin
    FErrorText := 'ヘッダ部に情報がありません．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;
end;

//-----------------------------------------------------------------------------
//   TplCommMem を利用して複数データを書き込んだ共有メモリからデータを読み込む関数
//  実際には引数の TMemoryStream に読み込む 
//
//  AMapName      　: 共有メモリ名
//  ADstStream    　: このメモリストリームに共有メモリから読み出したデータを格納
//  var AInfoArray : データの情報を格納した TCommMemHeader;レコード型の動的配列
//                   この関数実行前に ReadCommMemHeader 関数を実行すると取得できる
//  AIndex        : データの情報を格納した TCommMemHeader レコード型の動的配列の要素番号
//                  この要素番号のデータを引数の TMemoryStream に格納する
//  var ADataType : 読み出したデータの形式
//-----------------------------------------------------------------------------
function ReadCommMemData(AMapName: string; ADstStream: TMemoryStream;
  AInfoArray: TCommMemHeaderArray; AIndex: Integer; var ADataType: string): Boolean;
begin
  Result := False;

  if (AIndex < 0) or (AIndex > High(AInfoArray)) then begin
    FErrorText := '配列の要素番号が範囲外です．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;

  if not CheckInfoHeader(AMapName, AIndex) then begin
    FErrorText := '要素 ' + IntToStr(AIndex)
                + ' (  ' + IntToStr(AIndex + 1) + ' 番目 ) のデータがありません．';
    MessageBox(0, PChar(FErrorText), '情報', MB_ICONWARNING);
    Exit;
  end;

  ADataType :=  AInfoArray[AIndex].dType;
  ADstStream.Size := AInfoArray[AIndex].Size;
  Result := TplCommMem.Read(AMapName,
                            ADstStream.Memory,
                            AInfoArray[AIndex].Offset,
                            AInfoArray[AIndex].Size);
  ADstStream.Position := 0;
end;

//-----------------------------------------------------------------------------
//  String 型の文字列をメモリストリームに格納する関数
//
//  原則として
//  Create 等で生成するオブジェクトは関数の戻り値にはしない
//  クラス型のインスタンスを手続きや関数内で生成して引数や戻り値にはしない
//-----------------------------------------------------------------------------
function StringToMemStream(AText: string; ADstStream: TMemoryStream): Boolean;
begin
  ADstStream.Size := Length(AText) * SizeOf(Char);
  if ADstStream.Size > 0 then begin
    Move(AText[1], ADstStream.Memory^, ADstStream.Size);
  end;
  Result := ADstStream.Size > 0;
end;

//-----------------------------------------------------------------------------
//  メモリストリームのデータを String 型の文字列変数に代入する関数
//-----------------------------------------------------------------------------
function MemStreamToString(ASrcStream: TMemoryStream): string;
begin
  Result := '';
  if ASrcStream.Size = 0 then Exit;

  SetLength(Result, ASrcStream.Size div SizeOf(Char));
  Move(ASrcStream.Memory^, Result[1], ASrcStream.Size);
end;

end.