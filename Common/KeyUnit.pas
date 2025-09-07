unit KeyUnit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.Classes, System.Generics.Collections;

type
  // キー状態集合型
  TModifierState = set of (tksShift, tksCtrl, tksNonConvert);

  TKeyCombination = record
    Key: Byte;
    State: TModifierState;
  end;

  TKeyInfo = record
    IsPressed: Boolean;
    IsRepeat:  Boolean;
  end;

  TKeyController = class
  private
    FControlKey: TKeyCombination;
    FOperationKeyList: TList<TKeyCombination>;
    function GetControlKey: TKeyCombination;
    function GetOperationKey(Index: Integer): TKeyCombination;
    procedure SetControlKey(const Value: TKeyCombination);
    procedure SetOperationKey(Index: Integer; const Value: TKeyCombination);
    procedure ProcKeyUp(aKey: Byte);
    procedure ProcKeyDown(aKey: Byte);
    procedure ProcModifierKeyUp(aState: TModifierState);
    procedure ProcModifierKeyDown(aState: TModifierState);
    procedure ProcWriteModifierStateToStream(aState: TModifierState; aStream: TStream);
    procedure ProcReadModifierStateFromStream(out aState: TModifierState; aStream: TStream);
  public
    constructor Create;
    destructor Destroy; override;
    function OperationKeyCount: Integer;
    function KeyConvert(aCombination: TKeyCombination; aKeyInfo: TKeyInfo; var aTime: UInt64): Boolean;
    procedure Assigne(aSource: TKeyController);
    procedure Clear;
    procedure AddOperationKey(aCombinationKey: TKeyCombination);
    procedure MoveOperationKey(CurIndex, NewIndex: NativeInt);
    procedure LoadFromStream(aStream: TStream);
    procedure SaveToStream(aStream: TStream);
    property ControlKey: TKeyCombination read GetControlKey write SetControlKey;
    property OperationKey[Index: Integer]: TKeyCombination read GetOperationKey write SetOperationKey;
  end;

  TKeyControllerList = class
  private
    FKeyControllerList: TObjectList<TKeyController>;
    function GetItems(aIndex: Integer): TKeyController;
  public
    constructor Create;
    destructor Destroy; override;
    function Count: Integer;
    function KeyConvert(aCombination: TKeyCombination; aKeyInfo: TKeyInfo; var aTime: UInt64): Boolean;
    procedure Add(aKeyController: TKeyController);
    procedure Assign(aSource: TKeyControllerList);
    procedure Clear;
    procedure LoadFromStream(aStream: TStream);
    procedure SaveToStream(aStream: TStream);
    property Items[aIndex: Integer]: TKeyController read GetItems;
  end;

  TKeyApplication = record
    ClassName: string;
    KeyControllerList: TKeyControllerList;
  end;

  TKeyCondition = record
  private
    FCombination: TKeyCombination;
    FMaskKey: Byte;
    FIsMasked: Boolean;
    FIsPressed: Boolean;
    FIsRepeat: Boolean;
    FCount: Integer;
    //FModifierState: TModifierState;
    function GetIsExecuted: Boolean;
    function GetIsPressed: Boolean;
    function GetKeyCombination: TKeyCombination;
    function GetModifierState: TModifierState;
    procedure SetModifierState(const Value: TModifierState);
    procedure ProcSetModifierKey(aModifierKey: TModifierState);
    procedure ProcUpdateModifierState;
    function GetInfo: TKeyInfo;
  public
    procedure Initialize;
    procedure Update(wPar: WPARAM; lPar: LPARAM);
    procedure StartMask(aCount: Integer);
    procedure StopMask;
    property IsExecuted: Boolean read GetIsExecuted;
    property KeyCombination: TKeyCombination read GetKeyCombination;
    property IsPressed: Boolean read GetIsPressed;
    property ModifierState: TModifierState read GetModifierState write SetModifierState;
    property Info: TKeyInfo read GetInfo;
  end;

  TkeyApplicationList = class
  private
    FTime: UInt64;
    //FMaskKey: Byte;
    //FIsMask: Boolean;
    //FModifierState: TModifierState;
    FKeyCondition: TKeyCondition;
    FKeyApplicationList: TList<TKeyApplication>;
    function ProcGetActiveWindowClassName: string;
    //procedure ProcSetModifierKey(aIsKeyPressed: Boolean; akey: TModifierState);
    //procedure ProcUpdateModifierKey(aKey: Integer; aIsKeyPressed: Boolean);
    function GetItems(Index: Integer): TKeyApplication;
  public
    constructor Create;
    destructor Destroy; override;
    function Count: Integer;
    function KeyConvert(aCode: Integer; wPar: WPARAM; lPar: LPARAM): Boolean;
    procedure Add(aClassName: string; aKeyControllerList: TKeyControllerList);
    procedure Clear;
    procedure LoadFromStream(aStream: TStream);
    procedure SaveToStream(aStream: TStream);
    property Items[Index: Integer]: TKeyApplication read GetItems;
  end;

implementation

uses System.SysUtils, UnitUtils;

//**********************************************************************************************************************
{$REGION '// TKeyController '}
//**********************************************************************************************************************

constructor TKeyController.Create;
begin
  Self.FControlKey.Key := 0;
  Self.FControlKey.State := [];
  FOperationKeyList := TList<TKeyCombination>.Create;
end;

destructor TKeyController.Destroy;
begin
  FOperationKeyList.Free;
  inherited;
end;

function TKeyController.OperationKeyCount: Integer;
begin
  Result := Self.FOperationKeyList.Count;
end;

procedure TKeyController.Assigne(aSource: TKeyController);
begin
  Self.FControlKey := aSource.ControlKey;
  for var Cnt := 0 to aSource.OperationKeyCount-1 do begin
    Self.AddOperationKey(aSource.OperationKey[Cnt]);
  end;
end;

procedure TKeyController.Clear;
begin
  Self.FOperationKeyList.Clear;
end;

procedure TKeyController.AddOperationKey(aCombinationKey: TKeyCombination);
begin
  Self.FOperationKeyList.Add(aCombinationKey);
end;

function TKeyController.GetControlKey: TKeyCombination;
begin
  Result := Self.FControlKey;
end;

function TKeyController.GetOperationKey(Index: Integer): TKeyCombination;
begin
  Result := Self.FOperationKeyList.Items[Index];
end;

procedure TKeyController.MoveOperationKey(CurIndex, NewIndex: NativeInt);
begin
  Self.FOperationKeyList.Move(CurIndex, NewIndex);
end;

procedure TKeyController.SetControlKey(const Value: TKeyCombination);
begin
  Self.FControlKey := Value;
end;

procedure TKeyController.SetOperationKey(Index: Integer; const Value: TKeyCombination);
begin
  Self.FOperationKeyList.Items[Index] := Value;
end;

procedure TKeyController.ProcReadModifierStateFromStream(out aState: TModifierState; aStream: TStream);
var
  FState: Byte;
begin
  aStream.ReadBuffer(FState, SizeOf(FState));
  aState := [];
  if (FState and 1) <> 0 then Include(aState, tksShift);
  if (FState and 2) <> 0 then Include(aState, tksCtrl);
  if (FState and 4) <> 0 then Include(aState, tksNonConvert);
end;

procedure TKeyController.ProcWriteModifierStateToStream(aState: TModifierState; aStream: TStream);
begin
  var FState: Byte := 0;
  if tksShift in aState then FState := FState or 1;
  if tksCtrl in aState then FState := FState or 2;
  if tksNonConvert in aState then FState := FState or 4;
  aStream.WriteBuffer(FState, SizeOf(FState));
end;

procedure TKeyController.ProcKeyDown(aKey: Byte);
begin
  //WriteDebugLogFile('C:\temp\log.txt', Format('KeyDown %d', [aKey]));
  keybd_event(aKey, MapVirtualKey(aKey, 0),  0, 0);
end;

procedure TKeyController.ProcKeyUp(aKey: Byte);
begin
  //WriteDebugLogFile('C:\temp\log.txt', Format('KeyUp %d', [aKey]));
  keybd_event(aKey, MapVirtualKey(aKey, 0),  KEYEVENTF_KEYUP, 0);
end;

procedure TKeyController.ProcModifierKeyUp(aState: TModifierState);
begin
  if tksShift in aState then ProcKeyUp(VK_SHIFT);
  if tksCtrl in aState then ProcKeyUp(VK_CONTROL);
  if tksNonConvert in aState then ProcKeyUp(VK_NONCONVERT);
end;

procedure TKeyController.ProcModifierKeyDown(aState: TModifierState);
begin
  if tksShift in aState then ProcKeyDown(VK_SHIFT);
  if tksCtrl in aState then ProcKeyDown(VK_CONTROL);
  if tksNonConvert in aState then ProcKeyDown(VK_NONCONVERT);
end;

function TKeyController.KeyConvert(aCombination: TKeyCombination; aKeyInfo: TKeyInfo; var aTime: UInt64): Boolean;
begin
  Result := False;
  if (aCombination.State = Self.FControlKey.State) and (aCombination.Key = Self.FControlKey.Key) then begin
    // アプリケーションで押された「修飾キー」＋「コントロールキー」が登録されているとき
    var FGetTime := GetTickCount64;
    if (FGetTime - aTime) > 50 then begin
      ProcModifierKeyUp(aCombination.State);
      for var Cnt := 0 to OperationKeyCount-1 do begin
        ProcModifierKeyDown(OperationKey[Cnt].State);
        ProcKeyDown(OperationKey[Cnt].Key);
        ProcKeyUp(OperationKey[Cnt].Key);
        ProcModifierKeyUp(OperationKey[Cnt].State);
      end;
      ProcModifierKeyDown(aCombination.State);
      aTime := FGetTime;
    end;
    Result := True;
  end;
end;

procedure TKeyController.LoadFromStream(aStream: TStream);
var
  FCount: Integer;
  F: TKeyCombination;
begin
  aStream.ReadBuffer(Self.FControlKey.Key, SizeOf(Self.FControlKey.Key));
  ProcReadModifierStateFromStream(Self.FControlKey.State, aStream);
  aStream.ReadBuffer(FCount, SizeOf(FCount));
  for var Cnt := 0 to FCount-1 do begin
    aStream.ReadBuffer(F.Key, SizeOf(F.Key));
    ProcReadModifierStateFromStream(F.State, aStream);
    Self.FOperationKeyList.Add(F);
  end;
end;

procedure TKeyController.SaveToStream(aStream: TStream);
var
  FCount: Integer;
begin
  aStream.WriteBuffer(Self.FControlKey.Key, SizeOf(Self.FControlKey.Key));
  ProcWriteModifierStateToStream(Self.FControlKey.State, aStream);
  FCount := Self.OperationKeyCount;
  aStream.WriteBuffer(FCount, SizeOf(FCount));
  for var Cnt := 0 to FCount-1 do begin
    var F := Self.FOperationKeyList.Items[Cnt];
    aStream.WriteBuffer(F.Key, SizeOf(F.Key));
    ProcWriteModifierStateToStream(F.State, aStream);
  end;
end;

//**********************************************************************************************************************
{$ENDREGION}
//**********************************************************************************************************************

//**********************************************************************************************************************
{$REGION '// TKeyControllerList '}
//**********************************************************************************************************************

constructor TKeyControllerList.Create;
begin
  Self.FKeyControllerList := TObjectList<TKeyController>.Create;
end;

destructor TKeyControllerList.Destroy;
begin
  Self.FKeyControllerList.Free;
  inherited;
end;

function TKeyControllerList.GetItems(aIndex: Integer): TKeyController;
begin
  Result := Self.FKeyControllerList.Items[aIndex];
end;

procedure TKeyControllerList.Add(aKeyController: TKeyController);
begin
  Self.FKeyControllerList.Add(aKeyController);
end;

procedure TKeyControllerList.Assign(aSource: TKeyControllerList);
begin
  for var Cnt := 0 to aSource.Count-1 do begin
    var FKeyController := TKeyController.Create;
    FKeyController.Assigne(aSource.Items[Cnt]);
    Self.Add(FKeyController);
  end;
end;

procedure TKeyControllerList.Clear;
begin
  Self.FKeyControllerList.Free;
end;

function TKeyControllerList.Count: Integer;
begin
  Result := Self.FKeyControllerList.Count;
end;

function TKeyControllerList.KeyConvert(aCombination: TKeyCombination; aKeyInfo: TKeyInfo; var aTime: UInt64): Boolean;
begin
  // アプリケーションに登録されている
  Result := False;
  for var Cnt := 0 to Self.Count-1 do begin
    if Self.FKeyControllerList.Items[Cnt].KeyConvert(aCombination, aKeyInfo, aTime) then begin
      Result := True;
      Exit;
    end;
  end;
end;

procedure TKeyControllerList.LoadFromStream(aStream: TStream);
var
  FCount: Integer;
  F: TKeyController;
begin
  aStream.ReadBuffer(FCount, SizeOf(FCount));
  for var Cnt := 0 to FCount-1 do begin
    F := TKeyController.Create;
    F.LoadFromStream(aStream);
    Self.FKeyControllerList.Add(F);
  end;
end;

procedure TKeyControllerList.SaveToStream(aStream: TStream);
var
  FCount: Integer;
begin
  FCount := Self.Count;
  aStream.WriteBuffer(FCount, SizeOf(FCount));
  for var Cnt := 0 to FCount-1 do begin
    Self.FKeyControllerList.Items[Cnt].SaveToStream(aStream);
  end;
end;

//**********************************************************************************************************************
{$ENDREGION}
//**********************************************************************************************************************

//**********************************************************************************************************************
{$REGION '// TKeyCondition '}
//**********************************************************************************************************************

function TKeyCondition.GetIsExecuted: Boolean;
begin
  Result := False;
  //WriteDebugLogFile('c:\temp\log.txt',
  //  Format('FCount=%d Mask=%s Rep=%s Key=%d MaskKey=%d',
  //         [Self.FCount, BoolToStr(Self.FIsMasked, True), BoolToStr(Self.FIsRepeat, True), Self.FCombination.Key, Self.FMaskKey]));
  if Self.FCount <= 0 then begin
    if Self.IsPressed then begin
      Result := not Self.FIsMasked; // or (Self.FIsRepeat and (Self.FCombination.Key = Self.FMaskKey));
    end
    else begin
      if Self.FIsMasked and (Self.FCombination.Key = Self.FMaskKey) then begin
        Self.FIsMasked := False;
        Self.FMaskKey := 0;
      end;
    end;
  end;
end;

function TKeyCondition.GetInfo: TKeyInfo;
begin
  Result.IsPressed := Self.IsPressed;
  Result.IsRepeat := Self.FIsRepeat;
end;

function TKeyCondition.GetIsPressed: Boolean;
begin
  Result := Self.FIsPressed;
end;

function TKeyCondition.GetKeyCombination: TKeyCombination;
begin
  Result := Self.FCombination;
end;

function TKeyCondition.GetModifierState: TModifierState;
begin
  Result := Self.FCombination.State;
end;

procedure TKeyCondition.Initialize;
begin
  Self.FCombination.Key := 0;
  Self.FCombination.State := [];
  Self.FMaskKey := 0;
  Self.FIsMasked := False;
  Self.FIsPressed := False;
  Self.FIsRepeat := False;
  Self.FCount := 0;
end;

procedure TKeyCondition.SetModifierState(const Value: TModifierState);
begin
  Self.FCombination.State := Value;
end;

procedure TKeyCondition.StartMask(aCount: Integer);
begin
  Self.FIsMasked := True;
  Self.FMaskKey := Self.FCombination.Key;
  Self.FCount := aCount+1;
  //WriteDebugLogFile('C:\temp\log.txt', Format('[StartMask] IsMasked=True MaskKey=%d', [Self.FMaskKey]));
end;

procedure TKeyCondition.StopMask;
begin
//  WriteDebugLogFile('C:\temp\log.txt', Format('[StopMask] IsRepeat=%s Key=%d MaskKey=%d',
//                      [BoolToStr(FIsRepeat, True), FCombination.Key, FMaskKey]));
//  if not Self.FIsRepeat or (Self.FCombination.Key <> Self.FMaskKey) then begin
//    Self.FIsMasked := False;
//    Self.FMaskKey := 0;
//    WriteDebugLogFile('C:\temp\log.txt', '[StopMask] IsMasked=False');
//  end;
  if Self.FIsMasked and (Self.FCombination.Key = Self.FMaskKey) then begin
    Self.FIsMasked := False;
    Self.FMaskKey := 0;
    //WriteDebugLogFile('C:\temp\log.txt', '[StopMask] IsMasked=False');
  end;
end;

procedure TKeyCondition.ProcSetModifierKey(aModifierKey: TModifierState);
begin
  if Self.FIsPressed then
    Self.FCombination.State := Self.FCombination.State + aModifierKey  // Include(Self.FCombination.State, aModifierKey)
  else
    Self.FCombination.State := Self.FCombination.State - aModifierKey; // Exclude(Self.FCombination.State, aModifierKey)
end;

procedure TKeyCondition.ProcUpdateModifierState;
begin
  case Self.FCombination.Key of
    VK_CONTROL: ProcSetModifierKey([tksCtrl]);
    VK_SHIFT: ProcSetModifierKey([tksShift]);
    VK_NONCONVERT: ProcSetModifierKey([tksNonConvert]);
  end;
end;

procedure TKeyCondition.Update(wPar: WPARAM; lPar: LPARAM);

  function _State: Integer;
  begin
    Result := 0;
    if tksShift in Self.FCombination.State then Result := Result + 1;
    if tksCtrl in Self.FCombination.State then Result := Result + 2;
    if tksNonConvert in Self.FCombination.State then Result := Result + 4;
  end;

begin
  Self.FCombination.Key := wPar and $FF;
  Self.FIsPressed := ((lPar and (1 shl 31)) = 0); // lPar and 1000 0000 0000 0000 0000 0000 0000 0000
  Self.FIsRepeat := ((lPar and (1 shl 30)) >= 1); // lPar and 0100 0000 0000 0000 0000 0000 0000 0000
  ProcUpdateModifierState;
  if Self.FCount > 0 then Dec(Self.FCount);
  //WriteDebugLogFile('C:\temp\log.txt', Format('Key=%d IsPressed=%s IsRepeat=%s State=%d',
  //    [FCombination.Key, BoolToStr(FIsPressed, True), BoolToStr(FIsRepeat, True), _State]));
end;

//**********************************************************************************************************************
{$ENDREGION}
//**********************************************************************************************************************

//**********************************************************************************************************************
{$REGION '// TkeyApplicationList '}
//**********************************************************************************************************************

constructor TkeyApplicationList.Create;
begin
  Self.FTime := 0;
  Self.FKeyCondition.Initialize;
  Self.FKeyApplicationList := TList<TKeyApplication>.Create;
end;

destructor TkeyApplicationList.Destroy;
begin
  Self.Clear;
  Self.FKeyApplicationList.Free;
  inherited;
end;

function TkeyApplicationList.GetItems(Index: Integer): TKeyApplication;
begin
  Result := Self.FKeyApplicationList.Items[Index];
end;

procedure TkeyApplicationList.Add(aClassName: string; aKeyControllerList: TKeyControllerList);
var
  FKeyApplication: TKeyApplication;
begin
  FKeyApplication.ClassName := aClassName;
  FKeyApplication.KeyControllerList := aKeyControllerList;
  Self.FKeyApplicationList.Add(FKeyApplication);
end;

function TkeyApplicationList.Count: Integer;
begin
  Result := Self.FKeyApplicationList.Count;
end;

procedure TkeyApplicationList.Clear;
begin
  for var Cnt := 0 to Self.Count-1 do
    Self.FKeyApplicationList.Items[Cnt].KeyControllerList.Free;
  Self.FKeyApplicationList.Clear;
end;

function TkeyApplicationList.ProcGetActiveWindowClassName: string;
var
  FClassName: array[0..MAX_PATH-1] of Char;
begin
  Result := '';
  var FWnd := GetForegroundWindow;
  if FWnd <> 0 then begin
    FillChar(FClassName, SizeOf(FClassName), #0);
    GetClassName(FWnd, FClassName, SizeOf(FClassName));
    Result := FClassName;
  end;
end;

function TkeyApplicationList.KeyConvert(aCode: Integer; wPar: WPARAM; lPar: LPARAM): Boolean;

  function IntToBin(x: Integer): string;
  const
    N_BYTE: Integer = 4;
  begin
    Result := StringOfChar(' ', (N_BYTE*10-1));
    for var Cnt := N_BYTE*8 downto 1 do begin
      if ((x and 1) = 1) then
        Result[Cnt+((Cnt-1) shr 2)] := '1'
      else
        Result[Cnt+((Cnt-1) shr 2)] := '0';
      x := x shr 1;
    end;
  end;

begin
  Result := False;
  for var Cnt := 0 to Self.Count-1 do begin
    if (aCode = HC_ACTION) and (ProcGetActiveWindowClassName = Self.Items[Cnt].ClassName) then begin
      // キーを叩いたアプリが登録アプリケーションのとき
      FKeyCondition.Update(wPar, lPar);
      //WriteDebugLogFile('C:\temp\log.txt', Format('%s %d[%s] %s'
      //      , [FormatDateTime('hh:ss.zz', Now), wPar, BoolToStr(FKeyCondition.IsPressed, True), IntToBin(lPar)]));
      var FCount := 0;
      if FKeyCondition.IsPressed
          and Self.Items[Cnt].KeyControllerList.KeyConvert(FKeyCondition.KeyCombination, FKeyCondition.Info, FTime) then begin
        Result := True;
        Exit;
      end;
    end;
  end;
end;

procedure TkeyApplicationList.LoadFromStream(aStream: TStream);
var
  FCount: Integer;
  F: TKeyApplication;
begin
  var FState := Self.FKeyCondition.ModifierState;
  aStream.ReadBuffer(FState, SizeOf(FState));
  aStream.ReadBuffer(FCount, SizeOf(FCount));
  for var Cnt := 0 to FCount-1 do begin
    F.ClassName := ReadStreamString(aStream);
    F.KeyControllerList := TKeyControllerList.Create;
    F.KeyControllerList.LoadFromStream(aStream);
    Self.FKeyApplicationList.Add(F);
  end;
end;

procedure TkeyApplicationList.SaveToStream(aStream: TStream);
var
  FCount: Integer;
  F: TKeyApplication;
begin
  var FState := Self.FKeyCondition.ModifierState;
  aStream.WriteBuffer(FState, SizeOf(FState));
  FCount := Self.Count;
  aStream.WriteBuffer(FCount, SizeOf(FCount));
  for var Cnt := 0 to FCount-1 do begin
    F := Self.FKeyApplicationList.Items[Cnt];
    WriteStreamString(aStream, F.ClassName);
    F.KeyControllerList.SaveToStream(aStream);
  end;
end;

//**********************************************************************************************************************
{$ENDREGION}
//**********************************************************************************************************************

end.
