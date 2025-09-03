{$WARNINGS OFF}
//=============================================================================
//  特権の設定用ユニット
//  DEKOさんの[特権を有効にする]をユニットにしたもの
//  http://ht-deko.minim.ne.jp/tech043.html#tech089
//
//-----------------------------------------------------------------------------
//
//  【履歴】
//
//  2013年07月10日 ～
//
//-----------------------------------------------------------------------------
//
//  【動作確認環境】
//
//  Windows 7 U64(SP1) + Delphi XE Pro
//
//  Presented by Mr.XRAY
//  http://mrxray.on.coocan.jp/
//=============================================================================
unit Privilege;

interface

uses
  Winapi.Windows, Vcl.Dialogs;

function SetPrivilege(szPrivilege: String; aEnabled: Boolean): Boolean;

const
  SE_CREATE_TOKEN_NAME            = 'SeCreateTokenPrivilege';
  SE_ASSIGNPRIMARYTOKEN_NAME      = 'SeAssignPrimaryTokenPrivilege';
  SE_LOCK_MEMORY_NAME             = 'SeLockMemoryPrivilege';
  SE_INCREASE_QUOTA_NAME          = 'SeIncreaseQuotaPrivilege';
  SE_UNSOLICITED_INPUT_NAME       = 'SeUnsolicitedInputPrivilege';
  SE_MACHINE_ACCOUNT_NAME         = 'SeMachineAccountPrivilege';
  SE_TCB_NAME                     = 'SeTcbPrivilege';
  SE_SECURITY_NAME                = 'SeSecurityPrivilege';
  SE_TAKE_OWNERSHIP_NAME          = 'SeTakeOwnershipPrivilege';
  SE_LOAD_DRIVER_NAME             = 'SeLoadDriverPrivilege';
  SE_SYSTEM_PROFILE_NAME          = 'SeSystemProfilePrivilege';
  SE_SYSTEMTIME_NAME              = 'SeSystemtimePrivilege';
  SE_PROF_SINGLE_PROCESS_NAME     = 'SeProfileSingleProcessPrivilege';
  SE_INC_BASE_PRIORITY_NAME       = 'SeIncreaseBasePriorityPrivilege';
  SE_CREATE_PAGEFILE_NAME         = 'SeCreatePagefilePrivilege';
  SE_CREATE_PERMANENT_NAME        = 'SeCreatePermanentPrivilege';
  SE_BACKUP_NAME                  = 'SeBackupPrivilege';
  SE_RESTORE_NAME                 = 'SeRestorePrivilege';
  SE_SHUTDOWN_NAME                = 'SeShutdownPrivilege';
  SE_DEBUG_NAME                   = 'SeDebugPrivilege';
  SE_AUDIT_NAME                   = 'SeAuditPrivilege';
  SE_SYSTEM_ENVIRONMENT_NAME      = 'SeSystemEnvironmentPrivilege';
  SE_CHANGE_NOTIFY_NAME           = 'SeChangeNotifyPrivilege';
  SE_REMOTE_SHUTDOWN_NAME         = 'SeRemoteShutdownPrivilege';
  SE_UNDOCK_NAME                  = 'SeUndockPrivilege';
  SE_SYNC_AGENT_NAME              = 'SeSyncAgentPrivilege';
  SE_ENABLE_DELEGATION_NAME       = 'SeEnableDelegationPrivilege';
  SE_MANAGE_VOLUME_NAME           = 'SeManageVolumePrivilege';
  SE_IMPERSONATE_NAME             = 'SeImpersonatePrivilege';
  SE_CREATE_GLOBAL_NAME           = 'SeCreateGlobalPrivilege';
  SE_TRUSTED_CREDMAN_ACCESS_NAME  = 'SeTrustedCredManAccessPrivilege';
  SE_RELABEL_NAME                 = 'SeRelabelPrivilege';
  SE_INC_WORKING_SET_NAME         = 'SeIncreaseWorkingSetPrivilege';
  SE_TIME_ZONE_NAME               = 'SeTimeZonePrivilege';
  SE_CREATE_SYMBOLIC_LINK_NAME    = 'SeCreateSymbolicLinkPrivilege';


implementation


// 特権を変更する
// -----------------------------------------------------------------------------
// szPrivilege: システムを指定する特権文字列
// aEnabled   : 特権の有効(True) / 無効(False) を指定
// 戻り値     : 成功なら True, 失敗なら False
// -----------------------------------------------------------------------------
function SetPrivilege(szPrivilege: String; aEnabled: Boolean): Boolean;
var
  TokenHandle: THandle;
  NewState: TTokenPrivileges;
  ReturnLength: Cardinal;
begin
  result := False;
  if OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES, TokenHandle) then
    begin
      if LookupPrivilegeValue(nil, PChar(szPrivilege), NewState.Privileges[0].Luid) then
        begin
          NewState.PrivilegeCount := 1;
          if aEnabled then
            begin
              NewState.Privileges[0].Attributes := SE_PRIVILEGE_ENABLED;
              result := AdjustTokenPrivileges(TokenHandle, False, NewState, SizeOf(NewState), nil, ReturnLength);
            end
          else
            begin
              NewState.Privileges[0].Attributes := 0;
              result := AdjustTokenPrivileges(TokenHandle, True, NewState, 0, nil, ReturnLength);
            end;
        end;
      CloseHandle(TokenHandle);
    end;
end;

end.
