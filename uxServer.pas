unit uxServer;

interface

uses
  SysUtils,
  Winsock,
  Windows,
  Messages,
  ComObj,
  ActiveX,
  Classes,
  AnsiStrings, 
  IniFiles,
  DateUtils,
  Variants,
  
  imapsend,
  mimemess, 
  mimepart, 
  blcksock,
  synautil, 
  synachar,
  EncdDecd,
  SynaCode;

var
  Mail: TImapSend;
  Mime: TMimemess;
  Params: TIniFile;


function SvcStart:			boolean;
function SvcLoop:			boolean;
function SvcStop:			boolean;

type
  tWalk = class
    class procedure Walk(const Sender: TMimePart);
  end;

implementation

uses
  uxLogWriter,
  uxService;

var fCon: oleVariant;

function GetConStr: string;
begin
  Result := 'Provider='+Params.ReadString('SQL', 'Prov', '') + ';'; //SQLNCLI11;';
  Result := Result + 'Persist Security Info=False;';
  Result := Result + 'Data Source='+Params.ReadString('SQL', 'Serv', '') + ';';//192.168.44.100;';
  Result := Result + 'Initial Catalog='+Params.ReadString('SQL', 'Base', '') + ';';//vtk;';
  Result := Result + 'User ID='+Params.ReadString('SQL', 'User', '') + ';';//sa;';
  Result := Result + 'Application Name=' + ExtractFileName(ParamStr(0))+ ';';
  Result := Result + 'MultipleActiveResultSets=True;';
  Result := Result + 'Password='+Params.ReadString('SQL','Pass', '') + ';';//icq99802122;'
end;

function ConnectSQL(Var Con:OleVariant): boolean;
begin
    try
        Con := CreateOleObject('ADODB.Connection');
        Con.CursorLocation:= 3;
        Con.CommandTimeout := 60000;
        Con.ConnectionTimeout := 10;
        Con.Open(GetConStr);
        Con.Execute('set nocount on');
        Con.Execute(Format('select %d as userid into #tuser', [0]));
        Result := True;

        except
        on E:Exception do
        begin
            Debug('SQL connect error', E.Message);
            Debug('Connection string', GetConStr);
            Result := False;
        end;
    end;
end;

function CheckAllowed(const str: string): boolean;
var
	c: char;
begin
    Result := false;
    for C in str do
    begin
        if C in [#13, #10, #160] then Continue;

        if not (C in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.']) then Exit;
    end;
    Result:= true;
end;

function CheckMailFormat(const str: string): boolean;
var
	i:	integer;
	name_part, server_part : string;
begin
	Result := false;
    i := Pos('@', str);
    
	if i = 0 then
    	Exit;
    
    name_part:= Copy(str, 1, i - 1);
    server_part:= Copy(str, i + 1, Length(str));
        
    i := Pos('.', str);
    if i = 0 then
    	Exit;
        
    Result := CheckAllowed(name_part) and CheckAllowed(server_part);
end;

class procedure tWalk.Walk(const Sender: TMimePart);
var
mail:	string;
begin
	if Sender.Secondary<>'HTML' then Exit;
    if Sender.Encoding <> 'BASE64' then
    	mail := Trim(Sender.PartBody.Text)
    else
    	mail := Trim(DecodeString(Sender.PartBody.Text));

    mail := StringReplace(mail, #$a0, '', [rfReplaceAll]);
    mail := StringReplace(mail, #$0a, '', [rfReplaceAll]);
    mail := StringReplace(mail, #$0d, '', [rfReplaceAll]);
    mail := StringReplace(mail, '=', '', [rfReplaceAll]);
    Debug('Почта ' + mail, ' хочет отписаться от рассылки');

    if mail = '' then
        Exit;

    if CheckMailFormat(mail) = false then
    begin
        Debug('Неверный формат почты', '');
        Exit;
    end;

    try
    	ConnectSQL(fCon);
		fCon.mng_unsubscribe(mail);
        Debug('Почта ' + mail + ' была успешно удалена из рассылки.', '');
    except
    On E: Exception do Debug('Exception: ', E.Message);
    end;
end;

function IsNull(A, B: variant):variant;
begin
  if VarIsNull(A) then Result := B else Result := A;
end;

function SvcStart: boolean;
begin
    CoInitializeEx(nil, 0);
    Result := True;
    try
        Mail := TImapSend.Create;
        Mime := TMimeMess.Create;

        Mail.TargetHost:=Params.ReadString('MAILBOX', 'TargetHost', '');
        Mail.UserName:=Params.ReadString('MAILBOX', 'UserName', '');
        Mail.Password:=Params.ReadString('MAILBOX', 'Password', '');
        Mail.AutoTLS:=False;
        Mail.TargetPort:=Params.ReadString('MAILBOX', 'Port', '143');

        fLoopDelay := Params.ReadInteger('COMMON', 'LoopDelay', 300000);

    except
        on E: exception do
        begin
            Debug('Error', E.Message);
            if IsConsole then
            begin
                Debug('Press ENTER to exit', '');
                Readln;
            end;
            Result := False;
        end;
    end;
end;

function SvcLoop: boolean;
var
	i:  integer;
begin
	var sx := TStringList.Create;

    Debug('Loop start', Now);
    try
        try
            if Mail.Login and ConnectSQL(fCon) then
            begin
                Mail.List('', sx);
                if Mail.SelectFolder('inbox') then
                begin
                    Debug('Mail count', Mail.SelectedCount);

                    for i := 1 to Mail.SelectedCount do
                    begin
                        Mime.Clear;
                        Mail.FetchMess(i, mime.Lines);
                        mime.DecodeMessage;
                        Mime.MessagePart.OnWalkPart := tWalk.Walk;
                        Mime.MessagePart.WalkPart;
                        Debug('Mail from', Mime.Header.From);
                        Mail.CopyMess(I, 'Trash');
                        Mail.DeleteMess(I);
                    end;

                    Mail.CloseFolder;
                end;

                if Mail.SelectFolder('Trash') then
                begin
                    for i := 1 to Mail.SelectedCount do
                    begin
                        Mail.FetchMess(i, mime.Lines);
                        mime.DecodeMessage;
                        if DaysBetween(Now, mime.Header.Date) > 10 then Mail.DeleteMess(I);
                    end;

                    Mail.CloseFolder;
                end;

                Mail.Logout;
                fCon.Close;
            end;
            finally
            	sx.Free;
        end;
    except
    On E: Exception do Debug('Loop error', E.Message);
    end;
end;

function SvcStop: boolean;
begin
    Mail.Free;
    Mime.Free;
    Params.free;
    CoUninitialize;
end;

initialization
  Params := TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini'));
	FormatSettings.DecimalSeparator := '.';
finalization
  Params.Free;
end.

