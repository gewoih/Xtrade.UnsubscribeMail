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
  Registry,
  DateUtils,
  Variants,
  
  imapsend,
  mimemess, 
  mimepart, 
  blcksock,
  synautil, 
  synachar, 
  SynaCode;

var
  Mail: TImapSend;
  Mime: TMimemess;
  reg: TRegistry;


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
  reg.OpenKey('\SOFTWARE\Xtrade\XTrade.UnsubscribeMail.exe\sql', false);
  Result := 'Provider='+reg.ReadString('Prov') + ';'; //SQLNCLI11;';
  Result := Result + 'Persist Security Info=False;';
  Result := Result + 'Data Source='+reg.ReadString('Serv') + ';';//192.168.44.100;';
  Result := Result + 'Initial Catalog='+reg.ReadString('Base') + ';';//vtk;';
  Result := Result + 'User ID='+reg.ReadString('User') + ';';//sa;';
  Result := Result + 'Application Name=' + ExtractFileName(ParamStr(0))+ ';';
  Result := Result + 'MultipleActiveResultSets=True;';
  Result := Result + 'Password='+reg.ReadString('Pass') + ';';//icq99802122;'
  reg.CloseKey;
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
	i: integer;
begin
    Result := false;
    for i:= 1 to Length(str) do
    begin
        if not (str[i] in ['a'..'z', 'A'..'Z', '0'..'9', '_', '-', '.']) then
        	Exit;
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
    
    if (Length(str) = 0) or ((Length(str) < 5)) then
    	Exit;
        
    i := Pos('.', str);
    if i = 0 then
    	Exit;
        
    Result := CheckAllowed(name_part) and CheckAllowed(server_part);
end;

class procedure tWalk.Walk(const Sender: TMimePart);
var
S, sql: string;
rst: 	OleVariant;
mail:	string;
begin
    mail := Sender.PartBody[0];

    if CheckMailFormat(mail) = false then
    begin
        Debug('Неверный формат почты', '');
        Exit;
    end;
    
    with Sender do
    begin
        DecodePart;
//        if not (SameText(Disposition,'attachment') or SameText(Disposition,'inline')) then Exit;
//        if Length(Filename)<>24 then Exit;
//        if AnsiCompareText(ExtractFileExt(FileName),'.csv')<>0 then Exit;
//        if Encoding = 'BASE64' then S := DecodeBase64(PartBody.Text) else S := PartBody.Text;
//        sql := format('select dbo.AsStr(id, 6784, 0) ss from spr_data where nid=%s and metaid=1277', [copy(FileName,1,12)]);
//        rst := fCon.Execute(sql);
//        try
//            WriteLog('Имя файла ' + FileName);
//            Script.Clear;
//            Script.Parent := fsGlobalUnit;
//            InitFunctions(Script);
//            Script.Lines.Text := rst.Fields[0].Value;
//            Script.Variables['regid'] := copy(FileName, 1, 12);
//            Script.Variables['data_op'] := copy(FileName, 13, 8);
//            Script.Variables['file_content'] := S;
//
//            if Script.Compile then
//            begin
//                Script.Execute;
//                sql := Script.Variables['sql'];
//                fCon.Execute(sql);
//            end
//            else
//				WriteLog('Не компилируется ' + Script.ErrorMsg + ' ' + Script.ErrorPos);
//            except
//            on E:Exception do
//            begin
//                WriteLog('Исключение ' + Filename + ' ' + Mime.Header.From +
//                ' ' + DateToStr(Mime.Header.Date) + Script.ErrorMsg + ' ' + Script.ErrorPos);
//                if IsConsole then ReadLn;
//            end;
//        end;
    end;
    //sl.Free;
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
        try
          reg := TRegistry.Create;
          reg.RootKey := HKEY_CURRENT_USER;

          if not reg.OpenKey('\SOFTWARE\Xtrade\XTrade.UnsubscribeMail.exe\pop3', false)
          then raise Exception.Create('Key POP3 not found');

          Mail := TImapSend.Create;
          Mime := TMimeMess.Create;

          Debug('TargetHost', reg.ReadString('TargetHost'));
          Debug('UserName', reg.ReadString('UserName'));
          Debug('Password', reg.ReadString('Password'));

          Mail.TargetHost:=reg.ReadString('TargetHost');
          Mail.UserName:=reg.ReadString('UserName');
          Mail.Password:=reg.ReadString('Password');
          Mail.AutoTLS:=False;
          Mail.TargetPort:='143';
        finally
          reg.CloseKey;
        end;

        try
            reg := TRegistry.Create;
            reg.RootKey := HKEY_CURRENT_USER;

            if not reg.OpenKey('\SOFTWARE\Xtrade\XTrade.UnsubscribeMail.exe', false)
            then raise Exception.Create('Key not found');

            Debug('LoopDelay', reg.ReadInteger('LoopDelay'));
            fLoopDelay := reg.ReadInteger('LoopDelay');
        finally
        	reg.CloseKey;
        end;

        except
        on E: exception do
        begin
            Debug('Error', E.Message);
            if IsConsole then
            begin
                WriteLn('Press ENTER to exit');
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
                        //Mail.DeleteMess(I);
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
    reg.free;
    CoUninitialize;
end;

initialization
	FormatSettings.DecimalSeparator := '.';
end.

