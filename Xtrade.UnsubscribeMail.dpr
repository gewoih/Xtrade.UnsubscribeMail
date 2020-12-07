program Xtrade.UnsubscribeMail;

{$APPTYPE CONSOLE}
{$DEFINE NOFORMS}
{$R *.res}

uses
  System.SysUtils,
  uxService in 'uxService.pas',
  uxLogWriter in 'uxLogWriter.pas',
  uxServer in 'uxServer.pas';

begin
  ReportMemoryLeaksOnShutdown := True;
  PrepareProcessParams(SvcStart, SvcLoop, SvcStop);
end.
