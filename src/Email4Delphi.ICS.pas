unit Email4Delphi.ICS;

interface

uses
  Email4Delphi.Contract, System.Classes, System.SysUtils, System.StrUtils,
  OverbyteIcsWndControl, OverbyteIcsSmtpProt, OverbyteIcsSslBase, OverbyteIcsWSocket;

type
  TSendEmailICS = class(TInterfacedObject, ISendEmail)
  private
    FSslContext: TSslContext;
    FsslSmtp: TSslSmtpCli;
    FListTo: TStringList;
    FListCC: TStringList;
    FListBCC: TStringList;
    FLogExecute: TProc<string>;
    FEhloCount: Integer;

    function ClearAttachment: ISendEmail;
    function AddAttachment(const AString: string): ISendEmail;
    function AddBCC(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddCC(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddReplyTo(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddTo(const AEmail: string; const AName: string = ''): ISendEmail;
    function Auth(const AValue: Boolean): ISendEmail;
    function Clear: ISendEmail;
    function ClearRecipient: ISendEmail;
    function Connect: ISendEmail;
    function Disconnect: ISendEmail;
    function From(const AEmail: string; const AName: string = ''): ISendEmail;
    function Host(const AHost: string): ISendEmail;
    function IsConnected: boolean;
    function AddMessage(const AMessage: string; const IsBodyHTML: Boolean = True): ISendEmail;
    function Password(const APassword: string): ISendEmail;
    function Port(const APort: Integer): ISendEmail;
    function Send: ISendEmail;
    function SSL(const AValue: Boolean): ISendEmail; overload;
    function Subject(const ASubject: string): ISendEmail;
    function TLS(const AValue: Boolean): ISendEmail;
    function UserName(const AUserName: string): ISendEmail;
    function ConfirmReceipt(const AConfirmReceipt: Boolean): ISendEmail;
    function OnLog(const AExecute: TProc<string>): ISendEmail;
    function Priority(const APriority: TMessagePriority): ISendEmail;
    function TMessagePriorityToTSmtpPriority(const APriority: TMessagePriority): TSmtpPriority;
    function SetAgent(const AAgent: String): ISendEmail;

    procedure SmtpClientDisplay(Sender: TObject; Msg: string);
    procedure SslSmtpClientBeforeFileOpen(Sender: TObject;Idx: Integer; FileName: String; var Action: TSmtpBeforeOpenFileAction);
    procedure SslSmtpClientRequestDone(Sender: TObject; RqType: TSmtpRequest; Error: Word);
  public
    constructor Create;
    class function New: ISendEmail;
    destructor Destroy; override;
  end;

implementation

function TSendEmailICS.AddAttachment(const AString: string): ISendEmail;
begin
  Result := Self;
  FsslSmtp.EmailFiles.Add(AString.Trim);
end;

function TSendEmailICS.AddBCC(const AEmail: string; const AName: string = ''): ISendEmail;
begin
  Result := Self;
  FListBCC.Add(Format('"%s" <%s>', [AName, AEmail]));
end;

function TSendEmailICS.AddCC(const AEmail: string; const AName: string = ''): ISendEmail;
begin
  Result := Self;
  FListCC.Add(Format('"%s" <%s>', [AName, AEmail]));
end;

function TSendEmailICS.AddReplyTo(const AEmail: string; const AName: string = ''): ISendEmail;
begin
  Result := Self;
  FsslSmtp.HdrReplyTo := Format('"%s" <%s>', [AName, AEmail]);
end;

function TSendEmailICS.AddTo(const AEmail: string; const AName: string = ''): ISendEmail;
begin
  Result := Self;
  FListTo.Add(Format('"%s" <%s>', [AName, AEmail]));
end;

function TSendEmailICS.Clear: ISendEmail;
begin
  Result := Self;
  FListTo.Clear;
  FsslSmtp.HdrFrom := '';
  FsslSmtp.HdrTo := '';
  FsslSmtp.HdrCc := '';
  FsslSmtp.HdrReplyTo := '';
  FsslSmtp.HdrReturnPath := '';
  FsslSmtp.HdrSubject := '';
  FsslSmtp.HdrSender := '';
  FsslSmtp.FromName := '';
  FsslSmtp.RcptName.Clear;
  FsslSmtp.EmailFiles.Clear;
end;

function TSendEmailICS.ClearAttachment: ISendEmail;
begin
  Result := Self;
  FListTo.Clear;
  FsslSmtp.EmailFiles.Clear;
end;

function TSendEmailICS.ClearRecipient: ISendEmail;
begin
  Result := Self;
  FsslSmtp.HdrFrom := '';
  FsslSmtp.HdrTo := '';
  FsslSmtp.HdrCc := '';
  FsslSmtp.HdrReplyTo := '';
  FsslSmtp.HdrReturnPath := '';
  FsslSmtp.HdrSubject := '';
  FsslSmtp.HdrSender := '';
  FsslSmtp.FromName := '';
  FsslSmtp.RcptName.Clear;
end;

function TSendEmailICS.ConfirmReceipt(const AConfirmReceipt: Boolean): ISendEmail;
begin
  Result := Self;
  FsslSmtp.ConfirmReceipt := AConfirmReceipt;
end;

function TSendEmailICS.Connect: ISendEmail;
begin
  Result := Self;
  FsslSmtp.ConnectSync;
end;

constructor TSendEmailICS.Create;
begin
  {$REGION 'CONFIG FSslSmtpCli e FSslContext'}
  FSslContext := TSslContext.Create(nil);
  FsslSmtp := TSslSmtpCli.Create(nil);
  FsslSmtp.OnRequestDone := SslSmtpClientRequestDone;
  FsslSmtp.OnDisplay := SmtpClientDisplay;
  FsslSmtp.OnBeforeFileOpen := SslSmtpClientBeforeFileOpen;
  FsslSmtp.CharSet := 'utf-8';
  self.SetAgent('Email4Delphi');
  FsslSmtp.ContentType := smtpPlainText;
  FsslSmtp.WrapMsgMaxLineLen := 76;
  FsslSmtp.WrapMessageText := False;
  FsslSmtp.SslContext := FSslContext;
  {$ENDREGION}

  FEhloCount := 0;
  FLogExecute := nil;

  {$REGION 'Config Stringlist CC, BCC, LOG'}
  FListCC := TStringList.Create;
  FListCC.Delimiter := ';';

  FListTo := TStringList.Create;
  FListTo.Delimiter := ';';

  FListBCC := TStringList.Create;
  FListBCC.Delimiter := ';';
  {$ENDREGION}
end;

destructor TSendEmailICS.Destroy;
begin
  FLogExecute := nil;
  FreeAndNil(FSslContext);
  FreeAndNil(FsslSmtp);
  FreeAndNil(FListTo);
  FreeAndNil(FListBCC);
  FreeAndNil(FListCC);
  inherited;
end;

function TSendEmailICS.Auth(const AValue: Boolean): ISendEmail;
begin
  result := Self;
  if AValue then
    FsslSmtp.AuthType := TSmtpAuthType(2)
  else
    FsslSmtp.AuthType := TSmtpAuthType(0);
end;

function TSendEmailICS.From(const AEmail: string; const AName: string = ''): ISendEmail;
begin
  Result := Self;
  FsslSmtp.HdrFrom := Format('"%s" <%s>', [AName, AEmail]);
  FsslSmtp.FromName := FsslSmtp.HdrFrom;
end;

function TSendEmailICS.Disconnect: ISendEmail;
begin
  Result := Self;
  if Self.IsConnected then
    FsslSmtp.Quit;
end;

function TSendEmailICS.IsConnected: boolean;
begin
  Result := FsslSmtp.Connected;
end;

function TSendEmailICS.Host(const AHost: string): ISendEmail;
begin
  Result := Self;
  if not AHost.IsEmpty then
    FsslSmtp.Host := AHost;
end;

function TSendEmailICS.AddMessage(const AMessage: string; const IsBodyHTML: Boolean = True): ISendEmail;
begin
  Result := Self;
  FsslSmtp.MailMessage.Add(AMessage);
  if IsBodyHTML then
    FsslSmtp.ContentType := smtpHtml;
end;

class function TSendEmailICS.New: ISendEmail;
begin
  Result := TSendEmailICS.Create;
end;

function TSendEmailICS.OnLog(const AExecute: TProc<string>): ISendEmail;
begin
  Result := Self;
  FLogExecute := AExecute;
end;

function TSendEmailICS.Password(const APassword: string): ISendEmail;
begin
  Result := Self;
  if not APassword.IsEmpty then
    FsslSmtp.Password := APassword;
end;

function TSendEmailICS.Port(const APort: Integer): ISendEmail;
begin
  Result := Self;
  FsslSmtp.Port := IntToStr(APort);
end;

function TSendEmailICS.Priority(const APriority: TMessagePriority): ISendEmail;
begin
  if FsslSmtp.HdrPriority = TMessagePriorityToTSmtpPriority(APriority) then
    exit;

  FsslSmtp.HdrPriority := TMessagePriorityToTSmtpPriority(APriority);
end;

function TSendEmailICS.Send: ISendEmail;
begin
  Result := Self;
  FsslSmtp.RcptName.Clear;
  FsslSmtp.RcptNameAdd(FListTo.DelimitedText, FListCC.DelimitedText, FListBCC.DelimitedText);

  try
    if not Self.IsConnected then
      FsslSmtp.ConnectSync;
  except
    raise Exception.Create('Error during connection.');
  end;
end;

function TSendEmailICS.SetAgent(const AAgent: String): ISendEmail;
begin
  Result := Self;
  if not (AAgent.IsEmpty) then
    FsslSmtp.XMailer := AAgent;
end;

procedure TSendEmailICS.SmtpClientDisplay(Sender: TObject; Msg: string);
begin
  if Assigned(FLogExecute) then
    FLogExecute(Msg);
end;

function TSendEmailICS.SSL(const AValue: Boolean): ISendEmail;
begin
  Result := Self;
  if AValue then
    FsslSmtp.SslType := TSmtpSslType(1)
  else
    FsslSmtp.SslType := TSmtpSslType(0);
end;

procedure TSendEmailICS.SslSmtpClientBeforeFileOpen(Sender: TObject; Idx: Integer; FileName: String; var Action: TSmtpBeforeOpenFileAction);
begin
  Action := smtpBeforeOpenFileNone;
end;

procedure TSendEmailICS.SslSmtpClientRequestDone(Sender: TObject; RqType: TSmtpRequest; Error: Word);
begin
  if Error <> 0 then
  begin
    if Assigned(FLogExecute) then
      FLogExecute('RequestDone Rq=' + IntToStr(Ord(RqType)) + ' Error=' + FsslSmtp.ErrorMessage)
  end
  else
  begin
    if Assigned(FLogExecute) then
      FLogExecute('RequestDone Rq=' + IntToStr(Ord(RqType)) + ' Error=' + IntToStr(Error));
  end;

  if Error <> 0 then
  begin
    if Assigned(FLogExecute) then
      FLogExecute('Error, stopped Send');
    Exit;
  end;
  case RqType of
    smtpConnect:
      if FsslSmtp.AuthType = smtpAuthNone then
        FsslSmtp.Helo
      else
        FsslSmtp.Ehlo;
    smtpHelo: FsslSmtp.MailFrom;
    smtpEhlo:
      if FsslSmtp.SslType = smtpTlsExplicit then
      begin
        Inc(FEhloCount);
        if FEhloCount = 1 then
          FsslSmtp.StartTls
        else if FEhloCount > 1 then
          FsslSmtp.Auth;
      end
      else
        FsslSmtp.Auth;
    smtpStartTls: FsslSmtp.Ehlo;
    smtpAuth: FsslSmtp.MailFrom;
    smtpMailFrom: FsslSmtp.RcptTo;
    smtpRcptTo: FsslSmtp.Data;
    smtpData: FsslSmtp.Quit;
    smtpQuit: if Assigned(FLogExecute) then FLogExecute('Send Quit!');
  end;
end;

function TSendEmailICS.Subject(const ASubject: string): ISendEmail;
begin
  Result := Self;
  if not ASubject.IsEmpty then
    FsslSmtp.HdrSubject := ASubject;
end;

function TSendEmailICS.TLS(const AValue: Boolean): ISendEmail;
begin
  Result := Self;
  if AValue then
    FsslSmtp.SslType := smtpTlsExplicit
  else
    FsslSmtp.SslType := smtpTlsNone;
end;

function TSendEmailICS.TMessagePriorityToTSmtpPriority(const APriority: TMessagePriority): TSmtpPriority;
begin
  case APriority of
    pHighest: Result := smtpPriorityHighest;
    pHigh:    Result := smtpPriorityHigh;
    pNormal:  Result := smtpPriorityNormal;
    pLow:     Result := smtpPriorityLow;
    pLowest:  Result := smtpPriorityLowest;
  else
    Result := smtpPriorityNormal;
  end;
end;

function TSendEmailICS.UserName(const AUserName: string): ISendEmail;
begin
  Result := Self;
  if not AUserName.IsEmpty then
    FsslSmtp.Username := AUserName;
end;

end.
