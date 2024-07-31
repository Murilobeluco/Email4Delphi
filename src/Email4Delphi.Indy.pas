unit Email4Delphi.Indy;

interface

uses
  Email4Delphi.Contract, System.SysUtils, System.StrUtils, System.Classes, IdSMTP,
  IdSSL, IdSSLOpenSSL, IdSSLOpenSSLHeaders, IdExplicitTLSClientServerBase,
  IdIOHandler, IdMessage, IdMessageBuilder, idGlobal, IdComponent,
  IdAttachmentFile, System.TypInfo, System.Threading;

type
  TSendEmailIndy = class(TInterfacedObject, ISendEmail)
  private
    FIdSMTP: TIdSMTP;
    FIdSSLOpenSSL: TIdSSLIOHandlerSocketOpenSSL;
    FIdMessage: TIdMessage;
    FIdMessageBuilderHTML: TIdMessageBuilderHtml;
    FSSL: Boolean;
    FTLS: Boolean;
    FLogExecute: TProc<string>;
    FLogMode: TLogMode;
    FWorkBegin: TProc<Int64>;
    FWork: TProc<Int64>;
    FWorkEnd: TProc;
    FMessageStream: TMemoryStream;

    FConnectMaxReconnection: Integer;
    FConnectCountReconnect: Integer;
    FSendMaxReconnection: Integer;
    FSendCountReconnect: Integer;

    class var RW: TMultiReadExclusiveWriteSynchronizer;
    class var FInstance: ISendEmail;

    procedure Reconnect(AResend: Boolean = False);

    procedure LogSMTPStatus(ASender: TObject; const AStatus: TIdStatus; const AStatusText: string);
    procedure LogSSLStatus(const AMsg: string);
    procedure Log(const ALog: string; const AForced: Boolean = False);

    procedure WorkBegin(ASender: TObject; AWorkMode: TWorkMode; AWorkCountMax: Int64);
    procedure Work(ASender: TObject; AWorkMode: TWorkMode; AWorkCount: Int64);
    procedure WorkEnd(ASender: TObject; AWorkMode: TWorkMode);

    function AddAttachment(const AFileName: string; const ADisposition: TAttachmentDisposition = adAttachment): ISendEmail; overload;
    function AddAttachment(const AData: TStream; const AFileName: string; const ADisposition: TAttachmentDisposition = adAttachment): ISendEmail; overload;
    function AddBCC(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddCC(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddReceiptRecipient(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddReplyTo(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddTo(const AEmail: string; const AName: string = ''): ISendEmail;
    function Auth(const AValue: Boolean): ISendEmail;
    function Clear: ISendEmail;
    function ClearRecipient: ISendEmail;
    function Connect: ISendEmail;
    function Disconnect: ISendEmail;
    function From(const AEmail: string; const AName: string = ''): ISendEmail;
    function Host(const AHost: string): ISendEmail;
    function IsConnected: Boolean;
    function AddMessage(const AMessage: string; const IsBodyHTML: Boolean = True): ISendEmail;
    function OnLog(const AExecute: TProc<string>; const ALogMode: TLogMode = lmComponent): ISendEmail;
    function OnWork(const AExecute: TProc<Int64>): ISendEmail;
    function OnWorkBegin(const AExecute: TProc<Int64>): ISendEmail;
    function OnWorkEnd(const AExecute: TProc): ISendEmail;
    function Password(const APassword: string): ISendEmail;
    function Port(const APort: Integer): ISendEmail;
    function Send(const ADisconnectAfterSending: Boolean = True): ISendEmail;
    function SendAsync(const ACallBack: TProc<Boolean, string> = nil; const ADisconnectAfterSending: Boolean = True): ISendEmail;
    function SSL(const AValue: Boolean): ISendEmail;
    function Subject(const ASubject: string): ISendEmail;
    function TLS(const AValue: Boolean): ISendEmail;
    function UserName(const AUserName: string): ISendEmail;
    function TMessagePriorityToIdMessagePriority(const APriority: TMessagePriority): TIdMessagePriority;
    function Priority(const APriority: TMessagePriority): ISendEmail;
    function SetAgent(const AAgent: String): ISendEmail;
  public
    constructor Create;
    class function New: ISendEmail;
    destructor Destroy; override;
  end;

implementation

{ TSendEmail }

function TSendEmailIndy.AddAttachment(const AData: TStream; const AFileName: string; const ADisposition: TAttachmentDisposition): ISendEmail;
var
  LAdd: TIdMessageBuilderAttachment;
begin
  Result := Self;

  if not Assigned(AData) then
    Exit;

  if AFileName.Trim.IsEmpty then
    Exit;

  AData.Position := 0;
  LAdd := nil;

  case ADisposition of
    adAttachment:
      begin
        LAdd := FIdMessageBuilderHTML.Attachments.Add(AData, '', AFileName);
        Log(Format('Attachment(adAttachment)(%d): %s', [FIdMessageBuilderHTML.Attachments.Count, ExtractFileName(AFileName)]));
      end;

    adInline:
      begin
        LAdd := FIdMessageBuilderHTML.HtmlFiles.Add(AData, '', AFileName);
        Log(Format('Attachment(adInline)(%d): %s', [FIdMessageBuilderHTML.HtmlFiles.Count, ExtractFileName(AFileName)]));
      end;
  end;

  if not Assigned(LAdd) then
    Exit;

  LAdd.FileName := AFileName;
  LAdd.Name := ExtractFileName(AFileName);
  LAdd.WantedFileName := ExtractFileName(AFileName);
end;

function TSendEmailIndy.AddAttachment(const AFileName: string; const ADisposition: TAttachmentDisposition): ISendEmail;
begin
  Result := Self;

  if AFileName.Trim.IsEmpty then
    Exit;

  if not FileExists(AFileName) then
  begin
    Log('Attachment: Not found - ' + ExtractFileName(AFileName));
    Exit;
  end;

  case ADisposition of
    adAttachment:
      begin
        FIdMessageBuilderHTML.Attachments.Add(AFileName);
        Log(Format('Attachment(adAttachment)(%d): %s', [FIdMessageBuilderHTML.Attachments.Count, ExtractFileName(AFileName)]));
      end;

    adInline:
      begin
        FIdMessageBuilderHTML.HtmlFiles.Add(AFileName);
        Log(Format('Attachment(adInline)(%d): %s', [FIdMessageBuilderHTML.HtmlFiles.Count, ExtractFileName(AFileName)]));
      end;
  end;
end;

function TSendEmailIndy.AddBCC(const AEmail, AName: string): ISendEmail;
begin
  Result := Self;

  if AEmail.Trim.IsEmpty then
    Exit;

  with FIdMessage.BCCList.Add do
  begin
    name := AName;
    Address := AEmail;
  end;

  Log(Format('BCC(%d): %s', [FIdMessage.BCCList.Count, AEmail]));
end;

function TSendEmailIndy.AddCC(const AEmail, AName: string): ISendEmail;
begin
  Result := Self;

  if AEmail.Trim.IsEmpty then
    Exit;

  with FIdMessage.CCList.Add do
  begin
    name := AName;
    Address := AEmail;
  end;

  Log(Format('CC(%d): %s', [FIdMessage.CCList.Count, AEmail]));
end;

function TSendEmailIndy.AddReceiptRecipient(const AEmail, AName: string): ISendEmail;
begin
  Result := Self;

  if AEmail.Trim.IsEmpty then
    Exit;

  FIdMessage.ReceiptRecipient.Name := AName;
  FIdMessage.ReceiptRecipient.Address := AEmail;

  Log(Format('ReceiptRecipient: %s', [AEmail]));
end;

function TSendEmailIndy.AddReplyTo(const AEmail, AName: string): ISendEmail;
begin
  Result := Self;

  if AEmail.Trim.IsEmpty then
    Exit;

  with FIdMessage.ReplyTo.Add do
  begin
    name := AName;
    Address := AEmail;
  end;

  Log(Format('ReplyTo(%d): %s', [FIdMessage.ReplyTo.Count, AEmail]));
end;

function TSendEmailIndy.AddTo(const AEmail, AName: string): ISendEmail;
begin
  Result := Self;

  if AEmail.Trim.IsEmpty then
    Exit;

  with FIdMessage.Recipients.Add do
  begin
    name := AName;
    Address := AEmail;
  end;

  Log(Format('To(%d): %s', [FIdMessage.Recipients.Count, AEmail]));
end;

function TSendEmailIndy.Auth(const AValue: Boolean): ISendEmail;
begin
  Result := Self;

  if AValue then
    FIdSMTP.AuthType := satDefault
  else
    FIdSMTP.AuthType := satNone;

  Log(Format('Auth: %s', [IfThen(AValue, 'True', 'False')]));
end;

function TSendEmailIndy.Clear: ISendEmail;
begin
  Result := Self;
  Priority(TMessagePriority.pNormal);
  FIdMessage.ClearHeader;
  FIdMessage.Body.Clear;
  FIdMessage.MessageParts.Clear;
  FIdMessageBuilderHTML.Clear;
end;

function TSendEmailIndy.ClearRecipient: ISendEmail;
begin
  Result := Self;

  Priority(TMessagePriority.pNormal);
  FIdMessage.Recipients.Clear;
  FIdMessage.ReceiptRecipient.Text := '';
  FIdMessage.ReplyTo.Clear;
  FIdMessage.CCList.Clear;
  FIdMessage.BCCList.Clear;
  FIdMessage.Subject := '';
  FIdMessage.MessageParts.Clear;
  FIdMessageBuilderHTML.Clear;
end;

function TSendEmailIndy.Connect: ISendEmail;
var
  LLastResult: string;
begin
  Result := Self;

  if IsConnected then
    Exit;

  if FSSL or FTLS then
  begin
    if FSSL and FTLS then
      Log('Defining encryption: SSL/TLS')
    else
      if FSSL then
        Log('Defining encryption: SSL')
      else
        Log('Defining encryption: TLS');

    Log('Loading DLL');
    if not LoadOpenSSLLibrary then
    begin
      Log(WhichFailedToLoad);
      raise Exception.Create(Self.ClassName + ' > ' + WhichFailedToLoad);
    end;
    Log('Loaded DLL');

    if FSSL then
      FIdSSLOpenSSL.SSLOptions.Method := SslvSSLv23;

    if FTLS then
      FIdSSLOpenSSL.SSLOptions.Method := SslvTLSv1_2;

    FIdSSLOpenSSL.Destination := FIdSMTP.Host + ':' + FIdSMTP.Port.ToString;
    FIdSSLOpenSSL.Host := FIdSMTP.Host;
    FIdSSLOpenSSL.Port := FIdSMTP.Port;

    FIdSMTP.IOHandler := FIdSSLOpenSSL;

    if MatchText(FIdSMTP.Port.ToString, ['25', '587', '2587']) then
      FIdSMTP.UseTLS := utUseExplicitTLS
    else
      FIdSMTP.UseTLS := utUseImplicitTLS;
  end
  else
  begin
    Log('Defining encryption: None');
    FIdSMTP.IOHandler := nil;
    FIdSMTP.UseTLS := UtNoTLSSupport;
  end;

  try
    Log('Connecting');
    FIdSMTP.Connect;
    Log('Connected');

    if FIdSMTP.AuthType <> satNone then
    begin
      Log('Authenticating');
      FIdSMTP.Authenticate;
      Log('Authenticated');
    end;

    FConnectCountReconnect := 0;
  except
    on E: Exception do
    begin
      if FIdSMTP.LastCmdResult.ReplyExists then
      begin
        LLastResult := Format('Last Result: %s', [FIdSMTP.LastCmdResult.FormattedReply.Text]);

        if
          LLastResult.ToUpper.Contains('AUTHENTICATION SUCCEEDED') or
          LLastResult.Contains('250 OK')
        then
        begin
          Log(LLastResult, True);

          if FConnectCountReconnect < FConnectMaxReconnection then
          begin
            Sleep(100);
            Inc(FConnectCountReconnect);
            Reconnect(False);
            Exit;
          end
          else
          begin
            FConnectCountReconnect := 0;
            raise Exception.Create(E.Message);
          end;
        end;
      end;

      try
        if E.Message.ToUpper.Contains('INCORRECT AUTHENTICATION DATA') then
          raise Exception.Create('Incorrect authentication!');

        if E.Message.ToUpper.Contains('SOCKET ERROR # 10013') then
          raise Exception.Create('Firewall is blocking access to the internet!');

        if E.Message.ToUpper.Contains('SOCKET ERROR # 10054') then
          raise Exception.Create('Connection terminated!');

        if E.Message.ToUpper.Contains('SOCKET ERROR # 11001') then
          raise Exception.Create('Host not found!');

        if LLastResult.Contains(E.Message) and not LLastResult.Trim.IsEmpty then
          raise Exception.Create(LLastResult + sLineBreak + 'Message: ' + E.Message)
        else
          raise Exception.Create(E.Message);
      except
        on E: Exception do
        begin
          Log('Except: ' + E.Message, True);
          Log('Email not connected!');
          raise;
        end;
      end;
    end;
  end;
end;

constructor TSendEmailIndy.Create;
begin
  TSendEmailIndy.RW := TMultiReadExclusiveWriteSynchronizer.Create;

  FIdSMTP := TIdSMTP.Create;
  FIdSSLOpenSSL := TIdSSLIOHandlerSocketOpenSSL.Create;
  FIdMessageBuilderHTML := TIdMessageBuilderHtml.Create;
  FIdMessage := TIdMessage.Create;
  FMessageStream := TMemoryStream.Create;

  FLogExecute := nil;
  FLogMode := lmNone;

  FWorkBegin := nil;
  FWork := nil;
  FWorkEnd := nil;

  FConnectMaxReconnection := 5;
  FConnectCountReconnect := 0;
  FSendMaxReconnection := 5;
  FSendCountReconnect := 0;

  Self.Auth(True);
  Self.SSL(False);
  Self.TLS(False);
  Self.Priority(TMessagePriority.pNormal);
  Self.Port(587);

  {$REGION 'CONFIG FIdSMTP'}
  self.SetAgent('Email4Delphi');
  FIdSMTP.ConnectTimeout := 60000;
  FIdSMTP.ReadTimeout := 60000;
  FIdSMTP.UseEhlo := True;
  FIdSMTP.OnStatus := LogSMTPStatus;
  FIdSMTP.OnWorkBegin := WorkBegin;
  FIdSMTP.OnWork := Work;
  FIdSMTP.OnWorkEnd := WorkEnd;
  FIdSMTP.ManagedIOHandler := True;
  FIdSMTP.PipeLine := False;
  {$ENDREGION}

  {$REGION 'CONFIG FIdSSLOpenSSL'}
  FIdSSLOpenSSL.ConnectTimeout := 100000;
  FIdSSLOpenSSL.ReadTimeout := 100000;
  FIdSSLOpenSSL.PassThrough := True;
  FIdSSLOpenSSL.SSLOptions.SSLVersions := [SslvSSLv2, SslvSSLv23, SslvSSLv3, SslvTLSv1, SslvTLSv1_1, SslvTLSv1_2];
  FIdSSLOpenSSL.SSLOptions.Mode := SslmBoth;
  FIdSSLOpenSSL.SSLOptions.VerifyMode := [];
  FIdSSLOpenSSL.SSLOptions.VerifyDepth := 0;
  FIdSSLOpenSSL.OnStatus := LogSMTPStatus;
  FIdSSLOpenSSL.OnStatusInfo := LogSSLStatus;
  {$ENDREGION}

  {$REGION 'CONFIG FIdMessage'}
  FIdMessage.IsEncoded := True;
  FIdMessage.UseNowForDate := True;
  {$ENDREGION}
end;

destructor TSendEmailIndy.Destroy;
begin
  FWorkBegin := nil;
  FWork := nil;
  FWorkEnd := nil;
  FLogExecute := nil;

  FreeAndNil(FMessageStream);
  FreeAndNil(FIdMessageBuilderHTML);
  FreeAndNil(FIdMessage);
  FreeAndNil(FIdSSLOpenSSL);
  FreeAndNil(FIdSMTP);

  inherited;
end;

function TSendEmailIndy.Disconnect: ISendEmail;
begin
  Result := Self;

  if IsConnected then
    try
      Log('Disconnecting');
      FIdSMTP.Disconnect(False);
      Log('Disconnected');
    except
      Log('Except: Disconnected with error');
    end;
end;

function TSendEmailIndy.From(const AEmail, AName: string): ISendEmail;
begin
  Result := Self;

  if AEmail.Trim.IsEmpty then
    Exit;

  if IsConnected then
    if not FIdMessage.From.Address.Contains(AEmail) then
      Disconnect
    else
      Exit;

  FIdMessage.From.Name := AName;
  FIdMessage.From.Address := AEmail;

  Log(Format('From: %s', [AEmail]));
end;

function TSendEmailIndy.Host(const AHost: string): ISendEmail;
begin
  Result := Self;

  if IsConnected then
    if not FIdSMTP.Host.Contains(AHost) then
      Disconnect
    else
      Exit;

  FIdSMTP.Host := AHost;

  Log(Format('Host: %s', [AHost]));
end;


function TSendEmailIndy.IsConnected: Boolean;
begin
  Result := False;

  try
    Result := FIdSMTP.Connected;
  except
    on E: Exception do
    begin
      if E.Message.ToUpper.Contains('CLOSING CONNECTION') or
        E.Message.ToUpper.Contains('SSL3_GET_RECORD')
      then
        try
          Reconnect(False);
          Result := True;
        except
          Exit;
        end;
    end;
  end;
end;

procedure TSendEmailIndy.Log(const ALog: string; const AForced: Boolean);
begin
  if Assigned(FLogExecute) and ((FLogMode in [lmLib, lmAll]) or AForced) and not(FLogMode = lmNone) then
    FLogExecute('LIB: ' + ALog);
end;

procedure TSendEmailIndy.LogSMTPStatus(ASender: TObject; const AStatus: TIdStatus; const AStatusText: string);
begin
  if Assigned(FLogExecute) and (FLogMode in [lmComponent, lmAll]) then
    FLogExecute('SMTP: ' + AStatusText);
end;

procedure TSendEmailIndy.LogSSLStatus(const AMsg: string);
begin
  if Assigned(FLogExecute) and (FLogMode in [lmComponent, lmAll]) then
    FLogExecute('SSL: ' + AMsg.ToUpper.Replace('SSL STATUS: ', ''));
end;

function TSendEmailIndy.AddMessage(const AMessage: string; const IsBodyHTML: Boolean): ISendEmail;
begin
  Result := Self;

  if IsBodyHTML then
  begin
    FIdMessageBuilderHTML.Html.Text := AMessage;
    FIdMessageBuilderHTML.HtmlCharSet := 'utf-8';
    FIdMessageBuilderHTML.HtmlContentTransfer := 'base64';
  end
  else
  begin
    FIdMessageBuilderHTML.PlainText.Text := AMessage;
    FIdMessageBuilderHTML.PlainTextCharSet := 'utf-8';
    FIdMessageBuilderHTML.PlainTextContentTransfer := 'base64';
  end;

  Log(Format('Message: %s', [IfThen(IsBodyHTML, 'HTML', 'PlainText')]));
end;

class function TSendEmailIndy.New: ISendEmail;
begin
  Result := TSendEmailIndy.Create;
end;

function TSendEmailIndy.OnLog(const AExecute: TProc<string>; const ALogMode: TLogMode): ISendEmail;
begin
  Result := Self;

  FLogExecute := AExecute;
  FLogMode := ALogMode;
end;

function TSendEmailIndy.OnWork(const AExecute: TProc<Int64>): ISendEmail;
begin
  Result := Self;
  FWork := AExecute;
end;

function TSendEmailIndy.OnWorkBegin(const AExecute: TProc<Int64>): ISendEmail;
begin
  Result := Self;
  FWorkBegin := AExecute;
end;

function TSendEmailIndy.OnWorkEnd(const AExecute: TProc): ISendEmail;
begin
  Result := Self;
  FWorkEnd := AExecute;
end;

function TSendEmailIndy.Password(const APassword: string): ISendEmail;
begin
  Result := Self;

  if IsConnected then
    if not FIdSMTP.Password.Contains(APassword) then
      Disconnect
    else
      Exit;

  FIdSMTP.Password := APassword;

  Log(Format('Password: %s', ['********']));
end;


function TSendEmailIndy.Port(const APort: Integer): ISendEmail;
begin
  Result := Self;

  if IsConnected then
    if not FIdSMTP.Port = APort then
      Disconnect
    else
      Exit;

  FIdSMTP.Port := APort;

  Log(Format('Port: %d', [APort]));
end;

function TSendEmailIndy.Priority(const APriority: TMessagePriority): ISendEmail;
begin
  Result := Self;

  if FIdMessage.Priority = TMessagePriorityToIdMessagePriority(APriority) then
    Exit;

  FIdMessage.Priority := TMessagePriorityToIdMessagePriority(APriority);

  Log(Format('Priority: %s', [GetEnumName(TypeInfo(TMessagePriority), Integer(APriority))]));
end;

procedure TSendEmailIndy.Reconnect(AResend: Boolean);
begin
  if AResend then
    Log('Reconnecting: ' + FSendCountReconnect.ToString, True)
  else
    Log('Reconnecting: ' + FConnectCountReconnect.ToString, True);

  Disconnect;
  Connect;

  if AResend then
    Send;
end;

function TSendEmailIndy.Send(const ADisconnectAfterSending: Boolean): ISendEmail;
begin
  Result := Self;

  if not IsConnected then
    Connect;

  try
    try
      Log('Sending email');
      FIdMessageBuilderHTML.FillMessage(FIdMessage);
      FIdMessage.SaveToStream(FMessageStream);
      FIdSMTP.Send(FIdMessage);
      Log('Email sent');

      FSendCountReconnect := 0;
    except
      on E: Exception do
      begin
        Log('Except: ' + E.Message, True);

        if
          E.Message.ToUpper.Contains('CLOSING CONNECTION') or
          E.Message.ToUpper.Contains('TOO MANY MESSAGES') or
          E.Message.ToUpper.Contains('CONNECTION CLOSED')
        then
        begin
          if FSendCountReconnect < FSendMaxReconnection then
          begin
            Sleep(100);
            Inc(FSendCountReconnect);
            Reconnect(True);
            Exit;
          end
          else
          begin
            FSendCountReconnect := 0;
            raise Exception.Create(E.Message);
          end;
        end;

        if E.Message.ToUpper.Contains('NOT CONNECTED') then
          raise Exception.Create('Not connected to internet!');

        if
          E.Message.ToUpper.Contains('NO SUCH USER HERE') or
          E.Message.ToUpper.Contains('USER UNKNOWN') or
          E.Message.ToUpper.Contains('MAILBOX UNAVAILABLE')
        then
          raise Exception.Create('The recipient''s mailbox does not exist in the destination domain. It was probably typed incorrectly!');

        if
          E.Message.ToUpper.Contains('MAILBOX IS FULL') or
          E.Message.ToUpper.Contains('MAIL QUOTA EXCEEDED') or
          E.Message.ToUpper.Contains('MAILBOX FULL') or
          E.Message.ToUpper.Contains('DISK QUOTA EXCEEDED') or
          E.Message.ToUpper.Contains('USER IS OVER THE QUOTA')
        then
          raise Exception.Create('It means that the recipient''s inbox is full and cannot receive any more messages!');

        raise Exception.Create(E.Message);
      end;
    end;
  finally
    if ADisconnectAfterSending then
      Disconnect;
  end;
end;

function TSendEmailIndy.SendAsync(const ACallBack: TProc<Boolean, string>; const ADisconnectAfterSending: Boolean): ISendEmail;
begin
  Result := Self;

  TThread.CreateAnonymousThread(
    procedure
    var
      LMessage: string;
    begin
      RW.BeginWrite;
      try
        LMessage := '';

        try
          Send(ADisconnectAfterSending);
        except
          on E: Exception do
            LMessage := E.Message;
        end;
      finally
        RW.EndWrite;
      end;

      TThread.Synchronize(TThread.CurrentThread,
        procedure
        begin
          if Assigned(ACallBack) then
            ACallBack(not LMessage.Trim.IsEmpty, LMessage)
          else
            raise Exception.Create(LMessage);
        end);
    end).Start;
end;


function TSendEmailIndy.SetAgent(const AAgent: String): ISendEmail;
begin
  Result := Self;
  if not (AAgent.IsEmpty) then
    FIdSMTP.MailAgent := AAgent;
end;

function TSendEmailIndy.SSL(const AValue: Boolean): ISendEmail;
begin
  Result := Self;
  FSSL := AValue;
end;

function TSendEmailIndy.Subject(const ASubject: string): ISendEmail;
begin
  Result := Self;

  if not FIdMessage.Subject.Contains(ASubject) then
    FIdMessage.Subject := ASubject
  else
    Exit;

  Log(Format('Subject: %s', [ASubject]));
end;

function TSendEmailIndy.TLS(const AValue: Boolean): ISendEmail;
begin
  Result := Self;
  FTLS := AValue;
end;

function TSendEmailIndy.TMessagePriorityToIdMessagePriority(const APriority: TMessagePriority): TIdMessagePriority;
begin
  case APriority of
    pHighest: Result := mpHighest;
    pHigh:    Result := mpHigh;
    pNormal:  Result := mpNormal;
    pLow:     Result := mpLow;
    pLowest:  Result := mpLowest;
  else
    Result := mpNormal;
  end;
end;

function TSendEmailIndy.UserName(const AUserName: string): ISendEmail;
begin
  Result := Self;

  if IsConnected then
    if not FIdSMTP.UserName.Contains(AUserName) then
      Disconnect
    else
      Exit;

  FIdSMTP.UserName := AUserName;

  Log(Format('UserName: %s', [AUserName]));
end;

procedure TSendEmailIndy.Work(ASender: TObject; AWorkMode: TWorkMode; AWorkCount: Int64);
begin
  if Assigned(FWork) and (AWorkMode = wmWrite) then
    FWork(AWorkCount);
end;

procedure TSendEmailIndy.WorkBegin(ASender: TObject; AWorkMode: TWorkMode; AWorkCountMax: Int64);
begin
  Log('Starting transmission');

  if Assigned(FWorkBegin) and (AWorkMode = wmWrite) then
  begin
    FMessageStream.Position := 0;
    FWorkBegin(FMessageStream.Size);
  end;
end;

procedure TSendEmailIndy.WorkEnd(ASender: TObject; AWorkMode: TWorkMode);
begin
  Log('Transmission completed');

  FMessageStream.Clear;
  if Assigned(FWorkEnd) then
    FWorkEnd;
end;

end.

