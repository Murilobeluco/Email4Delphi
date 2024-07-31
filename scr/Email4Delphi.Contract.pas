unit Email4Delphi.Contract;

interface

uses
  System.SysUtils, System.StrUtils, System.Classes;

type
  TMessagePriority = (pHighest, pHigh, pNormal, pLow, pLowest);
  TAttachmentDisposition = (adAttachment, adInline);
  TLogMode = (lmComponent, lmLib, lmAll, lmNone);

  ISendEmail = interface;

  ISendEmail = interface
    ['{1BD51BA3-7FD7-4034-AB44-B712B9A09EAF}']
    function From(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddTo(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddReplyTo(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddCC(const AEmail: string; const AName: string = ''): ISendEmail;
    function AddBCC(const AEmail: string; const AName: string = ''): ISendEmail;
    function Priority(const APriority: TMessagePriority): ISendEmail;
    function SetAgent(const AAgent: String): ISendEmail;
    function AddMessage(const AMessage: string; const IsBodyHTML: Boolean = True): ISendEmail;

    {$IF DEFINED(EMAIL4D_INDY)}
    function AddReceiptRecipient(const AEmail: string; const AName: string = ''): ISendEmail;
    function OnWorkBegin(const AExecute: TProc<Int64>): ISendEmail;
    function OnWork(const AExecute: TProc<Int64>): ISendEmail;
    function OnWorkEnd(const AExecute: TProc): ISendEmail;
    function SendAsync(const ACallBack: TProc<Boolean, string> = nil; const ADisconnectAfterSending: Boolean = True): ISendEmail;

    function AddAttachment(const AFileName: string; const ADisposition: TAttachmentDisposition = adAttachment): ISendEmail; overload;
    function AddAttachment(const AData: TStream; const AFileName: string; const ADisposition: TAttachmentDisposition = adAttachment): ISendEmail; overload;

    function OnLog(const AExecute: TProc<string>; const ALogMode: TLogMode = lmComponent): ISendEmail;
    function Send(const ADisconnectAfterSending: Boolean = True): ISendEmail;
    {$ENDIF}

    {$IF DEFINED(EMAIL4D_ICS)}
    function AddAttachment(const AString: string): ISendEmail;
    function ClearAttachment: ISendEmail;
    function ConfirmReceipt(const AConfirmReceipt:Boolean): ISendEmail;
    function OnLog(const AExecute: TProc<string>): ISendEmail;
    function Send: ISendEmail;
    {$ENDIF}

    function SSL(const AValue: Boolean): ISendEmail;
    function TLS(const AValue: Boolean): ISendEmail;
    function Auth(const AValue: Boolean): ISendEmail;
    function Subject(const ASubject: string): ISendEmail;
    function Host(const AHost: string): ISendEmail;
    function Port(const APort: Integer): ISendEmail;
    function UserName(const AUserName: string): ISendEmail;
    function Password(const APassword: string): ISendEmail;
    function Clear: ISendEmail;
    function ClearRecipient: ISendEmail;
    function Connect: ISendEmail;
    function Disconnect: ISendEmail;
    function IsConnected: boolean;
  end;


implementation


end.
