unit Email4Delphi;

interface

uses
  Email4Delphi.contract, system.Classes;


type
  ISendEmail = Email4Delphi.contract.ISendEmail;

  TSendemail = class
  public
    class function New: ISendEmail;
  end;


implementation

uses
  {$IF DEFINED(EMAIL4D_ICS)}
    Email4Delphi.ics;
  {$ELSEIF DEFINED(EMAIL4D_INDY)}
    Email4Delphi.indy;
  {$ELSE}
    Email4Delphi.indy;
  {$ENDIF}

class function TSendemail.New: ISendEmail;
begin
  {$IF DEFINED(EMAIL4D_ICS)}
    Result := TSendEmailICS.New;
  {$ELSEIF DEFINED(EMAIL4D_INDY)}
    Result := TSendEmailIndy.New;
  {$ELSE}
    Result := TSendEmailIndy.New;
  {$ENDIF}
end;

end.
