unit RtfToken;

interface

type

  { TRtfToken }

  TRtfTokenType = (ttNone, ttKeyword, ttControl, ttText, ttEof, ttGroupStart,
    ttGroupEnd, ttWhitespace);

  TRtfToken = class
  private
    FTokenType: TRtfTokenType;
    FKey: String;
    FHasParam: Boolean;
    FParam: Integer;
  public
    constructor Create;
    function ControlIsText: Boolean;
    function IsText: Boolean;
  public
    property TokenType: TRtfTokenType read FTokenType write FTokenType default ttNone;
    property Key: String read FKey write FKey;
    property HasParameter: Boolean read FHasParam write FHasParam default False;
    property Parameter: Integer read FParam write FParam;
  end;

implementation

uses
  RtfClasses;

{ TRtfToken }

constructor TRtfToken.Create;
begin
  FTokenType := ttNone;
  FHasParam := False;
end;

function TRtfToken.ControlIsText: Boolean;
begin
  Result := (TokenType = ttControl) and (Key = TRtfChar.HexMarker);
end;

function TRtfToken.IsText: Boolean;
begin
  Result := TokenType = ttText;
end;

end.
