unit RtfLex;

interface

uses
  SysUtils, Classes, RtfToken, RtfClasses;

type

  { TRtfLex }

  TRtfLex = class;
  TRtfLexClasses = class of TRtfLex;

  TRtfLex = class
  private
    FStream: TStream;
    FChar: TRtfChar;
  protected
    function Read: Boolean;
    procedure MovePrev;
    procedure MoveNext;
    procedure ParseWhitespace(AToken: TRtfToken);
    procedure ParseKeyword(AToken: TRtfToken);
    procedure ParseText(AToken: TRtfToken);
  public
    constructor Create(AStream: TStream);
    destructor Destroy; override;
    function NextToken: TRtfToken;
  end;

var
  DefaultRtfLex: TRtfLexClasses = TRtfLex;

implementation

{ TRtfLex }

constructor TRtfLex.Create(AStream: TStream);
begin
  FStream := AStream;
  FChar := TRtfChar.Create;
end;

destructor TRtfLex.Destroy;
begin
  FreeAndNil(FChar);
  inherited;
end;

function TRtfLex.NextToken: TRtfToken;
begin
  Result := TRtfToken.Create;
  if Read then
  begin
    if FChar.IsBlockStart then
      Result.TokenType := ttGroupStart
    else if FChar.IsBlockEnd then
      Result.TokenType := ttGroupEnd
    else if FChar.IsWhitespace then
      ParseWhitespace(Result)
    else if FChar.IsKeywordMarker then
      ParseKeyword(Result)
    else
      ParseText(Result);
  end else
    Result.TokenType := ttEof;
end;

function TRtfLex.Read: Boolean;
begin
  Result := FChar.Read(FStream);
end;

procedure TRtfLex.MovePrev;
begin
  if FStream.Position > 0 then
    FStream.Position := FStream.Position - 1;
end;

procedure TRtfLex.MoveNext;
begin
  if FStream.Position < FStream.Size - 1 then
    FStream.Position := FStream.Position + 1;
end;

procedure TRtfLex.ParseWhitespace(AToken: TRtfToken);
begin
  AToken.TokenType := ttWhitespace;
  AToken.Key := '';
  while True do
  begin
    AToken.Key := AToken.Key + FChar.Char;
    if not Read or not FChar.IsWhitespace then
      Break;
  end;
  if not FChar.Eof then
    MovePrev;
end;

procedure TRtfLex.ParseKeyword(AToken: TRtfToken);
var
  AKey, ANb: String;
  C: Integer;

begin
  AKey := '';
  ANb := '';
  // Pick one character
  if Read then
  begin
    // is escape sequence for {, }, and \
    if FChar.IsEscapable then
    begin
      AToken.TokenType := ttText;
      AToken.Key := FChar.Char;
    end
    // is letter
    else if FChar.IsLetter then
    begin
      // pick for keyword key
      while True do
      begin
        AKey := AKey + FChar.Char;
        if not Read or not FChar.IsLetter then
          Break;
      end;
      AToken.TokenType := ttKeyword;
      AToken.Key := AKey;
      // pick for keyword parameter
      if not FChar.Eof and FChar.IsDigit then
      begin
        while True do
        begin
          ANb := ANb + FChar.Char;
          if not Read or not FChar.IsDigit then
            Break;
        end;
        AToken.HasParameter := True;
        AToken.Parameter := StrToIntDef(ANb, 0);
      end;
      // move back one character
      if not FChar.Eof and not FChar.IsChar(' ') then
        MovePrev;
    end
    else
    begin
      AToken.TokenType := ttControl;
      AToken.Key := FChar.Char;
      // is hexa character
      if FChar.IsHexEscapable then
      begin
        C := 2;
        while C > 0 do
        begin
          if not Read then
            Break;
          ANb := ANb + FChar.Char;
          Dec(C);
        end;
        if Length(ANb) = 2 then
        begin
          AToken.HasParameter := True;
          AToken.Parameter := StrToInt('$' + ANb);
        end;
      end
    end;
  end;
end;

procedure TRtfLex.ParseText(AToken: TRtfToken);
begin
  AToken.TokenType := ttText;
  AToken.Key := '';
  while True do
  begin
    AToken.Key := AToken.Key + FChar.Char;
    if not Read or FChar.IsEscapable or FChar.IsWhitespace then
      Break;
  end;
  if not FChar.Eof then
    MovePrev;
end;

end.
