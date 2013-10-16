unit RtfClasses;

interface

uses
  SysUtils, Classes, UITypes;

const
  TWIPS_UNIT    = 1440;

type

  { TRtfChar }

  TRtfChar = class
  strict private
    class var
      FBlockStart: Char;
      FBlockEnd: Char;
      FKeywordMarker: Char;
      FKeywordStop: Char;
      FHexMarker: Char;
  private
    FCh: Char;
    FEof: Boolean;
    class procedure Initialize;
  public
    function Read(Stream: TStream): Boolean;
    function IsBlockStart: Boolean; overload;
    function IsBlockEnd: Boolean; overload;
    function IsKeywordMarker: Boolean; overload;
    function IsKeywordStop: Boolean; overload;
    function IsHexMarker: Boolean; overload;
    function IsWhitespace: Boolean; overload;
    function IsLetter: Boolean; overload;
    function IsDigit: Boolean; overload;
    function IsEscapable: Boolean; overload;
    function IsHexEscapable: Boolean; overload;
    function IsChar(const Ch: Char): Boolean;
    function InSet(const Charsets: TSysCharSet): Boolean;
  public
    property Char: Char read FCh write FCh;
    property Eof: Boolean read FEof;
  public
    class function IsBlockStart(const Ch: Char): Boolean; overload;
    class function IsBlockEnd(const Ch: Char): Boolean; overload;
    class function IsKeywordMarker(const Ch: Char): Boolean; overload;
    class function IsKeywordStop(const Ch: Char): Boolean; overload;
    class function IsHexMarker(const Ch: Char): Boolean; overload;
    class function IsWhitespace(const Ch: Char): Boolean; overload;
    class function IsLetter(const Ch: Char): Boolean; overload;
    class function IsDigit(const Ch: Char): Boolean; overload;
    class function IsEscapable(const Ch: Char): Boolean; overload;
    class function IsHexEscapable(const Ch: Char): Boolean; overload;
    class function IsPlain(const Ch: Char): Boolean; overload;
    class function MustInHex(const Ch: Char): Boolean; overload;
    class property BlockStart: Char read FBlockStart;
    class property BlockEnd: Char read FBlockEnd;
    class property KeywordMarker: Char read FKeywordMarker;
    class property KeywordStop: Char read FKeywordStop;
    class property HexMarker: Char read FHexMarker;
  end;

  { TRtfDocumentProperty }

  TRtfDocumentProperty = class
  private
    FTitle: String;
    FSubject: String;
    FAuthor: String;
    FManager: String;
    FCompany: String;
    FOperator: String;
    FCategory: String;
    FKeywords: String;
    FComment: String;
    FDocComm: String;
    FHLinkBase: String;
    FCreatim: TDateTime;
    FRevtim: TDateTime;
    FPrintim: TDateTime;
    FBuptim: TDateTime;
    FVersion: Integer;
    FVern: Integer;
    FEdMins: Integer;
    FNOfPages: Integer;
    FNOfWords: Integer;
    FNOfChars: Integer;
    FId: Integer;
  public
    constructor Create;
    function ToString: String; override;
  public
    property Title: String read FTitle write FTitle;
    property Subject: String read FSubject write FSubject;
    property Author: String read FAuthor write FAuthor;
    property Manager: String read FManager write FManager;
    property Company: String read FCompany write FCompany;
    property Operator: String read FOperator write FOperator;
    property Category: String read FCategory write FCategory;
    property Keywords: String read FKeywords write FKeywords;
    property Comment: String read FComment write FComment;
    property DocComment: String read FDocComm write FDocComm;
    property HLinkBase: String read FHLinkBase write FHLinkBase;
    property CreationTime: TDateTime read FCreatim write FCreatim;
    property RevisionTime: TDateTime read FRevtim write FRevtim;
    property LastPrintTime: TDateTime read FPrintim write FPrintim;
    property BackupTime: TDateTime read FBuptim write FBuptim;
    property Version: Integer read FVersion write FVersion;
    property InternalVersion: Integer read FVern write FVern;
    property EditingTime: Integer read FEdMins write FEdMins;
    property NumberOfPages: Integer read FNOfPages write FNOfPages;
    property NumberOfWords: Integer read FNOfWords write FNOfWords;
    property NumberOfChars: Integer read FNOfChars write FNOfChars;
    property Id: Integer read FId write FId;
  end;

  { TRtfFontTable }

  TRtfFontTable = class
  private
    FFontIndex: Integer;
    FFontName: String;
  public
    constructor Create(const AFontIndex: Integer; const AFontName: String);
    property FontIndex: Integer read FFontIndex write FFontIndex;
    property FontName: String read FFontName write FFontName;
  end;

  { TRtfFontTables }

  TRtfFontTables = class
  private
    FItems: TList;
  public
    constructor Create;
    destructor Destroy; override;
  protected
    function GetItem(const AIndex: Integer): TRtfFontTable;
  public
    function Count: Integer;
    procedure Clear;
    procedure Add(AFont: TRtfFontTable);
    function IndexOf(AFont: TRtfFontTable): Integer; overload;
    function IndexOf(const AFontIndex: Integer): Integer; overload;
    function IndexOf(const AFontName: String): Integer; overload;
  public
    property Items[const AIndex: Integer]: TRtfFontTable read GetItem; default;
  end;

  { TRtfColorTable }

  TRtfColorTable = class
  private
    FColor: TColor;
  public
    constructor Create(const AColor: TColor);
    property Color: TColor read FColor write FColor;
  end;

  { TRtfColorTables }

  TRtfColorTables = class
  private
    FItems: TList;
  public
    constructor Create;
    destructor Destroy; override;
  protected
    function GetItem(const AIndex: Integer): TRtfColorTable;
  public
    function Count: Integer;
    procedure Clear;
    procedure Add(AColor: TRtfColorTable);
    function IndexOf(AColor: TRtfColorTable): Integer; overload;
    function IndexOf(const AColor: TColor): Integer; overload;
  public
    property Items[const AIndex: Integer]: TRtfColorTable read GetItem; default;
  end;

  { TRtfParagraph }

  TRtfTextAlignment = (taLeft, taRight, taCenter, taJustified);

  TRtfParagraph = class
  private
    FAlignment: TRtfTextAlignment;
    FLeftIndent: Double;
    FRightIndent: Double;
  public
    constructor Create;
  public
    property Alignment: TRtfTextAlignment read FAlignment write FAlignment default taLeft;
    property LeftIndent: Double read FLeftIndent write FLeftIndent;
    property RightIndent: Double read FRightIndent write FRightIndent;
  end;

  { TRtfUnit }

  TRtfMeasurement = (mNative, mInch, mCentimeter, mMilimeter, mPixel);

  TRtfUnit = class
  private
    FMeasurement: TRtfMeasurement;
  public
    constructor Create(const ADefault: TRtfMeasurement);
    class function ToNative(const Value: Double; const AFrom: TRtfMeasurement): Integer;
    class function FromNative(const Value: Integer; const AFrom: TRtfMeasurement): Double;
  public
    property Measurement: TRtfMeasurement read FMeasurement write FMeasurement;
  end;

  { TRtfPage }

  TRtfPage = class
  private
    FPageWidth: Double;
    FPageHeight: Double;
    FMarginTop: Double;
    FMarginBottom: Double;
    FMarginLeft: Double;
    FMarginRight: Double;
  public
    constructor Create;
  public
    property PageWidth: Double read FPageWidth write FPageWidth;
    property PageHeight: Double read FPageHeight write FPageHeight;
    property MarginTop: Double read FMarginTop write FMarginTop;
    property MarginBottom: Double read FMarginBottom write FMarginBottom;
    property MarginLeft: Double read FMarginLeft write FMarginLeft;
    property MarginRight: Double read FMarginRight write FMarginRight;
  end;

implementation

{ TRtfChar }

class procedure TRtfChar.Initialize;
begin
  FBlockStart := '{';
  FBlockEnd := '}';
  FKeywordMarker := '\';
  FKeywordStop := ' ';
  FHexMarker := '''';
end;

class function TRtfChar.IsBlockStart(const Ch: Char): Boolean;
begin
  Result := Ch = BlockStart;
end;

class function TRtfChar.IsBlockEnd(const Ch: Char): Boolean;
begin
  Result := Ch = BlockEnd;
end;

class function TRtfChar.IsKeywordMarker(const Ch: Char): Boolean;
begin
  Result := Ch = KeywordMarker;
end;

class function TRtfChar.IsKeywordStop(const Ch: Char): Boolean;
begin
  Result := Ch = KeywordStop;
end;

class function TRtfChar.IsHexMarker(const Ch: Char): Boolean;
begin
  Result := Ch = HexMarker;
end;

class function TRtfChar.IsWhitespace(const Ch: Char): Boolean;
begin
  Result := CharInSet(Ch, [#0, #7, #10, #13]);
end;

class function TRtfChar.IsLetter(const Ch: Char): Boolean;
begin
  Result := CharInSet(Ch, ['A'..'Z', 'a'..'z']);
end;

class function TRtfChar.IsDigit(const Ch: Char): Boolean;
begin
  Result := CharInSet(Ch, ['0'..'9', '-']);
end;

class function TRtfChar.IsEscapable(const Ch: Char): Boolean;
begin
  Result := CharInSet(Ch, [BlockStart, BlockEnd, KeywordMarker]);
end;

class function TRtfChar.IsHexEscapable(const Ch: Char): Boolean;
begin
  Result := CharInSet(Ch, [HexMarker]);
end;

class function TRtfChar.IsPlain(const Ch: Char): Boolean;
begin
  Result := CharInSet(Ch, [#32..#127]);
end;

class function TRtfChar.MustInHex(const Ch: Char): Boolean;
begin
  Result := CharInSet(Ch, [#0..#31, #128..#255]);
end;

function TRtfChar.Read(Stream: TStream): Boolean;
begin
  FEof := Stream.Read(FCh, 1) = 0;
  Result := not FEof;
end;

function TRtfChar.IsBlockStart: Boolean;
begin
  Result := TRtfChar.IsBlockStart(FCh);
end;

function TRtfChar.IsBlockEnd: Boolean;
begin
  Result := TRtfChar.IsBlockEnd(FCh);
end;

function TRtfChar.IsKeywordMarker: Boolean;
begin
  Result := TRtfChar.IsKeywordMarker(FCh);
end;

function TRtfChar.IsKeywordStop: Boolean;
begin
  Result := TRtfChar.IsKeywordStop(FCh);
end;

function TRtfChar.IsHexMarker: Boolean;
begin
  Result := TRtfChar.IsHexMarker(FCh);
end;

function TRtfChar.IsWhitespace: Boolean;
begin
  Result := TRtfChar.IsWhitespace(FCh);
end;

function TRtfChar.IsLetter: Boolean;
begin
  Result := TRtfChar.IsLetter(FCh);
end;

function TRtfChar.IsDigit: Boolean;
begin
  Result := TRtfChar.IsDigit(FCh);
end;

function TRtfChar.IsEscapable: Boolean;
begin
  Result := TRtfChar.IsEscapable(FCh);
end;

function TRtfChar.IsHexEscapable: Boolean;
begin
  Result := TRtfChar.IsHexEscapable(FCh);
end;

function TRtfChar.IsChar(const Ch: Char): Boolean;
begin
  Result := FCh = Ch;
end;

function TRtfChar.InSet(const Charsets: TSysCharSet): Boolean;
begin
  Result := CharInSet(FCh, Charsets);
end;

{ TRtfDocumentProperty }

constructor TRtfDocumentProperty.Create;
begin
  FVersion := -1;
  FVern := -1;
  FEdMins := -1;
  FNOfPages := -1;
  FNOfWords := -1;
  FNOfChars := -1;
  FId := -1;
end;

function TRtfDocumentProperty.ToString: String;
begin
  Result := '';
  Result := Result + 'Title     : ' + Title + #13#10;
  Result := Result + 'Subject   : ' + Subject + #13#10;
  Result := Result + 'Author    : ' + Author + #13#10;
  Result := Result + 'Manager   : ' + Manager + #13#10;
  Result := Result + 'Company   : ' + Company + #13#10;
  Result := Result + 'Operator  : ' + Operator + #13#10;
  Result := Result + 'Category  : ' + Category + #13#10;
  Result := Result + 'Keywords  : ' + Keywords + #13#10;
  Result := Result + 'Comment   : ' + Comment + #13#10;
  Result := Result + 'DComment  : ' + DocComment + #13#10;
  Result := Result + 'HLinkBase : ' + HlinkBase + #13#10;
  Result := Result + 'Created   : ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', CreationTime) + #13#10;
  Result := Result + 'Revised   : ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', RevisionTime) + #13#10;
  Result := Result + 'Printed   : ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', LastPrintTime) + #13#10;
  Result := Result + 'Backup    : ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', BackupTime) + #13#10;
  Result := Result + 'Version   : ' + IntToStr(Version) + #13#10;
  Result := Result + 'IVersion  : ' + IntToStr(InternalVersion) + #13#10;
  Result := Result + 'Editing   : ' + IntToStr(EditingTime) + #13#10;
  Result := Result + 'Num Pages : ' + IntToStr(NumberOfPages) + #13#10;
  Result := Result + 'Num Words : ' + IntToStr(NumberOfWords) + #13#10;
  Result := Result + 'Num Chars : ' + IntToStr(NumberOfChars) + #13#10;
  Result := Result + 'Id        : ' + IntToStr(Id) + #13#10;
end;

{ TRtfFontTable }

constructor TRtfFontTable.Create(const AFontIndex: Integer; const AFontName: String);
begin
  FFontIndex := AFontIndex;
  FFontName := AFontName;
end;

{ TRtfFontTables }

constructor TRtfFontTables.Create;
begin
  FItems := TList.Create;
end;

destructor TRtfFontTables.Destroy;
begin
  Clear;
  FreeAndNil(FItems);
  inherited;
end;

function TRtfFontTables.GetItem(const AIndex: Integer): TRtfFontTable;
begin
  Result := nil;
  if (AIndex > -1) and (AIndex < FItems.Count) then
    Result := TRtfFontTable(FItems[AIndex]);
end;

function TRtfFontTables.Count: Integer;
begin
  Result := FItems.Count;
end;

procedure TRtfFontTables.Clear;
begin
  while FItems.Count > 0 do
  begin
    TObject(FItems[0]).Free;
    FItems.Delete(0);
  end;
end;

procedure TRtfFontTables.Add(AFont: TRtfFontTable);
begin
  FItems.Add(AFont);
end;

function TRtfFontTables.IndexOf(AFont: TRtfFontTable): Integer;
begin
  Result := FItems.IndexOf(AFont);
end;

function TRtfFontTables.IndexOf(const AFontIndex: Integer): Integer;
var
  i: Integer;

begin
  Result := -1;
  for i := 0 to FItems.Count - 1 do
    if TRtfFontTable(FItems[i]).FontIndex = AFontIndex then
    begin
      Result := i;
      Break;
    end;
end;

function TRtfFontTables.IndexOf(const AFontName: String): Integer;
var
  i: Integer;

begin
  Result := -1;
  for i := 0 to FItems.Count - 1 do
    if TRtfFontTable(FItems[i]).FontName = AFontName then
    begin
      Result := i;
      Break;
    end;
end;

{ TRtfColorTable }

constructor TRtfColorTable.Create(const AColor: TColor);
begin
  FColor := AColor;
end;

{ TRtfColorTables }

constructor TRtfColorTables.Create;
begin
  FItems := TList.Create;
end;

destructor TRtfColorTables.Destroy;
begin
  Clear;
  FreeAndNil(FItems);
  inherited;
end;

function TRtfColorTables.GetItem(const AIndex: Integer): TRtfColorTable;
begin
  Result := nil;
  if (AIndex > -1) and (AIndex < FItems.Count) then
    Result := TRtfColorTable(FItems[AIndex]);
end;

function TRtfColorTables.Count: Integer;
begin
  Result := FItems.Count;
end;

procedure TRtfColorTables.Clear;
begin
  while FItems.Count > 0 do
  begin
    TObject(FItems[0]).Free;
    FItems.Delete(0);
  end;
end;

procedure TRtfColorTables.Add(AColor: TRtfColorTable);
begin
  FItems.Add(AColor);
end;

function TRtfColorTables.IndexOf(AColor: TRtfColorTable): Integer;
begin
  Result := FItems.IndexOf(AColor);
end;

function TRtfColorTables.IndexOf(const AColor: TColor): Integer;
var
  i: Integer;

begin
  Result := -1;
  for i := 0 to FItems.Count - 1 do
    if TRtfColorTable(FItems[i]).Color = AColor then
    begin
      Result := i;
      Break;
    end;
end;

{ TRtfParagraph }

constructor TRtfParagraph.Create;
begin
  FAlignment := taLeft;
  FLeftIndent := 0;
  FRightIndent := 0;
end;

{ TRtfUnit }

constructor TRtfUnit.Create(const ADefault: TRtfMeasurement);
begin
  FMeasurement := ADefault;
end;

class function TRtfUnit.ToNative(const Value: Double; const AFrom: TRtfMeasurement): Integer;
begin
  Result := 0;
  case AFrom of
  mNative:
    Result := Trunc(Value);
  mInch:
    Result := Trunc(Value * TWIPS_UNIT);
  mCentimeter:
    Result := Trunc(Value / 2.54 * TWIPS_UNIT);
  mMilimeter:
    Result := Trunc(Value / 25.4 * TWIPS_UNIT);
  mPixel:
    Result := Trunc(Value * 20);
  end;
end;

class function TRtfUnit.FromNative(const Value: Integer; const AFrom: TRtfMeasurement): Double;
begin
  Result := 0;
  case AFrom of
  mNative:
    Result := Value;
  mInch:
    Result := Value / TWIPS_UNIT;
  mCentimeter:
    Result := (Value / TWIPS_UNIT) * 2.54;
  mMilimeter:
    Result := (Value / TWIPS_UNIT) * 25.4;
  mPixel:
    Result := Value / 20;
  end;
end;

{ TRtfPage }

constructor TRtfPage.Create;
begin
  FPageWidth := 0;
  FPageHeight := 0;
  FMarginTop := 2;
  FMarginBottom := 2;
  FMarginLeft := 2;
  FMarginRight := 2;
end;

initialization
  TRtfChar.Initialize;

end.
