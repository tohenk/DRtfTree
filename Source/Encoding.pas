unit Encoding;

interface

uses
  Generics.Collections, SysUtils;

type

  //Provides buffer of dynamically allocated encodings
  TEncodingBuffer = class
  private type
    TEncodingDic = TObjectDictionary<Integer, TEncoding>;
  private
    FEncodings: TEncodingDic;
  public
    constructor Create;
    destructor Destroy; override;
    function GetEncoding(CodePage: Integer): TEncoding;
  end;

implementation

{ TEncodingBuffer }

constructor TEncodingBuffer.Create;
begin
  FEncodings:=TEncodingDic.Create([doOwnsValues]);
end;

destructor TEncodingBuffer.Destroy;
begin
  FEncodings.Free;
  inherited;
end;

function TEncodingBuffer.GetEncoding(CodePage: Integer): TEncoding;
begin
  if not FEncodings.TryGetValue(CodePage, Result) then
  begin
    Result:=TEncoding.GetEncoding(CodePage);
    FEncodings.Add(CodePage, Result);
  end;
end;

end.
