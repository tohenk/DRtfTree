unit RtfImg;

interface

uses
  Classes, SysUtils, StrUtils, Vcl.Graphics, RtfTree, RtfClasses;

type

  { TRtfImage }

  TRtfImage = class(TRtfObject)
  private
    FWidth: Integer;
    FHeight: Integer;
    FDesiredWidth: Integer;
    FDesiredHeight: Integer;
    FScaleX: Integer;
    FScaleY: Integer;
    FImageFormat: String;
    FPicture: TPicture;
  protected
    procedure Initialize; override;
    procedure Finalize; override;
    procedure ReadValues;
    procedure ReadImage;
    procedure LoadImage;
  public
    procedure SaveToFile(const FileName: String);
    function CreateNode(const AWidth, AHeight: Integer;
      const AScaleX: Integer = 0; const AScaleY: Integer = 0): TRtfTreeNode;
  public
    property Width: Integer read FWidth write FWidth;
    property Height: Integer read FHeight write FHeight;
    property DesiredWidth: Integer read FDesiredWidth write FDesiredWidth;
    property DesiredHeight: Integer read FDesiredHeight write FDesiredHeight;
    property ScaleX: Integer read FScaleX write FScaleX;
    property ScaleY: Integer read FScaleY write FScaleY;
    property ImageFormat: String read FImageFormat;
    property Picture: TPicture read FPicture;
  end;

implementation

uses
  Jpeg, PngImage;

{ TRtfImage }

procedure TRtfImage.Initialize;
begin
  FPicture := TPicture.Create;
  ReadValues;
  ReadImage;
  LoadImage;
end;

procedure TRtfImage.Finalize;
begin
  FreeAndNil(FPicture);
  inherited;
end;

procedure TRtfImage.SaveToFile(const FileName: String);
var
  Ext: String;
  Graphic: TGraphic;
  Bitmap: TBitmap;

begin
  if (FPicture.Graphic = nil) or FPicture.Graphic.Empty then
    Exit;
  Graphic := nil;
  Ext := ExtractFileExt(FileName);
  if SameText(Ext, '.jpg') then
    Graphic := TJPEGImage.Create
  else if SameText(Ext, '.png') then
    Graphic := TPngImage.Create
  else if SameText(Ext, '.wmf') or SameText(Ext, '.emf') then
    Graphic := TMetafile.Create
  else if SameText(Ext, '.bmp') then
    Graphic := TBitmap.Create;
  if Graphic <> nil then
  begin
    // check if image conversion is not needed
    if not FPicture.Graphic.Empty and FPicture.Graphic.ClassNameIs(Graphic.ClassName) then
      FPicture.Graphic.SaveToFile(FileName)
    else
    begin
      Bitmap := TBitmap.Create;
      try
        Bitmap.Assign(FPicture.Graphic);
        Graphic.Assign(Bitmap);
        Graphic.SaveToFile(FileName);
      finally
        FreeAndNil(Bitmap);
      end;
    end;
    FreeAndNil(Graphic);
  end;
end;

function TRtfImage.CreateNode(const AWidth, AHeight: Integer;
  const AScaleX: Integer = 0; const AScaleY: Integer = 0): TRtfTreeNode;
var
  Hex, Img: TStringStream;
  AFormat: String;

begin
  if (FPicture.Graphic = nil) or FPicture.Graphic.Empty then
    raise Exception.Create('Image is not loaded.');
  // get image format
  AFormat := '';
  if FPicture.Graphic is TPngImage then
    AFormat := 'pngblip'
  else if FPicture.Graphic is TJPEGImage then
    AFormat := 'jpegblip'
  else if FPicture.Graphic is TBitmap then
    AFormat := 'dibitmap'
  else if FPicture.Graphic is TMetafile then
    AFormat := 'wmetafile';
  if AFormat = '' then
    raise Exception.Create('Unsupported image ' + FPicture.Graphic.ClassName);
  Hex := TStringStream.Create;
  Img := TStringStream.Create;
  try
    // prepare image data
    FPicture.Graphic.SaveToStream(Img);
    ToHex(Img, Hex);
    // create image nodes
    Result := TRtfTreeNode.Create(ntGroup);
    Result.AppendChild(TRtfTreeNode.Create(ntKeyword, 'pict'));
    Result.AppendChild(TRtfTreeNode.Create(ntKeyword, AFormat));
    Result.AppendChild(TRtfTreeNode.Create(ntKeyword, 'picw', True,
      TRtfUnit.ToNative(FPicture.Width, mPixel)));
    Result.AppendChild(TRtfTreeNode.Create(ntKeyword, 'pich', True,
      TRtfUnit.ToNative(FPicture.Height, mPixel)));
    if AWidth > 0 then
      Result.AppendChild(TRtfTreeNode.Create(ntKeyword, 'picwgoal', True,
        TRtfUnit.ToNative(AWidth, mPixel)));
    if AHeight > 0 then
      Result.AppendChild(TRtfTreeNode.Create(ntKeyword, 'pichgoal', True,
        TRtfUnit.ToNative(AHeight, mPixel)));
    if AScaleX > 0 then
      Result.AppendChild(TRtfTreeNode.Create(ntKeyword, 'picscalex', True,
        AScaleX));
    if AScaleY > 0 then
      Result.AppendChild(TRtfTreeNode.Create(ntKeyword, 'picscaley', True,
        AScaleY));
    Result.AppendChild(TRtfTreeNode.Create(ntText, Hex.DataString.ToLower));
  finally
    FreeAndNil(Img);
    FreeAndNil(Hex);
  end;
end;

procedure TRtfImage.ReadValues;
begin
  ReadProp('picw', FWidth);
  ReadProp('pich', FHeight);
  ReadProp('picwgoal', FDesiredWidth);
  ReadProp('pichgoal', FDesiredHeight);
  ReadProp('picscalex', FScaleX);
  ReadProp('picscaley', FScaleY);
end;

procedure TRtfImage.ReadImage;
var
  Nodes: TRtfTreeNodeCollection;
  i: Integer;
  HexStr: String;

begin
  FindChild(['jpegblip', 'pngblip', 'emfblip', 'wmetafile', 'dibitmap', 'wbitmap'],
    FImageFormat);
  if Assigned(Node) then
  begin
    // (Word 97-2000): {\*\shppict {\pict\jpegblip <datos>}}{\nonshppict {\pict\wmetafile8 <datos>}}
    // (Wordpad)     : {\pict\wmetafile8 <datos>}
    HexStr := '';
    Nodes := Node.SelectChildNodes(ntText);
    for i := 0 to Nodes.Count - 1 do
      HexStr := HexStr + Nodes[i].NodeKey;
    if HexStr <> '' then
    begin
      TStringStream(HexData).WriteString(HexStr);
      ToBinary(HexData, BinaryData);
    end;
  end;
end;

procedure TRtfImage.LoadImage;
var
  Graphic: TGraphic;

begin
  if (FImageFormat <> '') and (BinaryData.Size > 0) then
  begin
    Graphic := nil;
    if FImageFormat = 'jpegblip' then
      Graphic := TJPEGImage.Create
    else if FImageFormat = 'pngblip' then
      Graphic := TPngImage.Create
    else if (FImageFormat = 'emfblip') or (FImageFormat = 'wmetafile') then
      Graphic := TMetafile.Create
    else if (FImageFormat = 'dibitmap') or (FImageFormat = 'wbitmap') then
      Graphic := TBitmap.Create;
    if Graphic <> nil then
    begin
      BinaryData.Position := 0;
      Graphic.LoadFromStream(BinaryData);
      FPicture.Assign(Graphic);
      FreeAndNil(Graphic);
    end;
  end;
end;

end.
