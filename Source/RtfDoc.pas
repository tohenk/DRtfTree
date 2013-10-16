unit RtfDoc;

interface

uses
  SysUtils, Vcl.Graphics, UITypes, RtfTree, RtfClasses;

type

  { TRtfDocument }

  TRtfDocument = class
  private
    FEncoding: TEncoding;
    FLocaleID: TLocaleID;
    FPage: TRtfPage;
    FFontTables: TRtfFontTables;
    FColorTables: TRtfColorTables;
    FParagraph: TRtfParagraph;
    FDefaultFont: TFont;
    FFont: TFont;
    FDocNode: TRtfTreeNode;
    FRtfUnit: TRtfUnit;
    FGenerator: String;
    FDirty: Boolean;
    FAddClosingParagraph: Boolean;
    FRtfTree: TRtfTree;
  protected
    function GetTree: TRtfTree;
    function GetText: String;
    function GetRtf: String;
    procedure CreateFont;
    procedure AddNode(ANode: TRtfTreeNode); overload;
    procedure AddNode(Nodes: TRtfTreeNodeCollection); overload;
    procedure BuildBaseTree(Tree: TRtfTree);
    procedure BuildDocData(Tree: TRtfTree);
    procedure BuildFontTables(Tree: TRtfTree);
    procedure BuildColorTables(Tree: TRtfTree);
    procedure BuildGenerator(Tree: TRtfTree);
    procedure BuildDocSettings(Tree: TRtfTree);
    procedure BuildText(const AText: String);
    procedure BuildFontName(AFont: TFont);
    procedure BuildFontColor(AFont: TFont);
    procedure BuildFontSize(AFont: TFont);
    procedure BuildFontBold(AFont: TFont);
    procedure BuildFontItalic(AFont: TFont);
    procedure BuildFontUnderline(AFont: TFont);
    procedure BuildParagraphAlignment(AParagraph: TRtfParagraph);
    procedure BuildParagraphLeftIndent(AParagraph: TRtfParagraph);
    procedure BuildParagraphRightIndent(AParagraph: TRtfParagraph);
  public
    constructor Create;
    destructor Destroy; override;
  public
    procedure Save(const FileName: String);
    function AddText(const AText: String; const AFont: TFont): TRtfDocument; overload;
    function AddText(const AText: String): TRtfDocument; overload;
    function AddNewLine(const N: Integer = 1): TRtfDocument; overload;
    function AddNewParagraph(const N: Integer = 1): TRtfDocument; overload;
    function AddNewParagraph(AParagraph: TRtfParagraph): TRtfDocument; overload;
    function AddImage(const FileName: String; const AWidth, AHeight: Integer;
      const AScaleX: Integer = 0; const AScaleY: Integer = 0): TRtfDocument; overload;
    function AddImage(const FileName: String; const AWidth, AHeight: Double;
      const AScaleX: Integer = 0; const AScaleY: Integer = 0): TRtfDocument; overload;
    function ResetChar: TRtfDocument;
    function ResetParagraph: TRtfDocument;
    function ResetFormat: TRtfDocument;
    function UpdatePage(APage: TRtfPage): TRtfDocument;
    function UpdateFont(AFont: TFont): TRtfDocument;
    function UpdateParagraph(AParagraph: TRtfParagraph): TRtfDocument;
    function SetParagraphAlignment(const AAlignment: TRtfTextAlignment): TRtfDocument;
    function SetParagraphLeftIndent(const AValue: Double): TRtfDocument;
    function SetParagraphRightIndent(const AValue: Double): TRtfDocument;
    function SetFontBold(const AValue: Boolean): TRtfDocument;
    function SetFontUnderline(const AValue: Boolean): TRtfDocument;
    function SetFontItalic(const AValue: Boolean): TRtfDocument;
    function SetFontColor(const AColor: TColor): TRtfDocument;
    function SetFontSize(const ASize: Integer): TRtfDocument;
    function SetFontName(const AName: String): TRtfDocument;
  public
    property Encoding: TEncoding read FEncoding write FEncoding;
    property LocaleID: TLocaleID read FLocaleID write FLocaleID;
    property Generator: String read FGenerator write FGenerator;
    property AddClosingParagraph: Boolean read FAddClosingParagraph write FAddClosingParagraph default False;
    property FontTables: TRtfFontTables read FFontTables;
    property ColorTables: TRtfColorTables read FColorTables;
    property Paragraph: TRtfParagraph read FParagraph;
    property Page: TRtfPage read FPage;
    property RtfUnit: TRtfUnit read FRtfUnit;
    property DefaultFont: TFont read FDefaultFont;
    property Text: String read GetText;
    property Rtf: String read GetRtf;
  end;

implementation

uses
  RtfImg;

const
  TextAlignmentKeywords: array[TRtfTextAlignment] of String = ('ql', 'qr', 'qc', 'qj');

{ TRtfDocument }

constructor TRtfDocument.Create;
begin
  FPage := TRtfPage.Create;
  FFontTables := TRtfFontTables.Create;
  FColorTables := TRtfColorTables.Create;
  FParagraph := TRtfParagraph.Create;
  FDocNode := TRtfTreeNode.Create(ntRoot);
  FRtfUnit := TRtfUnit.Create(mCentimeter);
  FEncoding := TEncoding.Default;
  FLocaleID := SysLocale.DefaultLCID;
  FGenerator := 'DRtfTree';
  FDirty := False;
  FAddClosingParagraph := False;
  FDefaultFont := TFont.Create;
  // add black as default color
  FColorTables.Add(TRtfColorTable.Create(TColors.Black));
end;

destructor TRtfDocument.Destroy;
begin
  FreeAndNil(FPage);
  FreeAndNil(FFontTables);
  FreeAndNil(FColorTables);
  FreeAndNil(FParagraph);
  FreeAndNil(FFont);
  FreeAndNil(FDocNode);
  FreeAndNil(FRtfUnit);
  FreeAndNil(FRtfTree);
  FreeAndNil(FDefaultFont);
  inherited;
end;

function TRtfDocument.GetTree: TRtfTree;
begin
  if FDirty then
  begin
    FDirty := False;
    FreeAndNil(FRtfTree);
    FRtfTree := TRtfTree.Create;
    BuildBaseTree(FRtfTree);
    BuildDocData(FRtfTree);
    BuildFontTables(FRtfTree);
    BuildColorTables(FRtfTree);
    BuildGenerator(FRtfTree);
    BuildDocSettings(FRtfTree);
  end;
  Result := FRtfTree;
end;

function TRtfDocument.GetText: String;
var
  Tree: TRtfTree;

begin
  Result := '';
  Tree := GetTree;
  if Assigned(Tree) then
    Result := Tree.Text;
end;

function TRtfDocument.GetRtf: String;
var
  Tree: TRtfTree;

begin
  Result := '';
  Tree := GetTree;
  if Assigned(Tree) then
    Result := Tree.Rtf;
end;

procedure TRtfDocument.CreateFont;
begin
  if FFont = nil then
    FFont := TFont.Create;
end;

procedure TRtfDocument.AddNode(ANode: TRtfTreeNode);
begin
  FDirty := True;
  FDocNode.AppendChild(ANode);
end;

procedure TRtfDocument.AddNode(Nodes: TRtfTreeNodeCollection);
begin
  FDirty := True;
  FDocNode.AppendChild(Nodes);
end;

procedure TRtfDocument.BuildBaseTree(Tree: TRtfTree);
var
  ANode: TRtfTreeNode;

begin
  ANode := TRtfTreeNode.Create(ntGroup);
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'rtf', True, 1));
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'ansi'));
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'ansicpg', True, FEncoding.CodePage));
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'deff', True, 0));
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'deflang', True, FLocaleID));
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'pard'));
  Tree.RootNode.AppendChild(ANode);
end;

procedure TRtfDocument.BuildDocData(Tree: TRtfTree);
var
  i: Integer;

begin
  for i := 0 to FDocNode.ChildNodes.Count - 1 do
    Tree.MainGroup.AppendChild(FDocNode.ChildNodes[i].CloneNode);
  if FAddClosingParagraph or (FDocNode.LastChild.NodeKey <> 'par') then
    Tree.MainGroup.AppendChild(TRtfTreeNode.Create(ntKeyword, 'par'));
end;

procedure TRtfDocument.BuildFontTables(Tree: TRtfTree);
var
  ANode, CNode: TRtfTreeNode;
  i: Integer;

begin
  ANode := TRtfTreeNode.Create(ntGroup);
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'fonttbl'));
  for i := 0 to FFontTables.Count - 1 do
  begin
    CNode := TRtfTreeNode.Create(ntGroup);
    CNode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'f', True, FFontTables[i].FontIndex));
    CNode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'fnil'));
    CNode.AppendChild(TRtfTreeNode.Create(ntText, FFontTables[i].FontName + ';'));
    ANode.AppendChild(CNode);
  end;
  Tree.MainGroup.InsertChild(5, ANode);
end;

procedure TRtfDocument.BuildColorTables(Tree: TRtfTree);
var
  ANode: TRtfTreeNode;
  i: Integer;

begin
  ANode := TRtfTreeNode.Create(ntGroup);
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'colortbl'));
  for i := 0 to FColorTables.Count - 1 do
  begin
    ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'red', True, FColorTables[i].Color and $FF));
    ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'green', True, (FColorTables[i].Color and $FF00) shr 8));
    ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'blue', True, (FColorTables[i].Color and $FF0000) shr 16));
    ANode.AppendChild(TRtfTreeNode.Create(ntText, ';'));
  end;
  Tree.MainGroup.InsertChild(6, ANode);
end;

procedure TRtfDocument.BuildGenerator(Tree: TRtfTree);
var
  ANode: TRtfTreeNode;

begin
  ANode := TRtfTreeNode.Create(ntGroup);
  ANode.AppendChild(TRtfTreeNode.Create(ntControl, '*'));
  ANode.AppendChild(TRtfTreeNode.Create(ntKeyword, 'generator'));
  ANode.AppendChild(TRtfTreeNode.Create(ntText, FGenerator + ';'));
  Tree.MainGroup.InsertChild(7, ANode);
end;

procedure TRtfDocument.BuildDocSettings(Tree: TRtfTree);
var
  Idx: Integer;

begin
  Idx := Tree.MainGroup.ChildNodes.IndexOf('pard');
  Tree.MainGroup.InsertChild(Idx, TRtfTreeNode.Create(ntKeyword, 'viewkind', True, 4));
  Inc(Idx);
  Tree.MainGroup.InsertChild(Idx, TRtfTreeNode.Create(ntKeyword, 'uc', True, 1));
  Inc(Idx);
  if FPage.PageWidth > 0 then
  begin
    Tree.MainGroup.InsertChild(Idx, TRtfTreeNode.Create(ntKeyword, 'paperw', True,
      RtfUnit.ToNative(FPage.PageWidth, RtfUnit.Measurement)));
    Inc(Idx);
  end;
  if FPage.PageHeight > 0 then
  begin
    Tree.MainGroup.InsertChild(Idx, TRtfTreeNode.Create(ntKeyword, 'paperh', True,
      RtfUnit.ToNative(FPage.PageHeight, RtfUnit.Measurement)));
    Inc(Idx);
  end;
  Tree.MainGroup.InsertChild(Idx, TRtfTreeNode.Create(ntKeyword, 'margl', True,
    RtfUnit.ToNative(FPage.MarginLeft, RtfUnit.Measurement)));
  Inc(Idx);
  Tree.MainGroup.InsertChild(Idx, TRtfTreeNode.Create(ntKeyword, 'margr', True,
    RtfUnit.ToNative(FPage.MarginRight, RtfUnit.Measurement)));
  Inc(Idx);
  Tree.MainGroup.InsertChild(Idx, TRtfTreeNode.Create(ntKeyword, 'margt', True,
    RtfUnit.ToNative(FPage.MarginTop, RtfUnit.Measurement)));
  Inc(Idx);
  Tree.MainGroup.InsertChild(Idx, TRtfTreeNode.Create(ntKeyword, 'margb', True,
    RtfUnit.ToNative(FPage.MarginBottom, RtfUnit.Measurement)));
end;

procedure TRtfDocument.BuildText(const AText: String);
var
  Nodes: TRtfTreeNodeCollection;

begin
  // use default font if none provided
  if FFont = nil then
    UpdateFont(DefaultFont);
  Nodes := TRtfTreeNode.CreateText(AText);
  try
    AddNode(Nodes);
  finally
    FreeAndNil(Nodes);
  end;
end;

procedure TRtfDocument.BuildFontName(AFont: TFont);
begin
  if FFontTables.IndexOf(AFont.Name) < 0 then
    FFontTables.Add(TRtfFontTable.Create(FFontTables.Count, AFont.Name));
  AddNode(TRtfTreeNode.Create(ntKeyword, 'f', True, FFontTables.IndexOf(AFont.Name)));
end;

procedure TRtfDocument.BuildFontColor(AFont: TFont);
begin
  if FColorTables.IndexOf(AFont.Color) < 0 then
    FColorTables.Add(TRtfColorTable.Create(AFont.Color));
  AddNode(TRtfTreeNode.Create(ntKeyword, 'cf', True, FColorTables.IndexOf(AFont.Color)));
end;

procedure TRtfDocument.BuildFontSize(AFont: TFont);
begin
  AddNode(TRtfTreeNode.Create(ntKeyword, 'fs', True, AFont.Size * 2));
end;

procedure TRtfDocument.BuildFontBold(AFont: TFont);
begin
  if fsBold in AFont.Style then
    AddNode(TRtfTreeNode.Create(ntKeyword, 'b'))
  else
    AddNode(TRtfTreeNode.Create(ntKeyword, 'b', True, 0));
end;

procedure TRtfDocument.BuildFontItalic(AFont: TFont);
begin
  if fsItalic in AFont.Style then
    AddNode(TRtfTreeNode.Create(ntKeyword, 'i'))
  else
    AddNode(TRtfTreeNode.Create(ntKeyword, 'i', True, 0));
end;

procedure TRtfDocument.BuildFontUnderline(AFont: TFont);
begin
  if fsUnderline in AFont.Style then
    AddNode(TRtfTreeNode.Create(ntKeyword, 'ul'))
  else
    AddNode(TRtfTreeNode.Create(ntKeyword, 'ulnone'));
end;

procedure TRtfDocument.BuildParagraphAlignment(AParagraph: TRtfParagraph);
begin
  AddNode(TRtfTreeNode.Create(ntKeyword, TextAlignmentKeywords[AParagraph.Alignment]));
end;

procedure TRtfDocument.BuildParagraphLeftIndent(AParagraph: TRtfParagraph);
begin
  AddNode(TRtfTreeNode.Create(ntKeyword, 'li', True,
    RtfUnit.ToNative(AParagraph.LeftIndent, RtfUnit.Measurement)));
end;

procedure TRtfDocument.BuildParagraphRightIndent(AParagraph: TRtfParagraph);
begin
  AddNode(TRtfTreeNode.Create(ntKeyword, 'ri', True,
    RtfUnit.ToNative(AParagraph.RightIndent, RtfUnit.Measurement)));
end;

procedure TRtfDocument.Save(const FileName: String);
var
  Tree: TRtfTree;

begin
  Tree := GetTree;
  if Assigned(Tree) then
    Tree.SaveToFile(FileName);
end;

function TRtfDocument.AddText(const AText: String; const AFont: TFont): TRtfDocument;
begin
  Result := Self;
  if AFont <> nil then
    UpdateFont(AFont);
  BuildText(AText);
end;

function TRtfDocument.AddText(const AText: String): TRtfDocument;
begin
  Result := Self;
  BuildText(AText);
end;

function TRtfDocument.AddNewLine(const N: Integer = 1): TRtfDocument;
var
  i: Integer;

begin
  Result := Self;
  for i := 1 to N do
    AddNode(TRtfTreeNode.Create(ntKeyword, 'line'));
end;

function TRtfDocument.AddNewParagraph(const N: Integer = 1): TRtfDocument;
var
  i: Integer;

begin
  Result := Self;
  for i := 1 to N do
    AddNode(TRtfTreeNode.Create(ntKeyword, 'par'));
end;

function TRtfDocument.AddNewParagraph(AParagraph: TRtfParagraph): TRtfDocument;
begin
  Result := Self;
  AddNewParagraph(1);
  UpdateParagraph(AParagraph);
end;

function TRtfDocument.AddImage(const FileName: String; const AWidth, AHeight: Integer;
  const AScaleX: Integer = 0; const AScaleY: Integer = 0): TRtfDocument;
var
  RtfImage: TRtfImage;

begin
  Result := Self;
  RtfImage := TRtfImage.Create(nil);
  try
    RtfImage.Picture.LoadFromFile(FileName);
    AddNode(RtfImage.CreateNode(AWidth, AHeight, AScaleX, AScaleY));
  finally
    FreeAndNil(RtfImage);
  end;
end;

function TRtfDocument.AddImage(const FileName: String; const AWidth, AHeight: Double;
  const AScaleX: Integer = 0; const AScaleY: Integer = 0): TRtfDocument;
var
  ImgWidth, ImgHeight: Integer;

begin
  ImgWidth := Round(RtfUnit.FromNative(RtfUnit.ToNative(AWidth, RtfUnit.Measurement),
    mPixel));
  ImgHeight := Round(RtfUnit.FromNative(RtfUnit.ToNative(AHeight, RtfUnit.Measurement),
    mPixel));
  Result := AddImage(FileName, ImgWidth, ImgHeight, AScaleX, AScaleY);
end;

function TRtfDocument.ResetChar: TRtfDocument;
begin
  Result := Self;
  AddNode(TRtfTreeNode.Create(ntKeyword, 'plain'));
end;

function TRtfDocument.ResetParagraph: TRtfDocument;
begin
  Result := Self;
  AddNode(TRtfTreeNode.Create(ntKeyword, 'pard'));
end;

function TRtfDocument.ResetFormat: TRtfDocument;
begin
  Result := Self;
  ResetParagraph;
  ResetChar;
end;

function TRtfDocument.UpdatePage(APage: TRtfPage): TRtfDocument;
begin
  Result := Self;
  FPage.PageWidth := APage.PageWidth;
  FPage.PageHeight := APage.PageHeight;
  FPage.MarginTop := APage.MarginTop;
  FPage.MarginBottom := APage.MarginBottom;
  FPage.MarginLeft := APage.MarginLeft;
  FPage.MarginRight := APage.MarginRight;
end;

function TRtfDocument.UpdateFont(AFont: TFont): TRtfDocument;
begin
  Result := Self;
  if AFont = nil then Exit;
  // check if font already in tables
  if FFontTables.IndexOf(AFont.Name) < 0 then
  begin
    if (FFont = nil) or (AFont.Color <> FFont.Color) then
      BuildFontColor(AFont);
    if (FFont = nil) or (AFont.Size <> FFont.Size) then
      BuildFontSize(AFont);
    if (FFont = nil) or (AFont.Name <> FFont.Name) then
      BuildFontName(AFont);
    if ((FFont = nil) and (fsBold in AFont.Style)) or
      ((FFont <> nil) and ((fsBold in AFont.Style) <> (fsBold in FFont.Style))) then
      BuildFontBold(AFont);
    if ((FFont = nil) and (fsItalic in AFont.Style)) or
      ((FFont <> nil) and ((fsItalic in AFont.Style) <> (fsItalic in FFont.Style))) then
      BuildFontItalic(AFont);
    if ((FFont = nil) and (fsUnderline in AFont.Style)) or
      ((FFont <> nil) and ((fsUnderline in AFont.Style) <> (fsUnderline in FFont.Style))) then
      BuildFontUnderline(AFont);
    CreateFont;
    FFont.Assign(AFont);
  end else
  begin
    SetFontColor(AFont.Color);
    SetFontSize(AFont.Size);
    SetFontName(AFont.Name);
    SetFontBold(fsBold in AFont.Style);
    SetFontItalic(fsItalic in AFont.Style);
    SetFontUnderline(fsUnderline in AFont.Style);
  end;
end;

function TRtfDocument.UpdateParagraph(AParagraph: TRtfParagraph): TRtfDocument;
begin
  Result := Self;
  SetParagraphAlignment(AParagraph.Alignment);
  SetParagraphLeftIndent(AParagraph.LeftIndent);
  SetParagraphRightIndent(AParagraph.RightIndent);
end;

function TRtfDocument.SetParagraphAlignment(const AAlignment: TRtfTextAlignment): TRtfDocument;
begin
  Result := Self;
  if FParagraph.Alignment <> AAlignment then
  begin
    FParagraph.Alignment := AAlignment;
    BuildParagraphAlignment(FParagraph);
  end;
end;

function TRtfDocument.SetParagraphLeftIndent(const AValue: Double): TRtfDocument;
begin
  Result := Self;
  if FParagraph.LeftIndent <> AValue then
  begin
    FParagraph.LeftIndent := AValue;
    BuildParagraphLeftIndent(FParagraph);
  end;
end;

function TRtfDocument.SetParagraphRightIndent(const AValue: Double): TRtfDocument;
begin
  Result := Self;
  if FParagraph.RightIndent <> AValue then
  begin
    FParagraph.RightIndent := AValue;
    BuildParagraphRightIndent(FParagraph);
  end;
end;

function TRtfDocument.SetFontBold(const AValue: Boolean): TRtfDocument;
begin
  Result := Self;
  CreateFont;
  if (fsBold in FFont.Style) <> AValue then
  begin
    if AValue then
      FFont.Style := FFont.Style + [fsBold]
    else
      FFont.Style := FFont.Style - [fsBold];
    BuildFontBold(FFont);
  end;
end;

function TRtfDocument.SetFontUnderline(const AValue: Boolean): TRtfDocument;
begin
  Result := Self;
  CreateFont;
  if (fsUnderline in FFont.Style) <> AValue then
  begin
    if AValue then
      FFont.Style := FFont.Style + [fsUnderline]
    else
      FFont.Style := FFont.Style - [fsUnderline];
    BuildFontUnderline(FFont);
  end;
end;

function TRtfDocument.SetFontItalic(const AValue: Boolean): TRtfDocument;
begin
  Result := Self;
  CreateFont;
  if (fsItalic in FFont.Style) <> AValue then
  begin
    if AValue then
      FFont.Style := FFont.Style + [fsItalic]
    else
      FFont.Style := FFont.Style - [fsItalic];
    BuildFontItalic(FFont);
  end;
end;

function TRtfDocument.SetFontColor(const AColor: TColor): TRtfDocument;
begin
  Result := Self;
  CreateFont;
  if FFont.Color <> AColor then
  begin
    FFont.Color := AColor;
    BuildFontColor(FFont);
  end;
end;

function TRtfDocument.SetFontSize(const ASize: Integer): TRtfDocument;
begin
  Result := Self;
  CreateFont;
  if FFont.Size <> ASize then
  begin
    FFont.Size := ASize;
    BuildFontSize(FFont);
  end;
end;

function TRtfDocument.SetFontName(const AName: String): TRtfDocument;
begin
  Result := Self;
  CreateFont;
  if FFont.Name <> AName then
  begin
    FFont.Name := AName;
    BuildFontName(FFont);
  end;
end;

end.
