unit RtfTree;

interface

uses
  Classes, SysUtils, StrUtils, DateUtils, UITypes, RtfToken, RtfLex, RtfClasses, Encoding;

type

  { Forward declaration }

  TRtfTree = class;
  TRtfTreeNodes = class;
  TRtfTreeNodeCollection = class;

  { TRtfTreeNode }

  TRtfTreeNodeType = (ntRoot, ntKeyword, ntControl, ntText, ntGroup, ntWhitespace,
    ntNone);

  TRtfTreeNode = class
  private
    FNodeType: TRtfTreeNodeType;
    FKey: String;
    FHasParam: Boolean;
    FParam: Integer;
    FChildren: TRtfTreeNodes;
    FParent: TRtfTreeNode;
    FRoot: TRtfTreeNode;
    FTree: TRtfTree;
    FGarbages: TList;
  protected
    constructor Create(AToken: TRtfToken); overload;
  public
    constructor Create; overload;
    constructor Create(const ANodeType: TRtfTreeNodeType); overload;
    constructor Create(const ANodeType: TRtfTreeNodeType; const AKey: String;
      const AHashParam: Boolean = False; const AParam: Integer = 0); overload;
    destructor Destroy; override;
  protected
    function GetChild(const AIndex: Integer): TRtfTreeNode;
    function GetFirstChild: TRtfTreeNode;
    function GetLastChild: TRtfTreeNode;
    function GetNextSibling: TRtfTreeNode;
    function GetPreviousSibling: TRtfTreeNode;
    function GetNextNode: TRtfTreeNode;
    function GetPreviousNode: TRtfTreeNode;
    function GetNodeIndex: Integer;
    function GetNodeText: String;
    function GetRawText: String;
    function EncodeText(const S: String): String;
    function GetRtf: String;
    function GetRtfInm(ANode, ANextNode: TRtfTreeNode): String;
    function AsRtf(ANext: TRtfTreeNode; const AEncode: Boolean = False): String;
    procedure UpdateNodeRoot(ANode: TRtfTreeNode);
    function GetText(const Raw: Boolean): String; overload;
    function GetText(const Raw: Boolean; const IgnoreNChars: Integer): String; overload;
    function IsGroupHasKey(const AKey: String; const IgnoreSpecial: Boolean): Boolean;
    function IsPlainText: Boolean;
    function MustUseKeywordStop: Boolean;
    function CreateCollection: TRtfTreeNodeCollection;
    procedure RemoveCollection(ACollection: TRtfTreeNodeCollection);
    function SelectChildNodesForText(const AText: String; const StartPos: Integer): TRtfTreeNodeCollection;
    function CombineNodesText(Nodes: TRtfTreeNodeCollection): TRtfTreeNode;
  public
    procedure AppendChild(ANode: TRtfTreeNode); overload;
    procedure AppendChild(ACollection: TRtfTreeNodeCollection); overload;
    procedure InsertChild(const AIndex: Integer; ANode: TRtfTreeNode);
    procedure RemoveChild(const AIndex: Integer); overload;
    procedure RemoveChild(ANode: TRtfTreeNode); overload;
    function CloneNode: TRtfTreeNode;
    function HasChildNodes: Boolean;
    function SelectSingleNode(const ANodeType: TRtfTreeNodeType): TRtfTreeNode; overload;
    function SelectSingleNode(const AKey: String): TRtfTreeNode; overload;
    function SelectSingleNode(const AKey: String; const AParam: Integer): TRtfTreeNode; overload;
    function SelectSingleGroup(const AKey: String): TRtfTreeNode; overload;
    function SelectSingleGroup(const AKey: String; const IgnoreSpecial: Boolean): TRtfTreeNode; overload;
    function SelectNodes(const ANodeType: TRtfTreeNodeType): TRtfTreeNodeCollection; overload;
    function SelectNodes(const AKey: String): TRtfTreeNodeCollection; overload;
    function SelectNodes(const AKey: String; const AParam: Integer): TRtfTreeNodeCollection; overload;
    function SelectGroups(const AKey: String): TRtfTreeNodeCollection; overload;
    function SelectGroups(const AKey: String; const IgnoreSpecial: Boolean): TRtfTreeNodeCollection; overload;
    function SelectSingleChildNode(const ANodeType: TRtfTreeNodeType): TRtfTreeNode; overload;
    function SelectSingleChildNode(const AKey: String): TRtfTreeNode; overload;
    function SelectSingleChildNode(const AKey: String; AParam: Integer): TRtfTreeNode; overload;
    function SelectSingleChildGroup(const AKey: String): TRtfTreeNode; overload;
    function SelectSingleChildGroup(const AKey: String; const IgnoreSpecial: Boolean): TRtfTreeNode; overload;
    function SelectChildNodes(const ANodeType: TRtfTreeNodeType): TRtfTreeNodeCollection; overload;
    function SelectChildNodes(const AKey: String): TRtfTreeNodeCollection; overload;
    function SelectChildNodes(const AKey: String; const AParam: Integer): TRtfTreeNodeCollection; overload;
    function SelectChildGroups(const AKey: String): TRtfTreeNodeCollection; overload;
    function SelectChildGroups(const AKey: String; const IgnoreSpecial: Boolean): TRtfTreeNodeCollection; overload;
    function SelectSibling(const ANodeType: TRtfTreeNodeType): TRtfTreeNode; overload;
    function SelectSibling(const AKey: String): TRtfTreeNode; overload;
    function SelectSibling(const AKey: String; const AParam: Integer): TRtfTreeNode; overload;
    function FindText(const AText: String): TRtfTreeNodeCollection;
    procedure ReplaceText(const OldValue, NewValue: String);
    function ReplaceTextEx(const AFrom, ATo: String): Boolean;
    function ToString: String; override;
    function IsEquals(const AKey: String): Boolean; overload;
    function IsEquals(const Keys: array of String): Boolean; overload;
    procedure Clear;
    class function DecodeChar(Code: Integer; Encoding: TEncoding): String;
    class function EscapeText(const S: String): String;
    class function CreateText(const AText: String): TRtfTreeNodeCollection;
  public
    property NodeType: TRtfTreeNodeType read FNodeType write FNodeType;
    property NodeKey: String read FKey write FKey;
    property HasParameter: Boolean read FHasParam write FHasParam;
    property Parameter: Integer read FParam write FParam;
    property ChildNodes: TRtfTreeNodes read FChildren;
    property ParentNode: TRtfTreeNode read FParent write FParent;
    property RootNode: TRtfTreeNode read FRoot write FRoot;
    property Tree: TRtfTree read FTree write FTree;
    property Child[const AKey: String]: TRtfTreeNode read SelectSingleChildNode;
    property Child[const AIndex: Integer]: TRtfTreeNode read GetChild; default;
    property FirstChild: TRtfTreeNode read GetFirstChild;
    property LastChild: TRtfTreeNode read GetLastChild;
    property NextSibling: TRtfTreeNode read GetNextSibling;
    property PreviousSibling: TRtfTreeNode read GetPreviousSibling;
    property NextNode: TRtfTreeNode read GetNextNode;
    property PreviousNode: TRtfTreeNode read GetPreviousNode;
    property Rtf: String read GetRtf;
    property Index: Integer read GetNodeIndex;
    property Text: String read GetNodeText;
    property RawText: String read GetRawText;
  end;

  { TRtfTreeNodes }

  TRtfTreeNodes = class
  private
    FItems: TList;
  public
    constructor Create;
    destructor Destroy; override;
  protected
    function GetNode(const AIndex: Integer): TRtfTreeNode;
  public
    function Count: Integer;
    procedure Clear;
    function Add(ANode: TRtfTreeNode): Integer;
    procedure Insert(const AIndex: Integer; ANode: TRtfTreeNode);
    function IndexOf(ANode: TRtfTreeNode): Integer; overload;
    function IndexOf(ANode: TRtfTreeNode; const StartIndex: Integer): Integer; overload;
    function IndexOf(const AKey: String): Integer; overload;
    function IndexOf(const AKey: String; const StartIndex: Integer): Integer; overload;
    procedure Remove(const AIndex: Integer; const ACount: Integer = 1);
  public
    property Items[const AIndex: Integer]: TRtfTreeNode read GetNode; default;
  end;

  { TRtfTreeNodeCollection }

  TRtfTreeNodeCollection = class
  private
    FItems: TList;
  public
    constructor Create;
    destructor Destroy; override;
  protected
    function GetItem(const AIndex: Integer): TRtfTreeNode;
  public
    function Count: Integer;
    procedure Add(ANode: TRtfTreeNode);
    procedure Clear;
    procedure Append(ACollection: TRtfTreeNodeCollection);
    procedure Delete(const AIndex: Integer);
    property Items[const AIndex: Integer]: TRtfTreeNode read GetItem; default;
  end;

  { TRtfObject }

  TRtfObject = class
  private
    FNode: TRtfTreeNode;
    FHexData: TStringStream;
    FBinaryData: TStringStream;
  protected
    function GetHexData: TStream;
    function GetBinaryData: TStream;
    function ReadProp(const AProp: String; var Value: Integer;
      const ADefault: Integer = -1): Boolean;
    function FindChild(const Keys: array of String; var S: String): Boolean;
    procedure Initialize; virtual;
    procedure Finalize; virtual;
    procedure ToBinary(HexStream, BinStream: TStream);
    procedure ToHex(BinStream, HexStream: TStream);
    property Node: TRtfTreeNode read FNode;
  public
    constructor Create(ANode: TRtfTreeNode);
    destructor Destroy; override;
  public
    property HexData: TStream read GetHexData;
    property BinaryData: TStream read GetBinaryData;
  end;


  { TRtfStyleSheet }

  TRtfStyleSheetType = (stNone, stCharacter, stParagraph, stSection, stTable);

  TRtfStyleSheet = class
  private
    FIndex: Integer;
    FName: String;
    FType: TRtfStyleSheetType;
    FAdditive: Boolean;
    FBasedOn: Integer;
    FNext: Integer;
    FAutoUpdate: Boolean;
    FHidden: Boolean;
    FLink: Integer;
    FLocked: Boolean;
    FPersonal: Boolean;
    FCompose: Boolean;
    FReply: Boolean;
    FStyrsid: Integer;
    FSemiHidden: Boolean;
    FKeyCode: TRtfTreeNodes;
    FFormatting: TRtfTreeNodes;
  public
    constructor Create;
    destructor Destroy; override;
  public
    property Index: Integer read FIndex write FIndex;
    property Name: String read FName write FName;
    property StyleSheetType: TRtfStyleSheetType read FType write FType default stParagraph;
    property Additive: Boolean read FAdditive write FAdditive default False;
    property BasedOn: Integer read FBasedOn write FBasedOn default -1;
    property Next: Integer read FNext write FNext default -1;
    property AutoUpdate: Boolean read FAutoUpdate write FAutoUpdate default False;
    property Hidden: Boolean read FHidden write FHidden default False;
    property Link: Integer read FLink write FLink default -1;
    property Locked: Boolean read FLocked write FLocked default False;
    property Personal: Boolean read FPersonal write FPersonal default False;
    property Compose: Boolean read FCompose write FCompose default False;
    property Reply: Boolean read FReply write FReply default False;
    property Styrsid: Integer read FStyrsid write FStyrsid default -1;
    property SemiHidden: Boolean read FSemiHidden write FSemiHidden default False;
    property KeyCode: TRtfTreeNodes read FKeyCode;
    property Formatting: TRtfTreeNodes read FFormatting;
  end;

  { TRtfStyleSheets }

  TRtfStyleSheets = class
  private
    FItems: TList;
  public
    constructor Create;
    destructor Destroy; override;
  protected
    function GetStyleSheet(const AIndex: Integer): TRtfStyleSheet;
  public
    function Count: Integer;
    procedure Clear;
    function Add(AStyleSheet: TRtfStyleSheet): Integer;
    procedure Insert(const AIndex: Integer; AStyleSheet: TRtfStyleSheet);
    function IndexOf(AStyleSheet: TRtfStyleSheet): Integer; overload;
    function IndexOf(const AName: String): Integer; overload;
    function IndexOf(const AIndex: Integer): Integer; overload;
    procedure Remove(AStyleSheet: TRtfStyleSheet); overload;
    procedure Remove(const AIndex: Integer); overload;
  public
    property Items[const AIndex: Integer]: TRtfStyleSheet read GetStyleSheet; default;
  end;

  { TRtfTree }

  TRtfTree = class
  private
    FRoot: TRtfTreeNode;
    FSource: TStream;
    FLex: TRtfLex;
    FLevel: Integer;
    FMergeSpecialCharacters: Boolean;
    FDocumentProp: TRtfDocumentProperty;
    FFontTables: TRtfFontTables;
    FColorTables: TRtfColorTables;
    FStyleSheets: TRtfStyleSheets;
    FEncoding: TEncoding;
    FIgnoreWhitespace: Boolean;
    FEncodingBuffer: TEncodingBuffer;
  protected
    procedure ResetAll;
    function ForceTokenAsText(AToken: TRtfToken): Boolean;
    function TryMergeText(ANode: TRtfTreeNode; AToken: TRtfToken;
      const Deep: Boolean): Boolean;
    function ParseRtfTree: Integer;
    procedure GetEncoding;
    procedure GetInfoGroup;
    procedure GetFontTable;
    procedure GetColorTable;
    procedure GetStyleSheetTable;
    function ToStringInm(ANode: TRtfTreeNode; const ALevel: Integer;
      const ShowNodeTypes: Boolean): String;
    function ParseDateTime(ANode: TRtfTreeNode): TDateTime;
    function ConvertToText: String;
    function ParseStyleSheet(ANode: TRtfTreeNode): TRtfStyleSheet;
    function GetMainGroup: TRtfTreeNode;
    function GetRtf: String;
  public
    constructor Create;
    destructor Destroy; override;
    function CloneTree: TRtfTree;
    function LoadFromStream(AStream: TStream): Integer;
    function LoadFromFile(const FileName: String): Integer;
    function LoadFromString(const S: String): Integer;
    procedure SaveToFile(const FileName: String);
    function ToString: String; override;
    function ToStringEx: String;
    procedure AddMainGroup;
  public
    property MergeSpecialCharacters: Boolean read FMergeSpecialCharacters write FMergeSpecialCharacters default False;
    property IgnoreWhitespace: Boolean read FIgnoreWhitespace write FIgnoreWhitespace default True;
    property RootNode: TRtfTreeNode read FRoot;
    property MainGroup: TRtfTreeNode read GetMainGroup;
    property Rtf: String read GetRtf;
    property Text: String read ConvertToText;
    property Props: TRtfDocumentProperty read FDocumentProp;
    property StyleSheets: TRtfStyleSheets read FStyleSheets;
    property FontTables: TRtfFontTables read FFontTables;
    property ColorTables: TRtfColorTables read FColorTables;
    property Encoding: TEncoding read FEncoding;
  end;

implementation

const
  HEX_BUF_SIZE    = 1024;

var
  NodeTypeStrings: array[TRtfTreeNodeType] of String = ('Root', 'Keyword',
    'Control', 'Text', 'Group', 'Whitespace', '');

{ TRtfTreeNode }

constructor TRtfTreeNode.Create(AToken: TRtfToken);
begin
  Create;
  case AToken.TokenType of
  ttNone:
    NodeType := ntNone;
  ttKeyword:
    NodeType := ntKeyword;
  ttControl:
    NodeType := ntControl;
  ttText:
    NodeType := ntText;
  ttWhitespace:
    NodeType := ntWhitespace;
  end;
  NodeKey := AToken.Key;
  HasParameter := AToken.HasParameter;
  Parameter := AToken.Parameter;
end;

constructor TRtfTreeNode.Create;
begin
  FChildren := TRtfTreeNodes.Create;
  FGarbages := TList.Create;
  NodeType := ntNone;
end;

constructor TRtfTreeNode.Create(const ANodeType: TRtfTreeNodeType);
begin
  Create;
  NodeType := ANodeType;
  if NodeType = ntRoot then
    RootNode := Self;
end;

constructor TRtfTreeNode.Create(const ANodeType: TRtfTreeNodeType; const AKey: String;
  const AHashParam: Boolean = False; const AParam: Integer = 0);
begin
  Create(ANodeType);
  NodeKey := AKey;
  HasParameter := AHashParam;
  Parameter := AParam;
end;

destructor TRtfTreeNode.Destroy;
begin
  Clear;
  FreeAndNil(FChildren);
  FreeAndNil(FGarbages);
  inherited;
end;

function TRtfTreeNode.GetChild(const AIndex: Integer): TRtfTreeNode;
begin
  Result := ChildNodes[AIndex];
end;

function TRtfTreeNode.GetFirstChild: TRtfTreeNode;
begin
  Result := nil;
  if ChildNodes.Count > 0 then
    Result := ChildNodes[0];
end;

function TRtfTreeNode.GetLastChild: TRtfTreeNode;
begin
  Result := nil;
  if ChildNodes.Count > 0 then
    Result := ChildNodes[ChildNodes.Count - 1];
end;

function TRtfTreeNode.GetNextSibling: TRtfTreeNode;
begin
  Result := nil;
  if Assigned(ParentNode) then
    Result := ParentNode.ChildNodes[ParentNode.ChildNodes.IndexOf(Self) + 1];
end;

function TRtfTreeNode.GetPreviousSibling: TRtfTreeNode;
begin
  Result := nil;
  if Assigned(ParentNode) then
    Result := ParentNode.ChildNodes[ParentNode.ChildNodes.IndexOf(Self) - 1];
end;

function TRtfTreeNode.GetNextNode: TRtfTreeNode;
begin
  Result := nil;
  if NodeType = ntRoot then
    Result := FirstChild
  else if Assigned(ParentNode) then
  begin
    if (NodeType = ntGroup) and (ChildNodes.Count > 0) then
      Result := FirstChild
    else if Index < ParentNode.ChildNodes.Count - 1 then
      Result := NextSibling
    else
      Result := ParentNode.NextSibling;
  end;
end;

function TRtfTreeNode.GetPreviousNode: TRtfTreeNode;
begin
  Result := nil;
  if (NodeType <> ntRoot) and Assigned(ParentNode) then
  begin
    if (Index > 0) then
    begin
      if PreviousSibling.NodeType = ntGroup then
        Result := PreviousSibling.LastChild
      else
        Result := PreviousSibling;
    end else
      Result := ParentNode;
  end;
end;

function TRtfTreeNode.GetNodeIndex: Integer;
begin
  Result := -1;
  if Assigned(ParentNode) then
    Result := ParentNode.ChildNodes.IndexOf(Self);
end;

function TRtfTreeNode.GetNodeText: String;
begin
  Result := GetText(False);
end;

function TRtfTreeNode.GetRawText: String;
begin
  Result := GetText(True);
end;

function TRtfTreeNode.EncodeText(const S: String): String;
var
  i: Integer;
  Nodes: TRtfTreeNodeCollection;

begin
  Result := '';
  Nodes := CreateText(S);
  try
    for i := 0 to Nodes.Count - 1 do
    begin
      if i < Nodes.Count - 1 then
        Result := Result + Nodes[i].AsRtf(Nodes[i + 1])
      else
        Result := Result + Nodes[i].AsRtf(nil);
    end;
  finally
    FreeAndNil(Nodes);
  end;
end;

function TRtfTreeNode.GetRtf: String;
begin
  Result := GetRtfInm(Self, nil);
end;

function TRtfTreeNode.GetRtfInm(ANode, ANextNode: TRtfTreeNode): String;
var
  i: Integer;

begin
  Result := '';
  if ANode.NodeType = ntRoot then
    // skip root node
  else if ANode.NodeType = ntWhitespace then
    Result := Result + ANode.NodeKey
  else if ANode.NodeType = ntGroup then
    Result := Result + TRtfChar.BlockStart
  else
    Result := ANode.AsRtf(ANextNode, True);
  for i := 0 to ANode.ChildNodes.Count - 1 do
    Result := Result + GetRtfInm(ANode.ChildNodes[i], ANode.ChildNodes[i].NextNode);
  if ANode.NodeType = ntGroup then
    Result := Result + TRtfChar.BlockEnd;
end;

function TRtfTreeNode.AsRtf(ANext: TRtfTreeNode; const AEncode: Boolean = False): String;
begin
  Result := '';
  if NodeType in [ntKeyword, ntControl] then
    Result := Result + TRtfChar.KeywordMarker;
  if AEncode then
    Result := Result + EncodeText(EscapeText(NodeKey))
  else
    Result := Result + NodeKey;
  if HasParameter then
  begin
    if NodeType = ntKeyword then
      Result := Result + IntToStr(Parameter)
    else if (NodeType = ntControl) and (NodeKey = TRtfChar.HexMarker) then
      Result := Result + LowerCase(IntToHex(Parameter, 2));
  end;
  // append keyword end delimeter (SPACE)
  if (NodeType in [ntKeyword]) and (ANext <> nil) and ANext.MustUseKeywordStop then
    Result := Result + TRtfChar.KeywordStop;
end;

procedure TRtfTreeNode.UpdateNodeRoot(ANode: TRtfTreeNode);
var
  i: Integer;

begin
  ANode.RootNode := RootNode;
  ANode.Tree := Tree;
  for i := 0 to ANode.ChildNodes.Count - 1 do
    UpdateNodeRoot(ANode.ChildNodes[i]);
end;

function TRtfTreeNode.GetText(const Raw: Boolean): String;
begin
  Result := GetText(Raw, 1);
end;

function TRtfTreeNode.GetText(const Raw: Boolean; const IgnoreNChars: Integer): String;
var
  ANode: TRtfTreeNode;
  i, uc: Integer;

begin
  Result := '';
  case NodeType of
  ntGroup:
    begin
      ANode := FirstChild;
      if ANode.IsEquals('*') then
        ANode := ANode.NextSibling;
      if (Raw or not ANode.IsEquals(['fonttbl', 'colortbl', 'stylesheet', 'generator',
        'info', 'pict', 'object', 'fldinst'])) then
      begin
        uc := IgnoreNChars;
        for i := 0 to ChildNodes.Count - 1 do
        begin
          Result := Result + ChildNodes[i].GetText(Raw, uc);
          if (ChildNodes[i].NodeType = ntKeyword) and ChildNodes[i].IsEquals('uc') then
            uc := ChildNodes[i].Parameter;
        end;
      end;
    end;
  ntControl:
    begin
      if NodeKey = TRtfChar.HexMarker then
        Result := Result + DecodeChar(Parameter, Tree.Encoding)
      else if NodeKey = '~' then
        Result := Result + ' ';
    end;
  ntText:
    begin
      if (PreviousNode.NodeType = ntKeyword) and PreviousNode.IsEquals('u') then
        Result := Result + RightStr(NodeKey, Length(NodeKey) - IgnoreNChars)
      else
        Result := Result + NodeKey;
    end;
  ntKeyword:
    begin
      if IsEquals('par') then
        Result := Result + #13#10
      else if IsEquals('tab') then
        Result := Result + #7
      else if IsEquals('line') then
        Result := Result + #13#10
      else if IsEquals('lquote') then
        Result := Result + '‘'
      else if IsEquals('rquote') then
        Result := Result + '’'
      else if IsEquals('ldblquote') then
        Result := Result + '“'
      else if IsEquals('rdblquote') then
        Result := Result + '”'
      else if IsEquals('emdash') then
        Result := Result + '—'
      else if IsEquals('u') and (Parameter > 255) then
        Result := Result + DecodeChar(Parameter, TEncoding.Unicode);
    end;
  end;
end;

function TRtfTreeNode.IsGroupHasKey(const AKey: String; const IgnoreSpecial: Boolean): Boolean;
begin
  Result := False;
  if (NodeType = ntGroup) and HasChildNodes and
    (
      (
        FirstChild.NodeKey = AKey
      ) or
      (
        IgnoreSpecial and
        (FirstChild.NodeKey = '*') and
        (ChildNodes.Count > 1) and
        (ChildNodes[1].NodeKey = AKey)
      )
    )
    then
    Result := True;
end;

function TRtfTreeNode.IsPlainText: Boolean;
begin
  Result := (NodeType = ntText);
  if Result and Assigned(ParentNode) and
    (ParentNode.FirstChild.NodeType = ntKeyword) and (ParentNode.FirstChild.NodeKey = '*') then
  begin
    Result := False;
  end;
end;

function TRtfTreeNode.MustUseKeywordStop: Boolean;
begin
  Result := True;
  if NodeType in [ntKeyword, ntControl, ntGroup] then
    Result := False
  else if (NodeType = ntText) and (NodeKey <> '') and (
    (NodeKey = ';') or TRtfChar.MustInHex(NodeKey[1]) or TRtfChar.IsEscapable(NodeKey[1])
  ) then
    Result := False;
end;

function TRtfTreeNode.CreateCollection: TRtfTreeNodeCollection;
begin
  Result := TRtfTreeNodeCollection.Create;
  FGarbages.Add(Result);
end;

procedure TRtfTreeNode.RemoveCollection(ACollection: TRtfTreeNodeCollection);
begin
  if Assigned(ACollection) then
  begin
    if FGarbages.IndexOf(ACollection) > -1 then
      FGarbages.Delete(FGarbages.IndexOf(ACollection));
    FreeAndNil(ACollection);
  end;
end;

function TRtfTreeNode.SelectChildNodesForText(const AText: String; const StartPos: Integer): TRtfTreeNodeCollection;
var
  i, StartIdx, EndIdx: Integer;
  S: String;

begin
  Result := nil;
  StartIdx := -1;
  EndIdx := -1;
  S := '';
  for i := 0 to ChildNodes.Count - 1 do
  begin
    // collect text
    S := S + ChildNodes[i].Text;
    // text
    if Length(S) >= StartPos then
    begin
      if StartIdx < 0 then
        StartIdx := i;
      if Pos(AText, S) = StartPos then
      begin
        EndIdx := i;
        Break;
      end;
    end;
  end;
  if StartIdx > -1 then
  begin
    // only single node matched and it is a group
    if (StartIdx = EndIdx) and (ChildNodes[StartIdx].NodeType = ntGroup) then
      Result := ChildNodes[StartIdx].SelectChildNodesForText(AText, StartPos)
    else
    begin
      Result := CreateCollection;
      for i := StartIdx to EndIdx do
        Result.Add(ChildNodes[i]);
    end;
  end;
end;

function TRtfTreeNode.CombineNodesText(Nodes: TRtfTreeNodeCollection): TRtfTreeNode;
var
  ANode: TRtfTreeNode;

begin
  Result := nil;
  if Assigned(Nodes) and (Nodes.Count > 1) then
  begin
    Result := Nodes[0];
    if Result.NodeType = ntGroup then
      Result := Result.LastChild;
    while Nodes.Count > 1 do
    begin
      ANode := Nodes[1];
      Result.NodeKey := Result.NodeKey + ANode.Text;
      if Assigned(ANode.ParentNode) then
        ANode.ParentNode.RemoveChild(ANode);
      Nodes.Delete(1);
    end;
  end;
end;

procedure TRtfTreeNode.AppendChild(ANode: TRtfTreeNode);
begin
  if ANode <> nil then
  begin
    ANode.ParentNode := Self;
    UpdateNodeRoot(ANode);
    ChildNodes.Add(ANode);
  end;
end;

procedure TRtfTreeNode.AppendChild(ACollection: TRtfTreeNodeCollection);
var
  i: Integer;

begin
  if Assigned(ACollection) then
    for i := 0 to ACollection.Count - 1 do
      AppendChild(ACollection[i]);
end;

procedure TRtfTreeNode.InsertChild(const AIndex: Integer; ANode: TRtfTreeNode);
begin
  if (ANode <> nil) and (AIndex > -1) and (AIndex <= FChildren.Count) then
  begin
    ANode.ParentNode := Self;
    UpdateNodeRoot(ANode);
    ChildNodes.Insert(AIndex, ANode);
  end;
end;

procedure TRtfTreeNode.RemoveChild(const AIndex: Integer);
begin
  ChildNodes.Remove(AIndex, 1);
end;

procedure TRtfTreeNode.RemoveChild(ANode: TRtfTreeNode);
begin
  ChildNodes.Remove(ChildNodes.IndexOf(ANode), 1);
end;

function TRtfTreeNode.CloneNode: TRtfTreeNode;
var
  i: Integer;

begin
  Result := TRtfTreeNode.Create;
  Result.NodeType := NodeType;
  Result.NodeKey := NodeKey;
  Result.HasParameter := HasParameter;
  Result.Parameter := Parameter;
  for i := 0 to ChildNodes.Count - 1 do
    Result.AppendChild(ChildNodes[i].CloneNode);
end;

function TRtfTreeNode.HasChildNodes: Boolean;
begin
  Result := ChildNodes.Count > 0;
end;

function TRtfTreeNode.SelectSingleNode(const ANodeType: TRtfTreeNodeType): TRtfTreeNode;
var
  i: Integer;

begin
  Result := nil;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].NodeType = ANodeType then
    begin
      Result := ChildNodes[i];
      Break;
    end else
    begin
      Result := ChildNodes[i].SelectSingleNode(ANodeType);
      if Result <> nil then
        Break;
    end;
end;

function TRtfTreeNode.SelectSingleNode(const AKey: String): TRtfTreeNode;
var
  i: Integer;

begin
  Result := nil;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].NodeKey = AKey then
    begin
      Result := ChildNodes[i];
      Break;
    end else
    begin
      Result := ChildNodes[i].SelectSingleNode(AKey);
      if Result <> nil then
        Break;
    end;
end;

function TRtfTreeNode.SelectSingleNode(const AKey: String; const AParam: Integer): TRtfTreeNode;
var
  i: Integer;

begin
  Result := nil;
  for i := 0 to ChildNodes.Count - 1 do
    if (ChildNodes[i].NodeKey = AKey) and (ChildNodes[i].Parameter = AParam) then
    begin
      Result := ChildNodes[i];
      Break;
    end else
    begin
      Result := ChildNodes[i].SelectSingleNode(AKey, AParam);
      if Result <> nil then
        Break;
    end;
end;

function TRtfTreeNode.SelectSingleGroup(const AKey: String): TRtfTreeNode;
begin
  Result := SelectSingleGroup(AKey, False);
end;

function TRtfTreeNode.SelectSingleGroup(const AKey: String; const IgnoreSpecial: Boolean): TRtfTreeNode;
var
  i: Integer;

begin
  Result := nil;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].IsGroupHasKey(AKey, IgnoreSpecial) then
    begin
      Result := ChildNodes[i];
      Break;
    end else
    begin
      Result := ChildNodes[i].SelectSingleGroup(AKey, IgnoreSpecial);
      if Result <> nil then
        Break;
    end;
end;

function TRtfTreeNode.SelectNodes(const ANodeType: TRtfTreeNodeType): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
  begin
    if ChildNodes[i].NodeType = ANodeType then
      Result.Add(ChildNodes[i]);
    Result.Append(ChildNodes[i].SelectNodes(ANodeType));
  end;
end;

function TRtfTreeNode.SelectNodes(const AKey: String): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
  begin
    if ChildNodes[i].NodeKey = AKey then
      Result.Add(ChildNodes[i]);
    Result.Append(ChildNodes[i].SelectNodes(AKey));
  end;
end;

function TRtfTreeNode.SelectNodes(const AKey: String; const AParam: Integer): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
  begin
    if (ChildNodes[i].NodeKey = AKey) and (ChildNodes[i].Parameter = AParam) then
      Result.Add(ChildNodes[i]);
    Result.Append(ChildNodes[i].SelectNodes(AKey, AParam));
  end;
end;

function TRtfTreeNode.SelectGroups(const AKey: String): TRtfTreeNodeCollection;
begin
  Result := SelectGroups(AKey, False);
end;

function TRtfTreeNode.SelectGroups(const AKey: String; const IgnoreSpecial: Boolean): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
  begin
    if ChildNodes[i].IsGroupHasKey(AKey, IgnoreSpecial) then
      Result.Add(ChildNodes[i]);
    Result.Append(ChildNodes[i].SelectGroups(AKey, IgnoreSpecial));
  end;
end;

function TRtfTreeNode.SelectSingleChildNode(const ANodeType: TRtfTreeNodeType): TRtfTreeNode;
var
  i: Integer;

begin
  Result := nil;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].NodeType = ANodeType then
    begin
      Result := ChildNodes[i];
      Break;
    end;
end;

function TRtfTreeNode.SelectSingleChildNode(const AKey: String): TRtfTreeNode;
var
  i: Integer;

begin
  Result := nil;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].NodeKey = AKey then
    begin
      Result := ChildNodes[i];
      Break;
    end;
end;

function TRtfTreeNode.SelectSingleChildNode(const AKey: String; AParam: Integer): TRtfTreeNode;
var
  i: Integer;

begin
  Result := nil;
  for i := 0 to ChildNodes.Count - 1 do
    if (ChildNodes[i].NodeKey = AKey) and (ChildNodes[i].Parameter = AParam) then
    begin
      Result := ChildNodes[i];
      Break;
    end;
end;

function TRtfTreeNode.SelectSingleChildGroup(const AKey: String): TRtfTreeNode;
begin
  Result := SelectSingleChildGroup(AKey, False);
end;

function TRtfTreeNode.SelectSingleChildGroup(const AKey: String; const IgnoreSpecial: Boolean): TRtfTreeNode;
var
  i: Integer;

begin
  Result := nil;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].IsGroupHasKey(AKey, IgnoreSpecial) then
    begin
      Result := ChildNodes[i];
      Break;
    end;
end;

function TRtfTreeNode.SelectChildNodes(const ANodeType: TRtfTreeNodeType): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].NodeType = ANodeType then
      Result.Add(ChildNodes[i]);
end;

function TRtfTreeNode.SelectChildNodes(const AKey: String): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].NodeKey = AKey then
      Result.Add(ChildNodes[i]);
end;

function TRtfTreeNode.SelectChildNodes(const AKey: String; const AParam: Integer): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
    if (ChildNodes[i].NodeKey = AKey) and (ChildNodes[i].Parameter = AParam) then
      Result.Add(ChildNodes[i]);
end;

function TRtfTreeNode.SelectChildGroups(const AKey: String): TRtfTreeNodeCollection;
begin
  Result := SelectChildGroups(AKey, False);
end;

function TRtfTreeNode.SelectChildGroups(const AKey: String; const IgnoreSpecial: Boolean): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].IsGroupHasKey(AKey, IgnoreSpecial) then
      Result.Add(ChildNodes[i]);
end;

function TRtfTreeNode.SelectSibling(const ANodeType: TRtfTreeNodeType): TRtfTreeNode;
var
  i, AIndex: Integer;

begin
  Result := nil;
  AIndex := Index;
  if AIndex > -1 then
    for i := AIndex + 1 to ParentNode.ChildNodes.Count - 1 do
      if ParentNode.ChildNodes[i].NodeType = ANodeType then
      begin
        Result := ParentNode.ChildNodes[i];
        Break;
      end;
end;

function TRtfTreeNode.SelectSibling(const AKey: String): TRtfTreeNode;
var
  i, AIndex: Integer;

begin
  Result := nil;
  AIndex := Index;
  if AIndex > -1 then
    for i := AIndex + 1 to ParentNode.ChildNodes.Count - 1 do
      if ParentNode.ChildNodes[i].NodeKey = AKey then
      begin
        Result := ParentNode.ChildNodes[i];
        Break;
      end;
end;

function TRtfTreeNode.SelectSibling(const AKey: String; const AParam: Integer): TRtfTreeNode;
var
  i, AIndex: Integer;

begin
  Result := nil;
  AIndex := Index;
  if AIndex > -1 then
    for i := AIndex + 1 to ParentNode.ChildNodes.Count - 1 do
      if (ParentNode.ChildNodes[i].NodeKey = AKey) and (ParentNode.ChildNodes[i].Parameter = AParam) then
      begin
        Result := ParentNode.ChildNodes[i];
        Break;
      end;
end;

function TRtfTreeNode.FindText(const AText: String): TRtfTreeNodeCollection;
var
  i: Integer;

begin
  Result := CreateCollection;
  for i := 0 to ChildNodes.Count - 1 do
  begin
    if (ChildNodes[i].NodeType = ntText) and (Pos(AText, ChildNodes[i].NodeKey) > 0) then
      Result.Add(ChildNodes[i]);
    Result.Append(ChildNodes[i].FindText(AText));
  end;
end;

procedure TRtfTreeNode.ReplaceText(const OldValue, NewValue: String);
var
  i: Integer;

begin
  for i := 0 to ChildNodes.Count - 1 do
  begin
    if (ChildNodes[i].NodeType = ntText) and (Pos(OldValue, ChildNodes[i].NodeKey) > 0) then
      ChildNodes[i].NodeKey := ReplaceStr(ChildNodes[i].NodeKey, OldValue, NewValue);
    ChildNodes[i].ReplaceText(OldValue, NewValue);
  end;
end;

function TRtfTreeNode.ReplaceTextEx(const AFrom, ATo: String): Boolean;
var
  i, P: Integer;
  ANodes: TRtfTreeNodeCollection;
  CNode: TRtfTreeNode;

begin
  Result := False;
  // found exact match on text node
  if (NodeType = ntText) then
  begin
    if ((ParentNode = nil) or IsPlainText) and (Pos(AFrom, NodeKey) > 0) then
    begin
      NodeKey := ReplaceStr(NodeKey, AFrom, ATo);
      Result := True;
      Exit;
    end;
  end;
  // found exact match on child
  for i := 0 to ChildNodes.Count - 1 do
    if ChildNodes[i].ReplaceTextEx(AFrom, ATo) then
    begin
      Result := True;
      Exit;
    end;
  // found match on across child node
  P := Pos(AFrom, Text);
  if (P > 0) then
  begin
    ANodes := SelectChildNodesForText(AFrom, P);
    if Assigned(ANodes) then
    begin
      CNode := CombineNodesText(ANodes);
      if CNode <> nil then
      begin
        CNode.ReplaceTextEx(AFrom, ATo);
        Result := True;
      end;
    end;
  end;
end;

function TRtfTreeNode.ToString: String;
begin
  Result := '[' + NodeTypeStrings[NodeType] + ', ' + NodeKey + ', ' + IfThen(HasParameter, 'True', 'False') + ', ' + IntToStr(Parameter) + ']';
end;

function TRtfTreeNode.IsEquals(const AKey: String): Boolean;
begin
  Result := SameText(AKey, NodeKey);
end;

function TRtfTreeNode.IsEquals(const Keys: array of String): Boolean;
var
  i: Integer;

begin
  Result := False;
  for i := Low(Keys) to High(Keys) do
    if IsEquals(Keys[i]) then
    begin
      Result := True;
      Break;
    end;
end;

procedure TRtfTreeNode.Clear;
begin
  // clean garbages
  while FGarbages.Count > 0 do
  begin
    TObject(FGarbages[0]).Free;
    FGarbages.Delete(0);
  end;
  // remove children
  ChildNodes.Clear;
end;

class function TRtfTreeNode.DecodeChar(Code: Integer; Encoding: TEncoding): String;
var
  Buff: TBytes;
  L: Integer;
  B: Byte;

begin
  Result := '';
  L := 0;
  while True do
  begin
    if Code = 0 then
      Break;
    Inc(L);
    B := Byte(Code and $FF);
    Code := Code shr 8;
    SetLength(Buff, L);
    Buff[L - 1] := B;
  end;
  Result := Encoding.GetString(Buff);
end;

class function TRtfTreeNode.EscapeText(const S: String): String;
var
  i: Integer;
  Ch: Char;

begin
  Result := '';
  for i := 1 to Length(S) do
  begin
    Ch := S[i];
    if TRtfChar.IsEscapable(Ch) then
      Result := Result + TRtfChar.KeywordMarker;
    Result := Result + Ch;
  end;
end;

class function TRtfTreeNode.CreateText(const AText: String): TRtfTreeNodeCollection;
var
  i, Len: Integer;
  Ch: Char;
  Bytes: TBytes;
  S: String;

begin
  Result := TRtfTreeNodeCollection.Create;
  i := 0;
  Len := Length(AText);
  while True do
  begin
    if i = Len then
      Break;
    Inc(i);
    Ch := AText[i];
    if Ch = #7 then
      Result.Add(TRtfTreeNode.Create(ntKeyword, 'tab'))
    else if (Ch = #13) and (i + 1 <= Len) and (AText[i + 1] = #10) then
    begin
      Result.Add(TRtfTreeNode.Create(ntKeyword, 'par'));
      Inc(i);
    end
    else if TRtfChar.MustInHex(Ch) then
      Result.Add(TRtfTreeNode.Create(ntControl, '''', True, Ord(Ch)))
    else if TRtfChar.IsPlain(Ch) then
    begin
      S := '';
      while True do
      begin
        S := S + Ch;
        if (i + 1 <= Len) and TRtfChar.IsPlain(AText[i + 1]) then
        begin
          Inc(i);
          Ch := AText[i];
          Continue;
        end;
        Break;
      end;
      Result.Add(TRtfTreeNode.Create(ntText, S));
    end else
    // unicode
    begin
      Bytes := TEncoding.Unicode.GetBytes(Ch);
      Result.Add(TRtfTreeNode.Create(ntKeyword, 'u', True, (Bytes[1] shl 8) + Bytes[0]));
      Result.Add(TRtfTreeNode.Create(ntText, '?'));
    end;
  end;
end;

{ TRtfTreeNodes }

constructor TRtfTreeNodes.Create;
begin
  FItems := TList.Create;
end;

destructor TRtfTreeNodes.Destroy;
begin
  FreeAndNil(FItems);
  inherited;
end;

function TRtfTreeNodes.GetNode(const AIndex: Integer): TRtfTreeNode;
begin
  Result := nil;
  if (AIndex > -1) and (AIndex < FItems.Count) then
    Result := TRtfTreeNode(FItems[AIndex]);
end;

function TRtfTreeNodes.Count: Integer;
begin
  Result := FItems.Count;
end;

procedure TRtfTreeNodes.Clear;
begin
  while FItems.Count > 0 do
  begin
    if TObject(FItems[0]) <> nil then
      TObject(FItems[0]).Free;
    FItems.Delete(0);
  end;
end;

function TRtfTreeNodes.Add(ANode: TRtfTreeNode): Integer;
begin
  FItems.Add(ANode);
  Result := FItems.Count - 1;
end;

procedure TRtfTreeNodes.Insert(const AIndex: Integer; ANode: TRtfTreeNode);
begin
  FItems.Insert(AIndex, ANode);
end;

function TRtfTreeNodes.IndexOf(ANode: TRtfTreeNode): Integer;
begin
  Result := FItems.IndexOf(ANode);
end;

function TRtfTreeNodes.IndexOf(ANode: TRtfTreeNode; const StartIndex: Integer): Integer;
var
  i: Integer;

begin
  Result := -1;
  for i := StartIndex to FItems.Count - 1 do
    if TRtfTreeNode(FItems[i]) = ANode then
    begin
      Result := i;
      Break;
    end;
end;

function TRtfTreeNodes.IndexOf(const AKey: String): Integer;
begin
  Result := IndexOf(AKey, 0);
end;

function TRtfTreeNodes.IndexOf(const AKey: String; const StartIndex: Integer): Integer;
var
  i: Integer;

begin
  Result := -1;
  for i := StartIndex to FItems.Count - 1 do
    if TRtfTreeNode(FItems[i]).NodeKey = AKey then
    begin
      Result := i;
      Break;
    end;
end;

procedure TRtfTreeNodes.Remove(const AIndex: Integer; const ACount: Integer = 1);
var
  Cnt: Integer;

begin
  Cnt := ACount;
  while Cnt > 0 do
  begin
    Dec(Cnt);
    if (AIndex > -1) and (AIndex < FItems.Count - 1) then
    begin
      TObject(FItems[AIndex]).Free;
      FItems.Delete(AIndex);
    end;
  end;
end;

{ TRtfTreeNodeCollection }

constructor TRtfTreeNodeCollection.Create;
begin
  FItems := TList.Create;
end;

destructor TRtfTreeNodeCollection.Destroy;
begin
  FItems.Free;
end;

function TRtfTreeNodeCollection.GetItem(const AIndex: Integer): TRtfTreeNode;
begin
  Result := nil;
  if (AIndex > -1) and (AIndex < FItems.Count) then
    Result := TRtfTreeNode(FItems[AIndex]);
end;

function TRtfTreeNodeCollection.Count: Integer;
begin
  Result := FItems.Count;
end;

procedure TRtfTreeNodeCollection.Add(ANode: TRtfTreeNode);
begin
  FItems.Add(ANode);
end;

procedure TRtfTreeNodeCollection.Clear;
begin
  FItems.Clear;
end;

procedure TRtfTreeNodeCollection.Append(ACollection: TRtfTreeNodeCollection);
var
  i: Integer;

begin
  for i := 0 to ACollection.Count - 1 do
    Add(ACollection[i]);
end;

procedure TRtfTreeNodeCollection.Delete(const AIndex: Integer);
begin
  if (AIndex > -1) and (AIndex < FItems.Count) then
    FItems.Delete(AIndex);
end;

{ TRtfObject }

constructor TRtfObject.Create(ANode: TRtfTreeNode);
begin
  FNode := ANode;
  FHexData := TStringStream.Create;
  FBinaryData := TStringStream.Create;
  Initialize;
end;

destructor TRtfObject.Destroy;
begin
  Finalize;
  FreeAndNil(FHexData);
  FreeAndNil(FBinaryData);
  inherited;
end;

function TRtfObject.GetHexData: TStream;
begin
  Result := FHexData;
end;

function TRtfObject.GetBinaryData: TStream;
begin
  Result := FBinaryData;
end;

function TRtfObject.ReadProp(const AProp: String; var Value: Integer;
  const ADefault: Integer = -1): Boolean;
var
  ANode: TRtfTreeNode;

begin
  ANode := nil;
  if FNode <> nil then
    ANode := FNode.SelectSingleChildNode(AProp);
  Result := ANode <> nil;
  if Result then
    Value := ANode.Parameter
  else
    Value := ADefault;
end;

function TRtfObject.FindChild(const Keys: array of String; var S: String): Boolean;
var
  i: Integer;
  ANode: TRtfTreeNode;

begin
  Result := False;
  if FNode = nil then Exit;
  for i := Low(Keys) to High(Keys) do
  begin
    ANode := FNode.SelectSingleChildNode(Keys[i]);
    if ANode <> nil then
    begin
      Result := True;
      S := Keys[i];
      Break;
    end;
  end;
end;

procedure TRtfObject.Initialize;
begin

end;

procedure TRtfObject.Finalize;
begin

end;

procedure TRtfObject.ToBinary(HexStream, BinStream: TStream);
var
  HexBuffer, Buffer: TBytes;
  Count: Integer;

begin
  HexStream.Position := 0;
  BinStream.Size := 0;
  SetLength(Buffer, HEX_BUF_SIZE);
  SetLength(HexBuffer, HEX_BUF_SIZE * 2);
  while True do
  begin
    Count := HexStream.Read(HexBuffer, HEX_BUF_SIZE * 2);
    if Count = 0 then
      Break;
    //HexToBin(HexBuffer, Buffer, Count div 2);
    BinStream.Write(Buffer, Count div 2);
  end;
end;

procedure TRtfObject.ToHex(BinStream, HexStream: TStream);
var
  Buffer, HexBuffer: TBytes;
  Count: Integer;

begin
  BinStream.Position := 0;
  HexStream.Size := 0;
  SetLength(HexBuffer, HEX_BUF_SIZE * 2);
  SetLength(Buffer, HEX_BUF_SIZE);
  while True do
  begin
    Count := BinStream.Read(Buffer, HEX_BUF_SIZE);
    if Count = 0 then
      Break;
//    BinToHex(Buffer, 0, HexBuffer, 0, Count);
//    HexStream.Write(HexBuffer, Count * 2);
  end;
end;

{ TRtfStyleSheet }

constructor TRtfStyleSheet.Create;
begin
  FKeyCode := TRtfTreeNodes.Create;
  FFormatting := TRtfTreeNodes.Create;
  FType := stParagraph;
  FAdditive := False;
  FBasedOn := -1;
  FNext := -1;
  FAutoUpdate := False;
  FHidden := False;
  FLink := -1;
  FLocked := False;
  FPersonal := False;
  FCompose := False;
  FReply := False;
  FStyrsid := -1;
  FSemiHidden := False;
end;

destructor TRtfStyleSheet.Destroy;
begin
  FreeAndNil(FKeyCode);
  FreeAndNil(FFormatting);
  inherited;
end;

{ TRtfStyleSheets }

constructor TRtfStyleSheets.Create;
begin
  FItems := TList.Create;
end;

destructor TRtfStyleSheets.Destroy;
begin
  Clear;
  FreeAndNil(FItems);
  inherited;
end;

function TRtfStyleSheets.GetStyleSheet(const AIndex: Integer): TRtfStyleSheet;
begin
  Result := nil;
  if (AIndex > -1) and (AIndex < FItems.Count) then
    Result := TRtfStyleSheet(FItems[AIndex]);
end;

function TRtfStyleSheets.Count: Integer;
begin
  Result := FItems.Count;
end;

procedure TRtfStyleSheets.Clear;
begin
  while FItems.Count > 0 do
  begin
    TObject(FItems[0]).Free;
    FItems.Delete(0);
  end;
end;

function TRtfStyleSheets.Add(AStyleSheet: TRtfStyleSheet): Integer;
begin
  FItems.Add(AStyleSheet);
  Result := FItems.Count - 1;
end;

procedure TRtfStyleSheets.Insert(const AIndex: Integer; AStyleSheet: TRtfStyleSheet);
begin
  FItems.Insert(AIndex, AStyleSheet);
end;

function TRtfStyleSheets.IndexOf(AStyleSheet: TRtfStyleSheet): Integer;
begin
  Result := FItems.IndexOf(AStyleSheet);
end;

function TRtfStyleSheets.IndexOf(const AName: String): Integer;
var
  i: Integer;

begin
  Result := -1;
  for i := 0 to FItems.Count - 1 do
    if TRtfStyleSheet(FItems[i]).Name = AName then
    begin
      Result := i;
      Break;
    end;
end;

function TRtfStyleSheets.IndexOf(const AIndex: Integer): Integer;
var
  i: Integer;

begin
  Result := -1;
  for i := 0 to FItems.Count - 1 do
    if TRtfStyleSheet(FItems[i]).Index = AIndex then
    begin
      Result := i;
      Break;
    end;
end;

procedure TRtfStyleSheets.Remove(AStyleSheet: TRtfStyleSheet);
begin
  if FItems.IndexOf(AStyleSheet) > -1 then
  begin
    FItems.Delete(FItems.IndexOf(AStyleSheet));
    AStyleSheet.Free;
  end;
end;

procedure TRtfStyleSheets.Remove(const AIndex: Integer);
begin
  if (AIndex > -1) and (AIndex < FItems.Count) then
  begin
    TObject(FItems[AIndex]).Free;
    FItems.Delete(AIndex);
  end;
end;

{ TRtfTree }

constructor TRtfTree.Create;
begin
  FEncodingBuffer:=TEncodingBuffer.Create;
  FRoot := TRtfTreeNode.Create(ntRoot, 'ROOT');
  FRoot.Tree := Self;
  FMergeSpecialCharacters := False;
  FDocumentProp := TRtfDocumentProperty.Create;
  FFontTables := TRtfFontTables.Create;
  FColorTables := TRtfColorTables.Create;
  FStyleSheets := TRtfStyleSheets.Create;
  FEncoding := TEncoding.Default;
  FIgnoreWhitespace := True;
end;

destructor TRtfTree.Destroy;
begin
  FreeAndNil(FRoot);
  FreeAndNil(FDocumentProp);
  FreeAndNil(FFontTables);
  FreeAndNil(FColorTables);
  FreeAndNil(FStyleSheets);
  FEncodingBuffer.Free;
  inherited;
end;

function TRtfTree.CloneTree: TRtfTree;
var
  I: Integer;

begin
  Result := TRtfTree.Create;
  Result.MergeSpecialCharacters := MergeSpecialCharacters;
  Result.IgnoreWhitespace := IgnoreWhitespace;
  for i := 0 to RootNode.ChildNodes.Count - 1 do
    Result.RootNode.AppendChild(RootNode.ChildNodes[i].CloneNode);
end;

function TRtfTree.LoadFromStream(AStream: TStream): Integer;
begin
  ResetAll;
  FSource := AStream;
  FSource.Position := 0;
  FLex := DefaultRtfLex.Create(FSource);
  try
    Result := ParseRtfTree;
  finally
    FreeAndNil(FLex);
  end;
  // is RTF valid?
  if Result = 0 then
  begin
    GetEncoding;
    GetInfoGroup;
    GetFontTable;
    GetColorTable;
    GetStyleSheetTable;
  end;
end;

function TRtfTree.LoadFromFile(const FileName: String): Integer;
var
  AStream: TFileStream;

begin
  AStream := TFileStream.Create(FileName, fmOpenRead or fmShareDenyWrite);
  try
    Result := LoadFromStream(AStream);
  finally
    AStream.Free;
  end;
end;

function TRtfTree.LoadFromString(const S: String): Integer;
var
  AStream: TStringStream;

begin
  AStream := TStringStream.Create(S);
  try
    Result := LoadFromStream(AStream);
  finally
    AStream.Free;
  end;
end;

procedure TRtfTree.SaveToFile(const FileName: String);
var
  Stream: TStringStream;

begin
  Stream := TStringStream.Create(Rtf);
  try
    Stream.SaveToFile(FileName);
  finally
    Stream.Free;
  end;
end;

function TRtfTree.ToString: String;
begin
  Result := ToStringInm(RootNode, 0, False);
end;

function TRtfTree.ToStringEx: String;
begin
  Result := ToStringInm(RootNode, 0, True);
end;

procedure TRtfTree.AddMainGroup;
begin
  if not RootNode.HasChildNodes then
    RootNode.AppendChild(TRtfTreeNode.Create(ntGroup));
end;

procedure TRtfTree.ResetAll;
begin
  FEncoding := TEncoding.Default;
  FFontTables.Clear;
  FColorTables.Clear;
  FStyleSheets.Clear;
  FRoot.Clear;
end;

function TRtfTree.ForceTokenAsText(AToken: TRtfToken): Boolean;
begin
  Result := AToken.ControlIsText;
  if Result then
  begin
    AToken.TokenType := ttText;
    AToken.Key := TRtfTreeNode.DecodeChar(AToken.Parameter, FEncoding);
    AToken.HasParameter := False;
  end;
end;

function TRtfTree.TryMergeText(ANode: TRtfTreeNode; AToken: TRtfToken;
  const Deep: Boolean): Boolean;
var
  CNode: TRtfTreeNode;

begin
  Result := False;
  if not Deep then
  begin
    // do not merge escapable char {, }, and \
    if AToken.IsText and (Length(AToken.Key) = 1) and TRtfChar.IsEscapable(AToken.Key[1]) then
      Exit;
    // do not merge hex coded text
    if AToken.ControlIsText then
      Exit;
  end;
  if AToken.IsText or AToken.ControlIsText then
  begin
    CNode := ANode.LastChild;
    if Assigned(CNode) and (CNode.NodeType = ntText) then
    begin
      // do not merge if previous node is an escapable char {, }, and \
      if not Deep and (Length(CNode.NodeKey) = 1) and TRtfChar.IsEscapable(CNode.NodeKey[1]) then
        Exit;
      Result := True;
      ForceTokenAsText(AToken);
      CNode.NodeKey := CNode.NodeKey + AToken.Key;
    end;
  end;
end;

function TRtfTree.ParseRtfTree: Integer;
var
  AToken: TRtfToken;
  ANode, NNode: TRtfTreeNode;
  Merged: Boolean;

begin
  Result := 0;
  FLevel := 0;
  ANode := RootNode;
  while True do
  begin
    AToken := FLex.NextToken;
    try
      if AToken.TokenType = ttEof then
        Break;
      case AToken.TokenType of
        ttGroupStart:
          begin
            NNode := TRtfTreeNode.Create(ntGroup, 'GROUP');
            ANode.AppendChild(NNode);
            ANode := NNode;
            Inc(FLevel);
          end;
        ttGroupEnd:
          begin
            ANode := ANode.ParentNode;
            Dec(FLevel);
          end;
        ttKeyword,
        ttControl,
        ttText:
          begin
            Merged := False;
            if MergeSpecialCharacters then
            begin
              if TryMergeText(ANode, AToken, True) then
                Merged := True
              else
                ForceTokenAsText(AToken);
            end
            else if IgnoreWhitespace and TryMergeText(ANode, AToken, False) then
              Merged := True;
            if not Merged then
            begin
              NNode := TRtfTreeNode.Create(AToken);
              ANode.AppendChild(NNode);
              if MergeSpecialCharacters then
              begin
                if (FLevel = 1) and (NNode.NodeType = ntKeyword) and (NNode.NodeKey = 'ansicpg') then
                  FEncoding := FEncodingBuffer.GetEncoding(NNode.Parameter);
              end;
            end;
          end;
        ttWhitespace:
          begin
            if not IgnoreWhitespace then
              ANode.AppendChild(TRtfTreeNode.Create(AToken));
          end;
        else
          Result := -1;
      end;
    finally
      FreeAndNil(AToken);
    end;
  end;
  // incomplete level
  if FLevel <> 0 then
    Result := -1;
end;

procedure TRtfTree.GetEncoding;
var
  ANode: TRtfTreeNode;

begin
  ANode := RootNode.SelectSingleNode('ansicpg');
  if Assigned(ANode) then
    FEncoding := FEncodingBuffer.GetEncoding(ANode.Parameter);
end;

procedure TRtfTree.GetInfoGroup;
var
  ANode: TRtfTreeNode;

begin
  if RootNode.SelectSingleNode('info') <> nil then
  begin
    // Title
    ANode := RootNode.SelectSingleNode('title');
    if ANode <> nil then
      FDocumentProp.Title := ANode.NextSibling.NodeKey;
    // Subject
    ANode := RootNode.SelectSingleNode('subject');
    if ANode <> nil then
      FDocumentProp.Subject := ANode.NextSibling.NodeKey;
    // Author
    ANode := RootNode.SelectSingleNode('author');
    if ANode <> nil then
      FDocumentProp.Author := ANode.NextSibling.NodeKey;
    // Manager
    ANode := RootNode.SelectSingleNode('manager');
    if ANode <> nil then
      FDocumentProp.Manager := ANode.NextSibling.NodeKey;
    // Company
    ANode := RootNode.SelectSingleNode('company');
    if ANode <> nil then
      FDocumentProp.Company := ANode.NextSibling.NodeKey;
    // Operator
    ANode := RootNode.SelectSingleNode('operator');
    if ANode <> nil then
      FDocumentProp.Operator := ANode.NextSibling.NodeKey;
    // Category
    ANode := RootNode.SelectSingleNode('category');
    if ANode <> nil then
      FDocumentProp.Category := ANode.NextSibling.NodeKey;
    // Keywords
    ANode := RootNode.SelectSingleNode('keywords');
    if ANode <> nil then
      FDocumentProp.Keywords := ANode.NextSibling.NodeKey;
    // Comments
    ANode := RootNode.SelectSingleNode('comment');
    if ANode <> nil then
      FDocumentProp.Comment := ANode.NextSibling.NodeKey;
    // Document comments
    ANode := RootNode.SelectSingleNode('doccomm');
    if ANode <> nil then
      FDocumentProp.DocComment := ANode.NextSibling.NodeKey;
    // Hlinkbase (The base address that is used for the path of all relative hyperlinks inserted in the document)
    ANode := RootNode.SelectSingleNode('hlinkbase');
    if ANode <> nil then
      FDocumentProp.HLinkBase := ANode.NextSibling.NodeKey;
    // Version
    ANode := RootNode.SelectSingleNode('version');
    if ANode <> nil then
      FDocumentProp.Version := ANode.Parameter;
    // Internal Version
    ANode := RootNode.SelectSingleNode('vern');
    if ANode <> nil then
      FDocumentProp.InternalVersion := ANode.Parameter;
    // Editing Time
    ANode := RootNode.SelectSingleNode('edmins');
    if ANode <> nil then
      FDocumentProp.EditingTime := ANode.Parameter;
    // Number of Pages
    ANode := RootNode.SelectSingleNode('nofpages');
    if ANode <> nil then
      FDocumentProp.NumberOfPages := ANode.Parameter;
    // Number of Chars
    ANode := RootNode.SelectSingleNode('nofchars');
    if ANode <> nil then
      FDocumentProp.NumberOfChars := ANode.Parameter;
    // Number of Words
    ANode := RootNode.SelectSingleNode('nofwords');
    if ANode <> nil then
      FDocumentProp.NumberOfWords := ANode.Parameter;
    // Id
    ANode := RootNode.SelectSingleNode('id');
    if ANode <> nil then
      FDocumentProp.Id := ANode.Parameter;
    // Creation DateTime
    ANode := RootNode.SelectSingleNode('creatim');
    if ANode <> nil then
      FDocumentProp.CreationTime := ParseDateTime(ANode.ParentNode);
    // Revision DateTime
    ANode := RootNode.SelectSingleNode('revtim');
    if ANode <> nil then
      FDocumentProp.RevisionTime := ParseDateTime(ANode.ParentNode);
    // Last Print Time
    ANode := RootNode.SelectSingleNode('printim');
    if ANode <> nil then
      FDocumentProp.LastPrintTime := ParseDateTime(ANode.ParentNode);
    // Backup Time
    ANode := RootNode.SelectSingleNode('buptim');
    if ANode <> nil then
      FDocumentProp.BackupTime := ParseDateTime(ANode.ParentNode);
  end;
end;

procedure TRtfTree.GetFontTable;
var
  ANode, CNode: TRtfTreeNode;
  i, j: Integer;
  FontIndex: Integer;
  FontName: String;

begin
  ANode := nil;
  if Assigned(RootNode.FirstChild) then
    ANode := RootNode.FirstChild.SelectSingleGroup('fonttbl');
  if Assigned(ANode) then
  begin
    for i := 1 to ANode.ChildNodes.Count - 1 do
    begin
      CNode := ANode.ChildNodes[i];
      FontIndex := -1;
      FontName := '';
      for j := 0 to CNode.ChildNodes.Count - 1 do
      begin
        if CNode.ChildNodes[j].NodeKey = 'f' then
          FontIndex := CNode.ChildNodes[j].Parameter;
        if CNode.ChildNodes[j].NodeType = ntText then
          FontName := LeftStr(CNode.ChildNodes[j].NodeKey, Length(CNode.ChildNodes[j].NodeKey) - 1);
      end;
      if FontName <> '' then
      begin
        if FontIndex = -1 then
          FontIndex := FFontTables.Count;
        FFontTables.Add(TRtfFontTable.Create(FontIndex, FontName));
      end;
    end;
  end;
end;

procedure TRtfTree.GetColorTable;
var
  ANode, CNode: TRtfTreeNode;
  i: Integer;
  RGB: Longint;

begin
  ANode := nil;
  if Assigned(RootNode.FirstChild) then
    ANode := RootNode.FirstChild.SelectSingleGroup('colortbl');
  if Assigned(ANode) then
  begin
    RGB := 0;
    for i := 0 to ANode.ChildNodes.Count - 1 do
    begin
      CNode := ANode.ChildNodes[i];
      if CNode.NodeType = ntKeyword then
      begin
        if CNode.IsEquals('red') then
          RGB := RGB or (Byte(CNode.Parameter))
        else if CNode.IsEquals('green') then
          RGB := RGB or (Byte(CNode.Parameter) shl 8)
        else if CNode.IsEquals('blue') then
          RGB := RGB or (Byte(CNode.Parameter) shl 16);
      end;
      if CNode.IsEquals(';') then
      begin
        FColorTables.Add(TRtfColorTable.Create(TColor(RGB)));
        RGB := 0;
      end;
    end;
  end;
end;

procedure TRtfTree.GetStyleSheetTable;
var
  ANode: TRtfTreeNode;
  i: Integer;

begin
  ANode := MainGroup.SelectSingleGroup('stylesheet');
  if ANode <> nil then
    for i := 1 to ANode.ChildNodes.Count - 1 do
      FStyleSheets.Add(ParseStyleSheet(ANode.ChildNodes[i]));
end;

function TRtfTree.ToStringInm(ANode: TRtfTreeNode; const ALevel: Integer;
  const ShowNodeTypes: Boolean): String;
var
  i: Integer;

begin
  Result := '';
  for i := 1 to ALevel do
    Result := Result + '  ';
  case ANode.NodeType of
  ntRoot:
    Result := Result + 'ROOT';
  ntGroup:
    Result := Result + 'GROUP';
  ntWhitespace:
    Result := Result + 'WHITESPACE';
  else
    if ShowNodeTypes then
      Result := Result + NodeTypeStrings[ANode.NodeType] + ': ';
    Result := Result + ANode.NodeKey;
    if ANode.HasParameter then
      Result := Result + ' ' + IntToStr(ANode.Parameter);
  end;
  if Result <> '' then
    Result := Result + #13#10;
  for i := 0 to ANode.ChildNodes.Count - 1 do
    Result := Result + ToStringInm(ANode.ChildNodes[i], ALevel + 1, ShowNodeTypes);
end;

function TRtfTree.ParseDateTime(ANode: TRtfTreeNode): TDateTime;
var
  Yr, Mo, Dy, Hr, Min, Sec: Word;
  i: Integer;
  CNode: TRtfTreeNode;

begin
  Yr := 0;
  Mo := 0;
  Dy := 0;
  Hr := 0;
  Min := 0;
  Sec := 0;
  for i := 0 to ANode.ChildNodes.Count - 1 do
  begin
    CNode := ANode.ChildNodes[i];
    if CNode.NodeKey = 'yr' then
      Yr := CNode.Parameter
    else if CNode.NodeKey = 'mo' then
      Mo := CNode.Parameter
    else if CNode.NodeKey = 'dy' then
      Dy := CNode.Parameter
    else if CNode.NodeKey = 'hr' then
      Hr := CNode.Parameter
    else if CNode.NodeKey = 'min' then
      Min := CNode.Parameter
    else if CNode.NodeKey = 'sec' then
      Sec := CNode.Parameter;
  end;
  Result := EncodeDateTime(Yr, Mo, Dy, Hr, Min, Sec, 0);
end;

function TRtfTree.ConvertToText: String;
var
  ANode, CNode: TRtfTreeNode;
  i: Integer;

begin
  Result := '';
  ANode := MainGroup.SelectSingleChildNode('pard');
  if ANode <> nil then
    for i := ANode.Index to MainGroup.ChildNodes.Count - 1 do
    begin
      CNode := MainGroup.ChildNodes[i];
      Result := Result + CNode.Text;
    end;
end;

function TRtfTree.ParseStyleSheet(ANode: TRtfTreeNode): TRtfStyleSheet;
var
  CNode: TRtfTreeNode;
  i, j: Integer;

begin
  Result := TRtfStyleSheet.Create;
  for i := 0 to ANode.ChildNodes.Count - 1 do
  begin
    CNode := ANode.ChildNodes[i];
    if False then
    else if CNode.NodeKey = 'cs' then
    begin
      Result.StyleSheetType := stCharacter;
      Result.Index := CNode.Parameter;
    end
    else if CNode.NodeKey = 's' then
    begin
      Result.StyleSheetType := stParagraph;
      Result.Index := CNode.Parameter;
    end
    else if CNode.NodeKey = 'ds' then
    begin
      Result.StyleSheetType := stSection;
      Result.Index := CNode.Parameter;
    end
    else if CNode.NodeKey = 'ts' then
    begin
      Result.StyleSheetType := stTable;
      Result.Index := CNode.Parameter;
    end
    else if CNode.NodeKey = 'additive' then
      Result.Additive := True
    else if CNode.NodeKey = 'sbasedon' then
      Result.BasedOn := CNode.Parameter
    else if CNode.NodeKey = 'snext' then
      Result.Next := CNode.Parameter
    else if CNode.NodeKey = 'sautoupd' then
      Result.AutoUpdate := True
    else if CNode.NodeKey = 'shidden' then
      Result.Hidden := True
    else if CNode.NodeKey = 'slink' then
      Result.Link := CNode.Parameter
    else if CNode.NodeKey = 'slocked' then
      Result.Locked := True
    else if CNode.NodeKey = 'spersonal' then
      Result.Personal := True
    else if CNode.NodeKey = 'scompose' then
      Result.Compose := True
    else if CNode.NodeKey = 'sreply' then
      Result.Reply := True
    else if CNode.NodeKey = 'styrsid' then
      Result.Styrsid := CNode.Parameter
    else if CNode.NodeKey = 'ssemihidden' then
      Result.SemiHidden := True
    else if (CNode.NodeType = ntGroup) and (CNode.ChildNodes.Count > 2) and
      (CNode.ChildNodes[0].NodeKey = '*') and (CNode.ChildNodes[1].NodeKey = 'keycode')
      then
    begin
      for j := 2 to CNode.ChildNodes.Count - 1 do
        Result.KeyCode.Add(CNode.ChildNodes[j].CloneNode);
    end
    else if (CNode.NodeType = ntText) then
      Result.Name := LeftStr(CNode.NodeKey, Length(CNode.NodeKey) - 1)
    else if (CNode.NodeKey <> '*') then
      Result.Formatting.Add(CNode.CloneNode);
  end;
end;

function TRtfTree.GetMainGroup: TRtfTreeNode;
begin
  Result := nil;
  if RootNode.HasChildNodes then
    Result := RootNode.FirstChild;
end;

function TRtfTree.GetRtf: String;
begin
  Result := RootNode.Rtf;
end;

end.
