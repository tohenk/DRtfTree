unit RtfObj;

interface

uses
  Classes, SysUtils, StrUtils, RtfTree, RtfClasses;

type

  { TRtfEmbeddedObject }

  TRtfEmbeddedObject = class(TRtfObject)
  private
    FObjectType: String;
    FObjectClass: String;
    FWidth: Integer;
    FHeight: Integer;
  protected
    procedure Initialize; override;
    procedure Finalize; override;
  public
    procedure SaveToFile(const FileName: String);
  public
    property ObjectType: String read FObjectType write FObjectType;
    property ObjectClass: String read FObjectClass write FObjectClass;
    property Width: Integer read FWidth write FWidth;
    property Height: Integer read FHeight write FHeight;
  end;

implementation

{ TRtfEmbeddedObject }

procedure TRtfEmbeddedObject.Initialize;
var
  ANode: TRtfTreeNode;
  Nodes: TRtfTreeNodeCollection;
  i: Integer;
  HexStr: String;

begin
  ReadProp('objw', FWidth);
  ReadProp('objh', FHeight);
  FindChild(['objemb', 'objlink', 'objautlink', 'objsub', 'objpub', 'objicemb',
    'objhtml', 'objhtml'], FObjectType);
  if Assigned(Node) then
  begin
    // {\*\objclass Paint.Picture}
    ANode := Node.SelectSingleNode('objclass');
    if Assigned(ANode) and Assigned(ANode.NextSibling) then
      FObjectClass := ANode.NextSibling.NodeKey;
    // '{' \object (<objtype> & <objmod>? & <objclass>? & <objname>? & <objtime>? & <objsize>? & <rsltmod>?) ('{\*' \objdata (<objalias>? & <objsect>?) <data> '}') <result> '}'
    ANode := Node.SelectSingleNode('objdata');
    if Assigned(ANode) and Assigned(ANode.ParentNode) then
    begin
      HexStr := '';
      Nodes := ANode.ParentNode.SelectChildNodes(ntText);
      for i := 0 to Nodes.Count - 1 do
        HexStr := HexStr + Nodes[i].NodeKey;
      if HexStr <> '' then
      begin
        TStringStream(HexData).WriteString(HexStr);
        ToBinary(HexData, BinaryData);
      end;
    end;
  end;
end;

procedure TRtfEmbeddedObject.Finalize;
begin

end;

procedure TRtfEmbeddedObject.SaveToFile(const FileName: String);
begin
  if BinaryData.Size > 0 then
    TStringStream(BinaryData).SaveToFile(FileName);
end;

end.
