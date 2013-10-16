unit TestRtfLoad;
{

  Delphi DUnit Test Case
  ----------------------
  This unit contains a skeleton test case class generated by the Test Case Wizard.
  Modify the generated code to correctly setup and call the methods from the unit
  being tested.

}

interface

uses
  TestFramework, StrUtils, Classes, SysUtils, TestRtfBase, RtfTree;

type

  TTestRtfLoad = class(TRtfTestCase)
  strict private
  public
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure TestLoadDoc;
    procedure TestLoadDocWhitespace;
    procedure TestLoadDocMergeSpecials;
    procedure TestLoadDocWithImage;
  end;

implementation

procedure TTestRtfLoad.SetUp;
begin
end;

procedure TTestRtfLoad.TearDown;
begin
end;

procedure TTestRtfLoad.TestLoadDoc;
var
  Rtf: TRtfTree;
  Res: Integer;

begin
  Rtf := TRtfTree.Create;
  try
    Res := Rtf.LoadFromFile(AssetPath + 'testdoc1.rtf');
    CheckEquals(0, Res, 'Loading rtf should return 0 if success');
    CheckEquals(False, Rtf.MergeSpecialCharacters, 'MergeSpecialCharacters is not False');
    Str.LoadFromFile(AssetPath + 'result1-1.txt');
    CheckEquals(Str.Text, Rtf.ToString, 'ToString not matched');
    Str.LoadFromFile(AssetPath + 'result1-2.txt');
    CheckEquals(Str.Text, Rtf.ToStringEx, 'ToStringEx not matched');
    Str.LoadFromFile(AssetPath + 'rtf1.txt');
    CheckEquals(Str.Text, Rtf.Rtf, 'Rtf property not matched');
    Str.LoadFromFile(AssetPath + 'text1.txt');
    CheckEquals(Str.Text, Rtf.Text, 'Text property not matched');
  finally
    FreeAndNil(Rtf);
  end;
end;

procedure TTestRtfLoad.TestLoadDocWhitespace;
var
  Rtf: TRtfTree;
  Res: Integer;

begin
  Rtf := TRtfTree.Create;
  try
    Rtf.IgnoreWhitespace := False;
    Res := Rtf.LoadFromFile(AssetPath + 'testdoc1.rtf');
    CheckEquals(0, Res, 'Loading rtf should return 0 if success');
    CheckEquals(False, Rtf.IgnoreWhitespace, 'IgnoreWhitespace is False');
    Str.LoadFromFile(AssetPath + 'result1-5.txt');
    CheckEquals(Str.Text, Rtf.ToString, 'ToString not matched');
    Str.LoadFromFile(AssetPath + 'result1-6.txt');
    CheckEquals(Str.Text, Rtf.ToStringEx, 'ToStringEx not matched');
    Str.LoadFromFile(AssetPath + 'testdoc1.rtf');
    CheckEquals(Str.Text, Rtf.Rtf, 'Rtf property not matched');
    Str.LoadFromFile(AssetPath + 'text1.txt');
    CheckEquals(Str.Text, Rtf.Text, 'Text property not matched');
  finally
    FreeAndNil(Rtf);
  end;
end;

procedure TTestRtfLoad.TestLoadDocMergeSpecials;
var
  Rtf: TRtfTree;
  Res: Integer;

begin
  Rtf := TRtfTree.Create;
  try
    Rtf.MergeSpecialCharacters := True;
    Res := Rtf.LoadFromFile(AssetPath + 'testdoc1.rtf');
    CheckEquals(0, Res, 'Loading rtf should return 0 if success');
    CheckEquals(True, Rtf.MergeSpecialCharacters, 'MergeSpecialCharacters is not True');
    Str.LoadFromFile(AssetPath + 'result1-3.txt');
    CheckEquals(Str.Text, Rtf.ToString, 'ToString not matched');
    Str.LoadFromFile(AssetPath + 'result1-4.txt');
    CheckEquals(Str.Text, Rtf.ToStringEx, 'ToStringEx not matched');
    Str.LoadFromFile(AssetPath + 'rtf1.txt');
    CheckEquals(Str.Text, Rtf.Rtf, 'Rtf property not matched');
    Str.LoadFromFile(AssetPath + 'text1.txt');
    CheckEquals(Str.Text, Rtf.Text, 'Text property not matched');
  finally
    FreeAndNil(Rtf);
  end;
end;

procedure TTestRtfLoad.TestLoadDocWithImage;
var
  Rtf: TRtfTree;
  Res: Integer;

begin
  Rtf := TRtfTree.Create;
  try
    Res := Rtf.LoadFromFile(AssetPath + 'testdoc3.rtf');
    CheckEquals(0, Res, 'Loading rtf should return 0 if success');
    CheckEquals(False, Rtf.MergeSpecialCharacters, 'MergeSpecialCharacters is not False');
    Str.LoadFromFile(AssetPath + 'rtf5.txt');
    CheckEquals(Str.Text, Rtf.Rtf, 'Rtf property not matched');
    Str.LoadFromFile(AssetPath + 'text2.txt');
    CheckEquals(Str.Text, Rtf.Text, 'Text property not matched');
  finally
    FreeAndNil(Rtf);
  end;
end;

initialization
  // Register any test cases with the test runner
  RegisterTest(TTestRtfLoad.Suite);

end.

