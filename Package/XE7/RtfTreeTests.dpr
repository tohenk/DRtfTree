program RtfTreeTests;
{

  Delphi DUnit Test Project
  -------------------------
  This project contains the DUnit test framework and the GUI/Console test runners.
  Add "CONSOLE_TESTRUNNER" to the conditional defines entry in the project options
  to use the console test runner.  Otherwise the GUI test runner will be used by
  default.

}

{$IFDEF CONSOLE_TESTRUNNER}
{$APPTYPE CONSOLE}
{$ENDIF}

uses
  DUnitTestRunner,
  RtfClasses in '..\..\Source\RtfClasses.pas',
  RtfToken in '..\..\Source\RtfToken.pas',
  RtfLex in '..\..\Source\RtfLex.pas',
  RtfTree in '..\..\Source\RtfTree.pas',
  RtfImg in '..\..\Source\RtfImg.pas',
  RtfObj in '..\..\Source\RtfObj.pas',
  RtfDoc in '..\..\Source\RtfDoc.pas',
  TestRtfBase in '..\..\Test\TestRtfBase.pas',
  TestRtfTreeNode in '..\..\Test\TestRtfTreeNode.pas',
  TestRtfTreeNodes in '..\..\Test\TestRtfTreeNodes.pas',
  TestRtfLoad in '..\..\Test\TestRtfLoad.pas',
  TestRtfSelection in '..\..\Test\TestRtfSelection.pas',
  TestRtfReplace in '..\..\Test\TestRtfReplace.pas',
  TestRtfHeaderSections in '..\..\Test\TestRtfHeaderSections.pas',
  TestRtfImage in '..\..\Test\TestRtfImage.pas',
  TestRtfObject in '..\..\Test\TestRtfObject.pas',
  TestRtfDoc in '..\..\Test\TestRtfDoc.pas';

{$R *.RES}

begin
  DUnitTestRunner.RunRegisteredTests;
end.

