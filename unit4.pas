// ***********************************************************************
// ***********************************************************************
// fbNotex 1.x
// Author and copyright: Massimo Nardello, Modena (Italy) 2020-2021.
// Free software released under GPL licence version 3 or later.

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version. You can read the version 3
// of the Licence in http://www.gnu.org/licenses/gpl-3.0.txt
// or in the file Licence.txt included in the files of the
// source code of this software.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// ***********************************************************************
// ***********************************************************************

unit Unit4;

{$mode objfpc}{$H+}
{$modeswitch objectivec1}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ComCtrls,
  ExtCtrls, LazUTF8, CocoaThemes, CocoaAll, CocoaUtils;

type

  { TfmOptions }

  TfmOptions = class(TForm)
    bnClose: TButton;
    bnStFontColorDef: TButton;
    bnStFontHighColorMod: TButton;
    bnStFontColorMod: TButton;
    bnStFontMarkMod: TButton;
    bnStFontTitleColorMod: TButton;
    edStMaxLess: TButton;
    edStSizePlus: TButton;
    edStSizeLess: TButton;
    edStMaxChar: TEdit;
    edStMaxPlus: TButton;
    edStSizeTitlesPlus: TButton;
    edStSizeTitlesLess: TButton;
    cbStFonts: TComboBox;
    cdColorDialog: TColorDialog;
    edBackup: TEdit;
    edFBgbakPath: TEdit;
    edFBLibPath: TEdit;
    edServer: TEdit;
    edPath: TEdit;
    edPort: TEdit;
    edStSize: TEdit;
    edStSizeTitles: TEdit;
    lbLineSpacing: TLabel;
    lbBackup: TLabel;
    lbFBgbakPath: TLabel;
    lbFBLibPath: TLabel;
    lbParaSpacing: TLabel;
    lbServer: TLabel;
    lbPath: TLabel;
    lbPort: TLabel;
    lbStSize: TLabel;
    lbStSizeTitles: TLabel;
    lbStMaxChar: TLabel;
    lnStFonts: TLabel;
    tbLineSpacing: TTrackBar;
    tbParaSpacing: TTrackBar;
    procedure bnCloseClick(Sender: TObject);
    procedure bnStFontHighColorModClick(Sender: TObject);
    procedure bnStFontColorDefClick(Sender: TObject);
    procedure bnStFontColorModClick(Sender: TObject);
    procedure bnStFontMarkModClick(Sender: TObject);
    procedure bnStFontTitleColorModClick(Sender: TObject);
    procedure cbStFontsChange(Sender: TObject);
    procedure edStMaxLessClick(Sender: TObject);
    procedure edStMaxPlusClick(Sender: TObject);
    procedure edStSizeChange(Sender: TObject);
    procedure edStSizeLessClick(Sender: TObject);
    procedure edStSizePlusClick(Sender: TObject);
    procedure edStSizeTitlesChange(Sender: TObject);
    procedure edStSizeTitlesLessClick(Sender: TObject);
    procedure edStSizeTitlesPlusClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure tbLineSpacingChange(Sender: TObject);
    procedure tbParaSpacingChange(Sender: TObject);
  private

  public

  end;

var
  fmOptions: TfmOptions;

implementation

uses Unit1;

{$R *.lfm}


{ TfmOptions }

procedure TfmOptions.FormCreate(Sender: TObject);
var
  fm: NSFontManager;
  FontFamilies, FontMembers: NSArray;
  i, n: integer;
  flItalics, flBold: boolean;
begin
  fm := NSFontManager.sharedFontManager;
  FontFamilies := fm.availableFontFamilies;
  for i := 0 to FontFamilies.Count - 1 do
  begin
    flBold := False;
    flItalics := False;
    FontMembers := fm.availableMembersOfFontFamily(
      FontFamilies.objectAtIndex(i).description);
    for n := 0 to FontMembers.Count - 1 do
    begin
      if fm.fontNamed_hasTraits(FontMembers.objectAtIndex(n).objectAtIndex(
        0).description, NSFontBoldTrait) = True then
      begin
        flBold := True;
      end;
      if fm.fontNamed_hasTraits(FontMembers.objectAtIndex(n).objectAtIndex(
        0).description, NSFontItalicTrait) = True then
      begin
        if flBold = True then
        begin
          flItalics := True;
        end;
      end;
    end;
    if ((flBold = True) and (flItalics = True)) then
    begin
      cbStFonts.Items.Add(NSStringToString(FontFamilies.objectAtIndex(
        i).description));
    end;
    if IsPaintDark = False then
    begin
      lnStFonts.Font.Color := $001ABC0F;
      lbStSize.Font.Color := $001ABC0F;
      lbStSizeTitles.Font.Color := $001ABC0F;
      lbLineSpacing.Font.Color := $001ABC0F;
      lbParaSpacing.Font.Color := $001ABC0F;
      lbStMaxChar.Font.Color := $001ABC0F;
      lbFBLibPath.Font.Color := $001ABC0F;
      lbFBgbakPath.Font.Color := $001ABC0F;
      lbServer.Font.Color := $001ABC0F;
      lbPath.Font.Color := $001ABC0F;
      lbPort.Font.Color := $001ABC0F;
      lbBackup.Font.Color := $001ABC0F;
    end
    else
    begin
      lnStFonts.Font.Color := clLime;
      lbStSize.Font.Color := clLime;
      lbStSizeTitles.Font.Color := clLime;
      lbLineSpacing.Font.Color := clLime;
      lbParaSpacing.Font.Color := clLime;
      lbStMaxChar.Font.Color := clLime;
      lbFBLibPath.Font.Color := clLime;
      lbFBgbakPath.Font.Color := clLime;
      lbServer.Font.Color := clLime;
      lbPath.Font.Color := clLime;
      lbPort.Font.Color := clLime;
      lbBackup.Font.Color := clLime;
    end;
  end;
  // Does not work
  if cbStFonts.Items.IndexOf('Avenir Next Condensed') > -1 then
  begin
    cbStFonts.Items.Delete(cbStFonts.Items.IndexOf('Avenir Next Condensed'));
  end;
  cbStFonts.ItemIndex := 0;
end;

procedure TfmOptions.FormShow(Sender: TObject);
begin
  if fmMain.zcConnection.Connected = False then
  begin
    edFBLibPath.ReadOnly := False;
    edFBgbakPath.ReadOnly := False;
    edServer.ReadOnly := False;
    edPort.ReadOnly := False;
    edPath.ReadOnly := False;
  end
  else
  begin
    edFBLibPath.ReadOnly := True;
    edFBgbakPath.ReadOnly := True;
    edServer.ReadOnly := True;
    edPort.ReadOnly := True;
    edPath.ReadOnly := True;
  end;
  edStSize.Text := IntToStr(fmMain.dbText.Font.Size);
  edStSizeTitles.Text := IntToStr(fmMain.sgTitles.Font.Size);
  tbLineSpacing.Position := iLineSpacing;
  tbParaSpacing.Position := iParagraphSpacing;
  edStMaxChar.Text := IntToStr(iSimpleTextFrom);
  edFBLibPath.Text := fmMain.zcConnection.LibraryLocation;
  edFBgbakPath.Text := stGBackDir;
  edServer.Text := fmMain.zcConnection.HostName;
  edPort.Text := IntToStr(fmMain.zcConnection.Port);
  edPath.Text := fmMain.zcConnection.Database;
  edBackup.Text := stBackupFile;
end;

procedure TfmOptions.tbLineSpacingChange(Sender: TObject);
begin
  iLineSpacing := tbLineSpacing.Position;
  fmMain.SetLineParagraph;
  fmMain.FormatMarkers(2);
end;

procedure TfmOptions.tbParaSpacingChange(Sender: TObject);
begin
  iParagraphSpacing := tbParaSpacing.Position;
  fmMain.SetLineParagraph;
  fmMain.FormatMarkers(2);
end;

procedure TfmOptions.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  with fmMain do
  begin
    if zcConnection.Connected = False then
    begin
      fmMain.zcConnection.LibraryLocation := edFBLibPath.Text;
      stGBackDir := edFBgbakPath.Text;
      zcConnection.HostName := edServer.Text;
      zcConnection.Port := StrToInt(edPort.Text);
      zcConnection.Database := edPath.Text;
    end;
    stBackupFile := edBackup.Text;
    if edStMaxChar.Text = '' then
    begin
      edStMaxChar.Text := '100000';
    end;
    iSimpleTextFrom := StrToInt(edStMaxChar.Text);
    if UTF8LowerCase(edServer.Text) <> 'localhost' then
    begin
      fmMain.miToolsBackup.Enabled := False;
      fmMain.miToolsRestore.Enabled := False;
      fmMain.miToolsCompact.Enabled := False;
      fmMain.bnBackup.Enabled := False;
      fmMain.bnRestore.Enabled := False;
    end
    else
    if fmMain.zcConnection.Connected = False then
    begin
      fmMain.miToolsBackup.Enabled := True;
      fmMain.miToolsRestore.Enabled := True;
      fmMain.miToolsCompact.Enabled := True;
      fmMain.bnBackup.Enabled := True;
      fmMain.bnRestore.Enabled := True;
    end;
  end;
end;

procedure TfmOptions.cbStFontsChange(Sender: TObject);
begin
  fmMain.dbText.Font.Name := cbStFonts.Text;
  fmMain.sgTitles.Font.Name := cbStFonts.Text;
  fmMain.FormatMarkers(2);
end;

procedure TfmOptions.edStSizeChange(Sender: TObject);
begin
  fmMain.dbText.Font.Size := StrToInt(edStSize.Text);
  fmMain.SetLineParagraph;
  fmMain.FormatMarkers(2);
end;

procedure TfmOptions.edStSizeLessClick(Sender: TObject);
begin
  if StrToInt(edStSize.Text) > 7 then
  begin
    edStSize.Text := IntToStr(StrToInt(edStSize.Text) - 1);
  end;
end;

procedure TfmOptions.edStSizePlusClick(Sender: TObject);
begin
  if StrToInt(edStSize.Text) < 256 then
  begin
    edStSize.Text := IntToStr(StrToInt(edStSize.Text) + 1);
  end;
end;

procedure TfmOptions.edStSizeTitlesChange(Sender: TObject);
begin
  fmMain.sgTitles.Font.Size := StrToInt(edStSizeTitles.Text);
end;

procedure TfmOptions.edStSizeTitlesLessClick(Sender: TObject);
begin
  if StrToInt(edStSizeTitles.Text) > 7 then
  begin
    edStSizeTitles.Text := IntToStr(StrToInt(edStSizeTitles.Text) - 1);
  end;
end;

procedure TfmOptions.edStSizeTitlesPlusClick(Sender: TObject);
begin
  if StrToInt(edStSizeTitles.Text) < 256 then
  begin
    edStSizeTitles.Text := IntToStr(StrToInt(edStSizeTitles.Text) + 1);
  end;
end;

procedure TfmOptions.edStMaxLessClick(Sender: TObject);
begin
  if StrToInt(edStMaxChar.Text) > 10000 then
  begin
    edStMaxChar.Text := IntToStr(StrToInt(edStMaxChar.Text) - 5000);
  end;
end;

procedure TfmOptions.edStMaxPlusClick(Sender: TObject);
begin
  if StrToInt(edStMaxChar.Text) < 1000000 then
  begin
    edStMaxChar.Text := IntToStr(StrToInt(edStMaxChar.Text) + 5000);
  end;
end;

procedure TfmOptions.bnStFontColorDefClick(Sender: TObject);
begin
  if IsPaintDark = True then
  begin
    fmMain.dbText.Font.Color := clWhite;
  end
  else
  begin
    fmMain.dbText.Font.Color := clDefault;
  end;
  fmMain.sgTitles.Font.Color := clDefault;
  clTitle := clRed;
  clHighlight := clYellow;
  clMark := clGray;
  fmMain.FormatMarkers(2);
end;

procedure TfmOptions.bnStFontColorModClick(Sender: TObject);
begin
  cdColorDialog.Color := fmMain.dbText.Font.Color;
  if cdColorDialog.Execute then
  begin
    fmMain.dbText.Font.Color := cdColorDialog.Color;
    fmMain.sgTitles.Font.Color := cdColorDialog.Color;
    fmMain.FormatMarkers(2);
  end;
end;

procedure TfmOptions.bnStFontMarkModClick(Sender: TObject);
begin
  cdColorDialog.Color := clMark;
  if cdColorDialog.Execute then
  begin
    clMark := cdColorDialog.Color;
    fmMain.FormatMarkers(2);
  end;
end;

procedure TfmOptions.bnStFontTitleColorModClick(Sender: TObject);
begin
  cdColorDialog.Color := clTitle;
  if cdColorDialog.Execute then
  begin
    clTitle := cdColorDialog.Color;
    fmMain.FormatMarkers(2);
    fmMain.sgTitles.Font.Color := clTitle;
  end;
end;

procedure TfmOptions.bnStFontHighColorModClick(Sender: TObject);
begin
  cdColorDialog.Color := clHighlight;
  if cdColorDialog.Execute then
  begin
    clHighlight := cdColorDialog.Color;
    fmMain.FormatMarkers(2);
  end;
end;

procedure TfmOptions.bnCloseClick(Sender: TObject);
begin
  Close;
end;

end.
