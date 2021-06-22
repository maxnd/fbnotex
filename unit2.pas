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

unit Unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, DBCtrls, DB,
  CocoaThemes;

type

  { TfmNotebooks }

  TfmNotebooks = class(TForm)
    bnCancel: TButton;
    bnOK: TButton;
    dbID: TDBEdit;
    dbTitle: TDBEdit;
    dbComments: TDBMemo;
    procedure bnCancelClick(Sender: TObject);
    procedure bnOKClick(Sender: TObject);
    procedure dbCommentsKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure dbTitleKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
  private

  public

  end;

var
  fmNotebooks: TfmNotebooks;


implementation

uses Unit1;

{$R *.lfm}


{ TfmNotebooks }

procedure TfmNotebooks.FormShow(Sender: TObject);
begin
  dbTitle.SetFocus;
  if IsPaintDark = False then
  begin
    dbComments.Font.Color := clBlack;
    dbID.Font.Color := $001ABC0F;
  end
  else
  begin
    dbComments.Font.Color := clWhite;
    dbID.Font.Color := clLime;
  end;
end;

procedure TfmNotebooks.bnOKClick(Sender: TObject);
begin
  if fmMain.dsNotebooks.State in [dsEdit, dsInsert] then
  begin
    fmMain.zqNotebooks.Post;
  end;
  if fmMain.zqNotebooks.UpdatesPending = True then
  begin
    fmMain.zqNotebooks.ApplyUpdates;
  end;
  Close;
end;

procedure TfmNotebooks.dbCommentsKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  if ((Key = 13) and (Shift = [ssMeta])) then
  begin
    key := 0;
    bnOKClick(nil);
  end;
end;

procedure TfmNotebooks.dbTitleKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    key := 0;
    dbComments.SetFocus;
  end;
end;

procedure TfmNotebooks.bnCancelClick(Sender: TObject);
begin
  if fmMain.dsNotebooks.State in [dsEdit, dsInsert] then
  begin
    fmMain.zqNotebooks.Cancel;
  end;
  if fmMain.zqNotebooks.UpdatesPending = True then
  begin
    fmMain.zqNotebooks.CancelUpdates;
  end;
  Close;
end;

end.



