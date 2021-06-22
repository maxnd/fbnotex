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

unit unit12;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, CocoaThemes;

type

  { TfmAdmin }

  TfmAdmin = class(TForm)
    bnCancel: TButton;
    bnOK: TButton;
    edSudoPwd: TEdit;
    edSysPwd: TEdit;
    edSysPwd1: TEdit;
    edSysPwd2: TEdit;
    lbSudoPwd: TLabel;
    lbSysPwd: TLabel;
    lbSysPwd1: TLabel;
    procedure bnCancelClick(Sender: TObject);
    procedure bnOKClick(Sender: TObject);
    procedure edSudoPwdKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure edSysPwd1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edSysPwd2KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edSysPwdKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
  private

  public

  end;

var
  fmAdmin: TfmAdmin;

resourcestring

  msgAdmin001 = 'The new passwords do not match.';

implementation

{$R *.lfm}

{ TfmAdmin }

procedure TfmAdmin.FormActivate(Sender: TObject);
begin
  edSudoPwd.SetFocus;
end;

procedure TfmAdmin.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if fmAdmin.edSysPwd1.Text = '' then
  begin
    edSudoPwd.Clear;
    edSysPwd.Clear;
    edSysPwd2.Clear;
    ModalResult := mrCancel;
  end
  else
  if fmAdmin.edSysPwd1.Text <> fmAdmin.edSysPwd2.Text then
  begin
    MessageDlg(msgAdmin001, mtError, [mbOK], 0);
    edSudoPwd.Clear;
    edSysPwd.Clear;
    edSysPwd1.Clear;
    edSysPwd2.Clear;
    ModalResult := mrCancel;
  end
  else
  begin
    ModalResult := mrOk;
  end;
end;

procedure TfmAdmin.bnOKClick(Sender: TObject);
begin
  Close;
end;

procedure TfmAdmin.bnCancelClick(Sender: TObject);
begin
  edSudoPwd.Clear;
  edSysPwd.Clear;
  edSysPwd1.Clear;
  edSysPwd2.Clear;
  ModalResult := mrCancel;
end;

procedure TfmAdmin.edSudoPwdKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    edSysPwd.SetFocus;
  end;
end;

procedure TfmAdmin.edSysPwdKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    edSysPwd1.SetFocus;
  end;
end;

procedure TfmAdmin.edSysPwd1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    edSysPwd2.SetFocus;
  end;
end;

procedure TfmAdmin.edSysPwd2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    Close;
  end;
end;

procedure TfmAdmin.FormCreate(Sender: TObject);
begin
  if IsPaintDark = False then
  begin
    lbSudoPwd.Font.Color := $001ABC0F;
    lbSysPwd.Font.Color := $001ABC0F;
    lbSysPwd1.Font.Color := $001ABC0F;
  end
  else
  begin
    lbSudoPwd.Font.Color := clLime;
    lbSysPwd.Font.Color := clLime;
    lbSysPwd1.Font.Color := clLime;
  end;
end;

procedure TfmAdmin.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  if key = 27 then
  begin
    edSudoPwd.Clear;
    edSysPwd.Clear;
    edSysPwd1.Clear;
    edSysPwd2.Clear;
    key := 0;
    ModalResult := mrCancel;
  end;
end;

end.



