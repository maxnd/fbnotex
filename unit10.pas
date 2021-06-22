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

unit Unit10;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, CocoaThemes;

type

  { TfmPassword }

  TfmPassword = class(TForm)
    bnCancel: TButton;
    bnOK: TButton;
    edSudoPwd: TEdit;
    edSysPwd: TEdit;
    lbSudoPwd: TLabel;
    lbSysPwd: TLabel;
    procedure edSudoPwdKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure edSysPwdKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
  private

  public

  end;

var
  fmPassword: TfmPassword;

implementation

{$R *.lfm}

{ TfmPassword }

procedure TfmPassword.edSudoPwdKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    edSysPwd.SetFocus;
  end;
end;

procedure TfmPassword.edSysPwdKeyDown(Sender: TObject; var Key: word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    ModalResult := mrOk;
  end;
end;

procedure TfmPassword.FormActivate(Sender: TObject);
begin
  edSudoPwd.SetFocus;
end;

procedure TfmPassword.FormCreate(Sender: TObject);
begin
  if IsPaintDark = False then
  begin
    lbSudoPwd.Font.Color := $001ABC0F;
    lbSysPwd.Font.Color := $001ABC0F;
  end
  else
  begin
    lbSudoPwd.Font.Color := clLime;
    lbSysPwd.Font.Color := clLime;
  end;
end;

procedure TfmPassword.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  if key = 27 then
  begin
    ModalResult := mrCancel;
  end;
end;

end.



