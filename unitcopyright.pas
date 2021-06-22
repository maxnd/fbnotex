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

unit unitcopyright;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  LCLIntf;

type

  { TfmCopyright }

  TfmCopyright = class(TForm)
    imLogo: TImage;
    imImagecopyright: TImage;
    lbCopyrightWeb: TLabel;
    lbOK: TLabel;
    lbCopyrightDesc: TLabel;
    lbCopyrightAuthor2: TLabel;
    lbCopyrighPowered: TLabel;
    lbCopyrightAuthor3: TLabel;
    lbCopyrightName: TLabel;
    rmCopyrightText: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure imLogoClick(Sender: TObject);
    procedure lbCopyrighPoweredClick(Sender: TObject);
    procedure lbCopyrightWebClick(Sender: TObject);
    procedure lbOKClick(Sender: TObject);
  private

  public

  end;

var
  fmCopyright: TfmCopyright;

implementation

{$R *.lfm}

{ TfmCopyright }

procedure TfmCopyright.FormCreate(Sender: TObject);
begin
  rmCopyrightText.Lines.Add(
     'This program is free software: you can redistribute it and/or modify it ' +
     'under the terms of the GNU General Public License as published by the Free ' +
     'Software Foundation, either version 3 of the License, or (at your option) ' +
     'any later version. You can read the version 3 of the Licence in ' +
     'http://www.gnu.org/licenses gpl-3.0.txt. For other information you can also ' +
     'see http://www.gnu.org/licenses.');
   rmCopyrightText.Lines.Add(
     'This program is distributed in the hope ' +
     'that it will be useful, but WITHOUT ANY WARRANTY; without even the implied ' +
     'warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU ' +
     'General Public License for more details.');
   rmCopyrightText.Lines.Add('This software has been written with Lazarus, and ' +
     'uses the Zeos components to access the Firebird database.');
end;

procedure TfmCopyright.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 27 then
  begin
    key := 0;
    Close;
  end;
end;

procedure TfmCopyright.imLogoClick(Sender: TObject);
begin
  OpenURL('https://www.gnu.org/licenses/gpl-3.0.html');
end;

procedure TfmCopyright.lbCopyrighPoweredClick(Sender: TObject);
begin
  OpenURL('https://firebirdsql.org/');
end;

procedure TfmCopyright.lbCopyrightWebClick(Sender: TObject);
begin
  OpenURL('https://github.com/maxnd/fbnotex/wiki');
end;

procedure TfmCopyright.lbOKClick(Sender: TObject);
begin
  Close;
end;

end.

