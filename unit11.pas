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

unit Unit11;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls;

type

  { TfmPicture }

  TfmPicture = class(TForm)
    imPicture: TImage;
    shPicture: TShape;
    tmPicture: TTimer;
    procedure FormHide(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure tmPictureTimer(Sender: TObject);
  private

  public

  end;

var
  fmPicture: TfmPicture;

implementation

uses Unit1;

{$R *.lfm}

{ TfmPicture }

procedure TfmPicture.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 27 then
  begin
    Close;
  end;
end;

procedure TfmPicture.FormHide(Sender: TObject);
begin
  imPicture.Picture.Clear;
  tmPicture.Enabled := False;
end;

procedure TfmPicture.FormShow(Sender: TObject);
begin
  tmPicture.Enabled := True;
end;

procedure TfmPicture.tmPictureTimer(Sender: TObject);
begin
  tmPicture.Enabled := False;
  Close;
  if ((fmMain.dbText.Visible = True) and
    (fmMain.pcPageControl.ActivePageIndex = 0)) then
  begin
    fmMain.dbText.SetFocus;
  end;
end;

end.

