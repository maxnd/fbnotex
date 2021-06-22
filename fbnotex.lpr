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

program fbnotex;

{$mode objfpc}{$H+}

uses
  Interfaces, // this includes the LCL widgetset
  Forms, zcomponent, Unit1, Unit2, Unit3, Unit4, Unit5,
  Unit6, Unit7, Unit8, Unit9, Unit10, Unit11, Unit12, unitcopyright,
  {$IFDEF Darwin}
    clocale, iosxlocale;
  {$ENDIF}

{$R *.res}

begin
  Application.Title:='fbNotex';
  Application.Scaled:=True;
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TfmNotebooks, fmNotebooks);
  Application.CreateForm(TfmSections, fmSections);
  Application.CreateForm(TfmCopyright, fmCopyright);
  Application.CreateForm(TfmOptions, fmOptions);
  Application.CreateForm(TfmRenameTags, fmRenameTags);
  Application.CreateForm(TfmInsertID, fmInsertID);
  Application.CreateForm(TfmShowAllTasks, fmShowAllTasks);
  Application.CreateForm(TfmSearch, fmSearch);
  Application.CreateForm(TfmBookmarks, fmBookmarks);
  Application.CreateForm(TfmAdmin, fmAdmin);
  Application.CreateForm(TfmPassword, fmPassword);
  Application.CreateForm(TfmPicture, fmPicture);
  Application.Run;
end.

