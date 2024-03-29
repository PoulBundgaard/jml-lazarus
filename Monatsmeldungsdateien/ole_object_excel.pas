unit ole_object_excel;

{$MODE objfpc}{$H+}

interface

uses
  locale_de, Classes, SysUtils,
  ActiveX,
  ComObj, variants, Windows,
  strutils, Dialogs;

function GetExcel_Kontrollsumme(FileName: string): currency;
function Getcsv_Kontrollsummen(FileName: string; var Preis, SPreis: currency): currency;
function GetLastTextFileLine(FileName: string): string;

implementation

uses unit1;

function GetExcel_Kontrollsumme(FileName: string): currency;
var
  V: variant;
  CloseExcel: boolean;
begin
  try
    Result := -1;
    {* laufende Excel Instanz verwenden, sonst neue Instanz starten *}
    v := GetActiveOLEObject('Excel.Application');
    CloseExcel := False;
    //V.Visible := true;
    //ShowWindow(ExcelApp.Hwnd, SW_ShowMaximized);
  except
    try
      V := CreateOleObject('Excel.Application');
      CloseExcel := True;
      //V.Visible := true;
    except
      raise EOleSysError.CreateFmt('Excel konnte nicht gestartet werden!', []);
    end;
  end;
  {* für OLE ist in Lazarus widestring erforderlich, siehe F1 für widestring *}
  V.Workbooks.Open('' + WideString(FileName) + '');
  try
    Result := V.ActiveWorkbook.Sheets[1].Range['Kontrollsumme'].Value;
  except
    Result := 0;
    ShowMessage(
      'Es gab einen Fehler, wahrscheinlich ist in der Excel Tabelle im ersten Blatt' +
      ' keine Zelle mit ''Kontrollsumme'' als Namen!' + NL +
      'Kontrollsumme erst ab 08/2012');
  end;
  V.ActiveWorkBook.Close;
  if CloseExcel then
    V.Quit;
  V := Unassigned;
end;

function Getcsv_Kontrollsummen(FileName: string; var Preis, SPreis: currency): currency;
var
  line: string;
begin

  (* in Spalte 14, 15 und letzter Zeile der 27-Felder Matrix steht der Preis, SPreis *)
  line := GetLastTextFileLine(FileName);

(* ExtractWord nimmt ;; als Fehler und nicht als leeren
   string, deshalb ExtractDelimited *)

  Preis := StrToCurrDef(ExtractDelimited(14, line, [';']), 0);
  SPreis := StrToCurrDef(ExtractDelimited(15, line, [';']), 0);

  Result := Preis + SPreis;
end;

(* Info zu: function GetLastTextFileLine:

  Siehe: http://www.delphifaq.com/faq/f87.shtml
  To read a text file backwards, you have to open it as a binary file, e.g. with FileOpen().
  The following procedure ReadBack()
  -> procedure ReadBack() wurde von mir in GetLastTextFileLine(FileName: string): string; umgewandelt
  reads one line backwards - up to the current position -.
  So, initially, you need to position to the end of the file.
*)
function GetLastTextFileLine(FileName: string): string;
const
  MAXLINELENGTH = 256;
var
  Line: string;
  f: integer;
  BeginOfFile: boolean;

  curr, Before: longint;
  Buffer: array[0..MAXLINELENGTH] of char;
  p: PChar;
begin
  Result := '';
  try
    f := FileOpen(FileName, 0);
    // move to end of file!
    FileSeek(f, 0, 2);

    // read all lines, backwards
    repeat

      // scan backwards to the last CR-LF
      curr := FileSeek(f, 0, 1);
      Before := curr - MAXLINELENGTH;
      if Before < 0 then
        Before := 0;
      FileSeek(f, Before, 0);
      FileRead(f, Buffer, curr - Before);
      Buffer[curr - Before] := #0;
      p := StrRScan(Buffer, #10);
      if p = nil then
      begin
        Line := StrPas(Buffer);
        FileSeek(f, 0, 0);
        BeginOfFile := True;
      end
      else
      begin
        Line := StrPas(p + 1);
        FileSeek(f, Before + longint(p) - longint(@Buffer), 0);
        BeginOfFile := False;
      end;

      // this will also work with Unix files (#10 only, no #13)
      if length(Line) > 0 then
        if Line[length(Line)] = #13 then
        begin
          SetLength(Line, length(Line) - 1);
        end;

    until BeginOfFile
      (* von mir ergänzt, ich will ja nur die letzte Zeile: *) or (trim(line) > '');


    Result := line;

  finally
    FileClose(f);
  end;
end;

end.
