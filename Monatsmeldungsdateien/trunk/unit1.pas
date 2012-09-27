unit Unit1;

{$mode objfpc}{$H+}

interface

uses
 locale_de, Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs,
 IniPropStorage, ExtCtrls, EditBtn, ComCtrls, Grids, StdCtrls, LCLType, Buttons,
 AsyncProcess, FileCtrl, StrUtils, PropertyStorage, Menus, FindFile1 (*,ActiveX,
  ComObj, variants, windows *);

type

  { TForm1 }

  TForm1 = class(TForm)
    AsyncProcess1: TAsyncProcess;
    BtnKontrollsumme: TBitBtn;
    BtnOpenWith: TBitBtn;
    BtnSearch: TButton;
    BtnMail: TButton;
    ComboBox1: TComboBox;
    DirectoryEdit1: TDirectoryEdit;
    FileSearch1: TFindFile;
    ImageList1: TImageList;
    IniPropStorage1: TIniPropStorage;
    EditFilter: TLabeledEdit;
    Label1: TLabel;
    Label2: TLabel;
    lbKontrollsumme: TLabel;
    ListBoxFilters: TListBox;
    MenuItem1: TMenuItem;
    OpenDialog1: TOpenDialog;
    PageControl1: TPageControl;
    Panel1: TPanel;
    Panel2: TPanel;
    PopupGrid: TPopupMenu;
    RadioGroup1: TRadioGroup;
    StatusBar1: TStatusBar;
    PageSettings: TTabSheet;
    PageFiles: TTabSheet;
    Grid: TStringGrid;
    procedure BtnKontrollsummeClick(Sender: TObject);
    procedure BtnMailClick(Sender: TObject);
    procedure BtnOpenWithClick(Sender: TObject);
    procedure BtnSearchClick(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox1DblClick(Sender: TObject);
    procedure EditFilterKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure GridClick(Sender: TObject);
    procedure GridDblClick(Sender: TObject);
    procedure IniPropStorage1RestoreProperties(Sender: TObject);
    procedure IniPropStorage1StoredValues4Restore(Sender: TStoredValue;
      var Value: TStoredType);
    procedure IniPropStorage1StoredValues4Save(Sender: TStoredValue;
      var Value: TStoredType);
    procedure ListBoxFiltersClick(Sender: TObject);
    procedure ListBoxFiltersDblClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
    function VonBisDatum(EingabeDatumStr: string;
      var MonatsErster, MonatsLetzter: TDateTime): boolean;
    function nth_WeekDayOfMonth(JDatum : TDateTime;IndexWochenTag,nth_Vorkommen : integer): TDateTime;
    function DatumsFilter(Eingabe: string): string;
    function EnableControls(TrueFalse: boolean): boolean;
  end;



var
  Form1: TForm1;
  ndx: integer = -1;
  OldValue: string;


implementation

uses ole_object_excel;

{$R *.lfm}

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
var
  x, mon: integer;
begin
  (* aktuellen Monat als Zahl ermitteln *)
  mon := StrToInt(FormatDateTime('MM', date()));
  (* vorhergehende Monate in Listbox eintragen *)
  for x := 1 to mon do
  begin
    ComboBox1.Items.add(AnsiToUTF8(
      FormatDateTime('mmmm yyyy', incMonth(Date, -1 * x))));
  end;
  (* Vormonat auswählen *)
  ComboBox1.ItemIndex := 0;

  PageControl1.ActivePage := PageFiles;

end;

procedure TForm1.BtnSearchClick(Sender: TObject);
var
  x: integer;
  //FileSearch1: TFindFile;
  erster, letzter: TDateTime;
begin
  try


    //FileSearch1 := TFindFile.Create(Self);
    Grid.RowCount := 1;
    //grid.Clean(0,1,0,Grid.RowCount -1,[gzNormal]) ;
    screen.Cursor := crHourGlass;
    application.ProcessMessages;

    if not DirectoryExists(DirectoryEdit1.Directory) then
      exit;

    FileSearch1.Directory := DirectoryEdit1.Directory;

    (* Monatsersten und Monatsletzen ermitteln *)
    if not VonBisDatum(ComboBox1.Text, erster, letzter) then
    begin
      ShowMessage('Aus ' + ComboBox1.Text +
        ' konnte kein gültiges Datum ermittelt werden');
      exit;
    end;

   (* Listbox mit den Dateifiltern durchlaufen und den darin enthaltennen Datumsfilter [D, M, Y]
      mit gewähltem Monat ausfüllen *)
    for x := 0 to ListBoxFilters.Count - 1 do
    begin
      (* Filterausdruck Filter1; Filter2; .... zusammenstellen *)
      if x = 0 then
      begin
        (* gibts einen Filterausdruck mit Datum ? -> [] *)
        if ((pos('[', ListBoxFilters.items[x]) > 0) and
          (pos(']', ListBoxFilters.items[x]) > 0)) then
          FileSearch1.Filter := DatumsFilter(ListBoxFilters.items[x])
        else
          FileSearch1.Filter := ListBoxFilters.items[x];
      end
      else
      begin
        if ((pos('[', ListBoxFilters.items[x]) > 0) and
          (pos(']', ListBoxFilters.items[x]) > 0)) then
          FileSearch1.Filter :=
            FileSearch1.Filter + ';' + DatumsFilter(ListBoxFilters.items[x])
        else
          FileSearch1.Filter := FileSearch1.Filter + ';' + ListBoxFilters.items[x];
      end;
    end;
    //ShowMessage(FileSearch1.Filter);

    (* jetzt die Dateien suchen *)
    FileSearch1.Execute;

    (* gefundene Datein ins Grid eintragen *)
    for x := 0 to FileSearch1.Files.Count - 1 do
    begin
      (* neue Zeile ins Grid *)
      Grid.RowCount := Grid.RowCount + 1;
      (* Datei eintragen *)
      grid.Cells[0, x + 1] := FileSearch1.Files[x];
    end;




  finally
    screen.Cursor := crDefault;
    application.ProcessMessages;
    //FreeAndNil(FileSearch1);
    StatusBar1.SimpleText := IntToStr(Grid.RowCount - 1) + ' Dateien gefunden';
    Grid.SortColRow(True, 0);
    if (Grid.RowCount - 1) > 0 then
      EnableControls(True)
    else
      EnableControls(False);
  end;

end;

procedure TForm1.BtnOpenWithClick(Sender: TObject);
var
  FName: string;
  Indx: integer;
begin

  Indx := RadioGroup1.ItemIndex;

  FName := Grid.Cells[0, Grid.Row];

  if not FileExists(FName) then
  begin
    ShowMessage('Die Datei ' + FName + ' existiert nicht');
    exit;
  end;

  if (Indx < 4 (* 4=Excel *)) then
    if not FileExists(IniPropStorage1.StoredValues[Indx].Value) then
    begin
      OpenDialog1.Title := RadioGroup1.Items[Indx] + ' suchen';
      if OpenDialog1.Execute then
      begin
        IniPropStorage1.StoredValues[Indx].Value := OpenDialog1.FileName;
      end
      else
      exit;
    end;

  case RadioGroup1.ItemIndex of
    4:
    begin
      (* mit Excel öffnen *)
      {$IFDEF MSWINDOWS}
      AsyncProcess1.CommandLine :=
        'rundll32.exe url.dll,FileProtocolHandler "' + FName + '"';
      AsyncProcess1.Active := True;
     {$ENDIF}
    end;
    5:
    begin
      (* mit Windows Explorer öffnen *)
      {$IFDEF MSWINDOWS}
      AsyncProcess1.CommandLine := 'Explorer.exe /n,/e,/select,"' + FName + '"';
      AsyncProcess1.Active := True;
     {$ENDIF}
    end
    else
    begin
      (* mit eingestellter Anwendung öffnen *)
       {$IFDEF MSWINDOWS}
      AsyncProcess1.CommandLine :=
        IniPropStorage1.StoredValues[Indx].Value + ' "' + FName + '"';
      AsyncProcess1.Active := True;
      {$ENDIF}

    end;
  end; (* case *)

end;

procedure TForm1.BtnMailClick(Sender: TObject);
var
  FName, f: string;
begin

  if FileExists(IniPropStorage1.StoredValue['ExcelMail']) then
  begin
    FName := IniPropStorage1.StoredValue['ExcelMail'];
    (* mit Excel öffnen *)
      {$IFDEF MSWINDOWS}
    AsyncProcess1.CommandLine :=
      'rundll32.exe url.dll,FileProtocolHandler "' + FName + '"';
    AsyncProcess1.Active := True;
     {$ENDIF}
  end
  else
  begin

    OpenDialog1.Title := 'Mailliste_Monatsmeldung.xls suchen';
    f := OpenDialog1.Filter;
    OpenDialog1.Filter := 'Excel ( *.xls)|*.xls';
    if OpenDialog1.Execute then
    begin
      IniPropStorage1.StoredValue['ExcelMail'] := OpenDialog1.FileName;
    end
    else
    exit;

    OpenDialog1.Filter := f;

    BtnMailClick(Sender);
  end;

end;


procedure TForm1.BtnKontrollsummeClick(Sender: TObject);
var Preis, SPreis, Summe : currency;
begin
  (* für die Excel Datei den Wert der Zelle mit dem Namen Kontrollsumme und für die
     27-Felder Matrix den Wert der Spalten Preis und SPreis anzeigen*)
  try
    //CurrencyString := AnsiToUTF8(CurrencyString);
    jei;
    lbKontrollsumme.Caption := '';
  if ExtractFileExt(Grid.Cells[0, Grid.Row]) = '.xls' then
    lbKontrollsumme.Caption :='Kontrollsumme in ''' + ExtractFileName(Grid.Cells[0, Grid.Row]) +
     ''' ist: ' + FormatCurr('#,##0.00 '+ AnsiToUTF8(CurrencyString), GetExcel_Kontrollsumme(Grid.Cells[0, Grid.Row]))
  else if ExtractFileExt(Grid.Cells[0, Grid.Row]) = '.csv' then
  begin
    Summe := Getcsv_Kontrollsummen(Grid.Cells[0, Grid.Row],Preis,SPreis);
    lbKontrollsumme.Caption := 'Kontrollsummen in ''' + ExtractFileName(Grid.Cells[0, Grid.Row]) +
    ''' sind: ' + FormatCurr(' Preis = #,##0.00 '+ AnsiToUTF8(CurrencyString), Preis) +
     FormatCurr(' SPreis = #,##0.00 '+ AnsiToUTF8(CurrencyString), SPreis) +
     FormatCurr(' Summe = #,##0.00 '+ AnsiToUTF8(CurrencyString), Summe)
  end
  else
   ShowMessage('Für die Datei ' + Grid.Cells[0, Grid.Row] + ' gibt es keine Kontrollsumme');
  finally
    Nei;
  end;
end;


procedure TForm1.ComboBox1Change(Sender: TObject);
begin
  BtnSearchClick(Sender);
end;

procedure TForm1.ComboBox1DblClick(Sender: TObject);
var
  erster, letzter: TDateTime;
begin
  (* nur so als Test, hat für die Funktion des Programms keine Bedeutung *)
  if VonBisDatum(ComboBox1.Text, erster, letzter) then
    ShowMessage('erster: ' + DateTimeToStr(erster) + NL + 'letzter: ' +
      DateTimeToStr(letzter));
end;

procedure TForm1.EditFilterKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
var
  NewValue: string;
begin
    (* nur in die Liste der Filter übernehmen, wenns den Filter noch nicht gibt
       bzw editierten Eintrag auf Wunsch ersetzen*)
  if (Key = VK_RETURN) then
  begin

    NewValue := trim(EditFilter.Text);

    if NewValue = '' then
      exit;

    if ((ndx = ListBoxFilters.ItemIndex) and (OldValue <> NewValue) and (ndx > -1)) then
    begin
      if Messagedlg('Soll der in der Liste markierte Filterausdruck ersetzt werden?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        ListBoxFilters.Items[ndx] := NewValue;
    end;

    if ListBoxFilters.Items.IndexOf(NewValue) = -1 then
        ListBoxFilters.Items.Add(trim(NewValue));
  end;

end;

procedure TForm1.GridClick(Sender: TObject);
begin
  if ExtractFileExt(Grid.Cells[0, Grid.Row]) = '.xls' then
    RadioGroup1.ItemIndex := 4;
end;

procedure TForm1.GridDblClick(Sender: TObject);
begin
  (* Bei Doppelklick Excel oder RMV_csv_viewer starten *)
  if ExtractFileExt(Grid.Cells[0, Grid.Row]) = '.xls' then
    RadioGroup1.ItemIndex := 4
  else
    RadioGroup1.ItemIndex := 0;

  RadioGroup1Click(Sender);

  BtnOpenWithClick(Sender);
end;

procedure TForm1.IniPropStorage1RestoreProperties(Sender: TObject);
begin
  (* Suche für den Vormonat starten *)
  BtnSearchClick(Sender);
  (* damit das Glyph von BtnOpenWith geladen wird *)
  RadioGroup1Click(Sender);

end;

procedure TForm1.IniPropStorage1StoredValues4Restore(Sender: TStoredValue;
  var Value: TStoredType);
begin
  (* Spaltenbreite des Grid  *)
  if StrToIntDef(Value, 0) > 0 then
    Grid.Columns[0].Width := StrToIntDef(Value, 0);

end;

procedure TForm1.IniPropStorage1StoredValues4Save(Sender: TStoredValue;
  var Value: TStoredType);
begin
  Value := IntToStr(Grid.Columns[0].Width);
end;

procedure TForm1.ListBoxFiltersClick(Sender: TObject);
begin
  OldValue := trim(ListBoxFilters.Items[ListBoxFilters.ItemIndex]);
  if OldValue <> '' then
  begin
    ndx := ListBoxFilters.ItemIndex;
    EditFilter.Text := OldValue;
  end;
end;

procedure TForm1.ListBoxFiltersDblClick(Sender: TObject);
begin
  (* Eintrag löschen *)
  ListBoxFilters.Items.Delete(ListBoxFilters.ItemIndex);
end;

procedure TForm1.MenuItem1Click(Sender: TObject);
begin
  try
  Jei;
  EditFilter.Text:=Grid.Cells[0, Grid.Row];
  EditFilter.SelStart :=0;
  EditFilter.SelLength := Length(EditFilter.Text);
  EditFilter.SelectAll;
  EditFilter.CopyToClipboard;
  //ShowMessage('Der Dateiname:' + NL + EditFilter.Text + NL +'wurde kopiert und kann z.B. im konverter Datei öffnen Dialog eingefügt werden.');
  EditFilter.Clear;
  finally
    Nei;
  end;
end;

procedure TForm1.RadioGroup1Click(Sender: TObject);
var
  bmp: TBitmap;
begin
  try

    bmp := TBitmap.Create;
    ImageList1.GetBitmap(RadioGroup1.ItemIndex, bmp);
    BtnOpenWith.Glyph.Assign(bmp);

    BtnOpenWith.Caption := 'öffnen mit ' + RadioGroup1.Items[RadioGroup1.ItemIndex];

    (* SetFocus, auch wenn beim Start ein Fehler auftritt *)
    try
      BtnOpenWith.SetFocus ;
    except
      on EInvalidOperation do
    end;

    {* Icon des Programms extrahieren, aus Delphi *}
    //icon.Handle := ExtractIconFromFile(ProgList.Cells[ColApp, x], 0);
  finally
    FreeAndNil(bmp);
  end;

end;

function TForm1.VonBisDatum(EingabeDatumStr: string;
  var MonatsErster, MonatsLetzter: TDateTime): boolean;
var
  x: integer;
begin
  Result := False;

  (*  aus "Januar 2011" den Index 1 des Monats ermitteln *)
  for x := 1 to 12 do
  begin
    (* der wievielte Eintrag in LongMonthNames entspricht dem gewählten Monat?  *)
    Result := trim(copy(EingabeDatumStr, 1, Length(EingabeDatumStr) - 4)) =
      AnsiToUTF8(LongMonthNames[x]);
    if Result then
      break;
  end;

  try
    (* Monatsersten und Monatsletzten als DateTime *)
    MonatsErster := StrToDateTime('1.' + IntToStr(x) + '.' +
      copy(EingabeDatumStr, Length(EingabeDatumStr) - 4, Length(EingabeDatumStr)));
    MonatsLetzter := IncMonth(MonatsErster, 1) - 1
  except
    on EConvertError do
      Result := False;
  end;
end;

function TForm1.nth_WeekDayOfMonth(JDatum: TDateTime; IndexWochenTag,
  nth_Vorkommen: integer): TDateTime;
  var VonDatum, BisDatum : TDateTime;
      x, z : integer;

begin

  (* Test
  JDatum := StrToDateTime('01.12.2011');
  IndexWochenTag := 1; (* Sonntag = 1 *)
  nth_Vorkommen := 2;
   *)

  (* Start und Ende der Prüfung ermitteln *)
  VonBisDatum(FormatDateTime('mmmm yyyy',JDatum), VonDatum, BisDatum);

  (* Zähler für das Vorkommen des Wochentages *)
  z := 0;
  Result :=  VonDatum;
  for x := 1 to trunc(BisDatum) - trunc(VonDatum) do
  begin

    (* Wochentag gefunden? *)
    if DayOfWeek(Result) =  IndexWochenTag then inc(z);
    if z = nth_Vorkommen then break;
    (* nächster Tag im Monat *)
    Result := Result + 1;
  end;

  if z < nth_Vorkommen then
  begin
     Result := 0;
     raise EConvertError.Create('Im Monat ' + FormatDateTime('mmmm yyyy',JDatum) + ' gibt es kein ' +
     IntToStr(nth_Vorkommen) + '.tes Vorkommen eines ' + LongDayNames[IndexWochenTag]);
  end;


end;

function TForm1.DatumsFilter(Eingabe: string): string;
var
  Filter, von, bis: string;
  x, vonpos, bispos: integer;
  erster, letzter: TDateTime;
begin
  (* Monatsersten und Monatsletzten als DateTime *)
  VonBisDatum(ComboBox1.Text, erster, letzter);
  (*  an welcher Position steht der Datumsfilterausdruck [] *)
  vonpos := pos('[', Eingabe);
  bispos := pos(']', Eingabe);
  Filter := copy(Eingabe, vonpos + 1, bispos - vonpos - 1);

  (* gibt es einen Datumsbreich [*-*] *)
  x := pos('-', Filter);
  if x > 0 then
  begin
    (* Datumsfilter als string *)
    von := FormatDateTime(copy(Filter, 1, x - 1), erster);
    bis := FormatDateTime(copy(Filter, x + 1, 20), letzter);
    (* In Eingabe den gewählten Monat eintragen *)
    Result := StringReplace(Eingabe, '[' + Filter + ']', von + '-' +
      bis, [rfReplaceAll]);
  end
  else
  begin
    (* Datumsfilter als string *)
    von := FormatDateTime(Filter, erster);
    (* In Eingabe den gewählten Monat eintragen *)
    Result := StringReplace(Eingabe, '[' + Filter + ']', von, [rfReplaceAll]);
  end;

end;

function TForm1.EnableControls(TrueFalse: boolean): boolean;
begin
  BtnOpenWith.Enabled := TrueFalse;
  RadioGroup1.Enabled := TrueFalse;
end;

end.

