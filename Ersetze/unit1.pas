unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  locale_de, SysUtils, SdfData, DB, Forms, Controls, Dialogs, ExtCtrls, EditBtn,
  StdCtrls, DBCtrls, DBGrids,
  IniPropStorage, ComCtrls, Menus, Grids, StrUtils, Classes;

type

  { TForm1 }

  TForm1 = class(TForm)
    BtnExecute: TButton;
    BtnSaveAs: TButton;
    cbFeldNamen: TCheckBox;
    Datasource1: TDatasource;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    EditSearch: TEdit;
    EditReplace: TEdit;
    EditDelimiter: TEdit;
    FileNameEdit1: TFileNameEdit;
    IniPropStorage1: TIniPropStorage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    lbRecordCount: TLabel;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    PopupDBGrid: TPopupMenu;
    ProgressBar1: TProgressBar;
    SaveDialog1: TSaveDialog;
    SdfDataSet1: TSdfDataSet;
    StatusBar1: TStatusBar;
    procedure BtnExecuteClick(Sender: TObject);
    procedure BtnSaveAsClick(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure DBGrid1ColEnter(Sender: TObject);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure FileNameEdit1AcceptFileName(Sender: TObject; var Value: string);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure IniPropStorage1RestoreProperties(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure MenuItem5Click(Sender: TObject);
    procedure MenuItem6Click(Sender: TObject);
    procedure PopupDBGridPopup(Sender: TObject);
    procedure SdfDataSet1AfterPost(DataSet: TDataSet);
    procedure SdfDataSet1UpdateInfo(DataSet: TDataSet);
    procedure SdfDataSet1FilterRecord(DataSet: TDataSet; var Accept: boolean);
  private
    { private declarations }
  public
    procedure AShowExample(Sender: TObject);
    procedure UpDateInfo;
    { public declarations }
  end;


type
  (* Siehe das dynamische Array Filters weiter unten *)
  TFilterRecord = record
    Filter: string;
    FieldIndex: integer;
  end;

var
  Form1: TForm1;
  ExePath: string;
  BackupFileName, DataSetFilter: string;
  SelIndx: longint = 0;
  (* dynamisches Array of records in dem für mehrere Spalten individuelle Filterkriterien gespeichert werden *)
  Filters: array of TFilterRecord;
  FirstRun: boolean = True;




const
  NL = Chr(10) + Chr(13);

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.FileNameEdit1AcceptFileName(Sender: TObject; var Value: string);
var
  x: integer;
begin

  if SdfDataSet1.Filtered then
  begin
    SdfDataSet1.Filtered := False;
    UpdateInfo;
  end;

  if SdfDataSet1.Active then
    SdfDataSet1.Close;

  SdfDataSet1.FieldDefs.Clear;
  SdfDataSet1.Delimiter := EditDelimiter.Text[1];
  SdfDataSet1.FileName := Value;
  SdfDataSet1.FirstLineAsSchema := cbFeldNamen.Checked;
  SdfDataSet1.Active := True;
  BtnExecute.Enabled := True;
  UpDateInfo;


  //ShowMessage(IntToStr(High(Filters)));

  (* Vorbereitung für die Sortierreihenfolge des StringGrid *)
  for x := 0 to DBGrid1.Columns.Count - 1 do
  begin
    DBGrid1.Columns[x].Tag := 0;
  end;

  if FirstRun then
  begin
    FirstRun := False;
    ShowMessage('Achtung, Änderung der Daten werden ohne Nachfrage in die gerade geöffnete Datei geschrieben!'
      + NL + 'Evtl. sollten Sie erst eine Sicherungskopie erstellen.');

  end;
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if SdfDataSet1.Modified then
  begin
    if Messagedlg(
      'Die Daten wurden geändert, wollen Sie Ihre Änderungen über die Schaltfläche in eine Datei schreiben?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      CanClose := False;
      BtnSaveAsClick(Sender);
      CanClose := True;
    end
    else
    begin
      //SdfDataSet1.Modified := False;
      SdfDataSet1.Close;
      CanClose := True;
    end;
  end;
end;

procedure TForm1.BtnExecuteClick(Sender: TObject);
var
  x, y, RecNr, r: integer;
  BM: TBookmark;
begin
  try
    y := DBGrid1.SelectedIndex;

    EditSearch.Text := trim(EditSearch.Text);
    EditReplace.Text := trim(EditReplace.Text);

    if Messagedlg('Ersetze:   ''' + EditSearch.Text + '''' + NL +
      'durch:      ''' + EditReplace.Text + '''' + NL + 'In Spalte: ''' +
      DBGrid1.SelectedField.FieldName + '''' + NL + NL +
      'Soll die Ersetzung für alle angezeigten Zeilen ausgeführt werden?',
      mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      exit;

    BM := SdfDataSet1.GetBookmark;
    Screen.Cursor := crHourglass;
    Application.ProcessMessages;



    SdfDataSet1.First;
    (* Zähler für Ersetzungen *)
    r := 0;
    (* aktuellen Datensatz merken *)
    RecNr := SdfDataSet1.RecNo;
    ProgressBar1.Position := 0;
    x := 0;
    ProgressBar1.Max := SdfDataSet1.RecordCount;

    SdfDataSet1.First;
    SdfDataSet1.DisableControls;

    while not SdfDataSet1.EOF do
    begin
      Application.ProcessMessages;
      if SdfDataSet1.Fields[y].AsString = EditSearch.Text then
      begin
        SdfDataSet1.Edit;
        SdfDataSet1.Fields[y].AsString :=
          AnsiReplaceText(SdfDataSet1.Fields[y].AsString, EditSearch.Text,
          EditReplace.Text);
        SdfDataSet1.Post;
        Inc(r);
      end;
      Application.ProcessMessages;
      (* ein Filter macht ja den editierten aktuellen Datensatz ggf. unsichtbar,
         deshalb entfällt SdfDataSet1.Next, wenn der aktuelle Datensatz unsichtbar wurde *)
      if RecNr = SdfDataSet1.RecNo then
      begin
        SdfDataSet1.Next;
        RecNr := SdfDataSet1.RecNo;
      end;
      Inc(x);
      ProgressBar1.Position := x;
    end;


    try
      SdfDataSet1.GotoBookmark(BM);
    except
      on E: EDataBaseError do
      begin
        beep;
        if SdfDataSet1.Filtered then
        begin
          SdfDataSet1.Filtered := False;
          SetLength(Filters, 0);
          lbRecordCount.Caption := '';
          SdfDataSet1.GotoBookmark(BM);
          Screen.Cursor := crDefault;
          ShowMessage('Der Filter wurde aufgehoben, um die Änderung anzuzeigen.');
        end;
      end;
    end;

  finally
    Screen.Cursor := crDefault;
    Application.ProcessMessages;
    SdfDataSet1.EnableControls;
    ProgressBar1.Position := 0;
    if SdfDataSet1.BookmarkValid(BM) then
      SdfDataSet1.FreeBookmark(BM);
    UpDateInfo;
    ShowMessage('In ' + IntToStr(r) + ' Fällen wurde ''' +
      EditSearch.Text + ''' durch ''' + EditReplace.Text + ''' ersetzt!');
  end;

end;

procedure TForm1.BtnSaveAsClick(Sender: TObject);
begin
  SaveDialog1.Filter := FileNameEdit1.Filter;
  SaveDialog1.FilterIndex := FileNameEdit1.FilterIndex;

  SaveDialog1.InitialDir := ExtractFilePath(FileNameEdit1.FileName);
  BackupFileName := ChangeFileExt(FileNameEdit1.FileName, '-' +
    FormatDateTime('dd.mm.yy-hh-nn-ss', Now())) +
    ExtractFileExt(FileNameEdit1.FileName);
  SaveDialog1.FileName := BackupFileName;
  if SaveDialog1.Execute then
  begin
    SdfDataSet1.SaveFileAs(SaveDialog1.FileName);
    //ShowMessage('Die Änderungen der Daten werden gespeichert, eine Sicherung finden Sie unter: ' + NL + NL + '''' + BackupFileName + '''');
  end
  else
  begin
      (* SdfDataSet1 daran hindern das Original zu überschreiben,
         auch wenn im SaveDialog Abbrechen gewählt wurde *)
    SdfDataSet1.SaveFileAs(BackupFileName);
    DeleteFile(BackupFileName);
  end;

end;

procedure TForm1.DBGrid1CellClick(Column: TColumn);
begin
  AShowExample(Column as TObject);
end;

procedure TForm1.DBGrid1ColEnter(Sender: TObject);
begin
  AShowExample(Sender);
end;

procedure TForm1.DBGrid1TitleClick(Column: TColumn);
var
  Data: PChar;
  col, row: integer;
  RestoreFilter: boolean;
  StringGrid1: TStringGrid;
begin
  try
    (* SdfDataSet1 hat keine Sortierfunction, deshalb die Daten in ein TStringGrid kopieren,
       dort sortieren und dann wieder zurück kopieren *)
    StringGrid1 := TStringGrid.Create(Application.MainForm);
    StringGrid1.FixedCols := 0;
    Screen.Cursor := crHourGlass;
    RestoreFilter := SdfDataSet1.Filtered;
    SdfDataSet1.Filtered := False;
    Application.ProcessMessages;


    SdfDataSet1.DisableControls;


    (* Sortierreihenfolge pro column speichern/umkehren *)
    if Column.Tag = 0 then
    begin
      Column.Tag := 1;
    end
    else
    begin
      Column.Tag := 0;
    end;

    (* SortOrder für StringGrid übernehmen *)
    if (Column.Tag = 0) then
      StringGrid1.SortOrder := soDescending
    else
      StringGrid1.SortOrder := soAscending;

    StringGrid1.Clear;
    StringGrid1.RowCount := SdfDataSet1.RecordCount;
    StringGrid1.ColCount := SdfDataSet1.FieldCount;
    SdfDataSet1.First;

    (* SdfDataSet1.RecordCount = StringGrid1.RowCount - StringGrid1.FixedRows, deshalb   *)
    if SdfDataSet1.FirstLineAsSchema then
      row := 1
    else
      row := 0;

    while not SdfDataSet1.EOF do
    begin
      for col := 0 to SdfDataSet1.FieldCount - 1 do
        StringGrid1.Cells[col, row] := SdfDataSet1.Fields[col].AsString;

      Inc(row);
      SdfDataSet1.Next;
    end;

    StringGrid1.SortColRow(True, Column.Index);

    SdfDataSet1.First;

    if SdfDataSet1.FirstLineAsSchema then
      row := 1
    else
      row := 0;

    while not SdfDataSet1.EOF do
    begin
      SdfDataSet1.Edit;

      for col := 0 to SdfDataSet1.FieldCount - 1 do
        SdfDataSet1.Fields[col].AsString := StringGrid1.Cells[col, row];

      SdfDataSet1.Post;

      Inc(row);
      SdfDataSet1.Next;
    end;




  finally
    Screen.Cursor := crDefault;
    SdfDataSet1.EnableControls;
    SdfDataSet1.Filtered := RestoreFilter;
    SdfDataSet1.Refresh;
    SdfDataSet1.First;
    FreeAndNil(StringGrid1);
    Application.ProcessMessages;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  ExePath := ExtractFilePath(Application.ExeName);
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  if not FileExists(FileNameEdit1.FileName) then
    FileNameEdit1.InitialDir := ExePath;

end;

procedure TForm1.IniPropStorage1RestoreProperties(Sender: TObject);
var
  p: integer;
begin
  (* den Dateiöffnen Filterdialog passend zur Dateierweiterung vorgeben *)
  try
    if trim(FileNameEdit1.FileName) <> '' then
    begin
      for p := 0 to WordCount(FileNameEdit1.Filter, ['|']) - 1 do
      begin
        if ExtractWord(p, FileNameEdit1.Filter, ['|']) = '*' +
          ExtractFileExt(FileNameEdit1.FileName) then
        begin
          FileNameEdit1.FilterIndex := p - 2;
          break;
        end;
      end;
    end;

  finally
  end;
end;

procedure TForm1.MenuItem1Click(Sender: TObject);
var
  x, y: integer;
  ndx: integer;
begin
  (* Filterkriterium für die aktuelle Spalte im dynamischen Array of records speichern *)
  ndx := -1;

  SelIndx := DBGrid1.SelectedIndex;
  (* das Filterkriterium *)
  DataSetFilter := SdfDataSet1.Fields[SelIndx].AsString;

  (* wurde für das aktuelle Field bereits ein Filter gespeichert? *)
  for x := low(Filters) to High(Filters) do
  begin
    if Filters[x].FieldIndex = SelIndx then
    begin
      ndx := SelIndx;
      break;
    end;
  end;

  Application.ProcessMessages;



  (* es gab also für dieses Feld noch kein Filterkriterium im array, also wird das array erweitert
     und das Kriterium und der FeldIndex im Array gespeichert*)
  if ndx = -1 then
  begin
    SetLength(Filters, length(Filters) + 1);
    Filters[length(Filters) - 1].Filter := DataSetFilter;
    Filters[length(Filters) - 1].FieldIndex := SelIndx;

    if trim(lbRecordCount.Caption) = '' then
      lbRecordCount.Caption :=
        'Filter: ' + SdfDataSet1.Fields[Filters[length(Filters) -
        1].FieldIndex].Fieldname + ' = ' + Filters[length(Filters) - 1].Filter
    else
      lbRecordCount.Caption :=
        lbRecordCount.Caption + ' UND ' +
        SdfDataSet1.Fields[Filters[length(Filters) - 1].FieldIndex].Fieldname +
        ' = ' + Filters[length(Filters) - 1].Filter;

  end;


  if not SdfDataSet1.Filtered then
    SdfDataSet1.Filtered := True;

  SdfDataSet1.Refresh;
  UpDateInfo;

end;

procedure TForm1.MenuItem2Click(Sender: TObject);
begin
  SdfDataSet1.Filtered := False;
  SdfDataSet1.Refresh;
  UpDateInfo;

end;

procedure TForm1.MenuItem3Click(Sender: TObject);
var
  sum: currency;
  BM: TBookmark;
begin
  try
    SdfDataSet1.DisableControls;
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    BM := SdfDataSet1.GetBookmark;
    sum := 0;

    SdfDataSet1.First;
    while not SdfDataSet1.EOF do
    begin
      sum := sum + SdfDataSet1.Fields[DBGrid1.SelectedIndex].AsCurrency;
      SdfDataSet1.Next;
    end;
    SdfDataSet1.GotoBookmark(BM);
  finally
    SdfDataSet1.FreeBookmark(BM);
    Screen.Cursor := crDefault;
    Application.ProcessMessages;
    SdfDataSet1.EnableControls;
    ShowMessage(FormatFloat('Summe: #,#0.00', sum));
  end;

end;

procedure TForm1.MenuItem4Click(Sender: TObject);
var
  BM: TBookmark;
  anz: longint;
  erg: string;
begin
  try
    SdfDataSet1.DisableControls;
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    anz := 0;
    BM := SdfDataSet1.GetBookmark;

    SdfDataSet1.First;
    while not SdfDataSet1.EOF do
    begin
      Inc(anz);
      SdfDataSet1.Next;
    end;
  finally
    SdfDataSet1.EnableControls;
    SdfDataSet1.GotoBookmark(BM);
    SdfDataSet1.FreeBookmark(BM);
    Screen.Cursor := crDefault;
    Application.ProcessMessages;
    //erg := FormatFloat('#,##0 Datensaetze',anz);
    erg := IntToStr(anz) + ' Datensätze';
    ShowMessage(erg);
  end;
end;

procedure TForm1.MenuItem5Click(Sender: TObject);
var
  col, SearchStr: string;
  BM: TBookmark;
  found: boolean;
begin
  try
    SearchStr := SdfDataSet1.Fields[DBGrid1.SelectedIndex].AsString;
    if InputQuery('Suche in aktueller Spalte', 'Suche nach:', False,
      SearchStr) = False then
      exit;

    SdfDataSet1.DisableControls;

    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;

    found := False;

    BM := SdfDataSet1.GetBookmark;

    col := SdfDataSet1.Fields[DBGrid1.SelectedIndex].FieldName;


    //SdfDataSet1.First;

    while not SdfDataSet1.EOF do
    begin
      found := UpperCase(SearchStr) =
        UpperCase(SdfDataSet1.Fields[DBGrid1.SelectedIndex].AsString);
      if found then
        break;
      SdfDataSet1.Next;
    end;



    if not found then
    begin
      SdfDataSet1.GotoBookmark(BM);
      Screen.Cursor := crDefault;
      Application.ProcessMessages;
      ShowMessage('''' + SearchStr + ''' konnte in Spalte ''' + col +
        ''' nicht gefunden werden.' + NL + NL +
        'Es wurde ab aktueller Zeile gesucht, also ggf. die erste Zeile selber auswählen');
    end;

  finally
    SdfDataSet1.EnableControls;
    SdfDataSet1.FreeBookmark(BM);
    Screen.Cursor := crDefault;
    Application.ProcessMessages;
  end;

end;

procedure TForm1.MenuItem6Click(Sender: TObject);
var
  BM: TBookmark;
  x, col: integer;
  memo: TMemo;
  delim, sline: string;
begin
  (* angezeigte Datensätze kopieren *)
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;

    SdfDataSet1.DisableControls;

    memo := TMemo.Create(Application.MainForm);
    memo.Parent := Application.MainForm;
    memo.Visible := False;
    memo.Clear;
    Application.ProcessMessages;

    //delim :=  trim(EditDelimiter.Text);
    // if trim(delim) = '' then
    delim := #9;


    x := 0;

    BM := SdfDataSet1.GetBookmark;


    (* erste Zeile: Spaltentitel *)
    for col := 0 to DBGrid1.Columns.Count - 1 do
      if col = 0 then
        sline := DBGrid1.Columns[col].Title.Caption
      else
        sline := sline + delim + DBGrid1.Columns[col].Title.Caption;

    (* auch das letzte Feld mit Feldtrennzeichen beenden *)
    sline := sline + delim;

    memo.Lines.add(sline);
    Application.ProcessMessages;

    SdfDataSet1.First;

    (* jetzt die eigentlichen Daten ins Memo kopieren *)
    while not SdfDataSet1.EOF do
    begin
      sline := '';
      for col := 0 to SdfDataSet1.FieldCount - 1 do
        if col = 0 then
          sline := SdfDataSet1.Fields[col].AsString
        else
          sline := sline + delim + SdfDataSet1.Fields[col].AsString;


      (* auch das letzte Feld mit Feldtrennzeichen beenden *)
      sline := sline + delim;

      memo.Lines.add(sline);

      Inc(x);
      SdfDataSet1.Next;
    end;

    SdfDataSet1.GotoBookmark(BM);

    Application.ProcessMessages;
    memo.SelectAll;
    Application.ProcessMessages;
    memo.CopyToClipboard;
    Application.ProcessMessages;


    Screen.Cursor := crDefault;
    Application.ProcessMessages;

    ShowMessage(IntToStr(x) +
      ' Datensätze wurden kopiert und können jetzt z.B. in Excel einegfügt werden.');

  finally
    SdfDataSet1.EnableControls;
    Screen.Cursor := crDefault;
    Application.ProcessMessages;
    SdfDataSet1.FreeBookmark(BM);
    FreeAndNil(memo);
  end;
end;

procedure TForm1.PopupDBGridPopup(Sender: TObject);
var
  curr: currency;
begin
  try
    if not SdfDataSet1.Active then
    begin
      PopupDBGrid.Close;
      exit;
    end;
    curr := SdfDataSet1.Fields[DBGrid1.SelectedIndex].AsCurrency;
    PopupDBGrid.Items[2].Enabled := True;
  except
    on E: EConvertError do
      PopupDBGrid.Items[2].Enabled := False;

  end;
end;

procedure TForm1.SdfDataSet1AfterPost(DataSet: TDataSet);
begin
  if DataSet.Filtered then
    DataSet.Refresh;
end;

procedure TForm1.SdfDataSet1UpdateInfo(DataSet: TDataSet);
begin
  UpDateInfo;
end;

procedure TForm1.SdfDataSet1FilterRecord(DataSet: TDataSet; var Accept: boolean);
var
  x: integer;
begin
  (* Auswahlfilter setzen *)
  Accept := False;
  (* Das dynamische Array durchlaufen und mit den Feldwerten des aktuellen Datensatzes vergleichen *)
  for x := low(Filters) to High(Filters) do
  begin
    if AnsiCompareText(DataSet.Fields[Filters[x].FieldIndex].AsString,
      Filters[x].Filter) = 0 then
      Accept := True
    else
      Accept := False;
    //ShowMessage(DataSet.Fields[Filters[x].FieldIndex].FieldName + ' mit '+
    //DataSet.Fields[Filters[x].FieldIndex].AsString + ' soll sein ' + Filters[x].Filter );

    (* Ablehnen, sobald ein Suchkriterium mit dem Feldwert NICHT übereinstimmt *)
    if not Accept then
      break;
  end;

end;

procedure TForm1.AShowExample(Sender: TObject);
begin
  SdfDataSet1.RemoveExtraColumns;
  if not SdfDataSet1.Active then
    exit;

  if pos(trim(EditSearch.Text), SdfDataSet1.Fields[DbGrid1.SelectedIndex].AsString) >
    0 then

    StatusBar1.SimpleText := 'Beispiel: ersetze in der Auswahl ''' +
      SdfDataSet1.Fields[DbGrid1.SelectedIndex].AsString + ''' durch ''' +
      AnsiReplaceText(SdfDataSet1.Fields[DbGrid1.SelectedIndex].AsString,
      trim(EditSearch.Text), trim(EditReplace.Text)) + ''''
  else
    StatusBar1.SimpleText := '';

end;

procedure TForm1.UpDateInfo;
begin
  lbRecordCount.Visible := SdfDataSet1.Filtered;

  if not SdfDataSet1.Filtered then
  begin
    SetLength(Filters, 0);
    lbRecordCount.Caption := '';
  end;

end;

end.

