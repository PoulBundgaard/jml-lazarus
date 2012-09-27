unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  ComCtrls, XMLPropStorage, StdCtrls;

type

  { TForm1 }

  TForm1 = class(TForm)
    ApplicationProperties1: TApplicationProperties;
    Button1: TButton;
    StatusBar1: TStatusBar;
    XMLPropStorage1: TXMLPropStorage;
    procedure ApplicationProperties1Hint(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
    procedure JEi;
    procedure NEi;

  end; 

var
  Form1: TForm1;

  ExePath : String;
  const NL = chr(10) + chr(13);


implementation

{ TForm1 }

procedure TForm1.FormCreate(Sender: TObject);
begin
  ExePath := ExtractFilePath(Application.ExeName);
end;

procedure TForm1.ApplicationProperties1Hint(Sender: TObject);
begin
    StatusBar1.SimpleText:= Application.Hint;
end;



procedure TForm1.JEi;
begin
  screen.cursor := crHourglass;
end;

procedure TForm1.NEi;
begin
  screen.cursor := crDefault;
end;

initialization
  {$I unit1.lrs}

end.

