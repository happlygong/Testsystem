program search;

uses
  Forms,
  searchu in 'searchu.pas' {Form1},
  aboutU in 'aboutU.pas' {AboutBox};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
//  Application.CreateForm(Tabout, about);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
end.
