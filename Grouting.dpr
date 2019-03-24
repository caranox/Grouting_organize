program Grouting;

uses
  Forms,
  Unit1 in '..\Grouting\Unit1.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '灌浆数据整理';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
