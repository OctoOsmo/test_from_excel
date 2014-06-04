program ProjectTestExcel;

uses
  Forms,
  TestFromExcel in 'TestFromExcel.pas' {FormMain};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
