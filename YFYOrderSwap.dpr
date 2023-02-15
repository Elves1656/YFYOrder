program YFYOrderSwap;

uses
  Vcl.Forms,
  uDataSwap in 'uDataSwap.pas' {fmDataSwap};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfmDataSwap, fmDataSwap);
  Application.Run;
end.
