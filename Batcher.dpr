program Batcher;

uses
  Vcl.Forms,
  BatcherU in 'BatcherU.pas' {frmBatcher};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmBatcher, frmBatcher);
  Application.Run;
end.
