unit BatcherU;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Winapi.ShellAPI, VCLTee.TeeSVGCanvas,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, WebImage,
  Vcl.ExtCtrls, Vcl.ToolWin, Vcl.ComCtrls, Vcl.StdCtrls, VCLTee.TeeProcs,
  VCLTee.TeeDraw3D, Vcl.ExtDlgs, Vcl.OleCtrls, SHDocVw, Vcl.ImgList;

type
  TfrmBatcher = class(TForm)
    pnlFiles: TPanel;
    pnlScript: TPanel;
    mmScript: TMemo;
    Splitter1: TSplitter;
    pnlWmf: TPanel;
    Splitter2: TSplitter;
    pnlSVG: TPanel;
    ToolBar2: TToolBar;
    pnlPreviewWmf: TPanel;
    Splitter3: TSplitter;
    lbWMF: TListBox;
    tbConvert: TToolButton;
    ToolButton4: TToolButton;
    tbTargetDir: TToolButton;
    ToolButton2: TToolButton;
    stSvgDir: TStaticText;
    picWMF: TDraw3D;
    chkPreviewSvg: TCheckBox;
    pnlPreviewSvg: TPanel;
    Splitter4: TSplitter;
    svgViewer: TWebBrowser;
    lbSvg: TListBox;
    ToolBar3: TToolBar;
    tbAddFile: TToolButton;
    tbAddFolder: TToolButton;
    ToolButton7: TToolButton;
    tbRemove: TToolButton;
    ToolButton9: TToolButton;
    chkPreviewWmf: TCheckBox;
    StatusBar1: TStatusBar;
    OpenPictureDialog1: TOpenPictureDialog;
    OpenDialog1: TOpenDialog;
    tbReset: TToolButton;
    ImageList1: TImageList;
    procedure tbAddFileClick(Sender: TObject);
    procedure tbAddFolderClick(Sender: TObject);
    procedure lbWMFClick(Sender: TObject);
    procedure chkPreviewWmfClick(Sender: TObject);
    procedure chkPreviewSvgClick(Sender: TObject);
    procedure tbTargetDirClick(Sender: TObject);
    procedure tbRemoveClick(Sender: TObject);
    procedure tbResetClick(Sender: TObject);
    procedure Splitter2Moved(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure tbConvertClick(Sender: TObject);
    procedure lbSvgClick(Sender: TObject);
  private
    { Private declarations }
    SplitRatio: double;
    ExePath: string;
  public
    { Public declarations }
    SvgDir: string;
    procedure UpdateCount;
  end;

var frmBatcher: TfrmBatcher;

implementation
{$R *.dfm}
uses Vcl.FileCtrl, System.Win.Registry;

function GetDosOutput(CommandLine: string; Work: string = 'C:\'): string;
var
  SA: TSecurityAttributes;
  SI: TStartupInfo;
  PI: TProcessInformation;
  StdOutPipeRead, StdOutPipeWrite: THandle;
  WasOK: Boolean;
  Buffer: array[0..255] of AnsiChar;
  BytesRead: Cardinal;
  WorkDir: string;
  Handle: Boolean;
begin
  Result := '';
  with SA do begin
    nLength := SizeOf(SA);
    bInheritHandle := True;
    lpSecurityDescriptor := nil;
  end;
  CreatePipe(StdOutPipeRead, StdOutPipeWrite, @SA, 0);
  try
    with SI do
    begin
      FillChar(SI, SizeOf(SI), 0);
      cb := SizeOf(SI);
      dwFlags := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
      wShowWindow := SW_HIDE;
      hStdInput := GetStdHandle(STD_INPUT_HANDLE); // don't redirect stdin
      hStdOutput := StdOutPipeWrite;
      hStdError := StdOutPipeWrite;
    end;
    WorkDir := Work;
    Handle := CreateProcess(nil, PChar('cmd.exe /C ' + CommandLine),
                            nil, nil, True, 0, nil,
                            PChar(WorkDir), SI, PI);
    CloseHandle(StdOutPipeWrite);
    if Handle then
      try
        repeat
          WasOK := ReadFile(StdOutPipeRead, Buffer, 255, BytesRead, nil);
          if BytesRead > 0 then
          begin
            Buffer[BytesRead] := #0;
            Result := Result + Buffer;
          end;
        until not WasOK or (BytesRead = 0);
        WaitForSingleObject(PI.hProcess, INFINITE);
      finally
        CloseHandle(PI.hThread);
        CloseHandle(PI.hProcess);
      end;
  finally
    CloseHandle(StdOutPipeRead);
  end;
end;

function Check4Java: boolean;
var reg: TRegistry;
begin
  reg := TRegistry.Create(KEY_READ);
  reg.RootKey := HKEY_LOCAL_MACHINE;
  Result:= reg.KeyExists('SOFTWARE\Classes\Applications\javaw.exe');
  reg.Free;
end;

procedure TfrmBatcher.tbConvertClick(Sender: TObject);
var                                // ext -> *.ext
  i, j: integer;
  dr, wfn, sfn,ss: string;
begin
  if not FileExists(ExePath+'wmf2svg-0.9.5.jar') then
    raise Exception.Create('Error: file <wmf2svg-0.9.5.jar> is missing.');
  for i := 0 to lbWmf.Items.Count-1 do
  begin
    wfn:= lbWmf.Items[i];
    sfn:= ChangeFileExt(ExtractFileName(wfn),'.svg');
    if DirectoryExists(SvgDir) then
      sfn:= IncludeTrailingBackslash(SvgDir)+sfn
    else
      sfn:= IncludeTrailingBackslash(ExtractFileDir(lbWmf.Items[i]))+sfn;
    lbWmf.ItemIndex:= i; lbWMFClick(nil); Application.ProcessMessages;
    if FileExists(sfn) then
       if MessageDlg('Target file "'+sfn+'" exists. Overwrite it? ',
                     mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrNo
         then continue;
    ss:= 'java -jar wmf2svg-0.9.5.jar '+wfn+' '+sfn;
    mmScript.Lines.Add(ss);
    mmScript.Lines.Add(GetDosOutput(ss, ExePath));
    lbSvg.Items.Add(sfn);
    lbSvg.ItemIndex:= lbSvg.Items.Count-1; lbSvgClick(nil);
  end;
end;

function GetFilesFromDirectory(dir, ext: string; Flist: TStrings): boolean;
var                                // ext -> *.ext
  i, j, FileAttrs: integer;
  sr: TSearchRec;
  dr, ss: string;
begin
  Result := False;
  FileAttrs := faArchive;
  if not DirectoryExists(dir) then
    exit;
  if FindFirst(dir + ext, FileAttrs, sr) = 0 then
  begin
    repeat
      Flist.Add(dir + sr.name);
    until FindNext(sr) <> 0;
    FindClose(sr);
  end;
  Result := True;
end;

procedure TfrmBatcher.UpdateCount;
begin
  chkPreviewWmf.Caption:= 'Preview   Count='+IntToStr(lbWmf.Items.Count);
end;

procedure TfrmBatcher.tbAddFileClick(Sender: TObject);
var
  i: integer;
  fn: string;
begin
  if not OpenPictureDialog1.Execute then
    exit;
  for fn in OpenPictureDialog1.Files do
  begin
    if not SameText(ExtractFileExt(fn), '.wmf') then
      continue;
    if lbWMF.Items.IndexOf(fn) = -1 then
      lbWMF.Items.Add(fn);
  end;
  UpdateCount;
end;

function PickDirectory(var dr: string): boolean;
begin
  Result:= false;
  with TFileOpenDialog.Create(nil) do
    try
      Title := 'Select Directory';
      Options := [fdoPickFolders, fdoPathMustExist, fdoForceFileSystem]; // YMMV
      OkButtonLabel := 'Select';
      DefaultFolder := dr;
      FileName := dr;
      if not Execute then exit;
      dr := FileName;
      Result:= true;
    finally
      Free;
    end;
end;

procedure TfrmBatcher.tbAddFolderClick(Sender: TObject);
var
  Flist: TStrings;
  dr, fn: string;
begin
  Flist := TStringList.Create;
  if not PickDirectory(dr) then exit;
  GetFilesFromDirectory(dr + '\', '*.WMF', Flist);
  for fn in Flist do
  begin
    if not SameText(ExtractFileExt(fn), '.wmf') then
      continue;
    if lbWMF.Items.IndexOf(fn) = -1 then
      lbWMF.Items.Add(fn);
  end;
  Flist.Free;
  UpdateCount;
end;

procedure TfrmBatcher.tbRemoveClick(Sender: TObject);
var  i, j: integer;
begin
  for i:= lbWMF.Items.Count-1 downto 0 do
    if lbWMF.Selected[i] then lbWMF.Items.Delete(i);
  UpdateCount;
end;

procedure TfrmBatcher.tbResetClick(Sender: TObject);
begin
  lbWMF.Items.Clear; lbSvg.Items.Clear;
  stSvgDir.Caption:= '  Dir: ';
  UpdateCount;
end;

procedure TfrmBatcher.tbTargetDirClick(Sender: TObject);
begin
  if not PickDirectory(SvgDir) then exit;
  stSvgDir.Caption:= '  Dir: '+SvgDir;
end;

procedure TfrmBatcher.chkPreviewSvgClick(Sender: TObject);
begin
  pnlPreviewSvg.Visible := chkPreviewSvg.Checked;
  Splitter4.Visible := pnlPreviewSvg.Visible;
end;

procedure TfrmBatcher.chkPreviewWmfClick(Sender: TObject);
begin
  pnlPreviewWmf.Visible := chkPreviewWmf.Checked;
  Splitter3.Visible := pnlPreviewWmf.Visible;
end;

procedure TfrmBatcher.FormCreate(Sender: TObject);
begin
  SplitRatio:= 0.5; FormResize(nil);
  ExePath:= IncludeTrailingBackslash(ExtractFileDir(Application.ExeName));
  if not Check4Java then
    raise Exception.Create('Java installation is nowhere to be found (hint: you need one)');
end;

procedure TfrmBatcher.FormResize(Sender: TObject);
begin
   pnlWmf.width:= round(width*SplitRatio);
end;

procedure LoadHTML2WebBrowser(WebBrowser: TWebBrowser; HTMLfile: string);
var Flags: OleVariant; ss: string;
begin
  Flags:= OleVariant(2+4+8);
  ss:= 'file://'+StringReplace(HTMLfile,'\','/',[rfReplaceAll]);
  WebBrowser.Navigate(ss,Flags);
end;

procedure TfrmBatcher.lbSvgClick(Sender: TObject);
begin
   LoadHTML2WebBrowser(svgViewer,lbSvg.Items[lbSvg.ItemIndex]);
end;

procedure TfrmBatcher.lbWMFClick(Sender: TObject);
begin
  picWMF.BackImage.LoadFromFile(lbWMF.Items[lbWMF.ItemIndex]);
end;

procedure TfrmBatcher.Splitter2Moved(Sender: TObject);
begin
  SplitRatio:= pnlWmf.width / width;
end;

end.
