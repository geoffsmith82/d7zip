unit Unit1;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  System.IOUtils,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.ExtCtrls,
  Vcl.FileCtrl,
  Vcl.ComCtrls,
  Vcl.StdCtrls,
  JvExStdCtrls,
  JvListBox,
  JvDriveCtrls,
  JvCombobox,
  sevenzip
  ;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    JvDriveCombo1: TJvDriveCombo;
    JvDirectoryListBox1: TJvDirectoryListBox;
    JvFileListBox1: TJvFileListBox;
    TreeView1: TTreeView;
    procedure FormCreate(Sender: TObject);
    procedure JvFileListBox1Change(Sender: TObject);
  private
    dllPath : string;
    archive : I7zInArchive;
    factory : IArchiveFactory;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  factory := TArchiveFactory.Create;
  dllPath := TPath.Combine( ExtractFilePath(ParamStr(0)),'..\..\..\..\bin64\7z.dll');
  dllPath := TPath.GetFullPath(dllPath);
end;

procedure TForm1.JvFileListBox1Change(Sender: TObject);
var
  i : Integer;

begin
  if factory = nil then Exit;
  if JvFileListBox1.FileName.IsEmpty then Exit;
  TreeView1.Items.Clear;

  archive := factory.CreateInArchive(JvFileListBox1.FileName, dllPath);
  archive.OpenFile(JvFileListBox1.FileName);
  TreeView1.LockDrawing;
  try
    for i := 0 to archive.NumberOfItems - 1 do
    begin
      if not archive.ItemIsFolder[i] then
      begin
        TreeView1.Items.AddChild(nil, archive.ItemPath[i]);
      end;
    end;
  finally
    TreeView1.UnlockDrawing;
  end;
end;

end.
