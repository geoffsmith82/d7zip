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
  Vcl.StdCtrls,
  Vcl.Menus,
  Vcl.ComCtrls,
  JvExStdCtrls,
  JvListBox,
  JvDriveCtrls,
  JvCombobox,
  sevenzip,
  VirtualTrees;

type
  PFileItem = ^TFileItem;
  TFileItem = record
    ItemName: string;
    IsFolder: Boolean;
  end;

  TForm1 = class(TForm)
    Panel1: TPanel;
    JvDriveCombo1: TJvDriveCombo;
    JvDirectoryListBox1: TJvDirectoryListBox;
    JvFileListBox1: TJvFileListBox;
    StatusBar: TStatusBar;
    MainMenu: TMainMenu;
    File1: TMenuItem;
    New1: TMenuItem;
    Open1: TMenuItem;
    Save1: TMenuItem;
    SaveAs1: TMenuItem;
    Print1: TMenuItem;
    PrintSetup1: TMenuItem;
    Exit1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    VirtualStringTree1: TVirtualStringTree;
    procedure Exit1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure JvFileListBox1Change(Sender: TObject);
    procedure VirtualStringTree1GetText(Sender: TBaseVirtualTree; Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
  private
    dllPath: string;
    archive: I7zInArchive;
    factory: IArchiveFactory;
    procedure PopulateVirtualTree;
    function FindOrAddNode(Parent: PVirtualNode; const Name: string; IsFolder: Boolean): PVirtualNode;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Exit1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  factory := TArchiveFactory.Create;
  dllPath := TPath.Combine(ExtractFilePath(ParamStr(0)), '..\..\..\..\bin64\7z.dll');
  dllPath := TPath.GetFullPath(dllPath);

  // Set up the VirtualStringTree
  VirtualStringTree1.NodeDataSize := SizeOf(TFileItem);
end;

procedure TForm1.JvFileListBox1Change(Sender: TObject);
begin
  if factory = nil then
    Exit;
  if JvFileListBox1.FileName.IsEmpty then
    Exit;

  archive := factory.CreateInArchive(JvFileListBox1.FileName, dllPath);
  archive.OpenFile(JvFileListBox1.FileName);

  // Clear existing data in the VirtualStringTree
  VirtualStringTree1.Clear;

  // Populate the VirtualStringTree with archive data
  PopulateVirtualTree;
end;

procedure TForm1.PopulateVirtualTree;
var
  i: Integer;
  PathParts: TArray<string>;
  ParentNode: PVirtualNode;
  Part: string;
begin
  VirtualStringTree1.BeginUpdate;
  try
    for i := 0 to archive.NumberOfItems - 1 do
    begin
      PathParts := archive.ItemPath[i].Split(['\', '/']);  // Split path into parts
      ParentNode := nil;
      for Part in PathParts do
      begin
        if (Part = '') then
          Continue;
        // Check if it's the last part (file), otherwise it's a folder
        ParentNode := FindOrAddNode(ParentNode, Part, (Part = PathParts[High(PathParts)]) and (not archive.ItemIsFolder[i]));
      end;
    end;
  finally
    VirtualStringTree1.EndUpdate;
  end;
end;

function TForm1.FindOrAddNode(Parent: PVirtualNode; const Name: string; IsFolder: Boolean): PVirtualNode;
var
  Node: PVirtualNode;
  Data: PFileItem;
begin
  // Iterate through the children of the Parent node and find if the node already exists
  Node := VirtualStringTree1.GetFirstChild(Parent);
  while Node <> nil do
  begin
    Data := VirtualStringTree1.GetNodeData(Node);
    if (Data.ItemName = Name) and (Data.IsFolder = IsFolder) then
      Exit(Node); // Return the node if it already exists
    Node := VirtualStringTree1.GetNextSibling(Node);
  end;

  // If not found, create a new node
  Node := VirtualStringTree1.AddChild(Parent);
  Data := VirtualStringTree1.GetNodeData(Node);
  Data.ItemName := Name;
  Data.IsFolder := IsFolder;

  Result := Node; // Return the newly created node
end;

procedure TForm1.VirtualStringTree1GetText(Sender: TBaseVirtualTree; Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
var
  Data: PFileItem;
begin
  Data := Sender.GetNodeData(Node);
  if Assigned(Data) then
  begin
    CellText := Data.ItemName; // Display the item name in the tree
  end;
end;

end.

