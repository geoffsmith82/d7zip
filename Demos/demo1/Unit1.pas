unit Unit1;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  System.IOUtils,
  System.ImageList,
  Vcl.ImgList,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.ExtCtrls,
  Vcl.FileCtrl,
  Vcl.StdCtrls,
  Vcl.Menus,
  Vcl.ComCtrls,
  Shellapi,
  JvExStdCtrls,
  JvListBox,
  JvDriveCtrls,
  JvCombobox,
  sevenzip,
  VirtualTrees
  ;

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
    ImageList1: TImageList;
    procedure Exit1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure JvFileListBox1Change(Sender: TObject);
    procedure VirtualStringTree1GetText(Sender: TBaseVirtualTree; Node: PVirtualNode; Column: TColumnIndex; TextType: TVSTTextType; var CellText: string);
    procedure VirtualStringTree1GetImageIndexEx(Sender: TBaseVirtualTree;
      Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
      var Ghosted: Boolean; var ImageIndex: TImageIndex;
      var ImageList: TCustomImageList);
  private
    dllPath: string;
    archive: I7zInArchive;
    factory: IArchiveFactory;
    procedure PopulateVirtualTree;
    function FindOrAddNode(Parent: PVirtualNode; const Name: string; IsFolder: Boolean): PVirtualNode;
    procedure LoadShellIcons;
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
  LoadShellIcons;

  // Set up the VirtualStringTree
  VirtualStringTree1.NodeDataSize := SizeOf(TFileItem);
end;


procedure TForm1.LoadShellIcons;
var
  FileInfo: SHFILEINFO;
  FolderIcon, FileIcon: TIcon;
begin
  // Set up ImageList to support 32-bit icons (with alpha transparency)
  ImageList1.ColorDepth := cd32Bit;
  ImageList1.Masked := True;

  // Create TIcon instances for folder and file icons
  FolderIcon := TIcon.Create;
  FileIcon := TIcon.Create;
  try
    // Retrieve the folder icon using SHGetFileInfo
    SHGetFileInfo('C:\WINDOWS\', FILE_ATTRIBUTE_DIRECTORY, FileInfo, SizeOf(FileInfo), SHGFI_ICON or SHGFI_SMALLICON or SHGFI_USEFILEATTRIBUTES);
    FolderIcon.Handle := FileInfo.hIcon;  // Assign the icon handle to the TIcon
    ImageList1.AddIcon(FolderIcon);       // Add the folder icon to the ImageList
    DestroyIcon(FileInfo.hIcon);          // Free the system icon handle

    // Retrieve the generic file icon using SHGetFileInfo
    SHGetFileInfo('C:\dummy.txt', FILE_ATTRIBUTE_NORMAL, FileInfo, SizeOf(FileInfo), SHGFI_ICON or SHGFI_SMALLICON or SHGFI_USEFILEATTRIBUTES);
    FileIcon.Handle := FileInfo.hIcon;    // Assign the icon handle to the TIcon
    ImageList1.AddIcon(FileIcon);         // Add the file icon to the ImageList
    DestroyIcon(FileInfo.hIcon);          // Free the system icon handle

  finally
    // Free the TIcon instances
    FolderIcon.Free;
    FileIcon.Free;
  end;
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

procedure TForm1.VirtualStringTree1GetImageIndexEx(Sender: TBaseVirtualTree;
  Node: PVirtualNode; Kind: TVTImageKind; Column: TColumnIndex;
  var Ghosted: Boolean; var ImageIndex: TImageIndex;
  var ImageList: TCustomImageList);
var
  Data: PFileItem;
begin
  ImageList := ImageList1;
  Data := Sender.GetNodeData(Node);
  if Kind = ikState then
    Exit;

  if Assigned(Data) then
  begin
    // Assign the correct icon index based on whether it's a folder or file

    if Data.IsFolder then
      ImageIndex := 1  // Folder icon (index 0 in ImageList1)
    else
      ImageIndex := 0; // File icon (index 1 in ImageList1)
  end;

end;

end.

