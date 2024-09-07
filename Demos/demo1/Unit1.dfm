object Form1: TForm1
  Left = 0
  Top = 0
  Caption = '7z.dll Demo'
  ClientHeight = 442
  ClientWidth = 856
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 233
    Height = 442
    Align = alLeft
    Caption = 'Panel1'
    TabOrder = 0
    ExplicitHeight = 441
    object JvDriveCombo1: TJvDriveCombo
      Left = 1
      Top = 1
      Width = 231
      Height = 22
      Align = alTop
      DriveTypes = [dtFixed, dtRemote, dtCDROM]
      Offset = 4
      TabOrder = 0
    end
    object JvDirectoryListBox1: TJvDirectoryListBox
      Left = 1
      Top = 23
      Width = 231
      Height = 202
      Align = alTop
      Directory = 'c:\program files (x86)\embarcadero\studio\22.0\bin'
      FileList = JvFileListBox1
      DriveCombo = JvDriveCombo1
      ItemHeight = 17
      TabOrder = 1
      ExplicitLeft = 0
      ExplicitTop = 17
    end
    object JvFileListBox1: TJvFileListBox
      Left = 1
      Top = 225
      Width = 231
      Height = 216
      Align = alClient
      ItemHeight = 15
      Mask = '*.zip; *.cab; *.7z; *.gzip; *.iso; *.wim; *.rar'
      TabOrder = 2
      OnChange = JvFileListBox1Change
      ForceFileExtensions = False
    end
  end
  object TreeView1: TTreeView
    Left = 233
    Top = 0
    Width = 305
    Height = 442
    Align = alLeft
    Indent = 19
    TabOrder = 1
    ExplicitLeft = 239
    ExplicitTop = 8
    ExplicitHeight = 273
  end
end
