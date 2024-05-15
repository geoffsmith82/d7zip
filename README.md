# 7-zip Delphi API

* Author: Henri Gourvest <hgourvest@progdigy.com>
* Licence: MPL1.1
* Date: 2011-11-29
* Version: 1.2
* [Original source at Google Code](https://code.google.com/archive/p/d7zip/source/default/commits)
* Extended version by Daniel Marschall 2024-05-15 with a lot of changes (see sevenzip.pas header for changelog)

This API use the 7-zip dll (7z.dll) to read and write all 7-zip supported archive formats.  The latest 32-bit and 64-bit version of the 7z.dll is included in the repository (currently 24.4).

Please note: Use 7z.dll (taken from the 32/64 bit package of 7zip), not 7za.dll from the extras package!

  
## Reading archive:
### Extract to path:

```pascal
 with CreateInArchive(CLSID_CFormatZip) do
 begin
   OpenFile('c:\test.zip');
   ExtractTo('c:\test');
 end;
```
### Get file list:
```Pascal
 with CreateInArchive(CLSID_CFormat7z) do
 begin
   OpenFile('c:\test.7z');
   for i := 0 to NumberOfItems - 1 do
    if not ItemIsFolder[i] then
      Writeln(ItemPath[i]);
 end;
```
### Extract to stream
```Pascal
 with CreateInArchive(CLSID_CFormat7z) do
 begin
   OpenFile('c:\test.7z');
   for i := 0 to NumberOfItems - 1 do
     if not ItemIsFolder[i] then
       ExtractItem(i, stream, false);
 end;
```
### Extract "n" Items
```Pascal
function GetStreamCallBack(sender: Pointer; index: Cardinal;
  var outStream: ISequentialOutStream): HRESULT; stdcall;
begin
  case index of ...
    outStream := T7zStream.Create(aStream, soReference);
  Result := S_OK;
end;

procedure TMainForm.ExtractClick(Sender: TObject);
var
  i: integer;
  items: array[0..2] of Cardinal;
begin
  with CreateInArchive(CLSID_CFormat7z) do
  begin
    OpenFile('c:\test.7z');
    // items must be sorted by index!
    items[0] := 0;
    items[1] := 1;
    items[2] := 2;
    ExtractItems(@items, Length(items), false, nil, GetStreamCallBack);
  end;
end;
```
### Open stream
```Pascal
 with CreateInArchive(CLSID_CFormatZip) do
 begin
   OpenStream(T7zStream.Create(TFileStream.Create('c:\test.zip', fmOpenRead), soOwned));
   OpenStream(aStream, soReference);
   ...
 end;
```
### Progress bar
```Pascal
 function ProgressCallback(sender: Pointer; total: boolean; value: int64): HRESULT; stdcall;
 begin
   if total then
     Mainform.ProgressBar.Max := value else
     Mainform.ProgressBar.Position := value;
   Result := S_OK;
 end;

 procedure TMainForm.ExtractClick(Sender: TObject);
 begin
   with CreateInArchive(CLSID_CFormatZip) do
   begin
     OpenFile('c:\test.zip');
     SetProgressCallback(nil, ProgressCallback);
     ...
   end;
 end;
```
### Password
```Pascal
 function PasswordCallback(sender: Pointer; var password: WideString): HRESULT; stdcall;
 begin
   // call a dialog box ...
   password := 'password';
   Result := S_OK;
 end;

 procedure TMainForm.ExtractClick(Sender: TObject);
 begin
   with CreateInArchive(CLSID_CFormatZip) do
   begin
     // using callback
     SetPasswordCallback(nil, PasswordCallback);
     // or setting password directly
     SetPassword('password');
     OpenFile('c:\test.zip');
     ...
   end;
 end;
```
### Writing archive
```Pascal
 procedure TMainForm.PackFilesClick(Sender: TObject);
 var
   Arch: I7zOutArchive;
 begin
   Arch := CreateOutArchive(CLSID_CFormat7z);
   // add a file
   Arch.AddFile('c:\test.bin', 'folder\test.bin');
   // add files using willcards and recursive search
   Arch.AddFiles('c:\test', 'folder', '*.pas;*.dfm', true);
   // add a stream
   Arch.AddStream(aStream, soReference, faArchive, CurrentFileTime, CurrentFileTime, 'folder\test.bin', false, false);
   // compression level
   SetCompressionLevel(Arch, 5);
   // compression method if <> LZMA
   SevenZipSetCompressionMethod(Arch, m7BZip2);
   // add a progress bar ...
   Arch.SetProgressCallback(...);
   // set a password if necessary
   Arch.SetPassword('password');
   // Save to file
   Arch.SaveToFile('c:\test.zip');
   // or a stream
   Arch.SaveToStream(aStream);
 end;
```
