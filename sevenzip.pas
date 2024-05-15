(********************************************************************************)
(*                        7-ZIP DELPHI API                                      *)
(*                                                                              *)
(* The contents of this file are subject to the Mozilla Public License Version  *)
(* 1.1 (the "License"); you may not use this file except in compliance with the *)
(* License. You may obtain a copy of the License at http://www.mozilla.org/MPL/ *)
(*                                                                              *)
(* Software distributed under the License is distributed on an "AS IS" basis,   *)
(* WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for *)
(* the specific language governing rights and limitations under the License.    *)
(*                                                                              *)
(* Unit owner : Henri Gourvest <hgourvest@gmail.com>                            *)
(* V1.2                                                                         *)
(********************************************************************************)

// Original code at https://code.google.com/archive/p/d7zip/
// Uploaded to GitHub at https://github.com/danielmarschall/d7zip

// Current version by Daniel Marschall, 15 May 2024 with the following changes:
// - Added format GUID: RAR5; https://github.com/geoffsmith82/d7zip/issues/7
// - Fix Range Check Exception in RINOK(); https://github.com/geoffsmith82/d7zip/pull/8
// - Avoid unhandled Delphi Exceptions crashing the DLL parent process; https://github.com/geoffsmith82/d7zip/pull/9
// - Implemented packing and unpacking of empty directories, folders which begin with a dot, and hidden files; https://github.com/geoffsmith82/d7zip/pull/10
// - Implemented restoring of the file attributes and modification times; https://github.com/geoffsmith82/d7zip/pull/11
// - Added ExtractItemToPath method, by ekot1; https://github.com/ekot1/d7zip/commit/a6e35cac4fe2306307372f7b4647a7a620d86cf8
// - Fixed wrong method name in README.md; https://github.com/r3code/d7zip/commit/b7f067436b1259177603cf0cc6e64a80bceaa68e
// - Show better error message when 7z.dll can not be loaded, by ekot1; https://github.com/ekot1/d7zip/commit/4facb0ef8b190c129d494c9237337918b3dbeece
// - Changes to match propids from 7z.dll v16.04, by ekot1 (this also adds new format GUIDs); https://github.com/ekot1/d7zip/commit/149de16032fe461796857e5eee22c70858cdb4b9
// - Readme: The source version was actually 1.2 from 2011, and not 1.1 from 2009; https://github.com/danielmarschall/d7zip/commit/18cd6d2e20e755f8a261a9195dd9aadb12ae59d0
// - Fixed: The method in IOutStreamFinish is called OutStreamFinish, not Flush!
// - Updated interface definitions (Cardinal=>UInt32, Int64 sometimes UInt32, etc.); partially added missing interfaces from the 7zip source headers
// - Extended own interfaces to support 64 bit file sizes: https://github.com/wang80919/d7zip/commit/b89d4d7a2bc26928a3e8a1de896feac1a79706ce
//   and removed GUIDs because the signatures changed and we don't need COM interoperability for our interfaces
// - Added GetItemCompressedSize
// - Added LZMA2 to 7z methods (not tested)

// TODO: Possible changes to look closer at...
// - Add SetProgressCallbackEx method to allow use of anonymous methods as callbacks; https://github.com/ekot1/d7zip/commit/d850b85a05dd58ad6ded2823a635ab28b8cb62ca
// - Add SetProgressExceptCallback https://github.com/wang80919/d7zip/commit/626ad160001bb62671c959e43d052bea5047e950
// - Add packages and add namespace to unit; https://github.com/grandchef/d7zip/commit/b043af03cd22729be5c3515e27dfba402d0251f5#diff-3eea9649c6b570534a69a6f393cf9e7e7382bf9b8e6812a687d916d575237e02
// - Compatible with POSIX: https://github.com/zedalaye/d7zip/commit/7163f9f743c32c5b89e8f7f496ee960c46c4136d
//                          But this is complex, because we have things like FFileTime, TPropVariant, etc.!


unit sevenzip;
{$ALIGN ON}
{$MINENUMSIZE 4}
{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, ActiveX, SysUtils, Classes, Contnrs, System.IOUtils, Math;

type
  PInt32 = ^Int32;
  PVarType = ^TVarType;
  PCardArray = ^TCardArray;
  TCardArray = array[0..MaxInt div SizeOf(Cardinal) - 1] of Cardinal;

{$IFNDEF UNICODE}
  UnicodeString = WideString;
{$ENDIF}

{$REGION 'PropID.h'}

//******************************************************************************
// PropID.h
// (Last checked 14 May 2024; https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/PropID.h)
//******************************************************************************

const
  kpidNoProperty            = 0;
  kpidMainSubfile           = 1;
  kpidHandlerItemIndex      = 2;
  kpidPath                  = 3;  // VT_BSTR
  kpidName                  = 4;  // VT_BSTR
  kpidExtension             = 5;  // VT_BSTR
  kpidIsDir                 = 6;  // VT_BOOL
  kpidSize                  = 7;  // VT_UI8
  kpidPackSize              = 8;  // VT_UI8
  kpidAttrib                = 9;  // VT_UI4
  kpidCTime                 = 10; // VT_FILETIME
  kpidATime                 = 11; // VT_FILETIME
  kpidMTime                 = 12; // VT_FILETIME
  kpidSolid                 = 13; // VT_BOOL
  kpidCommented             = 14; // VT_BOOL
  kpidEncrypted             = 15; // VT_BOOL
  kpidSplitBefore           = 16; // VT_BOOL
  kpidSplitAfter            = 17; // VT_BOOL
  kpidDictionarySize        = 18; // VT_UI4
  kpidCRC                   = 19; // VT_UI4
  kpidType                  = 20; // VT_BSTR
  kpidIsAnti                = 21; // VT_BOOL
  kpidMethod                = 22; // VT_BSTR
  kpidHostOS                = 23; // VT_BSTR
  kpidFileSystem            = 24; // VT_BSTR
  kpidUser                  = 25; // VT_BSTR
  kpidGroup                 = 26; // VT_BSTR
  kpidBlock                 = 27; // VT_UI4
  kpidComment               = 28; // VT_BSTR
  kpidPosition              = 29; // VT_UI4
  kpidPrefix                = 30; // VT_BSTR
  kpidNumSubDirs            = 31; // VT_UI4
  kpidNumSubFiles           = 32; // VT_UI4
  kpidUnpackVer             = 33; // VT_UI1
  kpidVolume                = 34; // VT_UI4
  kpidIsVolume              = 35; // VT_BOOL
  kpidOffset                = 36; // VT_UI8
  kpidLinks                 = 37; // VT_UI4
  kpidNumBlocks             = 38; // VT_UI4
  kpidNumVolumes            = 39; // VT_UI4
  kpidTimeType              = 40; // VT_UI4
  kpidBit64                 = 41; // VT_BOOL
  kpidBigEndian             = 42; // VT_BOOL
  kpidCpu                   = 43; // VT_BSTR
  kpidPhySize               = 44; // VT_UI8
  kpidHeadersSize           = 45; // VT_UI8
  kpidChecksum              = 46; // VT_UI4
  kpidCharacts              = 47; // VT_BSTR
  kpidVa                    = 48; // VT_UI8
  kpidId                    = 49;
  kpidShortName             = 50;
  kpidCreatorApp            = 51;
  kpidSectorSize            = 52;
  kpidPosixAttrib           = 53;
  kpidSymLink               = 54;
  kpidError                 = 55;
  kpidTotalSize             = 56;
  kpidFreeSpace             = 57;
  kpidClusterSize           = 58;
  kpidVolumeName            = 59;
  kpidLocalName             = 60;
  kpidProvider              = 61;
  kpidNtSecure              = 62;
  kpidIsAltStream           = 63;
  kpidIsAux                 = 64;
  kpidIsDeleted             = 65;
  kpidIsTree                = 66;
  kpidSha1                  = 67;
  kpidSha256                = 68;
  kpidErrorType             = 69;
  kpidNumErrors             = 70;
  kpidErrorFlags            = 71;
  kpidWarningFlags          = 72;
  kpidWarning               = 73;
  kpidNumStreams            = 74;
  kpidNumAltStreams         = 75;
  kpidAltStreamsSize        = 76;
  kpidVirtualSize           = 77;
  kpidUnpackSize            = 78;
  kpidTotalPhySize          = 79;
  kpidVolumeIndex           = 80;
  kpidSubType               = 81;
  kpidShortComment          = 82;
  kpidCodePage              = 83;
  kpidIsNotArcType          = 84;
  kpidPhySizeCantBeDetected = 85;
  kpidZerosTailIsAllowed    = 86;
  kpidTailSize              = 87;
  kpidEmbeddedStubSize      = 88;
  kpidNtReparse             = 89;
  kpidHardLink              = 90;
  kpidINode                 = 91;
  kpidStreamId              = 92;
  kpidReadOnly              = 93;
  kpidOutName               = 94;
  kpidCopyLink              = 95;
  kpidArcFileName           = 96;
  kpidIsHash                = 97;
  kpidChangeTime            = 98;
  kpidUserId                = 99;
  kpidGroupId               =100;
  kpidDeviceMajor           =101;
  kpidDeviceMinor           =102;
  kpidDevMajor              =103;
  kpidDevMinor              =104;

  kpidUserDefined       = $10000;

{$ENDREGION}

{$REGION 'IProgress.h interfaces ("23170F69-40C1-278A-0000-000000xx0000")'}

//******************************************************************************
// IProgress.h
// (Last checked 14 May 2024; https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/IProgress.h)
//******************************************************************************
type
  IProgress = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000000050000}']
    function SetTotal(total: UInt64): HRESULT; stdcall;
    function SetCompleted(completeValue: PUInt64): HRESULT; stdcall;
  end;

{$ENDREGION}

{$REGION 'IFolderArchive.h interfaces ("23170F69-40C1-278A-0000-000100xx0000") - not implemented'}

// not implemented
// https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/UI/Agent/IFolderArchive.h

{$ENDREGION}

{$REGION 'IFolder.h interfaces ("23170F69-40C1-278A-0000-000800xx0000") - not implemented'}

// not implemented
// https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/UI/FileManager/IFolder.h

{$ENDREGION}

{$REGION 'IFolder.h FOLDER_MANAGER_INTERFACE ("23170F69-40C1-278A-0000-000900xx0000") - not implemented'}

// not implemented
// https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/UI/FileManager/IFolder.h

{$ENDREGION}

{$REGION 'PluginInterface.h interfaces ("23170F69-40C1-278A-0000-000A00xx0000") - not implemented'}

// not implemented
// https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/UI/FileManager/PluginInterface.h

{$ENDREGION}

{$REGION 'IPassword.h interfaces ("23170F69-40C1-278A-0000-000500xx0000")'}

//******************************************************************************
// IPassword.h
// (Last checked 14 May 2024; https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/IPassword.h)
//******************************************************************************

  ICryptoGetTextPassword = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000500100000}']
    function CryptoGetTextPassword(var password: TBStr): HRESULT; stdcall;
    (*
    How to use output parameter (BSTR *password):

    in:  The caller is required to set BSTR value as NULL (no string).
         The callee (in 7-Zip code) ignores the input value stored in BSTR variable,

    out: The callee rewrites BSTR variable (*password) with new allocated string pointer.
         The caller must free BSTR string with function SysFreeString();
    *)
  end;

  ICryptoGetTextPassword2 = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000500110000}']
    function CryptoGetTextPassword2(passwordIsDefined: PInt32; var password: TBStr): HRESULT; stdcall;
    (*
    in:
      The caller is required to set BSTR value as NULL (no string).
      The caller is not required to set (*passwordIsDefined) value.

    out:
      Return code: != S_OK : error code
      Return code:    S_OK : success

      if (*passwordIsDefined == 1), the variable (*password) contains password string

      if (*passwordIsDefined == 0), the password is not defined,
         but the callee still could set (*password) to some allocated string, for example, as empty string.

      The caller must free BSTR string with function SysFreeString()
    *)
  end;

{$ENDREGION}

{$REGION 'IStream.h interfaces ("23170F69-40C1-278A-0000-000300xx0000")'}

//******************************************************************************
// IStream.h
// "23170F69-40C1-278A-0000-000300xx0000"
// (Last checked 14 May 2024; https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/IStream.h)
//******************************************************************************

  ISequentialInStream = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300010000}']
    function Read(data: Pointer; size: UInt32; processedSize: PUInt32): HRESULT; stdcall;
    (*
    The requirement for caller: (processedSize != NULL).
    The callee can allow (processedSize == NULL) for compatibility reasons.

    if (size == 0), this function returns S_OK and (*processedSize) is set to 0.

    if (size != 0)
    {
      Partial read is allowed: (*processedSize <= avail_size && *processedSize <= size),
        where (avail_size) is the size of remaining bytes in stream.
      If (avail_size != 0), this function must read at least 1 byte: (*processedSize > 0).
      You must call Read() in loop, if you need to read exact amount of data.
    }

    If seek pointer before Read() call was changed to position past the end of stream:
      if (seek_pointer >= stream_size), this function returns S_OK and (*processedSize) is set to 0.

    ERROR CASES:
      If the function returns error code, then (*processedSize) is size of
      data written to (data) buffer (it can be data before error or data with errors).
      The recommended way for callee to work with reading errors:
        1) write part of data before error to (data) buffer and return S_OK.
        2) return error code for further calls of Read().
    *)
  end;

  ISequentialOutStream = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300020000}']
    function Write(data: Pointer; size: UInt32; processedSize: PUint32): HRESULT; stdcall;
    (*
    The requirement for caller: (processedSize != NULL).
    The callee can allow (processedSize == NULL) for compatibility reasons.

    if (size != 0)
    {
      Partial write is allowed: (*processedSize <= size),
      but this function must write at least 1 byte: (*processedSize > 0).
      You must call Write() in loop, if you need to write exact amount of data.
    }

    ERROR CASES:
      If the function returns error code, then (*processedSize) is size of
      data written from (data) buffer.
    *)
  end;

  IInStream = interface(ISequentialInStream)
  ['{23170F69-40C1-278A-0000-000300030000}']
    function Seek(offset: Int64; seekOrigin: UInt32; newPosition: PUInt64): HRESULT; stdcall;
    (*
    If you seek to position before the beginning of the stream,
    Seek() function returns error code:
        Recommended error code is __HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK).
        or STG_E_INVALIDFUNCTION
    It is allowed to seek past the end of the stream.
    if Seek() returns error, then the value of *newPosition is undefined.
    *)
  end;

  IOutStream = interface(ISequentialOutStream)
  ['{23170F69-40C1-278A-0000-000300040000}']
    function Seek(offset: Int64; seekOrigin: UInt32; newPosition: PUInt64): HRESULT; stdcall;
    function SetSize(newSize: UInt64): HRESULT; stdcall;
  end;

  IStreamGetSize = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300060000}']
    function GetSize(size: PUInt64): HRESULT; stdcall;
  end;

  IOutStreamFinish = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300070000}']
    function OutStreamFinish: HRESULT; stdcall; // sic! This method is called OutStreamFinish and not Flush!
  end;

  IStreamGetProps = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300080000}']
    function GetProps(size: PUInt64; cTime: PFileTime; aTime: PFileTime;
      mTime: PFileTime; attrib: PUInt32): HRESULT; stdcall;
  end;

{$ENDREGION}

{$REGION 'IArchive.h interfaces ("23170F69-40C1-278A-0000-000600xx0000")'}

//******************************************************************************
// IArchive.h
// (Last checked 14 May 2024; https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/Archive/IArchive.h)
//******************************************************************************

// MIDL_INTERFACE("23170F69-40C1-278A-0000-000600xx0000")
//#define ARCHIVE_INTERFACE_SUB(i, base,  x) \
//DEFINE_GUID(IID_ ## i, \
//0x23170F69, 0x40C1, 0x278A, 0x00, 0x00, 0x00, 0x06, 0x00, x, 0x00, 0x00); \
//struct i: public base

//#define ARCHIVE_INTERFACE(i, x) ARCHIVE_INTERFACE_SUB(i, IUnknown, x)

type
// NFileTimeType
  NFileTimeType = (
    kWindows = 0,
    kUnix,
    kDOS
  );

// NArcInfoFlags
  NArcInfoFlags = (
    aifKeepName        = 1 shl 0,  // keep name of file in archive name
    aifAltStreams      = 1 shl 1,  // the handler supports alt streams
    aifNtSecure        = 1 shl 2,  // the handler supports NT security
    aifFindSignature   = 1 shl 3,  // the handler can find start of archive
    aifMultiSignature  = 1 shl 4,  // there are several signatures
    aifUseGlobalOffset = 1 shl 5,  // the seek position of stream must be set as global offset
    aifStartOpen       = 1 shl 6,  // call handler for each start position
    aifPureStartOpen   = 1 shl 7,  // call handler only for start of file
    aifBackwardOpen    = 1 shl 8,  // archive can be open backward
    aifPreArc          = 1 shl 9,  // such archive can be stored before real archive (like SFX stub)
    aifSymLinks        = 1 shl 10, // the handler supports symbolic links
    aifHardLinks       = 1 shl 11  // the handler supports hard links
  );

// NArchive::NHandlerPropID
  NHandlerPropID = (
    kName = 0,          // VT_BSTR
    kClassID,           // binary GUID in VT_BSTR
    kExtension,         // VT_BSTR
    kAddExtension,      // VT_BSTR
    kUpdate,            // VT_BOOL
    kKeepName,          // VT_BOOL
    kSignature,         // binary in VT_BSTR
    kMultiSignature,    // binary in VT_BSTR
    kSignatureOffset,   // VT_UI4
    kAltStreams,        // VT_BOOL
    kNtSecure,          // VT_BOOL
    kFlags,             // VT_UI4
    kTimeFlags          // VT_UI4
    // kVersion           // VT_UI4 ((VER_MAJOR << 8) | VER_MINOR)
  );

// NArchive::NExtract::NAskMode
  NAskMode = (
    kExtract = 0,
    kTest,
    kSkip,
    kReadExternal
  );

// NArchive::NExtract::NOperationResult
  NExtOperationResult = (
    kOK = 0,
    kUnSupportedMethod,
    kDataError,
    kCRCError,
    kUnavailable,
    kUnexpectedEnd,
    kDataAfterEnd,
    kIsNotArc,
    kHeadersError,
    kWrongPassword
    // , kMemError
  );

// NArchive::NEventIndexType
  NEventIndexType = (
    kNoIndex = 0,
    kInArcIndex,
    kBlockIndex,
    kOutArcIndex
  );

// NArchive::NUpdate::NOperationResult
  NUpdOperationResult = (
    kOK_   = 0,
    kError
    //,kError_FileChanged
  );

  IArchiveOpenCallback = interface
  ['{23170F69-40C1-278A-0000-000600100000}']
    function SetTotal(files, bytes: PUInt64): HRESULT; stdcall;
    function SetCompleted(files, bytes: PUInt64): HRESULT; stdcall;
    (*
    IArchiveExtractCallback::
    7-Zip doesn't call IArchiveExtractCallback functions
      GetStream()
      PrepareOperation()
      SetOperationResult()
    from different threads simultaneously.
    But 7-Zip can call functions for IProgress or ICompressProgressInfo functions
    from another threads simultaneously with calls for IArchiveExtractCallback interface.
    IArchiveExtractCallback::GetStream()
      UInt32 index - index of item in Archive
      Int32 askExtractMode  (Extract::NAskMode)
        if (askMode != NExtract::NAskMode::kExtract)
        {
          then the callee can not real stream: (*inStream == NULL)
        }
      Out:
          (*inStream == NULL) - for directories
          (*inStream == NULL) - if link (hard link or symbolic link) was created
          if (*inStream == NULL && askMode == NExtract::NAskMode::kExtract)
          {
            then the caller must skip extracting of that file.
          }
      returns:
        S_OK     : OK
        S_FALSE  : data error (for decoders)
    if (IProgress::SetTotal() was called)
    {
      IProgress::SetCompleted(completeValue) uses
        packSize   - for some stream formats (xz, gz, bz2, lzma, z, ppmd).
        unpackSize - for another formats.
    }
    else
    {
      IProgress::SetCompleted(completeValue) uses packSize.
    }
    SetOperationResult()
      7-Zip calls SetOperationResult at the end of extracting,
      so the callee can close the file, set attributes, timestamps and security information.
      Int32 opRes (NExtract::NOperationResult)
    *)
  end;

  IArchiveExtractCallback = interface(IProgress)
  ['{23170F69-40C1-278A-0000-000600200000}']
    function GetStream(index: UInt32; var outStream: ISequentialOutStream;
        askExtractMode: NAskMode): HRESULT; stdcall;
    // GetStream OUT: S_OK - OK, S_FALSE - skeep this file
    function PrepareOperation(askExtractMode: NAskMode): HRESULT; stdcall;
    function SetOperationResult(resultEOperationResult: NExtOperationResult): HRESULT; stdcall;
    (*
    IArchiveExtractCallbackMessage can be requested from IArchiveExtractCallback object
      by Extract() or UpdateItems() functions to report about extracting errors
    ReportExtractResult()
      UInt32 indexType (NEventIndexType)
      UInt32 index
      Int32 opRes (NExtract::NOperationResult)
    *)
  end;

  // before v23:
  IArchiveExtractCallbackMessage = interface
  ['{23170F69-40C1-278A-0000-000600210000}']
    function ReportExtractResult(indexType: NEventIndexType; index: UInt32; opRes: Int32): HRESULT; stdcall;
  end;

  IArchiveExtractCallbackMessage2 = interface
  ['{23170F69-40C1-278A-0000-000600220000}']
    // IArchiveExtractCallbackMessage2 can be requested from IArchiveExtractCallback object by Extract() or UpdateItems() functions to report about extracting errors
    function ReportExtractResult(indexType: NEventIndexType; index: UInt32; opRes: Int32): HRESULT; stdcall;
  end;

  IArchiveOpenVolumeCallback = interface
  ['{23170F69-40C1-278A-0000-000600300000}']
    function GetProperty(propID: PROPID; var value: OleVariant): HRESULT; stdcall;
    function GetStream(const name: PWideChar; var inStream: IInStream): HRESULT; stdcall;
  end;

  IInArchiveGetStream = interface
  ['{23170F69-40C1-278A-0000-000600400000}']
    function GetStream(index: UInt32; var stream: ISequentialInStream ): HRESULT; stdcall;
  end;

  IArchiveOpenSetSubArchiveName = interface
  ['{23170F69-40C1-278A-0000-000600500000}']
    function SetSubArchiveName(name: PWideChar): HRESULT; stdcall;
  end;

  IInArchive = interface
  ['{23170F69-40C1-278A-0000-000600600000}']
    function Open(stream: IInStream; const maxCheckStartPosition: PInt64;
        openArchiveCallback: IArchiveOpenCallback): HRESULT; stdcall;
    (*
    IInArchive::Open
        stream
          if (kUseGlobalOffset), stream current position can be non 0.
          if (!kUseGlobalOffset), stream current position is 0.
        if (maxCheckStartPosition == NULL), the handler can try to search archive start in stream
        if (*maxCheckStartPosition == 0), the handler must check only current position as archive start
    IInArchive::Extract:
      indices must be sorted
      numItems = (UInt32)(Int32)-1 = 0xFFFFFFFF means "all files"
      testMode != 0 means "test files without writing to outStream"
    IInArchive::GetArchiveProperty:
      kpidOffset  - start offset of archive.
          VT_EMPTY : means offset = 0.
          VT_UI4, VT_UI8, VT_I8 : result offset; negative values is allowed
      kpidPhySize - size of archive. VT_EMPTY means unknown size.
        kpidPhySize is allowed to be larger than file size. In that case it must show
        supposed size.
      kpidIsDeleted:
      kpidIsAltStream:
      kpidIsAux:
      kpidINode:
        must return VARIANT_TRUE (VT_BOOL), if archive can support that property in GetProperty.
    Notes:
      Don't call IInArchive functions for same IInArchive object from different threads simultaneously.
      Some IInArchive handlers will work incorrectly in that case.
    *)
    function Close: HRESULT; stdcall;
    function GetNumberOfItems(var numItems: UInt32): HRESULT; stdcall;
    function GetProperty(index: UInt32; propID: PROPID; var value: OleVariant): HRESULT; stdcall;
    function Extract(indices: PCardArray; numItems: UInt32;
        testMode: Integer; extractCallback: IArchiveExtractCallback): HRESULT; stdcall;
    // indices must be sorted
    // numItems = 0xFFFFFFFF means all files
    // testMode != 0 means "test files operation"

    function GetArchiveProperty(propID: PROPID; var value: OleVariant): HRESULT; stdcall;

    function GetNumberOfProperties(numProps: PUInt32): HRESULT; stdcall;
    function GetPropertyInfo(index: UInt32;
        name: PBSTR; propID: PPropID; varType: PVarType): HRESULT; stdcall;

    function GetNumberOfArchiveProperties(var numProps: UInt32): HRESULT; stdcall;
    function GetArchivePropertyInfo(index: UInt32;
        name: PBSTR; propID: PPropID; varType: PVARTYPE): HRESULT; stdcall;
  end;

  IArchiveOpenSeq = interface
  ['{23170F69-40C1-278A-0000-000600610000}']
    function OpenSeq(var stream: ISequentialInStream): HRESULT; stdcall;
  end;

  (*
  IArchiveOpen2 = interface
  ['{23170F69-40C1-278A-0000-000600620000}']
    function ArcOpen2(ISequentialInStream *stream, UInt32 flags, IArchiveOpenCallback *openCallback): HRESULT; stdcall;
  end;
  *)

// NParentType::
  NParentType = (
    kDir = 0,
    kAltStream
  );

// NPropDataType::
  NPropDataType = (
    kMask_ZeroEnd   = $10, // 1 shl 4,
    // kMask_BigEndian = 1 shl 5,
    kMask_Utf       = $40, // 1 shl 6,
    kMask_Utf8      = $40, // kMask_Utf or 0,
    kMask_Utf16     = $41, // kMask_Utf or 1,
    // kMask_Utf32 = $42, // kMask_Utf or 2,

    kNotDefined = 0,
    kRaw = 1,

    kUtf8z  = $50, // kMask_Utf8  or kMask_ZeroEnd,
    kUtf16z = $51  // kMask_Utf16 or kMask_ZeroEnd
  );

// NUpdateNotifyOp::
{
  NUpdateNotifyOp = (
    kAdd,
    kUpdate,
    kAnalyze,
    kReplicate,
    kRepack,
    kSkip,
    kDelete,
    kHeader,
    kHashRead,
    kInFileChanged
    // , kOpFinished
    // , kNumDefined
  );
}

  IArchiveGetRawProps = interface
  ['{23170F69-40C1-278A-0000-000600700000}']
    function GetParent(index: UInt32; var parent: UInt32; var parentType: UInt32): HRESULT; stdcall;
    function GetRawProp(index: UInt32; propID: PROPID; var data: Pointer; var dataSize: UInt32; var propType: UInt32): HRESULT; stdcall;
    function GetNumRawProps(var numProps: UInt32): HRESULT; stdcall;
    function GetRawPropInfo(index: UInt32; name: PBSTR; var propID: PROPID): HRESULT; stdcall;
  end;

  IArchiveGetRootProps = interface
  ['{23170F69-40C1-278A-0000-000600710000}']
    function GetRootProp(propID: PROPID; var value: PROPVARIANT): HRESULT; stdcall;
    function GetRootRawProp(propID: PROPID; var data: Pointer; var dataSize: UINT32; var propType: UInt32): HRESULT; stdcall;
  end;

  IArchiveUpdateCallback = interface(IProgress)
  ['{23170F69-40C1-278A-0000-000600800000}']
    function GetUpdateItemInfo(index: UInt32;
        newData: PInt32; // 1 - new data, 0 - old data
        newProps: PInt32; // 1 - new properties, 0 - old properties
        indexInArchive: PUInt32 // -1 if there is no in archive, or if doesn't matter
        ): HRESULT; stdcall;
    function GetProperty(index: UInt32; propID: PROPID; var value: OleVariant): HRESULT; stdcall;
    function GetStream(index: UInt32; var inStream: ISequentialInStream): HRESULT; stdcall;
    function SetOperationResult(operationResult: Int32): HRESULT; stdcall;
  end;

  IArchiveUpdateCallback2 = interface(IArchiveUpdateCallback)
  ['{23170F69-40C1-278A-0000-000600820000}']
    function GetVolumeSize(index: UInt32; size: PUInt64): HRESULT; stdcall;
    function GetVolumeStream(index: UInt32; var volumeStream: ISequentialOutStream): HRESULT; stdcall;
  end;

  (*
  IArchiveUpdateCallbackFile = interface
  ['{23170F69-40C1-278A-0000-000600830000}']
    function GetStream2(UInt32 index, ISequentialInStream **inStream, UInt32 notifyOp)): HRESULT; stdcall;
    function ReportOperation(UInt32 indexType, UInt32 index, UInt32 notifyOp)): HRESULT; stdcall;
  end;

  IArchiveGetDiskProperty = interface
  ['{23170F69-40C1-278A-0000-000600840000}']
    function GetDiskProperty(UInt32 index, PROPID propID, PROPVARIANT *value)): HRESULT; stdcall;
  end;

  IArchiveUpdateCallbackArcProp = interface
  ['{23170F69-40C1-278A-0000-000600850000}']
    function ReportProp(UInt32 indexType, UInt32 index, PROPID propID, const PROPVARIANT *value)): HRESULT; stdcall;
    function ReportRawProp(UInt32 indexType, UInt32 index, PROPID propID, const void *data, UInt32 dataSize, UInt32 propType)): HRESULT; stdcall;
    function ReportFinished(UInt32 indexType, UInt32 index, Int32 opRes)): HRESULT; stdcall;
    function DoNeedArcProp(PROPID propID, Int32 *answer)): HRESULT; stdcall;
  end;
  *)

  IOutArchive = interface
  ['{23170F69-40C1-278A-0000-000600A00000}']
    function UpdateItems(outStream: ISequentialOutStream; numItems: UInt32;
      updateCallback: IArchiveUpdateCallback): HRESULT; stdcall;
    function GetFileTimeType(type_: PUint32): HRESULT; stdcall;
  end;

  ISetProperties = interface
  ['{23170F69-40C1-278A-0000-000600030000}']
    function SetProperties(names: PPWideChar; values: PPROPVARIANT; numProps: UInt32): HRESULT; stdcall;
  end;

  IArchiveKeepModeForNextOpen = interface
  ['{23170F69-40C1-278A-0000-000600040000}']
    function KeepModeForNextOpen: HRESULT; stdcall;
  end;

  IArchiveAllowTail = interface
  ['{23170F69-40C1-278A-0000-000600050000}']
    function AllowTail(allowTail: Int32): HRESULT; stdcall;
  end;

{$ENDREGION}

{$REGION 'ICoder.h interfaces ("23170F69-40C1-278A-0000-000400xx0000")'}

//******************************************************************************
// ICoder.h
// "23170F69-40C1-278A-0000-000400xx0000"
// (Last checked 14 May 2024; https://github.com/mcmilk/7-Zip/blob/master/CPP/7zip/ICoder.h)
//******************************************************************************

  ICompressProgressInfo = interface
  ['{23170F69-40C1-278A-0000-000400040000}']
    function SetRatioInfo(inSize, outSize: PUInt64): HRESULT; stdcall;
  end;

  ICompressCoder = interface
  ['{23170F69-40C1-278A-0000-000400050000}']
  function Code(inStream, outStream: ISequentialInStream;
      inSize, outSize: PUInt64;
      progress: ICompressProgressInfo): HRESULT; stdcall;
  end;

  ICompressCoder2 = interface
  ['{23170F69-40C1-278A-0000-000400180000}']
  function Code(var inStreams: ISequentialInStream;
      var inSizes: PUInt64;
      numInStreams: UInt32;
      var outStreams: ISequentialOutStream;
      var outSizes: PUInt64;
      numOutStreams: UInt32;
      progress: ICompressProgressInfo): HRESULT; stdcall;
  end;

//NCoderPropID::
  NCoderPropID = (
    kDefaultProp = 0,
    kDictionarySize,
    kUsedMemorySize,
    kOrder,
    kBlockSize,
    kPosStateBits,
    kLitContextBits,
    kLitPosBits,
    kNumFastBytes,
    kMatchFinder,
    kMatchFinderCycles,
    kNumPasses,
    kAlgorithm,
    kNumThreads,
    kEndMarker,
    kLevel,
    kReduceSize // estimated size of data that will be compressed. Encoder can use this value to reduce dictionary size.
  );


type
  ICompressSetCoderPropertiesOpt = interface
  ['{23170F69-40C1-278A-0000-0004001F0000}']
    function SetCoderPropertiesOpt(propIDs: PPropID; props: PROPVARIANT; numProps: UInt32): HRESULT; stdcall;
  end;

  ICompressSetCoderProperties = interface
  ['{23170F69-40C1-278A-0000-000400200000}']
    function SetCoderProperties(propIDs: PPropID; props: PROPVARIANT; numProps: UInt32): HRESULT; stdcall;
  end;

  // name conflict?! Is commented out in https://github.com/keithjjones/7z/blob/master/CPP/7zip/ICoder.h
  (*
  ICompressSetCoderProperties = interface
  ['{23170F69-40C1-278A-0000-000400210000}']
    procedure SetDecoderProperties(inStream: ISequentialInStream);
  end;
  *)

  ICompressSetDecoderProperties2 = interface
  ['{23170F69-40C1-278A-0000-000400220000}']
    function SetDecoderProperties2(data: PByte; size: UInt32): HRESULT; stdcall;
    (*
    S_OK
    E_NOTIMP      : unsupported properties
    E_INVALIDARG  : incorrect (or unsupported) properties
    E_OUTOFMEMORY : memory allocation error
    *)
  end;

  ICompressWriteCoderProperties = interface
  ['{23170F69-40C1-278A-0000-000400230000}']
    function WriteCoderProperties(outStreams: ISequentialOutStream): HRESULT; stdcall;
  end;

  ICompressGetInStreamProcessedSize = interface
  ['{23170F69-40C1-278A-0000-000400240000}']
    function GetInStreamProcessedSize(value: PUInt64): HRESULT; stdcall;
  end;

  ICompressSetCoderMt = interface
  ['{23170F69-40C1-278A-0000-000400250000}']
    function SetNumberOfThreads(numThreads: UInt32): HRESULT; stdcall;
  end;

  ICompressSetFinishMode = interface
  ['{23170F69-40C1-278A-0000-000400260000}']
    function SetFinishMode(finishMode: UInt32): HRESULT; stdcall;
    (* finishMode:
    0 : partial decoding is allowed. It's default mode for ICompressCoder::Code(), if (outSize) is defined.
    1 : full decoding. The stream must be finished at the end of decoding.
    *)
  end;

  ICompressGetInStreamProcessedSize2 = interface
  ['{23170F69-40C1-278A-0000-000400270000}']
    function GetInStreamProcessedSize2(streamIndex: UInt32; value: PUInt64): HRESULT; stdcall;
  end;

  ICompressSetMemLimit = interface
  ['{23170F69-40C1-278A-0000-000400280000}']
    function SetMemLimit(memUsage: UInt64): HRESULT; stdcall;
  end;

  ICompressReadUnusedFromInBuf = interface
  ['{23170F69-40C1-278A-0000-000400290000}']
    function ReadUnusedFromInBuf(data: Pointer; size: UInt32; processedSize: PUInt32): HRESULT; stdcall;
    (*
    ICompressReadUnusedFromInBuf is supported by ICoder object
    call ReadUnusedFromInBuf() after ICoder::Code(inStream, ...).
    ICoder::Code(inStream, ...) decodes data, and the ICoder object is allowed
    to read from inStream to internal buffers more data than minimal data required for decoding.
    So we can call ReadUnusedFromInBuf() from same ICoder object to read unused input
    data from the internal buffer.
    in ReadUnusedFromInBuf(): the Coder is not allowed to use (ISequentialInStream *inStream) object, that was sent to ICoder::Code().
    *)
  end;

  ICompressGetSubStreamSize = interface
  ['{23170F69-40C1-278A-0000-000400300000}']
    function GetSubStreamSize(subStream: UInt64; value: PUInt64): HRESULT; stdcall;
  end;

  ICompressSetInStream = interface
  ['{23170F69-40C1-278A-0000-000400310000}']
    function SetInStream(inStream: ISequentialInStream): HRESULT; stdcall;
    function ReleaseInStream: HRESULT; stdcall;
  end;

  ICompressSetOutStream = interface
  ['{23170F69-40C1-278A-0000-000400320000}']
    function SetOutStream(outStream: ISequentialOutStream): HRESULT; stdcall;
    function ReleaseOutStream: HRESULT; stdcall;
  end;

  ICompressSetInStreamSize = interface
  ['{23170F69-40C1-278A-0000-000400330000}']
    function SetInStreamSize(inSize: PUInt64): HRESULT; stdcall;
  end;

  ICompressSetOutStreamSize = interface
  ['{23170F69-40C1-278A-0000-000400340000}']
    function SetOutStreamSize(outSize: PUInt64): HRESULT; stdcall;
  end;

  ICompressSetBufSize = interface
  ['{23170F69-40C1-278A-0000-000400350000}']
    function SetInBufSize(streamIndex: UInt32; size: UInt32): HRESULT; stdcall;
    function SetOutBufSize(streamIndex: UInt32; size: UInt32): HRESULT; stdcall;
  end;

  ICompressInitEncoder = interface
  ['{23170F69-40C1-278A-0000-000400360000}']
    function InitEncoder: HRESULT; stdcall;
    (*
    That function initializes encoder structures.
    Call this function only for stream version of encoder.
    *)
  end;

  ICompressSetInStream2 = interface
  ['{23170F69-40C1-278A-0000-000400370000}']
   function SetInStream2(streamIndex: UInt32; inStream: ISequentialInStream): HRESULT; stdcall;
   function ReleaseInStream2(streamIndex: UInt32): HRESULT; stdcall;
  end;

  ICompressSetOutStream2 = interface
  ['{23170F69-40C1-278A-0000-000400380000}']
   function SetOutStream2(streamIndex: UInt32; outStream: ISequentialOutStream): HRESULT; stdcall;
   function ReleaseOutStream2(streamIndex: UInt32): HRESULT; stdcall;
  end;

  ICompressSetInStreamSize2 = interface
  ['{23170F69-40C1-278A-0000-000400390000}']
   function SetInStreamSize2(streamIndex: UInt32; inSize: PUInt64): HRESULT; stdcall;
  end;

  ICompressInSubStreams = interface
  ['{23170F69-40C1-278A-0000-0004003A0000}']
   function GetNextInSubStream(streamIndexRes: PUInt64; var stream: ISequentialInStream): HRESULT; stdcall;
  end;

  ICompressOutSubStreams = interface
  ['{23170F69-40C1-278A-0000-0004003B0000}']
   function GetNextOutSubStream(streamIndexRes: PUInt64; var stream: ISequentialOutStream): HRESULT; stdcall;
  end;

  (*
  ICompressFilter
  Filter(Byte *data, UInt32 size)
  (size)
     converts as most as possible bytes required for fast processing.
     Some filters have (smallest_fast_block).
     For example, (smallest_fast_block == 16) for AES CBC/CTR filters.
     If data stream is not finished, caller must call Filter() for larger block:
     where (size >= smallest_fast_block).
     if (size >= smallest_fast_block)
     {
       The filter can leave some bytes at the end of data without conversion:
       if there are data alignment reasons or speed reasons.
       The caller can read additional data from stream and call Filter() again.
     }
     If data stream was finished, caller can call Filter() for (size < smallest_fast_block)

  (data) parameter:
     Some filters require alignment for any Filter() call:
        1) (stream_offset % alignment_size) == (data % alignment_size)
        2) (alignment_size == 2^N)
     where (stream_offset) - is the number of bytes that were already filtered before.
     The callers of Filter() are required to meet these requirements.
     (alignment_size) can be different:
           16 : for AES filters
       4 or 2 : for some branch convert filters
            1 : for another filters
     (alignment_size >= 16) is enough for all current filters of 7-Zip.
     But the caller can use larger (alignment_size).
     Recommended alignment for (data) of Filter() call is (alignment_size == 64).
     Also it's recommended to use aligned value for (size):
       (size % alignment_size == 0),
     if it's not last call of Filter() for current stream.

  returns: (outSize):
       if (outSize == 0) : Filter have not converted anything.
           So the caller can stop processing, if data stream was finished.
       if (outSize <= size) : Filter have converted outSize bytes
       if (outSize >  size) : Filter have not converted anything.
           and it needs at least outSize bytes to convert one block
           (it's for crypto block algorithms).
*/
*)

  ICompressFilter = interface
  ['{23170F69-40C1-278A-0000-000400400000}']
    function Init: HRESULT; stdcall;
    function Filter(data: PByte; size: UInt32): UInt32; stdcall;
    // Filter return outSize (UInt32)
    // if (outSize <= size): Filter have converted outSize bytes
    // if (outSize > size): Filter have not converted anything.
    //      and it needs at least outSize bytes to convert one block
    //      (it's for crypto block algorithms).
  end;

  ICompressCodecsInfo = interface
  ['{23170F69-40C1-278A-0000-000400600000}']
    function GetNumMethods(numMethods: PUInt32): HRESULT; stdcall;
    function GetProperty(index: UInt32; propID: PROPID; var value: PROPVARIANT): HRESULT; stdcall;
    function CreateDecoder(index: UInt32; const iid: PGUID; var coder: Pointer): HRESULT; stdcall;
    function CreateEncoder(index: UInt32; const iid: PGUID; var coder: Pointer): HRESULT; stdcall;
  end;

  ISetCompressCodecsInfo = interface
  ['{23170F69-40C1-278A-0000-000400610000}']
    function SetCompressCodecsInfo(compressCodecsInfo: ICompressCodecsInfo): HRESULT; stdcall;
  end;

  ICryptoProperties = interface
  ['{23170F69-40C1-278A-0000-000400800000}']
    function SetKey(Data: PByte; size: UInt32): HRESULT; stdcall;
    function SetInitVector(data: PByte; size: UInt32): HRESULT; stdcall;
  end;

  ICryptoResetSalt = interface
  ['{23170F69-40C1-278A-0000-000400880000}']
    function ResetSalt: HRESULT; stdcall;
  end;

  ICryptoResetInitVector = interface
  ['{23170F69-40C1-278A-0000-0004008C0000}']
    function ResetInitVector: HRESULT; stdcall;
    (*
    Call ResetInitVector() only for encoding.
    Call ResetInitVector() before encoding and before WriteCoderProperties().
    Crypto encoder can create random IV in that function.
    *)
  end;

  ICryptoSetPassword = interface
  ['{23170F69-40C1-278A-0000-000400900000}']
    function CryptoSetPassword(data: PByte; size: UInt32): HRESULT; stdcall;
  end;

  ICryptoSetCRC = interface
  ['{23170F69-40C1-278A-0000-000400A00000}']
    function CryptoSetCRC(crc: UInt32): HRESULT; stdcall;
  end;

  IHasher = interface
  ['{23170F69-40C1-278A-0000-000400C00000}']
    procedure Init; stdcall;
    procedure Update(const data: Pointer; size: UInt32); stdcall;
    procedure Final(digest: PByte); stdcall;
    function GetDigestSize: UInt32; stdcall;
  end;

  IHashers = interface
  ['{23170F69-40C1-278A-0000-000400C10000}']
    function GetNumHashers: UInt32;
    function GetHasherProp(index: UInt32; propID: PROPID; var value: PROPVARIANT): HRESULT; stdcall;
    function CreateHasher(index: UInt32; var hasher: IHasher): HRESULT; stdcall;
  end;

{$ENDREGION}

{$REGION 'Classes'}

//******************************************************************************
// CLASSES
//******************************************************************************

  T7zPasswordCallback = function(sender: Pointer; var password: UnicodeString): HRESULT; stdcall;
  T7zGetStreamCallBack = function(sender: Pointer; index: UInt32;
    var outStream: ISequentialOutStream): HRESULT; stdcall;
  T7zProgressCallback = function(sender: Pointer; total: boolean; value: int64): HRESULT; stdcall;

  // Note: This is "our" interface, not an COM interface from 7z.dll
  I7zInArchive = interface
    procedure OpenFile(const filename: string); stdcall;
    procedure OpenStream(stream: IInStream); stdcall;
    procedure Close; stdcall;
    function GetNumberOfItems: UInt32; stdcall;
    function GetItemPath(const index: integer): UnicodeString; stdcall;
    function GetItemName(const index: integer): UnicodeString; stdcall;
    function GetItemSize(const index: integer): Int64; stdcall;
    function GetItemCompressedSize(const index: integer): Int64; stdcall;
    function GetItemWriteTime(const index: integer): TDateTime; stdcall;
    function GetItemAttributes(const index: integer): DWORD; stdcall;
    function GetItemIsFolder(const index: integer): boolean; stdcall;
    function GetInArchive: IInArchive;
    procedure ExtractItem(const item: UInt32; Stream: TStream; test: longbool); stdcall;
    procedure ExtractItems(items: PCardArray; count: UInt32; test: longbool;
      sender: pointer; callback: T7zGetStreamCallBack); stdcall;
    procedure ExtractAll(test: longbool; sender: pointer; callback: T7zGetStreamCallBack); stdcall;
    procedure ExtractTo(const path: string); stdcall;
    procedure SetPasswordCallback(sender: Pointer; callback: T7zPasswordCallback); stdcall;
    procedure SetPassword(const password: UnicodeString); stdcall;
    procedure SetProgressCallback(sender: Pointer; callback: T7zProgressCallback); stdcall;
    procedure SetClassId(const classid: TGUID);
    function GetClassId: TGUID;
    property ClassId: TGUID read GetClassId write SetClassId;
    property NumberOfItems: UInt32 read GetNumberOfItems;
    property ItemPath[const index: integer]: UnicodeString read GetItemPath;
    property ItemName[const index: integer]: UnicodeString read GetItemName;
    property ItemSize[const index: integer]: Int64 read GetItemSize;
    property ItemCompressedSize[const index: integer]: Int64 read GetItemCompressedSize;
    property ItemIsFolder[const index: integer]: boolean read GetItemIsFolder;
    property InArchive: IInArchive read GetInArchive;
  end;

  // Note: This is "our" interface, not an COM interface from 7z.dll
  I7zOutArchive = interface
    procedure AddStream(Stream: TStream; Ownership: TStreamOwnership; Attributes: Cardinal;
      CreationTime, LastWriteTime: TFileTime; const Path: UnicodeString;
      IsFolder, IsAnti: boolean); stdcall;
    procedure AddFile(const Filename: TFileName; const Path: UnicodeString); stdcall;
    procedure AddFiles(const Dir, Path, Wildcard: string; recurse: boolean); stdcall;
    procedure SaveToFile(const FileName: TFileName); stdcall;
    procedure SaveToStream(stream: TStream); stdcall;
    procedure SetProgressCallback(sender: Pointer; callback: T7zProgressCallback); stdcall;
    procedure ClearBatch; stdcall;
    procedure SetPassword(const password: UnicodeString); stdcall;
    procedure SetPropertie(name: UnicodeString; value: OleVariant); stdcall;
    procedure SetClassId(const classid: TGUID);
    function GetClassId: TGUID;
    property ClassId: TGUID read GetClassId write SetClassId;
  end;

  // Note: This is "our" interface, not an COM interface from 7z.dll
  I7zCodec = interface

  end;


  T7zStream = class(TInterfacedObject, IInStream, IStreamGetSize,
    ISequentialOutStream, ISequentialInStream, IOutStream, IOutStreamFinish)
  private
    FStream: TStream;
    FOwnership: TStreamOwnership;
    FFileName: string;
    FWriteTime: TDateTime;
  protected
    function Read(data: Pointer; size: UInt32; processedSize: PUInt32): HRESULT; stdcall;
    function Seek(offset: Int64; seekOrigin: UInt32; newPosition: PUInt64): HRESULT; stdcall;
    function GetSize(size: PUInt64): HRESULT; stdcall;
    function SetSize(newSize: UInt64): HRESULT; stdcall;
    function Write(data: Pointer; size: UInt32; processedSize: PUInt32): HRESULT; stdcall;
    function OutStreamFinish: HRESULT; stdcall;
  public
    constructor Create(Stream: TStream; Ownership: TStreamOwnership = soReference; filename: string=''; writeTime: TDateTime=0);
    destructor Destroy; override;
  end;

  // I7zOutArchive property setters
type
  TZipCompressionMethod = (mzCopy, mzDeflate, mzDeflate64, mzBZip2, mzLZMA, mzPPMD);
  TZipEncryptionMethod = (emAES128, emAES192, emAES256, emZIPCRYPTO);
  T7zCompressionMethod = (m7Copy, m7LZMA, m7BZip2, m7PPMd, m7Deflate, m7Deflate64, m7LZMA2);

{$ENDREGION}

{$REGION 'Methods'}

  procedure SetCompressionLevel(Arch: I7zOutArchive; level: Cardinal);                        //   X   X   X   X
  procedure SetMultiThreading(Arch: I7zOutArchive; ThreadCount: Cardinal);                    //   X   X       X

  procedure SetCompressionMethod(Arch: I7zOutArchive; method: TZipCompressionMethod);         //   X
  procedure SetEncryptionMethod(Arch: I7zOutArchive; method: TZipEncryptionMethod);           //   X
  procedure SetDictionnarySize(Arch: I7zOutArchive; size: Cardinal); // < 32                  //   X           X
  procedure SetMemorySize(Arch: I7zOutArchive; size: Cardinal);                               //   X
  procedure SetDeflateNumPasses(Arch: I7zOutArchive; pass: Cardinal);                         //   X       X   X
  procedure SetNumFastBytes(Arch: I7zOutArchive; fb: Cardinal);                               //   X       X
  procedure SetNumMatchFinderCycles(Arch: I7zOutArchive; mc: Cardinal);                       //   X       X


  procedure SevenZipSetCompressionMethod(Arch: I7zOutArchive; method: T7zCompressionMethod);  //       X
  procedure SevenZipSetBindInfo(Arch: I7zOutArchive; const bind: UnicodeString);              //       X
  procedure SevenZipSetSolidSettings(Arch: I7zOutArchive; solid: boolean);                    //       X
  procedure SevenZipRemoveSfxBlock(Arch: I7zOutArchive; remove: boolean);                     //       X
  procedure SevenZipAutoFilter(Arch: I7zOutArchive; auto: boolean);                           //       X
  procedure SevenZipCompressHeaders(Arch: I7zOutArchive; compress: boolean);                  //       X
  procedure SevenZipCompressHeadersFull(Arch: I7zOutArchive; compress: boolean);              //       X
  procedure SevenZipEncryptHeaders(Arch: I7zOutArchive; Encrypt: boolean);                    //       X
  procedure SevenZipVolumeMode(Arch: I7zOutArchive; Mode: boolean);                           //       X

  // filetime util functions
  function DateTimeToFileTime(dt: TDateTime): TFileTime;
  function FileTimeToDateTime(ft: TFileTime): TDateTime;
  function CurrentFileTime: TFileTime;

  // constructors

  // Please note: Use 7z.dll (taken from the 32/64 bit package of 7zip), not 7za.dll from the extras package!
  function CreateInArchive(const classid: TGUID; const lib: string = '7z.dll'): I7zInArchive;
  function CreateOutArchive(const classid: TGUID; const lib: string = '7z.dll'): I7zOutArchive;

{$ENDREGION}

{$REGION 'Handler/Format GUIDs ("23170F69-40C1-278A-1000-000110xx0000")'}

const
  CLSID_CFormatZip      : TGUID = '{23170F69-40C1-278A-1000-000110010000}'; // [OUT] zip jar xpi
  CLSID_CFormatBZ2      : TGUID = '{23170F69-40C1-278A-1000-000110020000}'; // [OUT] bz2 bzip2 tbz2 tbz
  CLSID_CFormatRar      : TGUID = '{23170F69-40C1-278A-1000-000110030000}'; // [IN ] rar r00
  CLSID_CFormatArj      : TGUID = '{23170F69-40C1-278A-1000-000110040000}'; // [IN ] arj
  CLSID_CFormatZ        : TGUID = '{23170F69-40C1-278A-1000-000110050000}'; // [IN ] z taz
  CLSID_CFormatLzh      : TGUID = '{23170F69-40C1-278A-1000-000110060000}'; // [IN ] lzh lha
  CLSID_CFormat7z       : TGUID = '{23170F69-40C1-278A-1000-000110070000}'; // [OUT] 7z
  CLSID_CFormatCab      : TGUID = '{23170F69-40C1-278A-1000-000110080000}'; // [IN ] cab
  CLSID_CFormatNsis     : TGUID = '{23170F69-40C1-278A-1000-000110090000}'; // [IN ] nsis
  CLSID_CFormatLzma     : TGUID = '{23170F69-40C1-278A-1000-0001100A0000}'; // [IN ] lzma
  CLSID_CFormatLzma86   : TGUID = '{23170F69-40C1-278A-1000-0001100B0000}'; // [IN ] lzma 86
  CLSID_CFormatXz       : TGUID = '{23170F69-40C1-278A-1000-0001100C0000}'; // [OUT] xz
  CLSID_CFormatPpmd     : TGUID = '{23170F69-40C1-278A-1000-0001100D0000}'; // [IN ] ppmd

  CLSID_CFormatAVB      : TGUID = '{23170F69-40C1-278A-1000-000110C00000}';
  CLSID_CFormatLP       : TGUID = '{23170F69-40C1-278A-1000-000110C10000}';
  CLSID_CFormatSparse   : TGUID = '{23170F69-40C1-278A-1000-000110C20000}';
  CLSID_CFormatAPFS     : TGUID = '{23170F69-40C1-278A-1000-000110C30000}';
  CLSID_CFormatVhdx     : TGUID = '{23170F69-40C1-278A-1000-000110C40000}';
  CLSID_CFormatBase64   : TGUID = '{23170F69-40C1-278A-1000-000110C50000}';
  CLSID_CFormatCOFF     : TGUID = '{23170F69-40C1-278A-1000-000110C60000}';

  CLSID_CFormatExt      : TGUID = '{23170F69-40C1-278A-1000-000110C70000}'; // [IN ] ext
  CLSID_CFormatVMDK     : TGUID = '{23170F69-40C1-278A-1000-000110C80000}'; // [IN ] vmdk
  CLSID_CFormatVDI      : TGUID = '{23170F69-40C1-278A-1000-000110C90000}'; // [IN ] vdi
  CLSID_CFormatQcow     : TGUID = '{23170F69-40C1-278A-1000-000110CA0000}'; // [IN ] qcow
  CLSID_CFormatGPT      : TGUID = '{23170F69-40C1-278A-1000-000110CB0000}'; // [IN ] GPT
  CLSID_CFormatRar5     : TGUID = '{23170F69-40C1-278A-1000-000110CC0000}'; // [IN ] Rar5
  CLSID_CFormatIHex     : TGUID = '{23170F69-40C1-278A-1000-000110CD0000}'; // [IN ] IHex
  CLSID_CFormatHxs      : TGUID = '{23170F69-40C1-278A-1000-000110CE0000}'; // [IN ] Hxs
  CLSID_CFormatTE       : TGUID = '{23170F69-40C1-278A-1000-000110CF0000}'; // [IN ] TE
  CLSID_CFormatUEFIc    : TGUID = '{23170F69-40C1-278A-1000-000110D00000}'; // [IN ] UEFIc
  CLSID_CFormatUEFIs    : TGUID = '{23170F69-40C1-278A-1000-000110D10000}'; // [IN ] UEFIs
  CLSID_CFormatSquashFS : TGUID = '{23170F69-40C1-278A-1000-000110D20000}'; // [IN ] SquashFS
  CLSID_CFormatCramFS   : TGUID = '{23170F69-40C1-278A-1000-000110D30000}'; // [IN ] CramFS
  CLSID_CFormatAPM      : TGUID = '{23170F69-40C1-278A-1000-000110D40000}'; // [IN ] APM
  CLSID_CFormatMslz     : TGUID = '{23170F69-40C1-278A-1000-000110D50000}'; // [IN ] MsLZ
  CLSID_CFormatFlv      : TGUID = '{23170F69-40C1-278A-1000-000110D60000}'; // [IN ] FLV
  CLSID_CFormatSwf      : TGUID = '{23170F69-40C1-278A-1000-000110D70000}'; // [IN ] SWF
  CLSID_CFormatSwfc     : TGUID = '{23170F69-40C1-278A-1000-000110D80000}'; // [IN ] SWFC
  CLSID_CFormatNtfs     : TGUID = '{23170F69-40C1-278A-1000-000110D90000}'; // [IN ] NTFS
  CLSID_CFormatFat      : TGUID = '{23170F69-40C1-278A-1000-000110DA0000}'; // [IN ] FAT
  CLSID_CFormatMbr      : TGUID = '{23170F69-40C1-278A-1000-000110DB0000}'; // [IN ] MBR
  CLSID_CFormatVhd      : TGUID = '{23170F69-40C1-278A-1000-000110DC0000}'; // [IN ] VHD
  CLSID_CFormatPe       : TGUID = '{23170F69-40C1-278A-1000-000110DD0000}'; // [IN ] PE (Windows Exe)
  CLSID_CFormatElf      : TGUID = '{23170F69-40C1-278A-1000-000110DE0000}'; // [IN ] ELF (Linux Exe)
  CLSID_CFormatMachO    : TGUID = '{23170F69-40C1-278A-1000-000110DF0000}'; // [IN ] Mach-O
  CLSID_CFormatUdf      : TGUID = '{23170F69-40C1-278A-1000-000110E00000}'; // [IN ] iso
  CLSID_CFormatXar      : TGUID = '{23170F69-40C1-278A-1000-000110E10000}'; // [IN ] xar
  CLSID_CFormatMub      : TGUID = '{23170F69-40C1-278A-1000-000110E20000}'; // [IN ] mub
  CLSID_CFormatHfs      : TGUID = '{23170F69-40C1-278A-1000-000110E30000}'; // [IN ] HFS
  CLSID_CFormatDmg      : TGUID = '{23170F69-40C1-278A-1000-000110E40000}'; // [IN ] dmg
  CLSID_CFormatCompound : TGUID = '{23170F69-40C1-278A-1000-000110E50000}'; // [IN ] msi doc xls ppt
  CLSID_CFormatWim      : TGUID = '{23170F69-40C1-278A-1000-000110E60000}'; // [OUT] wim swm
  CLSID_CFormatIso      : TGUID = '{23170F69-40C1-278A-1000-000110E70000}'; // [IN ] iso
  CLSID_CFormatBkf      : TGUID = '{23170F69-40C1-278A-1000-000110E80000}'; // [IN ] BKF
  CLSID_CFormatChm      : TGUID = '{23170F69-40C1-278A-1000-000110E90000}'; // [IN ] chm chi chq chw hxs hxi hxr hxq hxw lit
  CLSID_CFormatSplit    : TGUID = '{23170F69-40C1-278A-1000-000110EA0000}'; // [IN ] 001
  CLSID_CFormatRpm      : TGUID = '{23170F69-40C1-278A-1000-000110EB0000}'; // [IN ] rpm
  CLSID_CFormatDeb      : TGUID = '{23170F69-40C1-278A-1000-000110EC0000}'; // [IN ] deb
  CLSID_CFormatCpio     : TGUID = '{23170F69-40C1-278A-1000-000110ED0000}'; // [IN ] cpio
  CLSID_CFormatTar      : TGUID = '{23170F69-40C1-278A-1000-000110EE0000}'; // [OUT] tar
  CLSID_CFormatGZip     : TGUID = '{23170F69-40C1-278A-1000-000110EF0000}'; // [OUT] gz gzip tgz tpz

{$ENDREGION}

implementation

const
  MAXCHECK : int64 = (1 shl 20);
  ZipCompressionMethod: array[TZipCompressionMethod] of UnicodeString = ('COPY', 'DEFLATE', 'DEFLATE64', 'BZIP2', 'LZMA', 'PPMD');
  ZipEncryptionMethod: array[TZipEncryptionMethod] of UnicodeString = ('AES128', 'AES192', 'AES256', 'ZIPCRYPTO');

  // TODO: DEFLATE, DEFLATE64 are not listed in https://www.7-zip.org/7z.html . Is it really supported?
  SevCompressionMethod: array[T7zCompressionMethod] of UnicodeString = ('COPY', 'LZMA', 'BZIP2', 'PPMD', 'DEFLATE', 'DEFLATE64', 'LZMA2');

function DateTimeToFileTime(dt: TDateTime): TFileTime;
var
  st: TSystemTime;
begin
  DateTimeToSystemTime(dt, st);
  if not (SystemTimeToFileTime(st, Result) and LocalFileTimeToFileTime(Result, Result))
    then RaiseLastOSError;
end;

function FileTimeToDateTime(ft: TFileTime): TDateTime;
var
  st: TSystemTime;
begin
  if not (FileTimeToLocalFileTime(ft, ft) and FileTimeToSystemTime(ft, st)) then
    RaiseLastOSError;
  Result := SystemTimeToDateTime(st);
end;

function CurrentFileTime: TFileTime;
begin
  GetSystemTimeAsFileTime(Result);
end;

procedure RINOK(const hr: HRESULT);
begin
  if hr <> S_OK then
    raise Exception.Create(SysErrorMessage(Cardinal(hr)));
end;

procedure SetCardinalProperty(arch: I7zOutArchive; const name: UnicodeString; card: Cardinal);
var
  value: OleVariant;
begin
  TPropVariant(value).vt := VT_UI4;
  TPropVariant(value).ulVal := card;
  arch.SetPropertie(name, value);
end;

procedure SetBooleanProperty(arch: I7zOutArchive; const name: UnicodeString; bool: boolean);
begin
  case bool of
    true: arch.SetPropertie(name, 'ON');
    false: arch.SetPropertie(name, 'OFF');
  end;
end;

procedure SetCompressionLevel(Arch: I7zOutArchive; level: Cardinal);
begin
  SetCardinalProperty(arch, 'X', level);
end;

procedure SetMultiThreading(Arch: I7zOutArchive; ThreadCount: Cardinal);
begin
  SetCardinalProperty(arch, 'MT', ThreadCount);
end;

procedure SetCompressionMethod(Arch: I7zOutArchive; method: TZipCompressionMethod);
begin
  Arch.SetPropertie('M', ZipCompressionMethod[method]);
end;

procedure SetEncryptionMethod(Arch: I7zOutArchive; method: TZipEncryptionMethod);
begin
  Arch.SetPropertie('EM', ZipEncryptionMethod[method]);
end;

procedure SetDictionnarySize(Arch: I7zOutArchive; size: Cardinal);
begin
  SetCardinalProperty(arch, 'D', size);
end;

procedure SetMemorySize(Arch: I7zOutArchive; size: Cardinal);
begin
  SetCardinalProperty(arch, 'MEM', size);
end;

procedure SetDeflateNumPasses(Arch: I7zOutArchive; pass: Cardinal);
begin
  SetCardinalProperty(arch, 'PASS', pass);
end;

procedure SetNumFastBytes(Arch: I7zOutArchive; fb: Cardinal);
begin
  SetCardinalProperty(arch, 'FB', fb);
end;

procedure SetNumMatchFinderCycles(Arch: I7zOutArchive; mc: Cardinal);
begin
  SetCardinalProperty(arch, 'MC', mc);
end;

procedure SevenZipSetCompressionMethod(Arch: I7zOutArchive; method: T7zCompressionMethod);
begin
  Arch.SetPropertie('0', SevCompressionMethod[method]);
end;

procedure SevenZipSetBindInfo(Arch: I7zOutArchive; const bind: UnicodeString);
begin
  arch.SetPropertie('B', bind);
end;

procedure SevenZipSetSolidSettings(Arch: I7zOutArchive; solid: boolean);
begin
  SetBooleanProperty(Arch, 'S', solid);
end;

procedure SevenZipRemoveSfxBlock(Arch: I7zOutArchive; remove: boolean);
begin
  SetBooleanProperty(Arch, 'RSFX', remove);
end;

procedure SevenZipAutoFilter(Arch: I7zOutArchive; auto: boolean);
begin
  SetBooleanProperty(Arch, 'F', auto);
end;

procedure SevenZipCompressHeaders(Arch: I7zOutArchive; compress: boolean);
begin
  SetBooleanProperty(Arch, 'HC', compress);
end;

procedure SevenZipCompressHeadersFull(Arch: I7zOutArchive; compress: boolean);
begin
  SetBooleanProperty(arch, 'HCF', compress);
end;

procedure SevenZipEncryptHeaders(Arch: I7zOutArchive; Encrypt: boolean);
begin
  SetBooleanProperty(arch, 'HE', Encrypt);
end;

procedure SevenZipVolumeMode(Arch: I7zOutArchive; Mode: boolean);
begin
  SetBooleanProperty(arch, 'V', Mode);
end;

type
  T7zPlugin = class(TInterfacedObject)
  private
    FHandle: THandle;
    FCreateObject: function(const clsid, iid :TGUID; var outObject): HRESULT; stdcall;
  public
    constructor Create(const lib: string); virtual;
    destructor Destroy; override;
    procedure CreateObject(const clsid, iid :TGUID; var obj);
  end;

  NMethodPropID = (
    kID,
    kMethodName, // kName
    kDecoder,
    kEncoder,
    kPackStreams,
    kUnpackStreams,
    kDescription,
    kDecoderIsAssigned,
    kEncoderIsAssigned,
    kDigestSize
  );

  T7zCodec = class(T7zPlugin, I7zCodec, ICompressProgressInfo)
  private
    FGetMethodProperty: function(index: UInt32; propID: NMethodPropID; var value: OleVariant): HRESULT; stdcall;
    FGetNumberOfMethods: function(numMethods: PUInt32): HRESULT; stdcall;
    function GetNumberOfMethods: UInt32;
    function GetMethodProperty(index: UInt32; propID: NMethodPropID): OleVariant;
    function GetName(const index: integer): string;
  protected
    function SetRatioInfo(inSize, outSize: PUInt64): HRESULT; stdcall;
  public
    function GetDecoder(const index: integer): ICompressCoder;
    function GetEncoder(const index: integer): ICompressCoder;
    constructor Create(const lib: string); override;
    property MethodProperty[index: UInt32; propID: NMethodPropID]: OleVariant read GetMethodProperty;
    property NumberOfMethods: UInt32 read GetNumberOfMethods;
    property Name[const index: integer]: string read GetName;
  end;

  T7zArchive = class(T7zPlugin)
  private
    FGetHandlerProperty: function(propID: NHandlerPropID; var value: OleVariant): HRESULT; stdcall;
    FClassId: TGUID;
    procedure SetClassId(const classid: TGUID);
    function GetClassId: TGUID;
  public
    function GetHandlerProperty(const propID: NHandlerPropID): OleVariant;
    function GetLibStringProperty(const Index: NHandlerPropID): string;
    function GetLibGUIDProperty(const Index: NHandlerPropID): TGUID;
    constructor Create(const lib: string); override;
    property HandlerProperty[const propID: NHandlerPropID]: OleVariant read GetHandlerProperty;
    property Name: string index kName read GetLibStringProperty;
    property ClassID: TGUID read GetClassId write SetClassId;
    property Extension: string index kExtension read GetLibStringProperty;
  end;

  T7zInArchive = class(T7zArchive, I7zInArchive, IProgress, IArchiveOpenCallback,
    IArchiveExtractCallback, ICryptoGetTextPassword, IArchiveOpenVolumeCallback,
    IArchiveOpenSetSubArchiveName)
  private
    FInArchive: IInArchive;
    FPasswordCallback: T7zPasswordCallback;
    FPasswordSender: Pointer;
    FProgressCallback: T7zProgressCallback;
    FProgressSender: Pointer;
    FStream: TStream;
    FPasswordIsDefined: Boolean;
    FPassword: UnicodeString;
    FSubArchiveMode: Boolean;
    FSubArchiveName: UnicodeString;
    FExtractCallBack: T7zGetStreamCallBack;
    FExtractSender: Pointer;
    FExtractPath: string;
    function GetInArchive: IInArchive;
    function GetItemProp(const Item: UInt32; prop: PROPID): OleVariant;
  protected
    // I7zInArchive
    procedure OpenFile(const filename: string); stdcall;
    procedure OpenStream(stream: IInStream); stdcall;
    procedure Close; stdcall;
    function GetNumberOfItems: UInt32; stdcall;
    function GetItemPath(const index: integer): UnicodeString; stdcall;
    function GetItemName(const index: integer): UnicodeString; stdcall;
    function GetItemSize(const index: integer): Int64; stdcall;
    function GetItemCompressedSize(const index: integer): Int64; stdcall;
    function GetItemWriteTime(const index: integer): TDateTime; stdcall;
    function GetItemAttributes(const index: integer): DWORD; stdcall;
    function GetItemIsFolder(const index: integer): boolean; stdcall;
    procedure ExtractItem(const item: UInt32; Stream: TStream; test: longbool); stdcall;
    procedure ExtractItemToPath(const item: UInt32; const path: string; test: longbool); stdcall;
    procedure ExtractItems(items: PCardArray; count: UInt32; test: longbool; sender: pointer; callback: T7zGetStreamCallBack); stdcall;
    procedure SetPasswordCallback(sender: Pointer; callback: T7zPasswordCallback); stdcall;
    procedure SetProgressCallback(sender: Pointer; callback: T7zProgressCallback); stdcall;
    procedure ExtractAll(test: longbool; sender: pointer; callback: T7zGetStreamCallBack); stdcall;
    procedure ExtractTo(const path: string); stdcall;
    procedure SetPassword(const password: UnicodeString); stdcall;
    // IArchiveOpenCallback
    function SetTotal(files, bytes: PUInt64): HRESULT; overload; stdcall;
    function SetCompleted(files, bytes: PUInt64): HRESULT; overload; stdcall;
    // IProgress
    function SetTotal(total: UInt64): HRESULT;  overload; stdcall;
    function SetCompleted(completeValue: PUInt64): HRESULT; overload; stdcall;
    // IArchiveExtractCallback
    function GetStream(index: UInt32; var outStream: ISequentialOutStream;
      askExtractMode: NAskMode): HRESULT; overload; stdcall;
    function PrepareOperation(askExtractMode: NAskMode): HRESULT; stdcall;
    function SetOperationResult(resultEOperationResult: NExtOperationResult): HRESULT; overload; stdcall;
    // ICryptoGetTextPassword
    function CryptoGetTextPassword(var password: TBStr): HRESULT; stdcall;
    // IArchiveOpenVolumeCallback
    function GetProperty(propID: PROPID; var value: OleVariant): HRESULT; overload; stdcall;
    function GetStream(const name: PWideChar; var inStream: IInStream): HRESULT; overload; stdcall;
    // IArchiveOpenSetSubArchiveName
    function SetSubArchiveName(name: PWideChar): HRESULT; stdcall;

  public
    constructor Create(const lib: string); override;
    destructor Destroy; override;
    property InArchive: IInArchive read GetInArchive;
  end;

  T7zOutArchive = class(T7zArchive, I7zOutArchive, IArchiveUpdateCallback, ICryptoGetTextPassword2)
  private
    FOutArchive: IOutArchive;
    FBatchList: TObjectList;
    FProgressCallback: T7zProgressCallback;
    FProgressSender: Pointer;
    FPassword: UnicodeString;
    function GetOutArchive: IOutArchive;
  protected
    // I7zOutArchive
    procedure AddStream(Stream: TStream; Ownership: TStreamOwnership;
      Attributes: Cardinal; CreationTime, LastWriteTime: TFileTime;
      const Path: UnicodeString; IsFolder, IsAnti: boolean); stdcall;
    procedure AddFile(const Filename: TFileName; const Path: UnicodeString); stdcall;
    procedure AddFiles(const Dir, Path, Wildcard: string; recurse: boolean); stdcall;
    procedure SaveToFile(const FileName: TFileName); stdcall;
    procedure SaveToStream(stream: TStream); stdcall;
    procedure SetProgressCallback(sender: Pointer; callback: T7zProgressCallback); stdcall;
    procedure ClearBatch; stdcall;
    procedure SetPassword(const password: UnicodeString); stdcall;
    procedure SetPropertie(name: UnicodeString; value: OleVariant); stdcall;
    // IProgress
    function SetTotal(total: UInt64): HRESULT; stdcall;
    function SetCompleted(completeValue: PUInt64): HRESULT; stdcall;
    // IArchiveUpdateCallback
    function GetUpdateItemInfo(index: UInt32;
        newData: PInt32; // 1 - new data, 0 - old data
        newProperties: PInt32; // 1 - new properties, 0 - old properties
        indexInArchive: PUInt32 // -1 if there is no in archive, or if doesn't matter
        ): HRESULT; stdcall;
    function GetProperty(index: UInt32; propID: PROPID; var value: OleVariant): HRESULT; stdcall;
    function GetStream(index: UInt32; var inStream: ISequentialInStream): HRESULT; stdcall;
    function SetOperationResult(operationResult: Integer): HRESULT; stdcall;
    // ICryptoGetTextPassword2
    function CryptoGetTextPassword2(passwordIsDefined: PInt32; var password: TBStr): HRESULT; stdcall;
  public
    constructor Create(const lib: string); override;
    destructor Destroy; override;
    property OutArchive: IOutArchive read GetOutArchive;
  end;

function CreateInArchive(const classid: TGUID; const lib: string): I7zInArchive;
begin
  Result := T7zInArchive.Create(lib);
  Result.ClassId := classid;
end;

function CreateOutArchive(const classid: TGUID; const lib: string): I7zOutArchive;
begin
  Result := T7zOutArchive.Create(lib);
  Result.ClassId := classid;
end;


{ T7zPlugin }

constructor T7zPlugin.Create(const lib: string);
begin
  inherited Create;
  FHandle := LoadLibrary(PChar(lib));
  if FHandle = 0 then
      RaiseLastOSError(GetLastError, Format(#13#10'Error loading library %s', [lib]));
  FCreateObject := GetProcAddress(FHandle, 'CreateObject');
  if not (Assigned(FCreateObject)) then
  begin
    FreeLibrary(FHandle);
    raise Exception.CreateFmt('%s is not a 7z library', [lib]);
  end;
end;

destructor T7zPlugin.Destroy;
begin
  FreeLibrary(FHandle);
  inherited;
end;

procedure T7zPlugin.CreateObject(const clsid, iid: TGUID; var obj);
var
  hr: HRESULT;
begin
  hr := FCreateObject(clsid, iid, obj);
  if failed(hr) then
    raise Exception.Create(SysErrorMessage(hr));
end;

{ T7zCodec }

constructor T7zCodec.Create(const lib: string);
begin
  inherited;
  FGetMethodProperty := GetProcAddress(FHandle, 'GetMethodProperty');
  FGetNumberOfMethods := GetProcAddress(FHandle, 'GetNumberOfMethods');
  if not (Assigned(FGetMethodProperty) and Assigned(FGetNumberOfMethods)) then
  begin
    FreeLibrary(FHandle);
    raise Exception.CreateFmt('%s is not a codec library', [lib]);
  end;
end;

function T7zCodec.GetDecoder(const index: integer): ICompressCoder;
var
  v: OleVariant;
begin
  v := MethodProperty[index, kDecoder];
  CreateObject(TPropVariant(v).puuid^, ICompressCoder, Result);
end;

function T7zCodec.GetEncoder(const index: integer): ICompressCoder;
var
  v: OleVariant;
begin
  v := MethodProperty[index, kEncoder];
  CreateObject(TPropVariant(v).puuid^, ICompressCoder, Result);
end;

function T7zCodec.GetMethodProperty(index: UInt32;
  propID: NMethodPropID): OleVariant;
var
  hr: HRESULT;
begin
  hr := FGetMethodProperty(index, propID, Result);
  if Failed(hr) then
    raise Exception.Create(SysErrorMessage(hr));
end;

function T7zCodec.GetName(const index: integer): string;
begin
  Result := MethodProperty[index, kMethodName];
end;

function T7zCodec.GetNumberOfMethods: UInt32;
var
  hr: HRESULT;
begin
  hr := FGetNumberOfMethods(@Result);
  if Failed(hr) then
    raise Exception.Create(SysErrorMessage(hr));
end;


function T7zCodec.SetRatioInfo(inSize, outSize: PUInt64): HRESULT;
begin
  Result := S_OK;
end;

{ T7zInArchive }

procedure T7zInArchive.Close; stdcall;
begin
  FPasswordIsDefined := false;
  FSubArchiveMode := false;
  FInArchive.Close;
  FInArchive := nil;
end;

constructor T7zInArchive.Create(const lib: string);
begin
  inherited;
  FPasswordCallback := nil;
  FPasswordSender := nil;
  FPasswordIsDefined := false;
  FSubArchiveMode := false;
  FExtractCallBack := nil;
  FExtractSender := nil;
end;

destructor T7zInArchive.Destroy;
begin
  FInArchive := nil;
  inherited;
end;

function T7zInArchive.GetInArchive: IInArchive;
begin
  if FInArchive = nil then
    CreateObject(ClassID, IInArchive, FInArchive);
  Result := FInArchive;
end;

function T7zInArchive.GetItemPath(const index: integer): UnicodeString; stdcall;
begin
  Result := UnicodeString(GetItemProp(index, kpidPath));
end;

function T7zInArchive.GetNumberOfItems: UInt32; stdcall;
begin
  RINOK(FInArchive.GetNumberOfItems(Result));
end;

procedure T7zInArchive.OpenFile(const filename: string); stdcall;
var
  strm: IInStream;
begin
  strm := T7zStream.Create(TFileStream.Create(filename, fmOpenRead or fmShareDenyNone), soOwned);
  try
    RINOK(
      InArchive.Open(
        strm,
          @MAXCHECK, self as IArchiveOpenCallBack
        )
      );
  finally
    strm := nil;
  end;
end;

procedure T7zInArchive.OpenStream(stream: IInStream); stdcall;
begin
  RINOK(InArchive.Open(stream, @MAXCHECK, self as IArchiveOpenCallBack));
end;

function T7zInArchive.GetItemAttributes(const index: integer): DWORD;
begin
  result := DWORD(GetItemProp(index, kpidAttrib));
end;

function T7zInArchive.GetItemIsFolder(const index: integer): boolean; stdcall;
begin
  Result := Boolean(GetItemProp(index, kpidIsDir));
end;

function T7zInArchive.GetItemProp(const Item: UInt32;
  prop: PROPID): OleVariant;
begin
  FInArchive.GetProperty(Item, prop, Result);
end;

procedure T7zInArchive.ExtractItem(const item: UInt32; Stream: TStream; test: longbool); stdcall;
begin
  FStream := Stream;
  try
    if test then
      RINOK(FInArchive.Extract(@item, 1, 1, self as IArchiveExtractCallback)) else
      RINOK(FInArchive.Extract(@item, 1, 0, self as IArchiveExtractCallback));
  finally
    FStream := nil;
  end;
end;

procedure T7zInArchive.ExtractItemToPath(const item: UInt32; const path: string; test: longbool); stdcall;
begin
  FExtractPath := IncludeTrailingPathDelimiter(path);
  try
    if test then
      RINOK(FInArchive.Extract(@item, 1, 1, self as IArchiveExtractCallback)) else
      RINOK(FInArchive.Extract(@item, 1, 0, self as IArchiveExtractCallback));
  finally
    FExtractPath := '';
  end;
end;

function T7zInArchive.GetStream(index: UInt32;
  var outStream: ISequentialOutStream; askExtractMode: NAskMode): HRESULT;
var
  path: string;
begin
  try
    if askExtractMode = kExtract then
      if FStream <> nil then
        outStream := T7zStream.Create(FStream, soReference) as ISequentialOutStream else
      if assigned(FExtractCallback) then
      begin
        Result := FExtractCallBack(FExtractSender, index, outStream);
        Exit;
      end else
      if FExtractPath <> '' then
      begin
        if not GetItemIsFolder(index) then
        begin
          path := FExtractPath + GetItemPath(index);
          ForceDirectories(ExtractFilePath(path));
          outStream := T7zStream.Create(TFileStream.Create(path, fmCreate), soOwned, path, GetItemWriteTime(index));
        end
        else
        begin
          path := FExtractPath + GetItemPath(index);
          ForceDirectories(path);
        end;
        SetFileAttributes(PChar(path), GetItemAttributes(index));
      end;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zInArchive.PrepareOperation(askExtractMode: NAskMode): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.SetCompleted(completeValue: PUInt64): HRESULT;
begin
  try
    if Assigned(FProgressCallback) and (completeValue <> nil) then
      Result := FProgressCallback(FProgressSender, false, completeValue^) else
      Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zInArchive.SetCompleted(files, bytes: PUInt64): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.SetOperationResult(
  resultEOperationResult: NExtOperationResult): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.SetTotal(total: UInt64): HRESULT;
begin
  try
    if Assigned(FProgressCallback) then
      Result := FProgressCallback(FProgressSender, true, total) else
      Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zInArchive.SetTotal(files, bytes: PUInt64): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.CryptoGetTextPassword(var password: TBStr): HRESULT;
var
  wpass: UnicodeString;
begin
  try
    if FPasswordIsDefined then
    begin
      password := SysAllocString(PWideChar(FPassword));
      if password = nil then
        Result := E_OUTOFMEMORY
      else
        Result := S_OK;
    end else
    if Assigned(FPasswordCallback) then
    begin
      Result := FPasswordCallBack(FPasswordSender, wpass);
      if Result = S_OK then
      begin
        password := SysAllocString(PWideChar(wpass));
        if password = nil then
        begin
          Result := E_OUTOFMEMORY;
          Exit;
        end;
        FPasswordIsDefined := True;
        FPassword := wpass;
      end;
    end else
      Result := S_FALSE;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zInArchive.GetProperty(propID: PROPID;
  var value: OleVariant): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.GetStream(const name: PWideChar;
  var inStream: IInStream): HRESULT;
begin
  Result := S_OK;
end;

procedure T7zInArchive.SetPasswordCallback(sender: Pointer;
  callback: T7zPasswordCallback); stdcall;
begin
  FPasswordSender := sender;
  FPasswordCallback := callback;
end;

function T7zInArchive.SetSubArchiveName(name: PWideChar): HRESULT;
begin
  FSubArchiveMode := true;
  FSubArchiveName := name;
  Result := S_OK;
end;

function T7zInArchive.GetItemName(const index: integer): UnicodeString; stdcall;
begin
  Result := UnicodeString(GetItemProp(index, kpidName));
end;

function T7zInArchive.GetItemSize(const index: integer): Int64; stdcall;
begin
  Result := Int64(GetItemProp(index, kpidSize));
end;

function T7zInArchive.GetItemCompressedSize(const index: integer): Int64; stdcall;
begin
  Result := Int64(GetItemProp(index, kpidPackSize));
end;

function T7zInArchive.GetItemWriteTime(const index: integer): TDateTime;
var
  v: OleVariant;
begin
  v := GetItemProp(index, kpidMTime);
  if TPropVariant(v).vt = VT_FILETIME then
  begin
    result := FileTimeToDateTime(TPropVariant(v).filetime);
  end
  else
  begin
    result := 0;
  end;
end;

procedure T7zInArchive.ExtractItems(items: PCardArray; count: UInt32; test: longbool;
  sender: pointer; callback: T7zGetStreamCallBack); stdcall;
begin
  FExtractCallBack := callback;
  FExtractSender := sender;
  try
    if test then
      RINOK(FInArchive.Extract(items, count, 1, self as IArchiveExtractCallback)) else
      RINOK(FInArchive.Extract(items, count, 0, self as IArchiveExtractCallback));
  finally
    FExtractCallBack := nil;
    FExtractSender := nil;
  end;
end;

procedure T7zInArchive.SetProgressCallback(sender: Pointer;
  callback: T7zProgressCallback); stdcall;
begin
  FProgressSender := sender;
  FProgressCallback := callback;
end;

procedure T7zInArchive.ExtractAll(test: longbool; sender: pointer;
  callback: T7zGetStreamCallBack);
begin
  FExtractCallBack := callback;
  FExtractSender := sender;
  try
    if test then
      RINOK(FInArchive.Extract(nil, $FFFFFFFF, 1, self as IArchiveExtractCallback)) else
      RINOK(FInArchive.Extract(nil, $FFFFFFFF, 0, self as IArchiveExtractCallback));
  finally
    FExtractCallBack := nil;
    FExtractSender := nil;
  end;
end;

procedure T7zInArchive.ExtractTo(const path: string);
begin
  FExtractPath := IncludeTrailingPathDelimiter(path);
  try
    RINOK(FInArchive.Extract(nil, $FFFFFFFF, 0, self as IArchiveExtractCallback));
  finally
    FExtractPath := '';
  end;
end;

procedure T7zInArchive.SetPassword(const password: UnicodeString);
begin
  FPassword := password;
  FPasswordIsDefined :=  FPassword <> '';
end;

{ T7zArchive }

constructor T7zArchive.Create(const lib: string);
begin
  inherited;
  FGetHandlerProperty := GetProcAddress(FHandle, 'GetHandlerProperty');
  if not Assigned(FGetHandlerProperty) then
  begin
    FreeLibrary(FHandle);
    raise Exception.CreateFmt('%s is not a Format library', [lib]);
  end;
  FClassId := GUID_NULL;
end;

function T7zArchive.GetClassId: TGUID;
begin
  Result := FClassId;
end;

function T7zArchive.GetHandlerProperty(const propID: NHandlerPropID): OleVariant;
var
  hr: HRESULT;
begin
  hr := FGetHandlerProperty(propID, Result);
  if Failed(hr) then
    raise Exception.Create(SysErrorMessage(hr));
end;

function T7zArchive.GetLibGUIDProperty(const Index: NHandlerPropID): TGUID;
var
  v: OleVariant;
begin
  v := HandlerProperty[index];
  Result := TPropVariant(v).puuid^;
end;

function T7zArchive.GetLibStringProperty(const Index: NHandlerPropID): string;
begin
  Result := HandlerProperty[Index];
end;

procedure T7zArchive.SetClassId(const classid: TGUID);
begin
  FClassId := classid;
end;

{ T7zStream }

constructor T7zStream.Create(Stream: TStream; Ownership: TStreamOwnership = soReference; filename: string=''; writeTime: TDateTime=0);
begin
  inherited Create;
  FStream := Stream;
  FOwnership := Ownership;
  FFileName := filename;
  FWriteTime := writeTime;
end;

destructor T7zStream.destroy;
begin
  if FOwnership = soOwned then
  begin
    FStream.Free;
    FStream := nil;
  end;

  // TODO: Shouldn't this be done somewhere else? Is the flush method the correct place (if flush=finished)?
  if (FFileName<>'') and (CompareValue(FWriteTime,0)<>0) then
    TFile.SetLastWriteTime(FFilename, FWriteTime);

  inherited;
end;

function T7zStream.OutStreamFinish: HRESULT;
begin
  Result := S_OK;
end;

function T7zStream.GetSize(size: PUInt64): HRESULT;
begin
  try
    if size <> nil then
      size^ := FStream.Size;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zStream.Read(data: Pointer; size: UInt32;
  processedSize: PUInt32): HRESULT;
var
  len: integer;
begin
  try
    len := FStream.Read(data^, size);
    if processedSize <> nil then
      processedSize^ := len;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zStream.Seek(offset: Int64; seekOrigin: UInt32;
  newPosition: PUInt64): HRESULT;
begin
  try
    FStream.Seek(offset, TSeekOrigin(seekOrigin));
    if newPosition <> nil then
      newPosition^ := FStream.Position;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zStream.SetSize(newSize: UInt64): HRESULT;
begin
  try
    FStream.Size := newSize;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zStream.Write(data: Pointer; size: UInt32;
  processedSize: PUInt32): HRESULT;
var
  len: integer;
begin
  try
    len := FStream.Write(data^, size);
    if processedSize <> nil then
      processedSize^ := len;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

type
  TSourceMode = (smStream, smFile);

  T7zBatchItem = class
    SourceMode: TSourceMode;
    Stream: TStream;
    Attributes: Cardinal;
    CreationTime, LastWriteTime, LastAccessTime: TFileTime;
    Path: UnicodeString;
    IsFolder, IsAnti: boolean;
    FileName: TFileName;
    Ownership: TStreamOwnership;
    Size: Int64;
    destructor Destroy; override;
  end;

destructor T7zBatchItem.Destroy;
begin
  if (Ownership = soOwned) and (Stream <> nil) then
    Stream.Free;
  inherited;
end;

{ T7zOutArchive }

procedure T7zOutArchive.AddFile(const Filename: TFileName; const Path: UnicodeString);
var
  item: T7zBatchItem;
  Handle: THandle;
  TempSize: Int64;
  TempSizeHi: Cardinal;
begin
  if not FileExists(Filename) then exit;
  item := T7zBatchItem.Create;
  Item.SourceMode := smFile;
  item.Stream := nil;
  item.FileName := Filename;
  item.Path := Path;
  Handle := FileOpen(Filename, fmOpenRead or fmShareDenyNone);
  GetFileTime(Handle, @item.CreationTime, nil, @item.LastWriteTime);
  Int64Rec(TempSize).Lo := GetFileSize(Handle, @TempSizeHi);
  Int64Rec(TempSize).Hi := TempSizeHi;
  item.Size := TempSize;
  CloseHandle(Handle);
  item.Attributes := GetFileAttributes(PChar(Filename));
  item.IsFolder := false;
  item.IsAnti := False;
  item.Ownership := soOwned;
  FBatchList.Add(item);
end;

procedure T7zOutArchive.AddFiles(const Dir, Path, Wildcard: string; recurse: boolean);
var
  lencut: integer;
  willlist: TStringList;
  zedir: string;
  procedure Traverse(p: string);
  var
    f: TSearchRec;
    i: integer;
    item: T7zBatchItem;
    alsoIncludeEmptyFolders: boolean;
  begin
    alsoIncludeEmptyFolders := false;
    for i := 0 to willlist.Count - 1 do
    begin
      if willlist[i] = '*' then
      begin
        alsoIncludeEmptyFolders := true;
      end;
    end;

    if recurse then
    begin
      if FindFirst(IncludeTrailingPathDelimiter(p) + '*', faDirectory or faReadOnly or faHidden or faSysFile or faArchive, f) = 0 then
      repeat
        if (f.Name <> '.') and (f.Name <> '..') then
        begin
          if DirectoryExists(IncludeTrailingPathDelimiter(p) + f.Name) then
          begin
            Traverse(IncludeTrailingPathDelimiter(p) + f.Name);

            if alsoIncludeEmptyFolders then
            begin
              item := T7zBatchItem.Create;
              Item.SourceMode := smFile;
              item.Stream := nil;
              item.FileName := IncludeTrailingPathDelimiter(p) + f.Name;
              item.Path := copy(item.FileName, lencut, length(item.FileName) - lencut + 1);
              if path <> '' then
                item.Path := IncludeTrailingPathDelimiter(path) + item.Path;
              item.CreationTime := f.FindData.ftCreationTime;
              item.LastWriteTime := f.FindData.ftLastWriteTime;
              item.LastAccessTime := f.FindData.ftLastAccessTime;
              item.Attributes := f.FindData.dwFileAttributes;
              item.Size := f.Size;
              item.IsFolder := true;
              item.IsAnti := False;
              item.Ownership := soOwned;
              FBatchList.Add(item);
            end;

          end;
        end;
      until FindNext(f) <> 0;
      SysUtils.FindClose(f);
    end;

    for i := 0 to willlist.Count - 1 do
    begin
      if FindFirst(IncludeTrailingPathDelimiter(p) + willlist[i], faReadOnly or faHidden or faSysFile or faArchive, f) = 0 then
      repeat
        if (f.Name <> '.') and (f.Name <> '..') then
        begin
          item := T7zBatchItem.Create;
          Item.SourceMode := smFile;
          item.Stream := nil;
          item.FileName := IncludeTrailingPathDelimiter(p) + f.Name;
          item.Path := copy(item.FileName, lencut, length(item.FileName) - lencut + 1);
          if path <> '' then
            item.Path := IncludeTrailingPathDelimiter(path) + item.Path;
          item.CreationTime := f.FindData.ftCreationTime;
          item.LastWriteTime := f.FindData.ftLastWriteTime;
          item.LastAccessTime := f.FindData.ftLastAccessTime;
          item.Attributes := f.FindData.dwFileAttributes;
          item.Size := f.Size;
          item.IsFolder := false;
          item.IsAnti := False;
          item.Ownership := soOwned;
          FBatchList.Add(item);
        end;
      until FindNext(f) <> 0;
      SysUtils.FindClose(f);
    end;
  end;
begin
  willlist := TStringList.Create;
  try
    willlist.Delimiter := ';';
    willlist.DelimitedText := Wildcard;
    zedir := IncludeTrailingPathDelimiter(Dir);
    lencut := Length(zedir) + 1;
    Traverse(zedir);
  finally
    willlist.Free;
  end;
end;

procedure T7zOutArchive.AddStream(Stream: TStream; Ownership: TStreamOwnership;
  Attributes: Cardinal; CreationTime, LastWriteTime: TFileTime;
  const Path: UnicodeString; IsFolder, IsAnti: boolean); stdcall;
var
  item: T7zBatchItem;
begin
  item := T7zBatchItem.Create;
  Item.SourceMode := smStream;
  item.Attributes := Attributes;
  item.CreationTime := CreationTime;
  item.LastWriteTime := LastWriteTime;
  item.LastAccessTime.dwLowDateTime := 0;
  item.LastAccessTime.dwHighDateTime := 0;
  item.Path := Path;
  item.IsFolder := IsFolder;
  item.IsAnti := IsAnti;
  item.Stream := Stream;
  item.Size := Stream.Size;
  item.Ownership := Ownership;
  FBatchList.Add(item);
end;

procedure T7zOutArchive.ClearBatch;
begin
  FBatchList.Clear;
end;

constructor T7zOutArchive.Create(const lib: string);
begin
  inherited;
  FBatchList := TObjectList.Create;
  FProgressCallback := nil;
  FProgressSender := nil;
end;

function T7zOutArchive.CryptoGetTextPassword2(passwordIsDefined: PInt32;
  var password: TBStr): HRESULT;
begin
  try
    if FPassword <> '' then
    begin
     passwordIsDefined^ := 1;
     password := SysAllocString(PWideChar(FPassword));
     if password = nil then
     begin
       Result := E_OUTOFMEMORY;
       Exit;
     end;
    end else
      passwordIsDefined^ := 0;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

destructor T7zOutArchive.Destroy;
begin
  FOutArchive := nil;
  FBatchList.Free;
  inherited;
end;

function T7zOutArchive.GetOutArchive: IOutArchive;
begin
  if FOutArchive = nil then
    CreateObject(ClassID, IOutArchive, FOutArchive);
  Result := FOutArchive;
end;

function T7zOutArchive.GetProperty(index: UInt32; propID: PROPID;
  var value: OleVariant): HRESULT;
var
  item: T7zBatchItem;
begin
  try
    item := T7zBatchItem(FBatchList[index]);
    case propID of
      kpidAttrib:
        begin
          TPropVariant(Value).vt := VT_UI4;
          TPropVariant(Value).ulVal := item.Attributes;
        end;
      kpidMTime:
        begin
          TPropVariant(value).vt := VT_FILETIME;
          TPropVariant(value).filetime := item.LastWriteTime;
        end;
      kpidATime:
        begin
          TPropVariant(value).vt := VT_FILETIME;
          TPropVariant(value).filetime := item.LastAccessTime;
        end;
      kpidPath:
        begin
          if item.Path <> '' then
            value := item.Path;
        end;
      kpidIsDir: Value := item.IsFolder;
      kpidSize:
        begin
          TPropVariant(Value).vt := VT_UI8;
          TPropVariant(Value).uhVal.QuadPart := item.Size;
        end;
      kpidCTime:
        begin
          TPropVariant(value).vt := VT_FILETIME;
          TPropVariant(value).filetime := item.CreationTime;
        end;
      kpidIsAnti: value := item.IsAnti;
    else
      // beep(0,0);
    end;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zOutArchive.GetStream(index: UInt32;
  var inStream: ISequentialInStream): HRESULT;
var
  item: T7zBatchItem;
begin
  try
    item := T7zBatchItem(FBatchList[index]);
    case item.SourceMode of
      smFile: inStream := T7zStream.Create(TFileStream.Create(item.FileName, fmOpenRead or fmShareDenyNone), soOwned);
      smStream:
        begin
          item.Stream.Seek(0, soFromBeginning);
          inStream := T7zStream.Create(item.Stream);
        end;
    end;
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zOutArchive.GetUpdateItemInfo(index: UInt32; newData,
  newProperties: PInt32; indexInArchive: PUInt32): HRESULT;
begin
  try
    newData^ := 1;
    newProperties^ := 1;
    indexInArchive^ := Cardinal(-1);
    Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

procedure T7zOutArchive.SaveToFile(const FileName: TFileName);
var
  f: TFileStream;
begin
  f := TFileStream.Create(FileName, fmCreate);
  try
    SaveToStream(f);
  finally
    f.free;
  end;
end;

procedure T7zOutArchive.SaveToStream(stream: TStream);
var
  strm: ISequentialOutStream;
begin
  strm := T7zStream.Create(stream);
  try
    RINOK(OutArchive.UpdateItems(strm, FBatchList.Count, self as IArchiveUpdateCallback));
  finally
    strm := nil;
  end;
end;

function T7zOutArchive.SetCompleted(completeValue: PUInt64): HRESULT;
begin
  try
    if Assigned(FProgressCallback) and (completeValue <> nil) then
      Result := FProgressCallback(FProgressSender, false, completeValue^) else
      Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

function T7zOutArchive.SetOperationResult(
  operationResult: Integer): HRESULT;
begin
  Result := S_OK;
end;

procedure T7zOutArchive.SetPassword(const password: UnicodeString);
begin
  FPassword := password;
end;

procedure T7zOutArchive.SetProgressCallback(sender: Pointer;
  callback: T7zProgressCallback);
begin
  FProgressCallback := callback;
  FProgressSender := sender;
end;

procedure T7zOutArchive.SetPropertie(name: UnicodeString;
  value: OleVariant);
var
  intf: ISetProperties;
  p: PWideChar;
begin
  intf := OutArchive as ISetProperties;
  p := PWideChar(name);
  RINOK(intf.SetProperties(@p, @TPropVariant(value), 1));
end;

function T7zOutArchive.SetTotal(total: UInt64): HRESULT;
begin
  try
    if Assigned(FProgressCallback) then
      Result := FProgressCallback(FProgressSender, true, total) else
      Result := S_OK;
  except
    // We must not throw an Delphi exception, because 7z.dll cannot handle it
    // (the unhandled Delphi Exception will be forwarded as 0x0eedfade and crashes
    // the calling process without possibility of recovery).
    on E: EAbort do
      Result := E_ABORT
    else
      Result := E_FAIL;
  end;
end;

end.
