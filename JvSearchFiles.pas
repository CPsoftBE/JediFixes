{-----------------------------------------------------------------------------
The contents of this file are subject to the Mozilla Public License
Version 1.1 (the "License"); you may not use this file except in compliance
with the License. You may obtain a copy of the License at
http://www.mozilla.org/MPL/MPL-1.1.html

Software distributed under the License is distributed on an "AS IS" basis,
WITHOUT WARRANTY OF ANY KIND, either expressed or implied. See the License for
the specific language governing rights and limitations under the License.

The Original Code is: JvSearchFiles.PAS, released on 2002-05-26.

The Initial Developer of the Original Code is Peter Thrnqvist [peter3 at sourceforge dot net]
Portions created by Peter Thrnqvist are Copyright (C) 2002 Peter Thrnqvist.
All Rights Reserved.

Contributor(s):
David Frauzel (DF)
Remko Bonte

You may retrieve the latest version of this file at the Project JEDI's JVCL home page,
located at http://jvcl.delphi-jedi.org

Description:
  Wrapper for a file search engine.

Known Issues:
-----------------------------------------------------------------------------}
// $Id$

Unit JvSearchFiles;

{$I jvcl.inc}
{$I windowsonly.inc}

Interface

Uses
  {$IFDEF UNITVERSIONING}
  JclUnitVersioning,
  {$ENDIF UNITVERSIONING}
  Classes, SysUtils,
  {$IFDEF MSWINDOWS}
  Windows,
  {$ENDIF MSWINDOWS}
  JvComponentBase, JvJCLUtils, JvWin32;

Type
  TJvAttrFlagKind = (tsMustBeSet, tsDontCare, tsMustBeUnSet);
  TJvDirOption = (doExcludeSubDirs, doIncludeSubDirs, doExcludeInvalidDirs,
    doExcludeCompleteInvalidDirs);
  { doExcludeSubDirs
      Only search in root directory.
    doIncludeSubDirs
      Search in root directory and it's sub-directories.
    doExcludeInvalidDirs
      Search in root directory and it's sub-directories; do not search in
      an invalid directory, but do search in the sub-directories of an
      invalid directory.
    doExcludeCompleteInvalidDirs
      Search in root directory and it's sub-directories; do not search in
      an invalid directory, and the sub-directories of an invalid directory.

    Invalid directory = directory with params that doesn't agree with the
      params specified by DirParams.
  }

  TJvSearchOption = (soAllowDuplicates, soCheckRootDirValid,
    soExcludeFilesInRootDir, soOwnerData, soSearchDirs, soSearchFiles, soSorted,
    soStripDirs, soIncludeSystemHiddenDirs, soIncludeSystemHiddenFiles);
  TJvSearchOptions = Set Of TJvSearchOption;
  { soAllowDuplicates
      Allow duplicate file/dir names in property Files and Directories.
    soCheckRootDirValid
      Check if the root-directory is valid; Must DirOption must be equal to
      doExcludeSubDirs or doExcludeCompleteInvalidDirs, otherwise this flag is
      ignored.
    soExcludeFilesInRootDir
      Do not search in the root directory.
    soOwnerData
      Do not fill property Files and Directories while searching
    soSearchDirs
      Search for directories; ie trigger OnFindDirectory event and update
      totals [TotalDirectories, TotalFileSize] when a valid directory is found.
    soSearchFiles
      Search for files; ie trigger OnFindFile event and update totals
      [TotalFileSize, TotalFiles] when a valid file is found.
    soSorted
      Keep the values in property Files and Directories sorted.
    soStripDirs
      Strip the path of a dir/file name before inserting it in property
      Files and Directories
    soIncludeSystemHiddenDirs
      Do NOT ignore directories that are both system and hidden.
      Examples of such directories are 'RECYCLER', 'System Volume Information' etc.
    soIncludeSystemHiddenFiles
      Do NOT ignore files that are both system and hidden.
      Examples of such files are 'pagefile.sys', 'IO.SYS' etc.

  }

  TJvSearchType = (stAttribute, stFileMask, stFileMaskCaseSensitive,
    stLastChangeAfter, stLastChangeBefore, stMaxSize, stMinSize);
  TJvSearchTypes = Set Of TJvSearchType;

  TJvFileSearchEvent = Procedure(Sender: TObject; Const AName: String) Of Object;
  TJvSearchFilesError = Procedure(Sender: TObject; Var Handled: Boolean) Of Object;
  TJvCheckEvent = Procedure(Sender: TObject; Var Result: Boolean) Of Object;

  TJvErrorResponse = (erAbort, erIgnore, erRaise);

  TJvSearchAttributes = Class(TPersistent)
  Private
    FIncludeAttr: DWORD;
    FExcludeAttr: DWORD;
    Function GetAttr(Const Index: Integer): TJvAttrFlagKind;
    Procedure SetAttr(Const Index: Integer; Value: TJvAttrFlagKind);
    Procedure ReadIncludeAttr(Reader: TReader);
    Procedure ReadExcludeAttr(Reader: TReader);
    Procedure WriteIncludeAttr(Writer: TWriter);
    Procedure WriteExcludeAttr(Writer: TWriter);
  Protected
    { DefineProperties is used to publish properties IncludeAttr and
      ExcludeAttr }
    Procedure DefineProperties(Filer: TFiler); Override;
  Public
    Procedure Assign(Source: TPersistent); Override;
    Property IncludeAttr: DWORD Read FIncludeAttr Write FIncludeAttr;
    Property ExcludeAttr: DWORD Read FExcludeAttr Write FExcludeAttr;
  Published
    property ReadOnly: TJvAttrFlagKind index FILE_ATTRIBUTE_READONLY read GetAttr
      write SetAttr stored False;
    property Hidden: TJvAttrFlagKind index FILE_ATTRIBUTE_HIDDEN
      read GetAttr write SetAttr stored False;
    property System: TJvAttrFlagKind index FILE_ATTRIBUTE_SYSTEM
      read GetAttr write SetAttr stored False;
    property Archive: TJvAttrFlagKind index FILE_ATTRIBUTE_ARCHIVE
      read GetAttr write SetAttr stored False;
    property Normal: TJvAttrFlagKind index FILE_ATTRIBUTE_NORMAL
      read GetAttr write SetAttr stored False;
    property Temporary: TJvAttrFlagKind index FILE_ATTRIBUTE_TEMPORARY
      read GetAttr write SetAttr stored False;
    property SparseFile: TJvAttrFlagKind index FILE_ATTRIBUTE_SPARSE_FILE
      read GetAttr write SetAttr stored False;
    property ReparsePoint: TJvAttrFlagKind index FILE_ATTRIBUTE_REPARSE_POINT
      read GetAttr write SetAttr stored False;
    property Compressed: TJvAttrFlagKind index FILE_ATTRIBUTE_COMPRESSED
      read GetAttr write SetAttr stored False;
    property OffLine: TJvAttrFlagKind index FILE_ATTRIBUTE_OFFLINE
      read GetAttr write SetAttr stored False;
    property NotContentIndexed: TJvAttrFlagKind index
      FILE_ATTRIBUTE_NOT_CONTENT_INDEXED read GetAttr write SetAttr stored False;
    property Encrypted: TJvAttrFlagKind index FILE_ATTRIBUTE_ENCRYPTED read
      GetAttr write SetAttr stored False;
  End;

  TJvSearchParams = Class(TPersistent)
  Private
    FMaxSizeHigh: Cardinal;
    FMaxSizeLow: Cardinal;
    FMinSizeHigh: Cardinal;
    FMinSizeLow: Cardinal;
    FLastChangeBefore: TDateTime;
    FLastChangeBeforeFT: TFileTime;
    FLastChangeAfter: TDateTime;
    FLastChangeAfterFT: TFileTime;
    FSearchTypes: TJvSearchTypes;
    FFileMasks: TStringList;
    FCaseFileMasks: TStringList;
    FFileMaskSeperator: Char;
    FAttributes: TJvSearchAttributes;
    Procedure FileMasksChange(Sender: TObject);
    Function GetFileMask: String;
    Function GetMaxSize: Int64;
    Function GetMinSize: Int64;
    Function GetFileMasks: TStrings;
    Function IsLastChangeAfterStored: Boolean;
    Function IsLastChangeBeforeStored: Boolean;
    Procedure SetAttributes(Const Value: TJvSearchAttributes);
    Procedure SetFileMasks(Const Value: TStrings);
    Procedure SetFileMask(Const Value: String);
    Procedure SetLastChangeAfter(Const Value: TDateTime);
    Procedure SetLastChangeBefore(Const Value: TDateTime);
    Procedure SetMaxSize(Const Value: Int64);
    Procedure SetMinSize(Const Value: Int64);
    Procedure SetSearchTypes(Const Value: TJvSearchTypes);
    Procedure UpdateCaseMasks;
  Public
    Constructor Create; Virtual;
    Destructor Destroy; Override;
    Procedure Assign(Source: TPersistent); Override;
    Function Check(Const AFindData: TWin32FindData): Boolean;
    Property FileMask: String Read GetFileMask Write SetFileMask;
    property FileMaskSeperator: Char read FFileMaskSeperator write
      FFileMaskSeperator default ';';
  Published
    Property Attributes: TJvSearchAttributes Read FAttributes Write SetAttributes;
    Property SearchTypes: TJvSearchTypes Read FSearchTypes Write SetSearchTypes Default [];
    Property MinSize: Int64 Read GetMinSize Write SetMinSize;
    Property MaxSize: Int64 Read GetMaxSize Write SetMaxSize;
    property LastChangeAfter: TDateTime read FLastChangeAfter write SetLastChangeAfter
      stored IsLastChangeAfterStored;
    property LastChangeBefore: TDateTime read FLastChangeBefore write SetLastChangeBefore
      stored IsLastChangeBeforeStored;
    Property FileMasks: TStrings Read GetFileMasks Write SetFileMasks;
  End;

  {$IFDEF RTL230_UP}
  [ComponentPlatformsAttribute(pidWin32 Or pidWin64 Or pidOSX32)]
  {$ENDIF RTL230_UP}
  TJvSearchFiles = Class(TJvComponent)
  Private
    FSearching: Boolean;
    FTotalDirectories: Integer;
    FTotalFiles: Integer;
    FTotalFileSize: Int64;
    FRootDirectory: String;
    FOnFindFile: TJvFileSearchEvent;
    FOnFindDirectory: TJvFileSearchEvent;
    FOptions: TJvSearchOptions;
    FOnAbort: TNotifyEvent;
    FOnError: TJvSearchFilesError;
    FOnProgress: TNotifyEvent;
    FDirectories: TStringList;
    FFiles: TStringList;
    FFindData: TWin32FindData;
    FAborting: Boolean;
    FErrorResponse: TJvErrorResponse;
    FOnCheck: TJvCheckEvent;
    FOnBeginScanDir: TJvFileSearchEvent;
    FDirOption: TJvDirOption;
    FDirParams: TJvSearchParams;
    FFileParams: TJvSearchParams;
    FRecurseDepth: Integer;
    FScanPath: String;
    FTotalFilesAdded: Integer;
    FCurrentDepth: Integer;
    Function GetIsRootDirValid: Boolean;
    Function GetIsDepthAllowed(Const ADepth: Integer): Boolean;
    Function GetDirectories: TStrings;
    Function GetFiles: TStrings;
    Function GetRootPath: String;
    Procedure SetDirParams(Const Value: TJvSearchParams);
    Procedure SetFileParams(Const Value: TJvSearchParams);
    Procedure SetOptions(Const Value: TJvSearchOptions);
    Procedure SetTotalFilesAdded(Const Value: Integer);
    Procedure SetCurrentDepth(Const Value: Integer);
  Protected
    Procedure DoBeginScanDir(Const ADirName: String); Virtual;
    Procedure DoFindFile(Const APath: String); Virtual;
    Procedure DoFindDir(Const APath: String); Virtual;
    Procedure DoAbort; Virtual;
    Procedure DoProgress; Virtual;
    Function DoCheckDir: Boolean; Virtual;
    Function DoCheckFile: Boolean; Virtual;
    Function HandleError: Boolean; Virtual;
    Procedure Init; Virtual;
    function EnumFiles(const ADirectoryName: string; Dirs: TStrings;
      const Search: Boolean): Boolean;
    function InternalSearch(const ADirectoryName: string;
      const Search: Boolean; var ADepth: Integer): Boolean; virtual;
  Public
    Constructor Create(AOwner: TComponent); Override;
    Destructor Destroy; Override;
    Procedure Abort;
    Function Search: Boolean;
    Property FindData: TWin32FindData Read FFindData;
    Property Files: TStrings Read GetFiles;
    Property Directories: TStrings Read GetDirectories;
    Property IsRootDirValid: Boolean Read GetIsRootDirValid;
    Property Searching: Boolean Read FSearching;
    Property TotalDirectories: Integer Read FTotalDirectories;
    Property TotalFileSize: Int64 Read FTotalFileSize;
    Property TotalFiles: Integer Read FTotalFiles;
    Property TotalFilesAdded: Integer Read FTotalFilesAdded Write SetTotalFilesAdded;
    Property RootPath: String Read GetRootPath;
    Property ScanPath: String Read FScanPath;
    Property CurrentDepth: Integer Read FCurrentDepth Write SetCurrentDepth Default 0;
  Published
    Property DirOption: TJvDirOption Read FDirOption Write FDirOption Default doIncludeSubDirs;
    // RecurseDepth sets the number of subfolders to search. If 0, all subfolders
    // are searched (as long as doIncludeSubDirs is true)
    Property RecurseDepth: Integer Read FRecurseDepth Write FRecurseDepth Default 0;
    Property RootDirectory: String Read FRootDirectory Write FRootDirectory;
    Property Options: TJvSearchOptions Read FOptions Write SetOptions Default [soSearchFiles];
    property ErrorResponse: TJvErrorResponse read FErrorResponse write
      FErrorResponse default erAbort;
    Property DirParams: TJvSearchParams Read FDirParams Write SetDirParams;
    Property FileParams: TJvSearchParams Read FFileParams Write SetFileParams;
    property OnBeginScanDir: TJvFileSearchEvent read FOnBeginScanDir write
      FOnBeginScanDir;
    Property OnFindFile: TJvFileSearchEvent Read FOnFindFile Write FOnFindFile;
    property OnFindDirectory: TJvFileSearchEvent read FOnFindDirectory write
      FOnFindDirectory;
    Property OnAbort: TNotifyEvent Read FOnAbort Write FOnAbort;
    Property OnError: TJvSearchFilesError Read FOnError Write FOnError;
    { Maybe add a flag to Options to disable OnCheck }
    Property OnCheck: TJvCheckEvent Read FOnCheck Write FOnCheck;
    // (rom) replaced ProcessMessages with OnProgress event
    Property OnProgress: TNotifyEvent Read FOnProgress Write FOnProgress;
  End;

{$IFDEF UNITVERSIONING}

Const
  UnitVersioning: TUnitVersionInfo = (
    RCSfile: '$URL$';
    Revision: '$Revision$';
    Date: '$Date$';
    LogPath: 'JVCL\run'
  );
{$ENDIF UNITVERSIONING}

Implementation

Uses
  JclStrings, JclDateTime;

{ Maybe TJvSearchFiles should be implemented with FindFirst, FindNext.
  There isn't a good reason to use FindFirstFile, FindNextFile instead of
  FindFirst, FindNext; except to prevent a little overhead perhaps. }

Const
  CDate1_1_1980 = 29221;

Function IsDotOrDotDot(P: PChar): Boolean;
Begin
  // check if a string is '.' (self) or '..' (parent)
  If P^ = '.' Then
  Begin
    Inc(P);
    Result := (P^ = #0) or ((P^ = '.') and ((P+1)^ = #0));
  End
  Else
    Result := False;
End;

Function IsSystemAndHidden(Const AFindData: TWin32FindData): Boolean;
Const
  cSystemHidden = FILE_ATTRIBUTE_SYSTEM Or FILE_ATTRIBUTE_HIDDEN;
Begin
  With AFindData Do
    Result := dwFileAttributes And cSystemHidden = cSystemHidden;
End;

//=== { TJvSearchFiles } =====================================================

Constructor TJvSearchFiles.Create(AOwner: TComponent);
Begin
  Inherited Create(AOwner);
  FFiles := TStringList.Create;
  FDirectories := TStringList.Create;
  FDirParams := TJvSearchParams.Create;
  FFileParams := TJvSearchParams.Create;

  { defaults }
  Options := [soSearchFiles];
  DirOption := doIncludeSubDirs;
  ErrorResponse := erAbort;
  //FFileParams.SearchTypes := [stFileMask];
End;

Destructor TJvSearchFiles.Destroy;
Begin
  FFiles.Free;
  FDirectories.Free;
  FFileParams.Free;
  FDirParams.Free;
  Inherited Destroy;
End;

Procedure TJvSearchFiles.Abort;
Begin
  If Not FSearching Then
    Exit;
  FAborting := True;
  DoAbort;
End;

Procedure TJvSearchFiles.DoAbort;
Begin
  If Assigned(FOnAbort) Then
    FOnAbort(Self);
End;

Procedure TJvSearchFiles.DoProgress;
Begin
  If Assigned(FOnProgress) Then
    FOnProgress(Self);
End;

Procedure TJvSearchFiles.DoBeginScanDir(Const ADirName: String);
Begin
  FScanPath := IncludeTrailingPathDelimiter(ADirName);
  // On Begin Scan Dir
  If Assigned(FOnBeginScanDir) Then
    FOnBeginScanDir(Self, ADirName);
End;

Function TJvSearchFiles.DoCheckDir: Boolean;
Begin
  If Assigned(FOnCheck) Then
  Begin
    // We always want to check the Dir Params
    Result := FDirParams.Check(FFindData);
    If Result Then
    Begin
      Result := False;
      FOnCheck(Self, Result);
    End;
  End
  Else
    Result := FDirParams.Check(FFindData)
End;

Function TJvSearchFiles.DoCheckFile: Boolean;
Begin
  if not (soIncludeSystemHiddenFiles in Options) and IsSystemAndHidden(FFindData) then
  Begin
    Result := False;
    Exit;
  End
  else
  if Assigned(FOnCheck) then
  Begin
    // We always want to check the File Params
    Result := FFileParams.Check(FFindData);
    If Result Then
    Begin
      Result := False;
      FOnCheck(Self, Result);
    End;
  End
  Else
    Result := FFileParams.Check(FFindData)
End;

Procedure TJvSearchFiles.DoFindDir(Const APath: String);
Var
  DirName: String;
  FileSize: Int64;
Begin
  Inc(FTotalDirectories);
  With FindData Do
  Begin
    If soStripDirs In Options Then
      DirName := cFileName
    Else
      DirName := APath + cFileName;

    if not (soOwnerData in Options) then
      Directories.Add(DirName);

    Int64Rec(FileSize).Lo := nFileSizeLow;
    Int64Rec(FileSize).Hi := nFileSizeHigh;
    Inc(FTotalFileSize, FileSize);

    { NOTE: soStripDirs also applies to the event }
    If Assigned(FOnFindDirectory) Then
      FOnFindDirectory(Self, DirName);
  End;
End;

Procedure TJvSearchFiles.DoFindFile(Const APath: String);
Var
  FileName: String;
  FileSize: Int64;
Begin
  Inc(FTotalFiles);

  With FindData Do
  Begin
    If soStripDirs In Options Then
      FileName := cFileName
    Else
      FileName := APath + cFileName;

    if not (soOwnerData in Options) then
      Files.Add(FileName);

    Int64Rec(FileSize).Lo := nFileSizeLow;
    Int64Rec(FileSize).Hi := nFileSizeHigh;
    Inc(FTotalFileSize, FileSize);

    { NOTE: soStripDirs also applies to the event }
    If Assigned(FOnFindFile) Then
      FOnFindFile(Self, FileName);
  End;
End;

function TJvSearchFiles.EnumFiles(const ADirectoryName: string;
  Dirs: TStrings; const Search: Boolean): Boolean;
Var
  Handle: THandle;
  Finished: Boolean;
  DirOK: Boolean;
Begin
  DoBeginScanDir(ADirectoryName);

  { Always scan the full directory - ie use * as mask - this seems faster
    then first using a mask, and then scanning the directory for subdirs }
  Handle := FindFirstFile(PChar(ADirectoryName + '*'), FFindData);
  Result := Handle <> INVALID_HANDLE_VALUE;
  If Not Result Then
  Begin
    Result := GetLastError In [ERROR_FILE_NOT_FOUND, ERROR_ACCESS_DENIED];;
    Exit;
  End;

  Finished := False;
  Try
    While Not Finished Do
    Begin
      // (p3) no need to bring in the Forms unit for this:
      If Not IsConsole Then
        DoProgress;
      { After DoProgress, the user can have called Abort,
        so check it }
      If FAborting Then
      Begin
        Result := False;
        Exit;
      End;

      With FFindData Do
        { Is it a directory? }
        If dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY > 0 Then
        Begin
          { Filter out '.' and '..'
            Other dir names can't begin with a '.' }

          {                         | Event | AddDir | SearchInDir
           -----------------------------------------------------------------
            doExcludeSubDirs        |
              True                  |   Y       N           N
              False                 |   N       N           N
            doIncludeSubDirs        |
              True                  |   Y       Y           Y
              False                 |   N       Y           Y
            doExcludeInvalidDirs    |
              True                  |   Y       Y           Y
              False                 |   N       Y           N
            doExcludeCompleteInvalidDirs |
              True                  |   Y       Y           Y
              False                 |   N       N           N
          }
          if not IsDotOrDotDot(cFileName) and
            ((soIncludeSystemHiddenDirs in Options) or not IsSystemAndHidden(FFindData)) then
            { Use case to prevent unnecessary calls to DoCheckDir }
            Case DirOption Of
              doExcludeSubDirs, doIncludeSubDirs:
                Begin
                  If Search And (soSearchDirs In Options) And DoCheckDir Then
                    DoFindDir(ADirectoryName);
                  If DirOption = doIncludeSubDirs Then
                    Dirs.AddObject(cFileName, TObject(True))
                End;
              doExcludeInvalidDirs, doExcludeCompleteInvalidDirs:
                Begin
                  DirOK := DoCheckDir;
                  If Search And (soSearchDirs In Options) And DirOK Then
                    DoFindDir(ADirectoryName);

                  If (DirOption = doExcludeInvalidDirs) Or DirOK Then
                    Dirs.AddObject(cFileName, TObject(DirOK));
                End;
            End;
        End
        else
        if Search and (soSearchFiles in Options) and DoCheckFile then
          DoFindFile(ADirectoryName);

      If Not FindNextFile(Handle, FFindData) Then
      Begin
        Finished := True;
        Result := GetLastError = ERROR_NO_MORE_FILES;
      End;
    End;
  Finally
    Result := FindClose(Handle) And Result;
  End;
End;

Function TJvSearchFiles.GetIsRootDirValid: Boolean;
Var
  Handle: THandle;
Begin
  Handle := FindFirstFile(PChar(ExcludeTrailingPathDelimiter(FRootDirectory)),
    FFindData);
  Result := Handle <> INVALID_HANDLE_VALUE;
  If Not Result Then
    Exit;

  Try
    With FFindData Do
      Result := (dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY > 0) and
        (cFileName[0] <> '.') and DoCheckDir;
  Finally
    FindClose(Handle);
  End;
End;

Function TJvSearchFiles.GetIsDepthAllowed(Const ADepth: Integer): Boolean;
Begin
  Result := (FRecurseDepth = 0) Or (ADepth <= FRecurseDepth);
End;

Function TJvSearchFiles.HandleError: Boolean;
Begin
  { ErrorResponse = erIgnore : Result = True
    ErrorResponse = erAbort  : Result = False
    ErrorResponse = erRaise  : The last error is raised.

    If a user implements an OnError event handler, these results can be
    overridden.
  }
  If FAborting Then
  Begin
    Result := False;
    Exit;
  End;

  Result := FErrorResponse = erIgnore;
  If Assigned(FOnError) Then
    FOnError(Self, Result);
  If (FErrorResponse = erRaise) And Not Result Then
    RaiseLastOSError;
End;

Function TJvSearchFiles.GetDirectories: TStrings;
Begin
  Result := FDirectories;
End;

Function TJvSearchFiles.GetFiles: TStrings;
Begin
  Result := FFiles;
End;

Procedure TJvSearchFiles.Init;
Begin
  FTotalFileSize := 0;
  FTotalDirectories := 0;
  FTotalFiles := 0;
  FTotalFilesAdded := 0;
  Directories.Clear;
  Files.Clear;
  FAborting := False;
End;

function TJvSearchFiles.InternalSearch(const ADirectoryName: string; const Search: Boolean;
  var ADepth: Integer): Boolean;
Var
  List: TStringList;
  DirSep: String;
  I: Integer;
Begin
  FCurrentDepth := ADepth;
  List := TStringList.Create;
  Try
    DirSep := IncludeTrailingPathDelimiter(ADirectoryName);

    Result := EnumFiles(DirSep, List, Search) Or HandleError;
    If Not Result Then
      Exit;

    { DO NOT set Result := False; the search should continue, this is not an error. }
    Inc(ADepth);
    If Not GetIsDepthAllowed(ADepth) Then
      Exit;

    { I think it would be better to do no recursion; Don't know if it can
      be easy implemented - if you want to keep the depth first search -
      and without doing a lot of TList moves }
    For I := 0 To List.Count - 1 Do
    Begin
      Result := InternalSearch(DirSep + List[I], List.Objects[I] <> Nil, ADepth);
      If Not Result Then
        Exit;
    End;
  Finally
    List.Free;
    Dec(ADepth);
  End;
End;

Function TJvSearchFiles.Search: Boolean;
Var
  SearchInRootDir: Boolean;
  ADepth: Integer;
Begin
  Result := False;
  If Searching Then
    Exit;

  Init;

  FSearching := True;
  Try
    { Search in root directory?

                            | soExcludeFiles | soCheckRootDirValid | Else
                            |  InRootDir     |                     |
                            |                |  Valid  | not Valid |
    --------------------------------------------------------------------------
    doExcludeSubDirs        |   No Search    |  True   | No Search | True
    doIncludeSubDirs        |   False        |  True   | False     | True
    doExcludeInvalidDirs    |   False        |  True   | False     | True
    doExcludeCompleteInvalidDirs |   False   |  True   | No Search | True
    }
    SearchInRootDir := not (soExcludeFilesInRootDir in Options) and
      (not (soCheckRootDirValid in Options) or IsRootDirValid);

    if not SearchInRootDir and ((DirOption = doExcludeSubDirs) or
      ((DirOption = doExcludeCompleteInvalidDirs) and
      (soCheckRootDirValid In Options))) Then
    Begin
      Result := True;
      Exit;
    End;

    ADepth := 0;
    Result := InternalSearch(FRootDirectory, SearchInRootDir, ADepth);
  Finally
    FSearching := False;
  End;
End;

Procedure TJvSearchFiles.SetDirParams(Const Value: TJvSearchParams);
Begin
  FDirParams.Assign(Value);
End;

Procedure TJvSearchFiles.SetFileParams(Const Value: TJvSearchParams);
Begin
  FFileParams.Assign(Value);
End;

Procedure TJvSearchFiles.SetCurrentDepth(Const Value: Integer);
Begin
  FCurrentDepth := Value;
End;

Procedure TJvSearchFiles.SetOptions(Const Value: TJvSearchOptions);
Var
  ChangedOptions: TJvSearchOptions;
Begin
  { I'm not sure, what to do when the user changes property Options, while
    the component is searching for files. As implemented now, the component
    just changes the options, and doesn't ensure that the properties hold
    for all data. For example unsetting flag soStripDirs while searching,
    results in a file list with values stripped, and other values not stripped.

    An other option could be to raise an exception when the user tries to
    change Options while the component is searching. But because no serious
    harm is caused - by changing Options, while searching - the component
    doen't do that.
  }
  { (p3) you could also do:
    if Searching then Exit;
  }
  // (rom) even better the search should use a local copy which stays unchanged

  If FOptions <> Value Then
  Begin
    ChangedOptions := FOptions + Value - (FOptions * Value);

    FOptions := Value;

    If soSorted In ChangedOptions Then
    Begin
      FDirectories.Sorted := soSorted In FOptions;
      FFiles.Sorted := soSorted In FOptions;
    End;

    If soAllowDuplicates In ChangedOptions Then
    Begin
      If soAllowDuplicates In FOptions Then
      Begin
        FDirectories.Duplicates := dupAccept;
        FFiles.Duplicates := dupAccept;
      End
      Else
      Begin
        FDirectories.Duplicates := dupIgnore;
        FFiles.Duplicates := dupIgnore;
      End;
    End;
    // soStripDirs; soIncludeSubDirs; soOwnerData
  End;
End;

Procedure TJvSearchFiles.SetTotalFilesAdded(Const Value: Integer);
Begin
  FTotalFilesAdded := Value;
End;

Function TJvSearchFiles.GetRootPath: String;
Begin
  Result := IncludeTrailingPathDelimiter(FRootDirectory);
End;

// === { TJvSearchAttributes } ================================================

Procedure TJvSearchAttributes.Assign(Source: TPersistent);
Begin
  If Source Is TJvSearchAttributes Then
  Begin
    IncludeAttr := TJvSearchAttributes(Source).IncludeAttr;
    ExcludeAttr := TJvSearchAttributes(Source).ExcludeAttr;
  End
  Else
    Inherited Assign(Source);
End;

Procedure TJvSearchAttributes.DefineProperties(Filer: TFiler);
Var
  Ancestor: TJvSearchAttributes;
  Attr: DWORD;
Begin
  Attr := 0;
  Ancestor := TJvSearchAttributes(Filer.Ancestor);
  If Assigned(Ancestor) Then
    Attr := Ancestor.FIncludeAttr;
  Filer.DefineProperty('IncludeAttr', ReadIncludeAttr, WriteIncludeAttr,
    Attr <> FIncludeAttr);
  If Assigned(Ancestor) Then
    Attr := Ancestor.FExcludeAttr;
  Filer.DefineProperty('ExcludeAttr', ReadExcludeAttr, WriteExcludeAttr,
    Attr <> FExcludeAttr);
End;

Function TJvSearchAttributes.GetAttr(Const Index: Integer): TJvAttrFlagKind;
Begin
  If FIncludeAttr And Index > 0 Then
    Result := tsMustBeSet
  else
  if FExcludeAttr and Index > 0 then
    Result := tsMustBeUnSet
  Else
    Result := tsDontCare;
End;

Procedure TJvSearchAttributes.ReadExcludeAttr(Reader: TReader);
Begin
  FExcludeAttr := Reader.ReadInteger;
End;

Procedure TJvSearchAttributes.ReadIncludeAttr(Reader: TReader);
Begin
  FIncludeAttr := Reader.ReadInteger;
End;

procedure TJvSearchAttributes.SetAttr(const Index: Integer;
  Value: TJvAttrFlagKind);
Begin
  Case Value Of
    tsMustBeSet:
      Begin
        FIncludeAttr := FIncludeAttr Or DWORD(Index);
        FExcludeAttr := FExcludeAttr And Not Index;
      End;
    tsMustBeUnSet:
      Begin
        FIncludeAttr := FIncludeAttr And Not Index;
        FExcludeAttr := FExcludeAttr Or DWORD(Index);
      End;
    tsDontCare:
      Begin
        FIncludeAttr := FIncludeAttr And Not Index;
        FExcludeAttr := FExcludeAttr And Not Index;
      End;
  End;
End;

Procedure TJvSearchAttributes.WriteExcludeAttr(Writer: TWriter);
Begin
  Writer.WriteInteger(FExcludeAttr);
End;

Procedure TJvSearchAttributes.WriteIncludeAttr(Writer: TWriter);
Begin
  Writer.WriteInteger(FIncludeAttr);
End;

//=== { TJvSearchParams } ====================================================

Constructor TJvSearchParams.Create;
Begin
  // (rom) added inherited Create
  Inherited Create;
  FAttributes := TJvSearchAttributes.Create;
  FFileMasks := TStringList.Create;
  FFileMasks.OnChange := FileMasksChange;
  FCaseFileMasks := TStringList.Create;

  { defaults }
  FFileMaskSeperator := ';';
  { Set to 1-1-1980 }
  FLastChangeBefore := CDate1_1_1980;
  FLastChangeAfter := CDate1_1_1980;
End;

Destructor TJvSearchParams.Destroy;
Begin
  FAttributes.Free;
  FFileMasks.Free;
  FCaseFileMasks.Free;
  Inherited Destroy;
End;

Procedure TJvSearchParams.Assign(Source: TPersistent);
Var
  Src: TJvSearchParams;
Begin
  If Source Is TJvSearchParams Then
  Begin
    Src := TJvSearchParams(Source);
    MaxSize := Src.MaxSize;
    MinSize := Src.MinSize;
    LastChangeBefore := Src.LastChangeBefore;
    LastChangeAfter := Src.LastChangeAfter;
    SearchTypes := Src.SearchTypes;
    FileMasks.Assign(Src.FileMasks);
    FileMaskSeperator := Src.FileMaskSeperator;
    Attributes.Assign(Src.Attributes);
  End
  Else
    Inherited Assign(Source);
End;

Function TJvSearchParams.Check(Const AFindData: TWin32FindData): Boolean;
Var
  I: Integer;
  FileName: String;
Begin
  Result := False;
  With AFindData Do
  Begin
    If stAttribute In FSearchTypes Then
    Begin
      { Note that if you set a flag in both ExcludeAttr and IncludeAttr
        the search always returns False }
      If dwFileAttributes And Attributes.ExcludeAttr > 0 Then
        Exit;
      If dwFileAttributes And Attributes.IncludeAttr <> Attributes.IncludeAttr Then
        Exit;
    End;

    If stMinSize In FSearchTypes Then
      if (nFileSizeHigh < FMinSizeHigh) or
        ((nFileSizeHigh = FMinSizeHigh) and (nFileSizeLow < FMinSizeLow)) then
        Exit;
    If stMaxSize In FSearchTypes Then
      if (nFileSizeHigh > FMaxSizeHigh) or
        ((nFileSizeHigh = FMaxSizeHigh) and (nFileSizeLow > FMaxSizeLow)) then
        Exit;
    If stLastChangeAfter In FSearchTypes Then
      If CompareFileTime(ftLastWriteTime, FLastChangeAfterFT) < 0 Then
        Exit;
    If stLastChangeBefore In FSearchTypes Then
      If CompareFileTime(ftLastWriteTime, FLastChangeBeforeFT) > 0 Then
        Exit;
    If (stFileMask In FSearchTypes) And (FFileMasks.Count > 0) Then
    Begin
      { StrMatches in JclStrings.pas is case-sensitive, thus for non case-
        sensitive search we have to do a little trick. The filename is
        upper-cased and compared with masks that are also upper-cased.
        This is a bit clumsy; a better solution would be to do this in
        StrMatches.

        I guess a lot of masks have the format 'mask*' or '*.ext'; so
        if you could specifiy to do a left or right scan in StrMatches
        would be better too. Note that if no char follows a '*', the
        result is always true; this isn't implemented so in StrMatches }

      If stFileMaskCaseSensitive In SearchTypes Then
        FileName := cFileName
      Else
        FileName := AnsiUpperCase(cFileName);

      I := 0;
      while (I < FFileMasks.Count) and
        not JclStrings.StrMatches(FCaseFileMasks[I], FileName) do
        Inc(I);
      If I >= FFileMasks.Count Then
        Exit;
    End;
  End;
  Result := True;
End;

Procedure TJvSearchParams.FileMasksChange(Sender: TObject);
Begin
  UpdateCaseMasks;
End;

Function TJvSearchParams.GetFileMask: String;
Begin
  Result := JclStrings.StringsToStr(FileMasks, FileMaskSeperator);
End;

Function TJvSearchParams.GetMaxSize: Int64;
Begin
  Int64Rec(Result).Lo := FMaxSizeLow;
  Int64Rec(Result).Hi := FMaxSizeHigh;
End;

Function TJvSearchParams.GetMinSize: Int64;
Begin
  Int64Rec(Result).Lo := FMinSizeLow;
  Int64Rec(Result).Hi := FMinSizeHigh;
End;

Function TJvSearchParams.GetFileMasks: TStrings;
Begin
  Result := FFileMasks;
End;

Function TJvSearchParams.IsLastChangeAfterStored: Boolean;
Begin
  Result := FLastChangeBefore <> CDate1_1_1980;
End;

Function TJvSearchParams.IsLastChangeBeforeStored: Boolean;
Begin
  Result := FLastChangeBefore <> CDate1_1_1980;
End;

Procedure TJvSearchParams.SetAttributes(Const Value: TJvSearchAttributes);
Begin
  FAttributes.Assign(Value);
End;

Procedure TJvSearchParams.SetFileMask(Const Value: String);
Begin
  JclStrings.StrToStrings(Value, FileMaskSeperator, FileMasks);
End;

Procedure TJvSearchParams.SetFileMasks(Const Value: TStrings);
Begin
  FFileMasks.Assign(Value);
End;

Procedure TJvSearchParams.SetLastChangeAfter(Const Value: TDateTime);
Var
  DosFileTime: Longint;
  LocalFileTime: TFileTime;
Begin
  { Value must be >= 1-1-1980 }
  DosFileTime := DateTimeToDosDateTime(Value);
  if not Windows.DosDateTimeToFileTime(LongRec(DosFileTime).Hi,
    LongRec(DosFileTime).Lo, LocalFileTime) or
    Not Windows.LocalFileTimeToFileTime(LocalFileTime, FLastChangeAfterFT) Then
    RaiseLastOSError;

  FLastChangeAfter := Value;
End;

Procedure TJvSearchParams.SetLastChangeBefore(Const Value: TDateTime);
Var
  DosFileTime: Longint;
  LocalFileTime: TFileTime;
Begin
  { Value must be >= 1-1-1980 }
  DosFileTime := DateTimeToDosDateTime(Value);
  if not Windows.DosDateTimeToFileTime(LongRec(DosFileTime).Hi,
    LongRec(DosFileTime).Lo, LocalFileTime) or
    Not Windows.LocalFileTimeToFileTime(LocalFileTime, FLastChangeBeforeFT) Then
    RaiseLastOSError;

  FLastChangeBefore := Value;
End;

Procedure TJvSearchParams.SetMaxSize(Const Value: Int64);
Begin
  FMaxSizeHigh := Int64Rec(Value).Hi;
  FMaxSizeLow := Int64Rec(Value).Lo;
End;

Procedure TJvSearchParams.SetMinSize(Const Value: Int64);
Begin
  FMinSizeHigh := Int64Rec(Value).Hi;
  FMinSizeLow := Int64Rec(Value).Lo;
End;

Procedure TJvSearchParams.SetSearchTypes(Const Value: TJvSearchTypes);
Var
  ChangedValues: TJvSearchTypes;
Begin
  If FSearchTypes = Value Then
    Exit;

  ChangedValues := FSearchTypes + Value - (FSearchTypes * Value);
  FSearchTypes := Value;

  If stFileMaskCaseSensitive In ChangedValues Then
    UpdateCaseMasks;
End;

Procedure TJvSearchParams.UpdateCaseMasks;
Var
  I: Integer;
Begin
  FCaseFileMasks.Assign(FileMasks);

  if not (stFileMaskCaseSensitive in SearchTypes) then
    For I := 0 To FCaseFileMasks.Count - 1 Do
      FCaseFileMasks[I] := AnsiUpperCase(FCaseFileMasks[I]);
End;

{$IFDEF UNITVERSIONING}

Initialization
  RegisterUnitVersion(HInstance, UnitVersioning);

Finalization
  UnregisterUnitVersion(HInstance);
{$ENDIF UNITVERSIONING}

End.
