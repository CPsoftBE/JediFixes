{ -----------------------------------------------------------------------------
  The contents of this file are subject to the Mozilla Public License
  Version 1.1 (the "License"); you may not use this file except in compliance
  with the License. You may obtain a copy of the License at
  http://www.mozilla.org/MPL/MPL-1.1.html

  Software distributed under the License is distributed on an "AS IS" basis,
  WITHOUT WARRANTY OF ANY KIND, either expressed or implied. See the License for
  the specific language governing rights and limitations under the License.

  The Original Code is: JvAppDBStorage.pas, released on 2004-02-04.

  The Initial Developer of the Original Code is Peter Thörnqvist
  Portions created by Peter Thörnqvist are Copyright (C) 2004 Peter Thörnqvist
  All Rights Reserved.

  Contributor(s):

  You may retrieve the latest version of this file at the Project JEDI's JVCL home page,
  located at http://jvcl.delphi-jedi.org

  Known Issues:
  ----------------------------------------------------------------------------- }
// $Id$

Unit JvAppDBStorage;

{$I jvcl.inc}

Interface

Uses
{$IFDEF UNITVERSIONING}
  JclUnitVersioning,
{$ENDIF UNITVERSIONING}
  SysUtils, Classes, DB, Variants, DBCtrls,
  JvAppStorage, JvTypes;

// DB table must contain 3 fields for the storage
// performance is probably improved if there is an index on the section and key fields (this can be unique)
// "section": string   - must support locate!
// "key": string      - must support locate!
// "value": string or memo

Type
  TJvDBStorageWriteEvent = Procedure(Sender: TObject; Const Section, Key, Value: String) Of Object;
  TJvDBStorageReadEvent = Procedure(Sender: TObject; Const Section, Key: String; Var Value: String) Of Object;
  EJvAppDBStorageError = Class(Exception);

  TJvCustomAppDBStorage = Class(TJvCustomAppStorage)
  Private
    FSectionLink: TFieldDataLink;
    FKeyLink: TFieldDataLink;
    FValueLink: TFieldDataLink;
    FOnRead: TJvDBStorageReadEvent;
    FOnWrite: TJvDBStorageWriteEvent;
    FBookmark: {$IFDEF RTL200_UP}TBookmark{$ELSE}TBookmarkStr{$ENDIF RTL200_UP};
    FDataSource: TDataSource;
    Procedure SetDataSource(Const Value: TDataSource);
    Function GetKeyField: String;
    Function GetSectionField: String;
    Function GetValueField: String;
    Procedure SetKeyField(Const Value: String);
    Procedure SetSectionField(Const Value: String);
    Procedure SetValueField(Const Value: String);
  Protected
    Function FieldsAssigned: Boolean;
    Procedure EnumFolders(Const Path: String; Const Strings: TStrings; Const ReportListAsValue: Boolean = True); Override;
    Procedure EnumValues(Const Path: String; Const Strings: TStrings; Const ReportListAsValue: Boolean = True); Override;
    Function PathExistsInt(Const Path: String): Boolean; Override;
    Function IsFolderInt(Const Path: String; ListIsValue: Boolean = True): Boolean; Override;
    Procedure RemoveValue(Const Section, Key: String);
    Procedure DeleteSubTreeInt(Const Path: String); Override;

    Function ValueStoredInt(Const Path: String): Boolean; Override;
    Procedure DeleteValueInt(Const Path: String); Override;
    Function DoReadInteger(Const Path: String; Default: Integer): Integer; Override;
    Procedure DoWriteInteger(Const Path: String; Value: Integer); Override;
    Function DoReadFloat(Const Path: String; Default: Extended): Extended; Override;
    Procedure DoWriteFloat(Const Path: String; Value: Extended); Override;
    Function DoReadString(Const Path: String; Const Default: String): String; Override;
    Procedure DoWriteString(Const Path: String; Const Value: String); Override;
    Function DoReadBinary(Const Path: String; Buf: TJvBytes; BufSize: Integer): Integer; Override;
    Procedure DoWriteBinary(Const Path: String; Const Buf: TJvBytes; BufSize: Integer); Override;
    Procedure Notification(AComponent: TComponent; Operation: TOperation); Override;
    Function SectionExists(Const Path: String; RestorePosition: Boolean; LocateOptions: TLocateOptions): Boolean;
    Function ValueExists(Const Section, Key: String; RestorePosition: Boolean): Boolean;
    Function ReadValue(Const Section, Key: String): String; Virtual;
    Procedure WriteValue(Const Section, Key, Value: String); Virtual;
    Procedure StoreDataset;
    Procedure RestoreDataset;
    Function GetPhysicalReadOnly: Boolean; Override;
  Public
    Constructor Create(AOwner: TComponent); Override;
    Destructor Destroy; Override;
    Procedure DeleteRuleTree(Const Path: String);
  Protected
    Property DataSource: TDataSource Read FDataSource Write SetDataSource;
    Property KeyField: String Read GetKeyField Write SetKeyField;
    Property SectionField: String Read GetSectionField Write SetSectionField;
    Property ValueField: String Read GetValueField Write SetValueField;
    Property OnRead: TJvDBStorageReadEvent Read FOnRead Write FOnRead;
    Property OnWrite: TJvDBStorageWriteEvent Read FOnWrite Write FOnWrite;
  End;

{$IFDEF RTL230_UP}
  [ComponentPlatformsAttribute(pidWin32 Or pidWin64)]
{$ENDIF RTL230_UP}

  TJvAppDBStorage = Class(TJvCustomAppDBStorage)
  Published
    Property ReadOnly;

    Property DataSource;
    Property FlushOnDestroy;
    Property KeyField;
    Property SectionField;
    Property SubStorages;
    Property ValueField;

    Property OnRead;
    Property OnWrite;
  End;

{$IFDEF UNITVERSIONING}

Const
  UnitVersioning: TUnitVersionInfo = (RCSfile: '$URL$'; Revision: '$Revision$'; Date: '$Date$'; LogPath: 'JVCL\run');
{$ENDIF UNITVERSIONING}

Implementation

Uses
{$IFDEF SUPPORTS_INLINE}
  Windows,
{$ENDIF SUPPORTS_INLINE}
  JclMime,
  JvJCLUtils, JvResources, JclStrings, JvJVCLUtils;

Constructor TJvCustomAppDBStorage.Create(AOwner: TComponent);
Begin
  // (p3) create these before calling inherited (AV's otherwise)
  FSectionLink := TFieldDataLink.Create;
  FKeyLink := TFieldDataLink.Create;
  FValueLink := TFieldDataLink.Create;
  Inherited Create(AOwner);
End;

Destructor TJvCustomAppDBStorage.Destroy;
Begin
  DataSource := Nil;
  FreeAndNil(FSectionLink);
  FreeAndNil(FKeyLink);
  FreeAndNil(FValueLink);
  Inherited Destroy;
End;

Procedure TJvCustomAppDBStorage.DeleteRuleTree(Const Path: String);
Begin
  If Not ReadOnly Then
    DeleteSubTreeInt(StrEnsureSuffix(PathDelim, Path));
End;

Procedure TJvCustomAppDBStorage.DeleteSubTreeInt(Const Path: String);
Begin
  If FieldsAssigned Then
  Begin
    StoreDataset;
    Try
      While SectionExists(StrEnsureNoPrefix(PathDelim, Path), False, [loCaseInsensitive, loPartialKey]) Do
        DataSource.DataSet.Delete;
    Finally
      RestoreDataset;
    End;
  End;
End;

Procedure TJvCustomAppDBStorage.DeleteValueInt(Const Path: String);
Var
  Section: String;
  Key: String;
Begin
  SplitKeyPath(Path, Section, Key);
  If FieldsAssigned Then
  Begin
    StoreDataset;
    Try
      While ValueExists(Section, Key, False) Do
        DataSource.DataSet.Delete;
    Finally
      RestoreDataset;
    End;
  End;
End;

Function TJvCustomAppDBStorage.DoReadBinary(Const Path: String; Buf: TJvBytes; BufSize: Integer): Integer;
Var
  Value: AnsiString;
Begin
  Raise EJvAppDBStorageError.CreateRes(@RsENotSupported);
  // TODO -cTESTING -oJVCL: NOT TESTED!!!
  Value := JclMime.MimeDecodeString(AnsiString(DoReadString(Path, ''))); // the cast to AnsiString converts with loss under D2009
  Result := Length(Value);
  If Result > BufSize Then
    Raise EJvAppDBStorageError.CreateResFmt(@RsEBufTooSmallFmt, [Result]);
  If Length(Value) > 0 Then
    Move(Value[1], Buf, Result * SizeOf(AnsiChar));
End;

Function TJvCustomAppDBStorage.DoReadFloat(Const Path: String; Default: Extended): Extended;
Begin
  // NOTE: StrToFloatDefIgnoreInvalidCharacters now called JvSafeStrToFloatDef:
  Result := JvSafeStrToFloatDef(DoReadString(Path, ''), Default);
End;

Function TJvCustomAppDBStorage.DoReadInteger(Const Path: String; Default: Integer): Integer;
Begin
  Result := StrToIntDef(DoReadString(Path, ''), Default);
End;

Function TJvCustomAppDBStorage.DoReadString(Const Path: String; Const Default: String): String;
Var
  Section: String;
  Key: String;
Begin
  SplitKeyPath(Path, Section, Key);
  If ValueExists(Section, Key, False) Then
    Result := ReadValue(Section, Key)
  Else
    Result := Default;
End;

Procedure TJvCustomAppDBStorage.DoWriteBinary(Const Path: String; Const Buf: TJvBytes; BufSize: Integer);
Var
  Value, Buf1: AnsiString;
Begin
  Raise EJvAppDBStorageError.CreateRes(@RsENotSupported);
  // TODO -cTESTING -oJVCL: NOT TESTED!!!
  SetLength(Value, BufSize);
  If BufSize > 0 Then
  Begin
    SetLength(Buf1, BufSize);
    Move(Buf, Buf1[1], BufSize);
    JclMime.MimeEncode(Buf1[1], BufSize, Value[1]);
    DoWriteString(Path, String(Value));
  End;
End;

Procedure TJvCustomAppDBStorage.DoWriteFloat(Const Path: String; Value: Extended);
Begin
  WriteBinary(Path, @Value, SizeOf(Value));
End;

Procedure TJvCustomAppDBStorage.DoWriteInteger(Const Path: String; Value: Integer);
Begin
  DoWriteString(Path, IntToStr(Value));
End;

Procedure TJvCustomAppDBStorage.DoWriteString(Const Path: String; Const Value: String);
Var
  Section: String;
  Key: String;
Begin
  SplitKeyPath(Path, Section, Key);
  WriteValue(Section, Key, Value);
End;

Procedure TJvCustomAppDBStorage.EnumFolders(Const Path: String; Const Strings: TStrings; Const ReportListAsValue: Boolean);
Begin
  Raise EJvAppDBStorageError.CreateRes(@RsENotSupported);
End;

Procedure TJvCustomAppDBStorage.EnumValues(Const Path: String; Const Strings: TStrings; Const ReportListAsValue: Boolean);
Begin
  Raise EJvAppDBStorageError.CreateRes(@RsENotSupported);
End;

Function TJvCustomAppDBStorage.FieldsAssigned: Boolean;
Begin
  Result := (FSectionLink.Field <> Nil) And (FKeyLink.Field <> Nil) And (FValueLink.Field <> Nil);
End;

Function TJvCustomAppDBStorage.GetKeyField: String;
Begin
  Result := FKeyLink.FieldName;
End;

Function TJvCustomAppDBStorage.GetSectionField: String;
Begin
  Result := FSectionLink.FieldName;
End;

Function TJvCustomAppDBStorage.GetValueField: String;
Begin
  Result := FValueLink.FieldName;
End;

Function TJvCustomAppDBStorage.IsFolderInt(Const Path: String; ListIsValue: Boolean): Boolean;
Begin
  { TODO -oJVCL -cTESTING : Is this correct implementation? }
  Result := SectionExists(StrEnsureNoPrefix(PathDelim, Path), True, [loCaseInsensitive]);
End;

Procedure TJvCustomAppDBStorage.Notification(AComponent: TComponent; Operation: TOperation);
Begin
  Inherited Notification(AComponent, Operation);
  If (Operation = opRemove) And Not(csDestroying In ComponentState) Then
    If AComponent = DataSource Then
      DataSource := Nil;
End;

Function TJvCustomAppDBStorage.PathExistsInt(Const Path: String): Boolean;
Begin
  { TODO -oJVCL -cTESTING : Is this correct implementation? }
  Result := SectionExists(StrEnsureNoPrefix(PathDelim, Path), True, [loCaseInsensitive]);
End;

Function TJvCustomAppDBStorage.ReadValue(Const Section, Key: String): String;
Begin
  If ValueExists(Section, Key, False) Then
    Result := FValueLink.Field.AsString
  Else
    Result := '';
  // always call event
  If Assigned(FOnRead) Then
    FOnRead(Self, Section, Key, Result);
End;

Procedure TJvCustomAppDBStorage.RemoveValue(Const Section, Key: String);
Begin
  { TODO -oJVCL -cTESTING : NOT TESTED!!! }
  If ValueExists(Section, Key, False) Then
    FValueLink.Field.Clear;
End;

Procedure TJvCustomAppDBStorage.RestoreDataset;
Begin
  If FBookmark = {$IFDEF RTL200_UP}Nil{$ELSE}''{$ENDIF RTL200_UP} Then
    Exit;
  If FieldsAssigned Then
    DataSource.DataSet.Bookmark := FBookmark;
  FBookmark := {$IFDEF RTL200_UP}Nil{$ELSE}''{$ENDIF RTL200_UP};
End;

Function TJvCustomAppDBStorage.GetPhysicalReadOnly: Boolean;
Begin
  If FieldsAssigned Then
    Result := False
  Else
    Result := Not DataSource.DataSet.CanModify;
End;

Function TJvCustomAppDBStorage.SectionExists(Const Path: String; RestorePosition: Boolean; LocateOptions: TLocateOptions): Boolean;
Begin
  Result := FieldsAssigned And DataSource.DataSet.Active;
  If Result Then
  Begin
    If RestorePosition Then
      StoreDataset;
    Try
      Result := DataSource.DataSet.Locate(SectionField, Path, LocateOptions);
    Finally
      If RestorePosition Then
        RestoreDataset;
    End;
  End;
End;

Procedure TJvCustomAppDBStorage.SetDataSource(Const Value: TDataSource);
Begin
  If Assigned(FSectionLink) And Not(FSectionLink.DataSourceFixed And (csLoading In ComponentState)) Then
  Begin
    FSectionLink.DataSource := Value;
    FKeyLink.DataSource := Value;
    FValueLink.DataSource := Value;
  End;
  ReplaceComponentReference(Self, Value, TComponent(FDataSource));
End;

Procedure TJvCustomAppDBStorage.SetKeyField(Const Value: String);
Begin
  FKeyLink.FieldName := Value;
End;

Procedure TJvCustomAppDBStorage.SetSectionField(Const Value: String);
Begin
  FSectionLink.FieldName := Value;
End;

Procedure TJvCustomAppDBStorage.SetValueField(Const Value: String);
Begin
  FValueLink.FieldName := Value;
End;

Procedure TJvCustomAppDBStorage.StoreDataset;
Begin
  If FBookmark <> {$IFDEF RTL200_UP}Nil{$ELSE}''{$ENDIF RTL200_UP} Then
    RestoreDataset;
  If FieldsAssigned And DataSource.DataSet.Active Then
  Begin
    FBookmark := DataSource.DataSet.Bookmark;
    DataSource.DataSet.DisableControls;
  End;
End;

Function TJvCustomAppDBStorage.ValueExists(Const Section, Key: String; RestorePosition: Boolean): Boolean;
Begin
  Result := FieldsAssigned And DataSource.DataSet.Active;
  If Result Then
  Begin
    If RestorePosition Then
      StoreDataset;
    Try
      Result := DataSource.DataSet.Locate(Format('%s;%s', [SectionField, KeyField]), VarArrayOf([Section, Key]), [loCaseInsensitive]);
    Finally
      If RestorePosition Then
        RestoreDataset;
    End;
  End;
End;

Function TJvCustomAppDBStorage.ValueStoredInt(Const Path: String): Boolean;
Var
  Section: String;
  Key: String;
Begin
  SplitKeyPath(Path, Section, Key);
  Result := ValueExists(Section, Key, True);
End;

Procedure TJvCustomAppDBStorage.WriteValue(Const Section, Key, Value: String);
Begin
  If FieldsAssigned Then
  Begin
    If ValueExists(Section, Key, False) Then
    Begin
      If AnsiSameStr(FValueLink.Field.AsString, Value) Then
        Exit; // don't save if it's the same value (NB: this also skips the event)
      DataSource.DataSet.Edit
    End
    Else
      DataSource.DataSet.Append;
    FSectionLink.Field.AsString := Section;
    FKeyLink.Field.AsString := Key;
    FValueLink.Field.AsString := Value;
    DataSource.DataSet.Post;
  End;
  // always call event
  If Assigned(FOnWrite) Then
    FOnWrite(Self, Section, Key, Value);
End;

{$IFDEF UNITVERSIONING}

Initialization

RegisterUnitVersion(HInstance, UnitVersioning);

Finalization

UnregisterUnitVersion(HInstance);
{$ENDIF UNITVERSIONING}

End.
