{ -----------------------------------------------------------------------------

  Project JEDI Visible Component Library (J-VCL)

  The contents of this file are subject to the Mozilla Public License Version
  1.1 (the "License"); you may not use this file except in compliance with the
  License. You may obtain a copy of the License at http://www.mozilla.org/MPL/

  Software distributed under the License is distributed on an "AS IS" basis,
  WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for
  the specific language governing rights and limitations under the License.

  The Initial Developer of the Original Code is Marcel Bestebroer
  <marcelb att zeelandnet dott nl>.
  Portions created by Marcel Bestebroer are Copyright (C) 2000 - 2002 mbeSoft.
  All Rights Reserved.

  ******************************************************************************

  Event scheduling component. Allows to schedule execution of events, with
  optional recurring schedule options.

  You may retrieve the latest version of this file at the Project JEDI home
  page, located at http://www.delphi-jedi.org
  ----------------------------------------------------------------------------- }
// $Id$

Unit JvScheduledEvents;

{$I jvcl.inc}

Interface

Uses
{$IFDEF UNITVERSIONING}
  JclUnitVersioning,
{$ENDIF UNITVERSIONING}
  SysUtils, Classes, Contnrs, SyncObjs,
{$IFDEF MSWINDOWS}
  Windows,
{$ENDIF MSWINDOWS}
  Messages, Forms,
  JclSchedule,
  JvAppStorage;

Const
  CM_EXECEVENT = WM_USER + $1000;

Type
  TJvCustomScheduledEvents = Class;
  TJvEventCollection = Class;
  TJvEventCollectionItem = Class;

  TScheduledEventState = (sesNotInitialized, sesWaiting, sesTriggered, sesExecuting, sesPaused, sesEnded);

  TScheduledEventStateInfo = Record
    { Common }
    ARecurringType: TScheduleRecurringKind;
    AStartDate: TTimeStamp;
    AEndType: TScheduleEndKind;
    AEndDate: TTimeStamp;
    AEndCount: Cardinal;
    ALastTriggered: TTimeStamp;

    { DayFrequency }
    DayFrequence: Record
      ADayFrequencyStartTime: Cardinal;
      ADayFrequencyEndTime: Cardinal;
      ADayFrequencyInterval: Cardinal;
    End;

    { Daily }
    Daily: Record
      ADayEveryWeekDay: Boolean;
      ADayInterval: Cardinal;
    End;

    { Weekly }
    Weekly: Record
      AWeekInterval: Cardinal;
      AWeekDaysOfWeek: TScheduleWeekDays;
    End;

    { Monthly }
    Monthly: Record
      AMonthIndexKind: TScheduleIndexKind;
      AMonthIndexValue: Cardinal;
      AMonthDay: Cardinal;
      AMonthInterval: Cardinal;
    End;

    { Yearly }
    Yearly: Record
      AYearIndexKind: TScheduleIndexKind;
      AYearIndexValue: Cardinal;
      AYearDay: Cardinal;
      AYearMonth: Cardinal;
      AYearInterval: Cardinal;
    End;
  End;

  TScheduledEventExecute = Procedure(Sender: TJvEventCollectionItem; Const IsSnoozeEvent: Boolean) Of Object;

  TJvCustomScheduledEvents = Class(TComponent)
  Private
    FAppStorage: TJvCustomAppStorage;
    FAppStoragePath: String;
    FAutoSave: Boolean;
    FEvents: TJvEventCollection;
    FPostedEvents: TList;
    FEventsPosted: Boolean;
    FOnStartEvent: TNotifyEvent;
    FOnEndEvent: TNotifyEvent;
    FWnd: THandle;
  Protected
    Procedure Notification(AComponent: TComponent; Operation: TOperation); Override;
    Procedure DoEndEvent(Const Event: TJvEventCollectionItem);
    Procedure DoStartEvent(Const Event: TJvEventCollectionItem);
    Procedure SetAppStorage(Value: TJvCustomAppStorage);
    Function GetEvents: TJvEventCollection;
    Procedure PostEvent(Event: TJvEventCollectionItem);
    Procedure RemovePostedEvent(Event: TJvEventCollectionItem);
    Procedure InitEvents;
    Procedure Loaded; Override;
    Procedure LoadSingleEvent(Sender: TJvCustomAppStorage; Const Path: String; Const List: TObject; Const Index: Integer;
      Const ItemName: String);
    Procedure SaveSingleEvent(Sender: TJvCustomAppStorage; Const Path: String; Const List: TObject; Const Index: Integer;
      Const ItemName: String);
    Procedure DeleteSingleEvent(Sender: TJvCustomAppStorage; Const Path: String; Const List: TObject; Const First, Last: Integer;
      Const ItemName: String);
    Procedure SetEvents(Value: TJvEventCollection);
    Procedure WndProc(Var Msg: TMessage); Virtual;
    Procedure CMExecEvent(Var Msg: TMessage); Message CM_EXECEVENT;
    Property AutoSave: Boolean Read FAutoSave Write FAutoSave;
    Property OnStartEvent: TNotifyEvent Read FOnStartEvent Write FOnStartEvent;
    Property OnEndEvent: TNotifyEvent Read FOnEndEvent Write FOnEndEvent;
    Property AppStorage: TJvCustomAppStorage Read FAppStorage Write SetAppStorage;
    Property AppStoragePath: String Read FAppStoragePath Write FAppStoragePath;
  Public
{$IFDEF SUPPORTS_CLASS_CTORDTORS}
    Class Destructor Destroy;
{$ENDIF SUPPORTS_CLASS_CTORDTORS}
    Constructor Create(AOwner: TComponent); Override;
    Destructor Destroy; Override;
    Property Handle: THandle Read FWnd;
    Property Events: TJvEventCollection Read GetEvents Write SetEvents;
    Procedure LoadEventStates(Const ClearBefore: Boolean = True);
    Procedure SaveEventStates;
    Procedure StartAll;
    Procedure StopAll;
    Procedure PauseAll;
  End;

{$IFDEF RTL230_UP}
  [ComponentPlatformsAttribute(pidWin32 Or pidWin64)]
{$ENDIF RTL230_UP}

  TJvScheduledEvents = Class(TJvCustomScheduledEvents)
  Published
    Property AppStorage;
    Property AppStoragePath;
    Property AutoSave;
    Property Events;
    Property OnStartEvent;
    Property OnEndEvent;
  End;

  TJvEventCollection = Class(TOwnedCollection)
  Protected
    Function GetItem(Index: Integer): TJvEventCollectionItem;
    Procedure SetItem(Index: Integer; Value: TJvEventCollectionItem);
    Procedure Notify(Item: TCollectionItem; Action: TCollectionNotification); Override;
  Public
    Constructor Create(AOwner: TPersistent);
    Function Add: TJvEventCollectionItem;
    Function Insert(Index: Integer): TJvEventCollectionItem;
    Property Items[Index: Integer]: TJvEventCollectionItem Read GetItem Write SetItem; Default;
  End;

  TJvEventCollectionItem = Class(TCollectionItem)
  Private
    FCountMissedEvents: Boolean;
    FName: String;
    FState: TScheduledEventState;
    FData: Pointer;
    FOnExecute: TScheduledEventExecute;
    FSchedule: IJclSchedule;
    FLastSnoozeInterval: TSystemTime;
    FScheduleFire: TTimeStamp;
    FSnoozeFire: TTimeStamp;
    FReqTriggerTime: TTimeStamp;
    FActualTriggerTime: TTimeStamp;
    Procedure Triggered;
  Protected
    Procedure DefineProperties(Filer: TFiler); Override;
    Procedure DoExecute(Const IsSnoozeFire: Boolean);
    Function GetDisplayName: String; Override;
    Function GetNextFire: TTimeStamp;
    Procedure Execute; Virtual;
    // schedule property readers/writers
    Procedure PropDateRead(Reader: TReader; Var Stamp: TTimeStamp);
    Procedure PropDateWrite(Writer: TWriter; Const Stamp: TTimeStamp);
    Procedure PropDailyEveryWeekDayRead(Reader: TReader);
    Procedure PropDailyEveryWeekDayWrite(Writer: TWriter);
    Procedure PropDailyIntervalRead(Reader: TReader);
    Procedure PropDailyIntervalWrite(Writer: TWriter);
    Procedure PropEndCountRead(Reader: TReader);
    Procedure PropEndCountWrite(Writer: TWriter);
    Procedure PropEndDateRead(Reader: TReader);
    Procedure PropEndDateWrite(Writer: TWriter);
    Procedure PropEndTypeRead(Reader: TReader);
    Procedure PropEndTypeWrite(Writer: TWriter);
    Procedure PropFreqEndTimeRead(Reader: TReader);
    Procedure PropFreqEndTimeWrite(Writer: TWriter);
    Procedure PropFreqIntervalRead(Reader: TReader);
    Procedure PropFreqIntervalWrite(Writer: TWriter);
    Procedure PropFreqStartTimeRead(Reader: TReader);
    Procedure PropFreqStartTimeWrite(Writer: TWriter);
    Procedure PropMonthlyDayRead(Reader: TReader);
    Procedure PropMonthlyDayWrite(Writer: TWriter);
    Procedure PropMonthlyIndexKindRead(Reader: TReader);
    Procedure PropMonthlyIndexKindWrite(Writer: TWriter);
    Procedure PropMonthlyIndexValueRead(Reader: TReader);
    Procedure PropMonthlyIndexValueWrite(Writer: TWriter);
    Procedure PropMonthlyIntervalRead(Reader: TReader);
    Procedure PropMonthlyIntervalWrite(Writer: TWriter);
    Procedure PropRecurringTypeRead(Reader: TReader);
    Procedure PropRecurringTypeWrite(Writer: TWriter);
    Procedure PropStartDateRead(Reader: TReader);
    Procedure PropStartDateWrite(Writer: TWriter);
    Procedure PropWeeklyDaysOfWeekRead(Reader: TReader);
    Procedure PropWeeklyDaysOfWeekWrite(Writer: TWriter);
    Procedure PropWeeklyIntervalRead(Reader: TReader);
    Procedure PropWeeklyIntervalWrite(Writer: TWriter);
    Procedure PropYearlyDayRead(Reader: TReader);
    Procedure PropYearlyDayWrite(Writer: TWriter);
    Procedure PropYearlyIndexKindRead(Reader: TReader);
    Procedure PropYearlyIndexKindWrite(Writer: TWriter);
    Procedure PropYearlyIndexValueRead(Reader: TReader);
    Procedure PropYearlyIndexValueWrite(Writer: TWriter);
    Procedure PropYearlyIntervalRead(Reader: TReader);
    Procedure PropYearlyIntervalWrite(Writer: TWriter);
    Procedure PropYearlyMonthRead(Reader: TReader);
    Procedure PropYearlyMonthWrite(Writer: TWriter);
    Procedure SetName(Value: String);
  Public
    Constructor Create(Collection: TCollection); Override;
    Destructor Destroy; Override;
    Procedure Assign(Source: TPersistent); Override;
    Procedure LoadState(Const TriggerStamp: TTimeStamp; Const TriggerCount, DayCount: Integer; Const SnoozeStamp: TTimeStamp;
      Const ALastSnoozeInterval: TSystemTime; Const AEventInfo: TScheduledEventStateInfo); Virtual;
    Procedure Pause;
    Procedure SaveState(Out TriggerStamp: TTimeStamp; Out TriggerCount, DayCount: Integer; Out SnoozeStamp: TTimeStamp;
      Out ALastSnoozeInterval: TSystemTime; Out AEventInfo: TScheduledEventStateInfo); Virtual;
    Procedure Snooze(Const MSecs: Word; Const Secs: Word = 0; Const Mins: Word = 0; Const Hrs: Word = 0; Const Days: Word = 0);
    Procedure Start;
    Procedure Stop;
    Property Data: Pointer Read FData Write FData;
    Property LastSnoozeInterval: TSystemTime Read FLastSnoozeInterval;
    Property NextFire: TTimeStamp Read GetNextFire;
    Property State: TScheduledEventState Read FState;
    Property NextScheduleFire: TTimeStamp Read FScheduleFire;
    Property RequestedTriggerTime: TTimeStamp Read FReqTriggerTime;
    Property ActualTriggerTime: TTimeStamp Read FActualTriggerTime;
  Published
    Property CountMissedEvents: Boolean Read FCountMissedEvents Write FCountMissedEvents Default False;
    Property Name: String Read FName Write SetName;
    Property Schedule: IJclSchedule Read FSchedule Write FSchedule Stored False;
    Property OnExecute: TScheduledEventExecute Read FOnExecute Write FOnExecute;
  End;

{$IFDEF UNITVERSIONING}

Const
  UnitVersioning: TUnitVersionInfo = (RCSfile: '$URL$'; Revision: '$Revision$'; Date: '$Date$'; LogPath: 'JVCL\run');
{$ENDIF UNITVERSIONING}

Implementation

Uses
  TypInfo,
  JclDateTime, JclRTTI,
  JvJVCLUtils, JvResources, JvTypes;

Const
  cEventPrefix = 'Event ';

  // === { TScheduleThread } ====================================================

Type
  TScheduleThread = Class(TJvCustomThread)
  Private
    FCritSect: TCriticalSection;
    FEnded: Boolean;
    FEventComponents: TComponentList;
    FEventIdx: Integer;
  Protected
    Procedure Execute; Override;
  Public
    Constructor Create;
    Destructor Destroy; Override;
    Procedure BeforeDestruction; Override;
    Procedure AddEventComponent(Const AComp: TJvCustomScheduledEvents);
    Procedure RemoveEventComponent(Const AComp: TJvCustomScheduledEvents);
    Procedure Lock;
    Procedure Unlock;
    Property Ended: Boolean Read FEnded;
  End;

Constructor TScheduleThread.Create;
Begin
  Inherited Create(True);
  FCritSect := TCriticalSection.Create;
  FEventComponents := TComponentList.Create(False);
End;

Destructor TScheduleThread.Destroy;
Begin
  Inherited Destroy;
  FreeAndNil(FCritSect);
End;

Procedure TScheduleThread.Execute;
Var
  TskColl: TJvEventCollection;
  I: Integer;
  SysTime: TSystemTime;
  NowStamp: TTimeStamp;
  SchedEvents: TJvCustomScheduledEvents;
Begin
  NameThread(ThreadName);
  Try
    FEnded := False;
    While Not Terminated Do
    Begin
      If (FCritSect <> Nil) And (FEventComponents <> Nil) Then
      Begin
        FCritSect.Enter;
        Try
          FEventIdx := FEventComponents.Count - 1;
          While (FEventIdx > -1) And Not Terminated Do
          Begin
            GetLocalTime(SysTime);
            NowStamp := DateTimeToTimeStamp(Now);
            NowStamp.Time := SysTime.wHour * 3600000 + SysTime.wMinute * 60000 + SysTime.wSecond * 1000 + SysTime.wMilliseconds;
            SchedEvents := TJvCustomScheduledEvents(FEventComponents[FEventIdx]);
            TskColl := SchedEvents.Events;
            I := 0;
            While (I < TskColl.Count) And Not Terminated Do
            Begin
              If (TskColl[I].State = sesWaiting) And (CompareTimeStamps(NowStamp, TskColl[I].NextFire) >= 0) Then
              Begin
                TskColl[I].Triggered;
                SchedEvents.PostEvent(TskColl[I]);
              End;
              Inc(I);
            End;
            Dec(FEventIdx);
          End;
        Finally
          FCritSect.Leave;
        End;
      End;
      If Not Terminated Then
        Sleep(1);
    End;
  Except
  End;
  FEnded := True;
End;

Procedure TScheduleThread.BeforeDestruction;
Begin
  If (FCritSect = Nil) Or (FEventComponents = Nil) Then
    Exit;
  FCritSect.Enter;
  Try
    FreeAndNil(FEventComponents);
  Finally
    FCritSect.Leave;
  End;
  Inherited BeforeDestruction;
End;

Procedure TScheduleThread.AddEventComponent(Const AComp: TJvCustomScheduledEvents);
Begin
  If (FCritSect = Nil) Or (FEventComponents = Nil) Then
    Exit;
  FCritSect.Enter;
  Try
    If FEventComponents.IndexOf(AComp) = -1 Then
    Begin
      FEventComponents.Add(AComp);
      If Suspended Then
        Suspended := False;
    End;
  Finally
    FCritSect.Leave;
  End;
End;

Procedure TScheduleThread.RemoveEventComponent(Const AComp: TJvCustomScheduledEvents);
Begin
  If (FCritSect = Nil) Or (FEventComponents = Nil) Then
    Exit;
  FCritSect.Enter;
  Try
    FEventComponents.Remove(AComp);
  Finally
    FCritSect.Leave;
  End;
End;

Procedure TScheduleThread.Lock;
Begin
  FCritSect.Enter;
End;

Procedure TScheduleThread.Unlock;
Begin
  FCritSect.Leave;
End;

{ TScheduleThread instance }

Var
  GScheduleThread: TScheduleThread = Nil;

Procedure FinalizeScheduleThread;
Begin
  If GScheduleThread <> Nil Then
  Begin
    If GScheduleThread.Suspended Then
    Begin
      GScheduleThread.Suspended := False;
      // In order for the thread to actually start (and respond to Terminate)
      // we must indicate to the system that we want to be paused. This way
      // the thread can start and will start working.
      // if we don't do this, the threadproc in classes.pas will directly see
      // that Terminated is set to True and never call Execute
      SleepEx(10, True);
    End;
    GScheduleThread.FreeOnTerminate := False;
    GScheduleThread.Terminate;
    While Not GScheduleThread.Ended Do
    Begin
      SleepEx(10, True);
      Application.ProcessMessages;
    End;
    FreeAndNil(GScheduleThread);
  End;
End;

Function ScheduleThread: TScheduleThread;
Begin
  If GScheduleThread = Nil Then
    GScheduleThread := TScheduleThread.Create;
  Result := GScheduleThread;
End;

// === { THackWriter } ========================================================

Type
  TReaderAccessProtected = Class(TReader);

Type
  THackWriter = Class(TWriter)
  Protected
    Procedure WriteSet(SetType: Pointer; Value: Integer);
  End;

  // Copied from D5 Classes.pas and modified a bit.

Procedure THackWriter.WriteSet(SetType: Pointer; Value: Integer);
Var
  I: Integer;
  BaseType: PTypeInfo;
Begin
  BaseType := GetTypeData(SetType)^.CompType^;
  WriteValue(vaSet);
  For I := 0 To SizeOf(TIntegerSet) * 8 - 1 Do
    If I In TIntegerSet(Value) Then
{$IFDEF RTL200_UP}WriteUTF8Str{$ELSE}WriteStr{$ENDIF RTL200_UP}(GetEnumName(BaseType, I));
  WriteStr('');
End;

// === { TJvCustomScheduledEvents } ===========================================

{$IFDEF SUPPORTS_CLASS_CTORDTORS}

Class Destructor TJvCustomScheduledEvents.Destroy;
Begin
  FinalizeScheduleThread;
End;
{$ENDIF SUPPORTS_CLASS_CTORDTORS}

Constructor TJvCustomScheduledEvents.Create(AOwner: TComponent);
Begin
  Inherited Create(AOwner);
  FPostedEvents := TList.Create;
  FEvents := TJvEventCollection.Create(Self);
  FWnd := AllocateHWndEx(WndProc);
  If Not(csDesigning In ComponentState) And Not(csLoading In ComponentState) Then
  Begin
    If AutoSave Then
      LoadEventStates;
    InitEvents;
  End;
  ScheduleThread.AddEventComponent(Self);
End;

Destructor TJvCustomScheduledEvents.Destroy;
Begin
  If Not(csDesigning In ComponentState) Then
  Begin
    ScheduleThread.RemoveEventComponent(Self);
    If AutoSave Then
      SaveEventStates;
    If FWnd <> 0 Then
      DeallocateHWndEx(FWnd);
  End;
  FEvents.Free;
  FPostedEvents.Free;
  Inherited Destroy;
End;

Procedure TJvCustomScheduledEvents.SetAppStorage(Value: TJvCustomAppStorage);
Begin
  ReplaceComponentReference(Self, Value, TComponent(FAppStorage));
End;

Procedure TJvCustomScheduledEvents.Notification(AComponent: TComponent; Operation: TOperation);
Begin
  Inherited Notification(AComponent, Operation);
  If (AComponent = AppStorage) And (Operation = opRemove) Then
    AppStorage := Nil;
End;

Procedure TJvCustomScheduledEvents.DoEndEvent(Const Event: TJvEventCollectionItem);
Begin
  If Assigned(FOnEndEvent) Then
    FOnEndEvent(Event);
End;

Procedure TJvCustomScheduledEvents.DoStartEvent(Const Event: TJvEventCollectionItem);
Begin
  If Assigned(FOnStartEvent) Then
    FOnStartEvent(Event);
End;

Function TJvCustomScheduledEvents.GetEvents: TJvEventCollection;
Begin
  Result := FEvents;
End;

Procedure TJvCustomScheduledEvents.InitEvents;
Var
  I: Integer;
Begin
  For I := 0 To FEvents.Count - 1 Do
    If FEvents[I].State = sesNotInitialized Then
      FEvents[I].Start;
End;

Procedure TJvCustomScheduledEvents.Loaded;
Begin
  If Not(csDesigning In ComponentState) Then
  Begin
    If AutoSave Then
      LoadEventStates;
    InitEvents;
  End;
End;

Procedure TJvCustomScheduledEvents.LoadSingleEvent(Sender: TJvCustomAppStorage; Const Path: String; Const List: TObject;
  Const Index: Integer; Const ItemName: String);
Var
  Stamp: TTimeStamp;
  TriggerCount: Integer;
  DayCount: Integer;
  Snooze: TTimeStamp;
  SnoozeInterval: TSystemTime;
  EventName: String;
  Event: TJvEventCollectionItem;

  AInt: Cardinal;
  EventInfo: TScheduledEventStateInfo;
Begin
  EventName := Sender.ReadString(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'Eventname']));
  If EventName <> '' Then
  Begin
    Stamp.Date := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'Stamp.Date']));
    Stamp.Time := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'Stamp.Time']));
    TriggerCount := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'TriggerCount']));
    DayCount := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayCount']));
    Snooze.Date := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'Snooze.Date']));
    Snooze.Time := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'Snooze.Time']));
    SnoozeInterval.wYear := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wYear']));
    SnoozeInterval.wMonth := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wMonth']));
    SnoozeInterval.wDay := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wDay']));
    SnoozeInterval.wHour := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wHour']));
    SnoozeInterval.wMinute := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wMinute']));
    SnoozeInterval.wSecond := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wSecond']));
    SnoozeInterval.wMilliseconds :=
      Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wMilliseconds']));
    { Common }
    With EventInfo Do
    Begin
      AInt := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'RecurringType']));
      ARecurringType := TScheduleRecurringKind(AInt);
      AStartDate.Time := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'StartDate_time']));
      AStartDate.Date := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'StartDate_date']));
      AInt := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'EndType']));
      AEndType := TScheduleEndKind(AInt);
      AEndDate.Time := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'EndDate_time']));
      AEndDate.Date := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'EndDate_date']));
      AEndCount := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'EndCount']));
      ALastTriggered.Time := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'LastTriggered_time']));
      ALastTriggered.Date := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'LastTriggered_date']));
    End;
    { DayFrequency }
    With EventInfo.DayFrequence Do
    Begin
      ADayFrequencyStartTime := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayFrequencyStartTime']));
      ADayFrequencyEndTime := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayFrequencyEndTime']));
      ADayFrequencyInterval := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayFrequencyInterval']));
    End;
    { Daily }
    With EventInfo.Daily Do
    Begin
      ADayEveryWeekDay := Sender.ReadBoolean(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayEveryWeekDay']));
      ADayInterval := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayInterval']));
    End;
    { Weekly }
    With EventInfo.Weekly Do
    Begin
      AWeekInterval := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'WeekInterval']));
      AppStorage.ReadSet(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'WeekDaysOfWeek']), TypeInfo(TScheduleWeekDays), [],
        AWeekDaysOfWeek);
    End;
    { Monthly }
    With EventInfo.Monthly Do
    Begin
      AInt := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'MothIndexKind']));
      AMonthIndexKind := TScheduleIndexKind(AInt);
      AMonthIndexValue := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'MonthIndexValue']));
      AMonthDay := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'MonthDay']));
      AMonthInterval := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'MonthInterval']));
    End;
    { Yearly }
    With EventInfo.Yearly Do
    Begin
      AInt := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearIndexKind']));
      AYearIndexKind := TScheduleIndexKind(AInt);
      AYearIndexValue := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearIndexValue']));
      AYearDay := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearDay']));
      AYearMonth := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearMonth']));
      AYearInterval := Sender.ReadInteger(Sender.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearInterval']));
    End;
    Event := TJvEventCollection(List).Add;
    Event.Name := EventName;
    Event.LoadState(Stamp, TriggerCount, DayCount, Snooze, SnoozeInterval, EventInfo);
  End;
End;

Procedure TJvCustomScheduledEvents.LoadEventStates(Const ClearBefore: Boolean = True);
Begin
  If ClearBefore Then
    FEvents.Clear;
  If Assigned(AppStorage) Then
    If AppStorage.PathExists(AppStoragePath) Then
      AppStorage.ReadList(AppStoragePath, FEvents, LoadSingleEvent, cEventPrefix);
End;

Procedure TJvCustomScheduledEvents.SaveSingleEvent(Sender: TJvCustomAppStorage; Const Path: String; Const List: TObject;
  Const Index: Integer; Const ItemName: String);
Var
  Stamp: TTimeStamp;
  TriggerCount: Integer;
  DayCount: Integer;
  StampDate: Integer;
  StampTime: Integer;
  SnoozeStamp: TTimeStamp;
  SnoozeInterval: TSystemTime;
  SnoozeDate: Integer;
  SnoozeTime: Integer;
  EventInfo: TScheduledEventStateInfo;
Begin
  TJvEventCollection(List)[Index].SaveState(Stamp, TriggerCount, DayCount, SnoozeStamp, SnoozeInterval, EventInfo);
  StampDate := Stamp.Date;
  StampTime := Stamp.Time;
  SnoozeDate := SnoozeStamp.Date;
  SnoozeTime := SnoozeStamp.Time;
  AppStorage.WriteString(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'Eventname']), FEvents[Index].Name);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'Stamp.Date']), StampDate);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'Stamp.Time']), StampTime);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'TriggerCount']), TriggerCount);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayCount']), DayCount);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'Snooze.Date']), SnoozeDate);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'Snooze.Time']), SnoozeTime);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wYear']), SnoozeInterval.wYear);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wMonth']), SnoozeInterval.wMonth);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wDay']), SnoozeInterval.wDay);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wHour']), SnoozeInterval.wHour);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wMinute']), SnoozeInterval.wMinute);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wSecond']), SnoozeInterval.wSecond);
  AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'SnoozeInterval.wMilliseconds']),
    SnoozeInterval.wMilliseconds);
  { Common }
  With EventInfo Do
  Begin
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'RecurringType']), Integer(ARecurringType));
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'StartDate_time']), AStartDate.Time);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'StartDate_date']), AStartDate.Date);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'EndType']), Integer(AEndType));
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'EndDate_time']), AEndDate.Time);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'EndDate_date']), AEndDate.Date);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'EndCount']), AEndCount);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'LastTriggered_time']), ALastTriggered.Time);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'LastTriggered_date']), ALastTriggered.Date);
  End;
  { DayFrequency }
  With EventInfo.DayFrequence Do
  Begin
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayFrequencyStartTime']), ADayFrequencyStartTime);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayFrequencyEndTime']), ADayFrequencyEndTime);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayFrequencyInterval']), ADayFrequencyInterval);
  End;
  { Daily }
  With EventInfo.Daily Do
  Begin
    AppStorage.WriteBoolean(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayEveryWeekDay']), ADayEveryWeekDay);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'DayInterval']), ADayInterval);
  End;
  { Weekly }
  With EventInfo.Weekly Do
  Begin
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'WeekInterval']), AWeekInterval);
    AppStorage.WriteSet(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'WeekDaysOfWeek']), TypeInfo(TScheduleWeekDays),
      AWeekDaysOfWeek);
  End;
  { Monthly }
  With EventInfo.Monthly Do
  Begin
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'MothIndexKind']), Integer(AMonthIndexKind));
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'MonthIndexValue']), AMonthIndexValue);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'MonthDay']), AMonthDay);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'MonthInterval']), AMonthInterval);
  End;
  { Yearly }
  With EventInfo.Yearly Do
  Begin
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearIndexKind']), Integer(AYearIndexKind));
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearIndexValue']), AYearIndexValue);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearDay']), AYearDay);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearMonth']), AYearMonth);
    AppStorage.WriteInteger(AppStorage.ConcatPaths([Path, ItemName + IntToStr(Index), 'YearInterval']), AYearInterval);
  End;
End;

Procedure TJvCustomScheduledEvents.DeleteSingleEvent(Sender: TJvCustomAppStorage; Const Path: String; Const List: TObject;
  Const First, Last: Integer; Const ItemName: String);
Var
  I: Integer;
Begin
  For I := First To Last Do
    Sender.DeleteSubTree(Sender.ConcatPaths([Path, ItemName + IntToStr(I)]));
End;

Procedure TJvCustomScheduledEvents.SaveEventStates;
Begin
  If Assigned(AppStorage) Then
    AppStorage.WriteList(AppStoragePath, FEvents, FEvents.Count, SaveSingleEvent, DeleteSingleEvent, cEventPrefix);
End;

Procedure TJvCustomScheduledEvents.StartAll;
Var
  I: Integer;
Begin
  For I := 0 To FEvents.Count - 1 Do
    If FEvents[I].State In [sesPaused, sesNotInitialized] Then
      FEvents[I].Start;
End;

Procedure TJvCustomScheduledEvents.StopAll;
Var
  I: Integer;
Begin
  For I := 0 To FEvents.Count - 1 Do
    FEvents[I].Stop;
End;

Procedure TJvCustomScheduledEvents.PauseAll;
Var
  I: Integer;
Begin
  For I := 0 To FEvents.Count - 1 Do
    FEvents[I].Pause;
End;

Procedure TJvCustomScheduledEvents.SetEvents(Value: TJvEventCollection);
Begin
  FEvents.Assign(Value);
End;

Procedure TJvCustomScheduledEvents.WndProc(Var Msg: TMessage);
Var
  List: TList;
  I: Integer;
Begin
  With Msg Do
    Case Msg Of
      CM_EXECEVENT:
        Dispatch(Msg);
      WM_TIMECHANGE:
        Begin
          // Mantis 3355: Time has changed, mark all running schedules as
          // "to be restarted", stop and then restart them.
          List := TList.Create;
          Try
            ScheduleThread.Lock;
            Try
              For I := 0 To FEvents.Count - 1 Do
              Begin
                // https://issuetracker.delphi-jedi.org/view.php?id=3355
                If FEvents[I].State <> sesNotInitialized Then // In [sesTriggered, sesExecuting, sesPaused] Then // CPsoft v2019.7.28.0
                Begin
                  List.Add(FEvents[I]);
                  FEvents[I].Stop;
                End;
              End;
              For I := 0 To List.Count - 1 Do
                TJvEventCollectionItem(List[I]).Start;
            Finally
              ScheduleThread.Unlock;
            End;
          Finally
            List.Free;
          End;
        End;
    Else
      Result := DefWindowProc(Handle, Msg, WParam, LParam);
    End;
End;

Procedure TJvCustomScheduledEvents.PostEvent(Event: TJvEventCollectionItem);
Begin
  ScheduleThread.Lock;
  Try
    FPostedEvents.Add(Event);
    If Not FEventsPosted Then
    Begin
      // Post one message for all posted events
      FEventsPosted := True;
      PostMessage(Handle, CM_EXECEVENT, 0, 0);
    End;
  Finally
    ScheduleThread.Unlock;
  End;
End;

Procedure TJvCustomScheduledEvents.RemovePostedEvent(Event: TJvEventCollectionItem);
Begin
  If Not(csDestroying In ComponentState) And (GScheduleThread <> Nil) Then
  Begin
    ScheduleThread.Lock;
    Try
      Event.FState := sesEnded;
      While FPostedEvents.Remove(Event) <> -1 Do;
    Finally
      ScheduleThread.Unlock;
    End;
  End;
End;

Procedure TJvCustomScheduledEvents.CMExecEvent(Var Msg: TMessage);
Var
  Event: TJvEventCollectionItem;
Begin
  Try
    ScheduleThread.Lock;
    Try
      While FPostedEvents.Count > 0 Do
      Begin
        Event := FPostedEvents[0];
        FPostedEvents.Delete(0);

        ScheduleThread.Unlock; // the user code must not be protected by the critical section
        Try
          Try
            DoStartEvent(Event);
            Event.Execute;
            DoEndEvent(Event);
          Except
            //
          End;
        Finally
          ScheduleThread.Lock;
        End;
      End;
    Finally
      FEventsPosted := False;
      ScheduleThread.Unlock;
    End;
  Except
  End;
  Msg.Result := 1;
End;

// === { TJvEventCollection } =================================================

Constructor TJvEventCollection.Create(AOwner: TPersistent);
Begin
  Inherited Create(AOwner, TJvEventCollectionItem);
End;

Function TJvEventCollection.GetItem(Index: Integer): TJvEventCollectionItem;
Begin
  Result := TJvEventCollectionItem(Inherited Items[Index]);
End;

Procedure TJvEventCollection.SetItem(Index: Integer; Value: TJvEventCollectionItem);
Begin
  Inherited Items[Index] := Value;
End;

Function TJvEventCollection.Add: TJvEventCollectionItem;
Begin
  Result := TJvEventCollectionItem(Inherited Add);
End;

Function TJvEventCollection.Insert(Index: Integer): TJvEventCollectionItem;
Begin
  Result := TJvEventCollectionItem(Inherited Insert(Index));
End;

Procedure TJvEventCollection.Notify(Item: TCollectionItem; Action: TCollectionNotification);
Begin
  Inherited Notify(Item, Action);
  If Action In [cnExtracting, cnDeleting] Then
    (Owner As TJvCustomScheduledEvents).RemovePostedEvent(Item As TJvEventCollectionItem);
End;

// === { TJvEventCollectionItem } =============================================

Constructor TJvEventCollectionItem.Create(Collection: TCollection);
Var
  NewName: String;
  I: Integer;
  J: Integer;

  Function NewNameIsUnique: Boolean;
  Begin
    With TJvEventCollection(Collection) Do
    Begin
      J := Count - 1;
      While (J >= 0) And Not AnsiSameText(Items[J].Name, NewName + IntToStr(I)) Do
        Dec(J);
      Result := J < 0;
    End;
  End;

  Procedure CreateNewName;
  Begin
    NewName := 'Event';
    I := 0;
    Repeat
      Inc(I);
    Until NewNameIsUnique;
  End;

Begin
  ScheduleThread.Lock;
  Try
    If csDesigning In TComponent(TJvEventCollection(Collection).GetOwner).ComponentState Then
      CreateNewName
    Else
      NewName := '';
    Inherited Create(Collection);
    FSchedule := CreateSchedule;
    FSnoozeFire := NullStamp;
    FScheduleFire := NullStamp;
    If NewName <> '' Then
      Name := NewName + IntToStr(I);
  Finally
    ScheduleThread.Unlock;
  End;
End;

Destructor TJvEventCollectionItem.Destroy;
Begin
  ScheduleThread.Lock;
  Try
    Stop;
    Inherited Destroy;
  Finally
    ScheduleThread.Unlock;
  End;
End;

Procedure TJvEventCollectionItem.Assign(Source: TPersistent);
Begin
  If Source Is TJvEventCollectionItem Then
  Begin
    Name := TJvEventCollectionItem(Source).Name;
    CountMissedEvents := TJvEventCollectionItem(Source).CountMissedEvents;
    Schedule := TJvEventCollectionItem(Source).Schedule;
    OnExecute := TJvEventCollectionItem(Source).OnExecute;
  End
  Else
    Inherited Assign(Source);
End;

Procedure TJvEventCollectionItem.Triggered;
Begin
  FState := sesTriggered;
End;

Procedure TJvEventCollectionItem.DefineProperties(Filer: TFiler);
Var
  SingleShot: Boolean;
  DailySched: Boolean;
  WeeklySched: Boolean;
  MonthlySched: Boolean;
  YearlySched: Boolean;
  MIK: TScheduleIndexKind;
  YIK: TScheduleIndexKind;
Begin
  // Determine settings to determine writing properties.
  SingleShot := Schedule.RecurringType = srkOneShot;
  DailySched := Schedule.RecurringType = srkDaily;
  WeeklySched := Schedule.RecurringType = srkWeekly;
  MonthlySched := Schedule.RecurringType = srkMonthly;
  YearlySched := Schedule.RecurringType = srkYearly;
  If MonthlySched Then
    MIK := (Schedule As IJclMonthlySchedule).IndexKind
  Else
    MIK := sikNone;
  If YearlySched Then
    YIK := (Schedule As IJclYearlySchedule).IndexKind
  Else
    YIK := sikNone;

  // Standard properties
  Filer.DefineProperty('StartDate', PropStartDateRead, PropStartDateWrite, True);
  Filer.DefineProperty('RecurringType', PropRecurringTypeRead, PropRecurringTypeWrite, Not SingleShot);
  Filer.DefineProperty('EndType', PropEndTypeRead, PropEndTypeWrite, Not SingleShot);
  Filer.DefineProperty('EndDate', PropEndDateRead, PropEndDateWrite, Not SingleShot And (Schedule.EndType = sekDate));
  Filer.DefineProperty('EndCount', PropEndCountRead, PropEndCountWrite, Not SingleShot And
    (Schedule.EndType In [sekTriggerCount, sekDayCount]));

  // Daily frequency properties
  Filer.DefineProperty('Freq_StartTime', PropFreqStartTimeRead, PropFreqStartTimeWrite, Not SingleShot);
  Filer.DefineProperty('Freq_EndTime', PropFreqEndTimeRead, PropFreqEndTimeWrite, Not SingleShot);
  Filer.DefineProperty('Freq_Interval', PropFreqIntervalRead, PropFreqIntervalWrite, Not SingleShot);

  // Daily schedule properties
  Filer.DefineProperty('Daily_EveryWeekDay', PropDailyEveryWeekDayRead, PropDailyEveryWeekDayWrite, DailySched);
  Filer.DefineProperty('Daily_Interval', PropDailyIntervalRead, PropDailyIntervalWrite, DailySched And Not(Schedule As IJclDailySchedule)
    .EveryWeekDay);

  // Weekly schedule properties
  Filer.DefineProperty('Weekly_DaysOfWeek', PropWeeklyDaysOfWeekRead, PropWeeklyDaysOfWeekWrite, WeeklySched);
  Filer.DefineProperty('Weekly_Interval', PropWeeklyIntervalRead, PropWeeklyIntervalWrite, WeeklySched);

  // Monthly schedule properties
  Filer.DefineProperty('Monthly_IndexKind', PropMonthlyIndexKindRead, PropMonthlyIndexKindWrite, MonthlySched);
  Filer.DefineProperty('Monthly_IndexValue', PropMonthlyIndexValueRead, PropMonthlyIndexValueWrite,
    MonthlySched And (MIK In [sikDay .. sikSunday]));
  Filer.DefineProperty('Monthly_Day', PropMonthlyDayRead, PropMonthlyDayWrite, MonthlySched And (MIK In [sikNone]));
  Filer.DefineProperty('Monthly_Interval', PropMonthlyIntervalRead, PropMonthlyIntervalWrite, MonthlySched);

  // Yearly schedule properties
  Filer.DefineProperty('Yearly_IndexKind', PropYearlyIndexKindRead, PropYearlyIndexKindWrite, YearlySched);
  Filer.DefineProperty('Yearly_IndexValue', PropYearlyIndexValueRead, PropYearlyIndexValueWrite,
    YearlySched And (YIK In [sikDay .. sikSunday]));
  Filer.DefineProperty('Yearly_Day', PropYearlyDayRead, PropYearlyDayWrite, YearlySched And (YIK In [sikNone, sikDay]));
  Filer.DefineProperty('Yearly_Month', PropYearlyMonthRead, PropYearlyMonthWrite, YearlySched);
  Filer.DefineProperty('Yearly_Interval', PropYearlyIntervalRead, PropYearlyIntervalWrite, YearlySched);
End;

Procedure TJvEventCollectionItem.DoExecute(Const IsSnoozeFire: Boolean);
Begin
  If Assigned(FOnExecute) Then
    FOnExecute(Self, IsSnoozeFire);
End;

Function TJvEventCollectionItem.GetDisplayName: String;
Begin
  Result := Name;
End;

Function TJvEventCollectionItem.GetNextFire: TTimeStamp;
Begin
  If IsNullTimeStamp(FSnoozeFire) Or (Not IsNullTimeStamp(FScheduleFire) And (CompareTimeStamps(FSnoozeFire, FScheduleFire) > 0)) Then
    Result := FScheduleFire
  Else
    Result := FSnoozeFire;
End;

Procedure TJvEventCollectionItem.Execute;
Var
  IsSnoozeFire: Boolean;
Begin
  If State <> sesTriggered Then
    Exit; // Ignore this message, something is wrong.
  FActualTriggerTime := DateTimeToTimeStamp(Now);
  IsSnoozeFire := Not IsNullTimeStamp(FSnoozeFire) And (CompareTimeStamps(FActualTriggerTime, FSnoozeFire) >= 0);
  If IsSnoozeFire And Not IsNullTimeStamp(FScheduleFire) And (CompareTimeStamps(FActualTriggerTime, FScheduleFire) >= 0) Then
  Begin
    { We can't have both, the schedule will win (other possibility: generate two succesive events
      from this method, one as a snooze, the other as a schedule) }
    FSnoozeFire := NullStamp;
    IsSnoozeFire := False;
  End;
  FState := sesExecuting;
  Try
    FReqTriggerTime := NextFire;
    If Not IsSnoozeFire Then
      FScheduleFire := Schedule.NextEventFromNow(CountMissedEvents);
    FSnoozeFire := NullStamp;
    DoExecute(IsSnoozeFire);
  Finally
    If IsNullTimeStamp(NextFire) Then
      FState := sesEnded
    Else
      FState := sesWaiting;
  End;
End;

Procedure TJvEventCollectionItem.PropDateRead(Reader: TReader; Var Stamp: TTimeStamp);
Var
  Str: String;
  Y: Integer;
  M: Integer;
  D: Integer;
  H: Integer;
  Min: Integer;
  MSecs: Integer;
Begin
  Str := Reader.ReadString;
  Y := StrToInt(Copy(Str, 1, 4));
  M := StrToInt(Copy(Str, 6, 2));
  D := StrToInt(Copy(Str, 9, 2));
  H := StrToInt(Copy(Str, 12, 2));
  Min := StrToInt(Copy(Str, 15, 2));
  MSecs := StrToInt(Copy(Str, 18, 2)) * 1000 + StrToInt(Copy(Str, 21, 3));

  Stamp := DateTimeToTimeStamp(EncodeDate(Y, M, D));
  Stamp.Time := H * 3600000 + Min * 60000 + MSecs;
End;

Procedure TJvEventCollectionItem.PropDateWrite(Writer: TWriter; Const Stamp: TTimeStamp);
Var
  TmpDate: TDateTime;
  Y: Word;
  M: Word;
  D: Word;
  MSecs: Integer;
Begin
  TmpDate := TimeStampToDateTime(Stamp);
  DecodeDate(TmpDate, Y, M, D);
  MSecs := Stamp.Time;
  Writer.WriteString(Format('%.4d/%.2d/%.2d %.2d:%.2d:%.2d.%.3d', [Y, M, D, (MSecs Div 3600000) Mod 24, (MSecs Div 60000) Mod 60,
    (MSecs Div 1000) Mod 60, MSecs Mod 1000]));
End;

Procedure TJvEventCollectionItem.PropDailyEveryWeekDayRead(Reader: TReader);
Begin
  (Schedule As IJclDailySchedule).EveryWeekDay := Reader.ReadBoolean;
End;

Procedure TJvEventCollectionItem.PropDailyEveryWeekDayWrite(Writer: TWriter);
Begin
  Writer.WriteBoolean((Schedule As IJclDailySchedule).EveryWeekDay);
End;

Procedure TJvEventCollectionItem.PropDailyIntervalRead(Reader: TReader);
Begin
  (Schedule As IJclDailySchedule).Interval := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropDailyIntervalWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclDailySchedule).Interval);
End;

Procedure TJvEventCollectionItem.PropEndCountRead(Reader: TReader);
Begin
  Schedule.EndCount := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropEndCountWrite(Writer: TWriter);
Begin
  Writer.WriteInteger(Schedule.EndCount);
End;

Procedure TJvEventCollectionItem.PropEndDateRead(Reader: TReader);
Var
  TmpStamp: TTimeStamp;
Begin
  PropDateRead(Reader, TmpStamp);
  Schedule.EndDate := TmpStamp;
End;

Procedure TJvEventCollectionItem.PropEndDateWrite(Writer: TWriter);
Begin
  PropDateWrite(Writer, Schedule.EndDate);
End;

Procedure TJvEventCollectionItem.PropEndTypeRead(Reader: TReader);
Begin
  Schedule.EndType := TScheduleEndKind(GetEnumValue(TypeInfo(TScheduleEndKind), Reader.ReadIdent));
End;

Procedure TJvEventCollectionItem.PropEndTypeWrite(Writer: TWriter);
Begin
  Writer.WriteIdent(GetEnumName(TypeInfo(TScheduleEndKind), Ord(Schedule.EndType)));
End;

Procedure TJvEventCollectionItem.PropFreqEndTimeRead(Reader: TReader);
Begin
  (Schedule As IJclScheduleDayFrequency).EndTime := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropFreqEndTimeWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclScheduleDayFrequency).EndTime);
End;

Procedure TJvEventCollectionItem.PropFreqIntervalRead(Reader: TReader);
Begin
  (Schedule As IJclScheduleDayFrequency).Interval := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropFreqIntervalWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclScheduleDayFrequency).Interval);
End;

Procedure TJvEventCollectionItem.PropFreqStartTimeRead(Reader: TReader);
Begin
  (Schedule As IJclScheduleDayFrequency).StartTime := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropFreqStartTimeWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclScheduleDayFrequency).StartTime);
End;

Procedure TJvEventCollectionItem.PropMonthlyDayRead(Reader: TReader);
Begin
  (Schedule As IJclMonthlySchedule).Day := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropMonthlyDayWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclMonthlySchedule).Day);
End;

Procedure TJvEventCollectionItem.PropMonthlyIndexKindRead(Reader: TReader);
Begin
  (Schedule As IJclMonthlySchedule).IndexKind := TScheduleIndexKind(GetEnumValue(TypeInfo(TScheduleIndexKind), Reader.ReadIdent));
End;

Procedure TJvEventCollectionItem.PropMonthlyIndexKindWrite(Writer: TWriter);
Begin
  Writer.WriteIdent(GetEnumName(TypeInfo(TScheduleIndexKind), Ord((Schedule As IJclMonthlySchedule).IndexKind)));
End;

Procedure TJvEventCollectionItem.PropMonthlyIndexValueRead(Reader: TReader);
Begin
  (Schedule As IJclMonthlySchedule).IndexValue := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropMonthlyIndexValueWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclMonthlySchedule).IndexValue);
End;

Procedure TJvEventCollectionItem.PropMonthlyIntervalRead(Reader: TReader);
Begin
  (Schedule As IJclMonthlySchedule).Interval := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropMonthlyIntervalWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclMonthlySchedule).Interval);
End;

Procedure TJvEventCollectionItem.PropRecurringTypeRead(Reader: TReader);
Begin
  Schedule.RecurringType := TScheduleRecurringKind(GetEnumValue(TypeInfo(TScheduleRecurringKind), Reader.ReadIdent));
End;

Procedure TJvEventCollectionItem.PropRecurringTypeWrite(Writer: TWriter);
Begin
  Writer.WriteIdent(GetEnumName(TypeInfo(TScheduleRecurringKind), Ord(Schedule.RecurringType)));
End;

Procedure TJvEventCollectionItem.PropStartDateRead(Reader: TReader);
Var
  TmpStamp: TTimeStamp;
Begin
  PropDateRead(Reader, TmpStamp);
  Schedule.StartDate := TmpStamp;
End;

Procedure TJvEventCollectionItem.PropStartDateWrite(Writer: TWriter);
Begin
  PropDateWrite(Writer, Schedule.StartDate);
End;

Procedure TJvEventCollectionItem.PropWeeklyDaysOfWeekRead(Reader: TReader);
Var
  TempVal: TScheduleWeekDays;
Begin
  JclIntToSet(TypeInfo(TScheduleWeekDays), TempVal, TReaderAccessProtected(Reader).ReadSet(TypeInfo(TScheduleWeekDays)));
  (Schedule As IJclWeeklySchedule).DaysOfWeek := TempVal;
End;

Procedure TJvEventCollectionItem.PropWeeklyDaysOfWeekWrite(Writer: TWriter);
Var
  TempVar: TScheduleWeekDays;
Begin
  TempVar := (Schedule As IJclWeeklySchedule).DaysOfWeek;
  THackWriter(Writer).WriteSet(TypeInfo(TScheduleWeekDays), JclSetToInt(TypeInfo(TScheduleWeekDays), TempVar));
End;

Procedure TJvEventCollectionItem.PropWeeklyIntervalRead(Reader: TReader);
Begin
  (Schedule As IJclWeeklySchedule).Interval := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropWeeklyIntervalWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclWeeklySchedule).Interval);
End;

Procedure TJvEventCollectionItem.PropYearlyDayRead(Reader: TReader);
Begin
  (Schedule As IJclYearlySchedule).Day := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropYearlyDayWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclYearlySchedule).Day);
End;

Procedure TJvEventCollectionItem.PropYearlyIndexKindRead(Reader: TReader);
Begin
  (Schedule As IJclYearlySchedule).IndexKind := TScheduleIndexKind(GetEnumValue(TypeInfo(TScheduleIndexKind), Reader.ReadIdent));
End;

Procedure TJvEventCollectionItem.PropYearlyIndexKindWrite(Writer: TWriter);
Begin
  Writer.WriteIdent(GetEnumName(TypeInfo(TScheduleIndexKind), Ord((Schedule As IJclYearlySchedule).IndexKind)));
End;

Procedure TJvEventCollectionItem.PropYearlyIndexValueRead(Reader: TReader);
Begin
  (Schedule As IJclYearlySchedule).IndexValue := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropYearlyIndexValueWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclYearlySchedule).IndexValue);
End;

Procedure TJvEventCollectionItem.PropYearlyIntervalRead(Reader: TReader);
Begin
  (Schedule As IJclYearlySchedule).Interval := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropYearlyIntervalWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclYearlySchedule).Interval);
End;

Procedure TJvEventCollectionItem.PropYearlyMonthRead(Reader: TReader);
Begin
  (Schedule As IJclYearlySchedule).Month := Reader.ReadInteger;
End;

Procedure TJvEventCollectionItem.PropYearlyMonthWrite(Writer: TWriter);
Begin
  Writer.WriteInteger((Schedule As IJclYearlySchedule).Month);
End;

Procedure TJvEventCollectionItem.SetName(Value: String);
Begin
  If FName <> Value Then
  Begin
    FName := Value;
    Changed(False);
  End;
End;

Procedure TJvEventCollectionItem.LoadState(Const TriggerStamp: TTimeStamp; Const TriggerCount, DayCount: Integer;
  Const SnoozeStamp: TTimeStamp; Const ALastSnoozeInterval: TSystemTime; Const AEventInfo: TScheduledEventStateInfo);
Var
  IDayFrequency: IJclScheduleDayFrequency;
  IDay: IJclDailySchedule;
  IWeek: IJclWeeklySchedule;
  IMonth: IJclMonthlySchedule;
  IYear: IJclYearlySchedule;
Begin
  With AEventInfo Do
  Begin
    Schedule.RecurringType := ARecurringType;
    Schedule.StartDate := AStartDate;
    Schedule.EndType := AEndType;
    If ARecurringType = srkOneShot Then
    Begin
      Schedule.EndType := sekDate;
      Schedule.EndDate := AStartDate;
      IDayFrequency := Schedule As IJclScheduleDayFrequency;
      IDayFrequency.StartTime := AStartDate.Time;
      IDayFrequency.EndTime := AEndDate.Time;
      IDayFrequency.Interval := 1;
    End
    Else
    Begin
      Case Schedule.EndType Of
        sekDate:
          Schedule.EndDate := AEndDate;
        sekTriggerCount:
          Schedule.EndCount := AEndCount;
        sekDayCount:
          Schedule.EndCount := AEndCount;
      End;
      IDayFrequency := Schedule As IJclScheduleDayFrequency;
      With AEventInfo.DayFrequence Do
      Begin
        IDayFrequency.StartTime := ADayFrequencyStartTime;
        IDayFrequency.EndTime := ADayFrequencyEndTime;
        IDayFrequency.Interval := ADayFrequencyInterval;
      End;
    End;
    Case ARecurringType Of
      srkOneShot:
        Begin
        End;
      srkDaily:
        Begin
          { IJclDailySchedule }
          IDay := Schedule As IJclDailySchedule;
          With AEventInfo.Daily Do
          Begin
            IDay.EveryWeekDay := ADayEveryWeekDay;
            If Not ADayEveryWeekDay Then
              IDay.Interval := ADayInterval;
          End;
        End;
      srkWeekly:
        Begin
          { IJclWeeklySchedule }
          IWeek := Schedule As IJclWeeklySchedule;
          With AEventInfo.Weekly Do
          Begin
            IWeek.DaysOfWeek := AWeekDaysOfWeek;
            IWeek.Interval := AWeekInterval;
          End;
        End;
      srkMonthly:
        Begin
          { IJclMonthlySchedule }
          IMonth := Schedule As IJclMonthlySchedule;
          With AEventInfo.Monthly Do
          Begin
            IMonth.IndexKind := AMonthIndexKind;
            If AMonthIndexKind <> sikNone Then
              IMonth.IndexValue := AMonthIndexValue;
            If AMonthIndexKind = sikNone Then
              IMonth.Day := AMonthDay;
            IMonth.Interval := AMonthInterval;
          End;
        End;
      srkYearly:
        Begin
          { IJclYearlySchedule }
          IYear := Schedule As IJclYearlySchedule;
          With AEventInfo.Yearly Do
          Begin
            IYear.IndexKind := AYearIndexKind;
            If AYearIndexKind <> sikNone Then
              IYear.IndexValue := AYearIndexValue
            Else
              IYear.Day := AYearDay;
            IYear.Month := AYearMonth;
            IYear.Interval := AYearInterval;
          End;
        End;
    End;
    Schedule.InitToSavedState(TriggerStamp, TriggerCount, DayCount);
    FScheduleFire := TriggerStamp;
    FSnoozeFire := SnoozeStamp;
    FLastSnoozeInterval := ALastSnoozeInterval;
    If IsNullTimeStamp(NextFire) Or (CompareTimeStamps(NextFire, DateTimeToTimeStamp(Now)) < 0) Then
      Schedule.NextEventFromNow(CountMissedEvents);
    If IsNullTimeStamp(NextFire) Then
      FState := sesEnded
    Else
      FState := sesNotInitialized; // sesWaiting; // CPsoft v2019.8.4.0
  End;
End;

Procedure TJvEventCollectionItem.Pause;
Begin
  If FState = sesWaiting Then
    FState := sesPaused;
End;

Procedure TJvEventCollectionItem.SaveState(Out TriggerStamp: TTimeStamp; Out TriggerCount, DayCount: Integer; Out SnoozeStamp: TTimeStamp;
  Out ALastSnoozeInterval: TSystemTime; Out AEventInfo: TScheduledEventStateInfo);
Var
  IDayFrequency: IJclScheduleDayFrequency;
  IDay: IJclDailySchedule;
  IWeek: IJclWeeklySchedule;
  IMonth: IJclMonthlySchedule;
  IYear: IJclYearlySchedule;
Begin
  { Common properties }
  With AEventInfo Do
  Begin
    AEndType := FSchedule.EndType;
    AEndDate := FSchedule.EndDate;
    AEndCount := FSchedule.EndCount;
    ALastTriggered := FSchedule.LastTriggered;
    AStartDate := FSchedule.StartDate;
    ARecurringType := FSchedule.RecurringType;
    { IJclScheduleDayFrequency }
    If ARecurringType <> srkOneShot Then
    Begin
      IDayFrequency := FSchedule As IJclScheduleDayFrequency;
      With AEventInfo.DayFrequence Do
      Begin
        ADayFrequencyStartTime := IDayFrequency.StartTime;
        ADayFrequencyEndTime := IDayFrequency.EndTime;
        ADayFrequencyInterval := IDayFrequency.Interval;
      End;
    End;
    Case ARecurringType Of
      srkOneShot:
        Begin
        End;
      srkDaily:
        Begin
          { IJclDailySchedule }
          IDay := FSchedule As IJclDailySchedule;
          With AEventInfo.Daily Do
          Begin
            ADayInterval := IDay.Interval;
            ADayEveryWeekDay := IDay.EveryWeekDay;
          End;
        End;
      srkWeekly:
        Begin
          { IJclWeeklySchedule }
          IWeek := FSchedule As IJclWeeklySchedule;
          With AEventInfo.Weekly Do
          Begin
            AWeekInterval := IWeek.Interval;
            AWeekDaysOfWeek := IWeek.DaysOfWeek;
          End;
        End;
      srkMonthly:
        Begin
          { IJclMonthlySchedule }
          IMonth := FSchedule As IJclMonthlySchedule;
          With AEventInfo.Monthly Do
          Begin
            AMonthIndexKind := IMonth.IndexKind;
            If AMonthIndexKind <> sikNone Then
              AMonthIndexValue := IMonth.IndexValue;
            AMonthDay := IMonth.Day;
            AMonthInterval := IMonth.Interval;
          End;
        End;
      srkYearly:
        Begin
          { IJclYearlySchedule }
          IYear := FSchedule As IJclYearlySchedule;
          With AEventInfo.Yearly Do
          Begin
            AYearIndexKind := IYear.IndexKind;
            If AYearIndexKind <> sikNone Then
              AYearIndexValue := IYear.IndexValue;
            AYearDay := IYear.Day;
            AYearMonth := IYear.Month;
            AYearInterval := IYear.Interval;
          End;
        End;
    End;
    { Old part }
    TriggerStamp := FScheduleFire;
    TriggerCount := Schedule.TriggerCount;
    DayCount := Schedule.DayCount;
    SnoozeStamp := FSnoozeFire;
    ALastSnoozeInterval := LastSnoozeInterval;
  End;
End;

Procedure TJvEventCollectionItem.Snooze(Const MSecs: Word; Const Secs: Word = 0; Const Mins: Word = 0; Const Hrs: Word = 0;
  Const Days: Word = 0);
Var
  IntervalMSecs: Integer;
  SnoozeStamp: TTimeStamp;
Begin
  // Update last snooze interval
  FLastSnoozeInterval.wDay := Days;
  FLastSnoozeInterval.wHour := Hrs;
  FLastSnoozeInterval.wMinute := Mins;
  FLastSnoozeInterval.wSecond := Secs;
  FLastSnoozeInterval.wMilliseconds := MSecs;
  // Calculate next event
  IntervalMSecs := MSecs + 1000 * (Secs + 60 * Mins + 1440 * Hrs);
  SnoozeStamp := DateTimeToTimeStamp(Now);
  SnoozeStamp.Time := SnoozeStamp.Time + IntervalMSecs;
  If SnoozeStamp.Time >= HoursToMSecs(24) Then
  Begin
    SnoozeStamp.Date := SnoozeStamp.Date + (SnoozeStamp.Time Div HoursToMSecs(24));
    SnoozeStamp.Time := SnoozeStamp.Time Mod HoursToMSecs(24);
  End;
  Inc(SnoozeStamp.Date, Days);
  FSnoozeFire := SnoozeStamp;
End;

Procedure TJvEventCollectionItem.Start;
Begin
  If FState In [sesTriggered, sesExecuting] Then
    Exit;
  If State = sesPaused Then
  Begin
    FScheduleFire := Schedule.NextEventFromNow(CountMissedEvents);
    If IsNullTimeStamp(NextFire) Then
      FState := sesEnded
    Else
      FState := sesWaiting;
  End
  Else
  Begin
    FState := sesNotInitialized;
    Schedule.Reset;
    FScheduleFire := Schedule.NextEventFromNow(CountMissedEvents);
    If IsNullTimeStamp(NextFire) Then
      FState := sesEnded
    Else
      FState := sesWaiting;
  End;
End;

Procedure TJvEventCollectionItem.Stop;
Begin
  If State <> sesNotInitialized Then
    FState := sesNotInitialized;
End;

(*
  procedure TJvEventCollectionItem.LoadFromStreamBin(const S: TStream);
  begin
  ScheduledEventStore_Stream(S, True, False).LoadSchedule(Self);
  end;

  procedure TJvEventCollectionItem.SaveToStreamBin(const S: TStream);
  begin
  ScheduledEventStore_Stream(S, True, False).SaveSchedule(Self);
  end;
*)

Initialization

{$IFDEF UNITVERSIONING}
  RegisterUnitVersion(HInstance, UnitVersioning);
{$ENDIF UNITVERSIONING}

Finalization

{$IFNDEF SUPPORTS_CLASS_CTORDTORS}
  FinalizeScheduleThread;
{$ENDIF ~SUPPORTS_CLASS_CTORDTORS}
{$IFDEF UNITVERSIONING}
UnregisterUnitVersion(HInstance);
{$ENDIF UNITVERSIONING}

End.
