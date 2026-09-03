Attribute VB_Name = "mod_CalendarEngine"
Option Explicit

'===============================================================================
' MODULE : mod_CalendarEngine
' DOMAINE / DOMAIN : Calendar
'
' FR
' Applique les calendriers 5D/6D aux decalages, durees et recherches de jours ouvrables.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Applies 5D/6D calendars to offsets, durations and working-day searches.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : NormalizeCalendarType, IsValidCalendarType, IsWorkingDay, NextWorkingDay, PreviousWorkingDay, AddWorkingDays, SubtractWorkingDays, DateDiffWorkingDays
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Public Const CALENDAR_7D As String = "7j/7"
Public Const CALENDAR_6D As String = "6j/7"
Public Const CALENDAR_5D As String = "5j/7"

'------------------------------------------------------------------------------
' FR: Normalise Calendar Type dans un format exploitable.
' EN: Normalizes Calendar Type into a usable format.
'------------------------------------------------------------------------------
Public Function NormalizeCalendarType(ByVal rawValue As Variant) As String

    Dim txt As String

    txt = LCase$(Trim$(CStr(rawValue)))
    txt = Replace(txt, " ", vbNullString)

    Select Case txt
        Case vbNullString, "7j/7", "7/7", "7j", "7d/7", "7d"
            NormalizeCalendarType = CALENDAR_7D

        Case "6j/7", "6/7", "6j", "6d/7", "6d"
            NormalizeCalendarType = CALENDAR_6D

        Case "5j/7", "5/7", "5j", "5d/7", "5d"
            NormalizeCalendarType = CALENDAR_5D

        Case Else
            Err.Raise vbObjectError + 7601, "NormalizeCalendarType", _
                "Invalid calendar type: " & CStr(rawValue)
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Indique si Valid Calendar Type est vrai pour le contexte courant.
' EN: Returns whether Valid Calendar Type is true for the current context.
'------------------------------------------------------------------------------
Public Function IsValidCalendarType(ByVal rawValue As Variant) As Boolean

    On Error GoTo InvalidValue
    NormalizeCalendarType rawValue
    IsValidCalendarType = True
    Exit Function

InvalidValue:
    IsValidCalendarType = False

End Function

'------------------------------------------------------------------------------
' FR: Indique si Working Day est vrai pour le contexte courant.
' EN: Returns whether Working Day is true for the current context.
'------------------------------------------------------------------------------
Public Function IsWorkingDay( _
    ByVal dateValue As Variant, _
    Optional ByVal calendarType As Variant = "") As Boolean

    Dim calType As String
    Dim weekdayIndex As Long

    If Not HasValue(dateValue) Then Exit Function

    calType = NormalizeCalendarType(calendarType)
    weekdayIndex = Weekday(CDate(dateValue), vbMonday)

    Select Case calType
        Case CALENDAR_7D
            IsWorkingDay = True

        Case CALENDAR_6D
            IsWorkingDay = (weekdayIndex <= 6)

        Case CALENDAR_5D
            IsWorkingDay = (weekdayIndex <= 5)
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Next Working Day sans modifier les donnees d'entree.
' EN: Returns the Next Working Day value without mutating input data.
'------------------------------------------------------------------------------

Public Function NextWorkingDay( _
    ByVal dateValue As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim d As Date
    Dim calType As String

    If Not HasValue(dateValue) Then
        NextWorkingDay = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    d = CDate(dateValue) + 1

    Do While Not IsWorkingDay(d, calType)
        d = d + 1
    Loop

    NextWorkingDay = CDbl(d)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Previous Working Day sans modifier les donnees d'entree.
' EN: Returns the Previous Working Day value without mutating input data.
'------------------------------------------------------------------------------

Public Function PreviousWorkingDay( _
    ByVal dateValue As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim d As Date
    Dim calType As String

    If Not HasValue(dateValue) Then
        PreviousWorkingDay = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    d = CDate(dateValue) - 1

    Do While Not IsWorkingDay(d, calType)
        d = d - 1
    Loop

    PreviousWorkingDay = CDbl(d)

End Function

'------------------------------------------------------------------------------
' FR: Ajoute Working Days a la structure cible.
' EN: Adds Working Days to the target structure.
'------------------------------------------------------------------------------
Public Function AddWorkingDays( _
    ByVal startDate As Variant, _
    ByVal durationDays As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim calType As String
    Dim remaining As Long
    Dim d As Date
    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("AddWorkingDays", "Calendar")

    If Not HasValue(startDate) Then
        AddWorkingDays = Empty
        Exit Function
    End If

    If Not HasValue(durationDays) Then
        AddWorkingDays = Empty
        Exit Function
    End If

    If Not IsNumeric(durationDays) Then
        AddWorkingDays = Empty
        Exit Function
    End If

    remaining = CLng(durationDays)
    If remaining <= 0 Then
        AddWorkingDays = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    d = NormalizeToWorkingDayForward(startDate, calType)
    remaining = remaining - 1

    Do While remaining > 0
        d = CDate(NextWorkingDay(d, calType))
        remaining = remaining - 1
    Loop

    AddWorkingDays = CDbl(d)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Subtract Working Days sans modifier les donnees d'entree.
' EN: Returns the Subtract Working Days value without mutating input data.
'------------------------------------------------------------------------------

Public Function SubtractWorkingDays( _
    ByVal finishDate As Variant, _
    ByVal durationDays As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim calType As String
    Dim remaining As Long
    Dim d As Date
    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("SubtractWorkingDays", "Calendar")

    If Not HasValue(finishDate) Then
        SubtractWorkingDays = Empty
        Exit Function
    End If

    If Not HasValue(durationDays) Then
        SubtractWorkingDays = Empty
        Exit Function
    End If

    If Not IsNumeric(durationDays) Then
        SubtractWorkingDays = Empty
        Exit Function
    End If

    remaining = CLng(durationDays)
    If remaining <= 0 Then
        SubtractWorkingDays = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    d = NormalizeToWorkingDayBackward(finishDate, calType)
    remaining = remaining - 1

    Do While remaining > 0
        d = CDate(PreviousWorkingDay(d, calType))
        remaining = remaining - 1
    Loop

    SubtractWorkingDays = CDbl(d)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Date Diff Working Days sans modifier les donnees d'entree.
' EN: Returns the Date Diff Working Days value without mutating input data.
'------------------------------------------------------------------------------

Public Function DateDiffWorkingDays( _
    ByVal startDate As Variant, _
    ByVal finishDate As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim calType As String
    Dim startSerial As Long
    Dim finishSerial As Long
    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("DateDiffWorkingDays", "Calendar")

    If Not HasValue(startDate) Then
        DateDiffWorkingDays = Empty
        Exit Function
    End If

    If Not HasValue(finishDate) Then
        DateDiffWorkingDays = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)
    startSerial = CLng(CDbl(CDate(startDate)))
    finishSerial = CLng(CDbl(CDate(finishDate)))

    If finishSerial < startSerial Then
        DateDiffWorkingDays = Empty
        Exit Function
    End If

    Select Case calType
        Case CALENDAR_7D
            DateDiffWorkingDays = finishSerial - startSerial + 1
        Case CALENDAR_6D
            DateDiffWorkingDays = CountWorkingDaysFast(startSerial, finishSerial, 6)
        Case CALENDAR_5D
            DateDiffWorkingDays = CountWorkingDaysFast(startSerial, finishSerial, 5)
        Case Else
            Err.Raise vbObjectError + 7603, "DateDiffWorkingDays", _
                "Unsupported calendar type: " & calType
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Actualise Apply Lag sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Lag without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Function ApplyLag( _
    ByVal baseDate As Variant, _
    ByVal lagDays As Variant, _
    Optional ByVal calendarType As Variant = "", _
    Optional ByVal linkType As String = "FS") As Variant

    Dim calType As String
    Dim lagVal As Long
    Dim anchorDate As Variant
    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("ApplyLag", "Calendar")

    If Not HasValue(baseDate) Then
        ApplyLag = Empty
        Exit Function
    End If

    If Not HasValue(lagDays) Then
        lagVal = 0
    ElseIf IsNumeric(lagDays) Then
        lagVal = CLng(lagDays)
    Else
        ApplyLag = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)

    Select Case UCase$(Trim$(linkType))
        Case "FS", ""
            anchorDate = NextWorkingDay(baseDate, calType)
        Case "SS", "FF"
            anchorDate = NormalizeToWorkingDayForward(baseDate, calType)
        Case Else
            Err.Raise vbObjectError + 7602, "ApplyLag", _
                "Unsupported link type for calendar lag: " & linkType
    End Select

    ApplyLag = ShiftWorkingDays(anchorDate, lagVal, calType)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Offset Working Days sans modifier les donnees d'entree.
' EN: Returns the Offset Working Days value without mutating input data.
'------------------------------------------------------------------------------

Public Function OffsetWorkingDays( _
    ByVal dateValue As Variant, _
    ByVal offsetDays As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim offsetVal As Long
    Dim calType As String

    If Not HasValue(dateValue) Then
        OffsetWorkingDays = Empty
        Exit Function
    End If

    If HasValue(offsetDays) Then
        offsetVal = CLng(CDbl(offsetDays))
    Else
        offsetVal = 0
    End If

    calType = NormalizeCalendarType(calendarType)
    OffsetWorkingDays = ShiftWorkingDays(dateValue, offsetVal, calType)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Signed Working Day Offset sans modifier les donnees d'entree.
' EN: Returns the Signed Working Day Offset value without mutating input data.
'------------------------------------------------------------------------------

Public Function SignedWorkingDayDelta( _
    ByVal startDate As Variant, _
    ByVal finishDate As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    Dim diffVal As Variant
    Dim calType As String

    If Not HasValue(startDate) Then
        SignedWorkingDayDelta = Empty
        Exit Function
    End If

    If Not HasValue(finishDate) Then
        SignedWorkingDayDelta = Empty
        Exit Function
    End If

    calType = NormalizeCalendarType(calendarType)

    If CLng(CDbl(CDate(finishDate))) >= CLng(CDbl(CDate(startDate))) Then
        diffVal = DateDiffWorkingDays(startDate, finishDate, calType)
        If HasValue(diffVal) Then
            SignedWorkingDayDelta = CLng(diffVal) - 1
            If CLng(CDbl(CDate(finishDate))) > CLng(CDbl(CDate(startDate))) Then
                If CLng(SignedWorkingDayDelta) = 0 Then SignedWorkingDayDelta = 1
            End If
        End If
    Else
        diffVal = DateDiffWorkingDays(finishDate, startDate, calType)
        If HasValue(diffVal) Then
            SignedWorkingDayDelta = -(CLng(diffVal) - 1)
            If CLng(SignedWorkingDayDelta) = 0 Then SignedWorkingDayDelta = -1
        End If
    End If

End Function

Public Function SignedWorkingDayOffset( _
    ByVal requiredDate As Variant, _
    ByVal actualDate As Variant, _
    Optional ByVal calendarType As Variant = "") As Variant

    SignedWorkingDayOffset = SignedWorkingDayDelta(requiredDate, actualDate, calendarType)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Shift Working Days sans modifier les donnees d'entree.
' EN: Returns the Shift Working Days value without mutating input data.
'------------------------------------------------------------------------------

Private Function ShiftWorkingDays( _
    ByVal dateValue As Variant, _
    ByVal offsetDays As Long, _
    ByVal calendarType As String) As Variant

    Dim d As Date
    Dim i As Long

    If Not HasValue(dateValue) Then
        ShiftWorkingDays = Empty
        Exit Function
    End If

    If offsetDays >= 0 Then
        d = NormalizeToWorkingDayForward(dateValue, calendarType)
        For i = 1 To offsetDays
            d = CDate(NextWorkingDay(d, calendarType))
        Next i
    Else
        d = NormalizeToWorkingDayBackward(dateValue, calendarType)
        For i = 1 To Abs(offsetDays)
            d = CDate(PreviousWorkingDay(d, calendarType))
        Next i
    End If

    ShiftWorkingDays = CDbl(d)

End Function

'------------------------------------------------------------------------------
' FR: Normalise To Working Day Forward dans un format exploitable.
' EN: Normalizes To Working Day Forward into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeToWorkingDayForward( _
    ByVal dateValue As Variant, _
    ByVal calendarType As String) As Date

    Dim d As Date

    d = CDate(dateValue)
    Do While Not IsWorkingDay(d, calendarType)
        d = d + 1
    Loop

    NormalizeToWorkingDayForward = d

End Function

'------------------------------------------------------------------------------
' FR: Normalise To Working Day Backward dans un format exploitable.
' EN: Normalizes To Working Day Backward into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeToWorkingDayBackward( _
    ByVal dateValue As Variant, _
    ByVal calendarType As String) As Date

    Dim d As Date

    d = CDate(dateValue)
    Do While Not IsWorkingDay(d, calendarType)
        d = d - 1
    Loop

    NormalizeToWorkingDayBackward = d

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Working Days Fast sans exposer de mutateur sur l'etat source.
' EN: Returns the Working Days Fast value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function CountWorkingDaysFast( _
    ByVal startSerial As Long, _
    ByVal finishSerial As Long, _
    ByVal workDaysPerWeek As Long) As Long

    Dim totalDays As Long
    Dim fullWeeks As Long
    Dim remainderDays As Long
    Dim i As Long
    Dim d As Date

    totalDays = finishSerial - startSerial + 1
    fullWeeks = totalDays \ 7
    remainderDays = totalDays Mod 7

    CountWorkingDaysFast = fullWeeks * workDaysPerWeek

    For i = 0 To remainderDays - 1
        d = CDate(startSerial + i)
        If Weekday(d, vbMonday) <= workDaysPerWeek Then
            CountWorkingDaysFast = CountWorkingDaysFast + 1
        End If
    Next i

End Function
