Attribute VB_Name = "mod_Settings"
Option Explicit

'===============================================================================
' MODULE : mod_Settings
' DOMAINE / DOMAIN : Settings / Language
'
' FR
' Possede les reglages persistants, les langues par domaine et les controles du panneau Settings.
' Ne decide aucune politique de calcul planning.
'
' EN
' Owns persisted settings, per-domain languages and Settings panel controls.
' Does not decide planning calculation policy.
'
' CONTRATS / CONTRACTS : Settings_Initialize, Settings_HydrateRuntimeState, Settings_GetOwnerLanguage, Settings_ToggleInfoMessages, Settings_ToggleGlobalLanguage, Settings_ToggleGlobalActivated
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const SETTINGS_SHEET As String = "SETTINGS"
Private Const SETTINGS_PREFIX As String = "SET_"

Private Const CELL_GLOBAL_LANGUAGE As String = "X2"
Private Const CELL_GLOBAL_ACTIVATED As String = "X3"
Private Const CELL_DASHBOARD_LANGUAGE As String = "X4"
Private Const CELL_GANTT_LANGUAGE As String = "X5"
Private Const CELL_SCURVE_LANGUAGE As String = "X6"
Private Const CELL_WBS_LANGUAGE As String = "X7"
Private Const CELL_EVENT_LANGUAGE As String = "X8"
Private Const CELL_INFO_ENABLED As String = "X9"
Private Const CELL_CONSTRAINTS_LANGUAGE As String = "X10"
Private Const CELL_DATE_DISPLAY_MODE As String = "X11"
Private Const SETTINGS_STORAGE_VERSION As String = "SETTINGS_STORAGE_V2"

Private Const DATE_MODE_DMY As String = "DMY"
Private Const DATE_MODE_MDY As String = "MDY"
Private Const DATE_MODE_ISO As String = "ISO"

Private Const DATE_FORMAT_DMY As String = "dd/mm/yyyy"
Private Const DATE_FORMAT_MDY As String = "mm/dd/yyyy"
Private Const DATE_FORMAT_ISO As String = "yyyy-mm-dd"
Private Const DATETIME_FORMAT_DMY As String = "dd/mm/yyyy hh:mm:ss"
Private Const DATETIME_FORMAT_MDY As String = "mm/dd/yyyy hh:mm:ss"
Private Const DATETIME_FORMAT_ISO As String = "yyyy-mm-dd hh:mm:ss"
Private Const DATE_COMPACT_FORMAT_DMY As String = "dd/mm"
Private Const DATE_COMPACT_FORMAT_MDY As String = "mm/dd"
Private Const DATE_COMPACT_FORMAT_ISO As String = "yyyy-mm-dd"

Private Const DATE_CONTROL_LABEL_NAME As String = "SET_DATE_FORMAT_VALUES"
Private Const DATE_CONTROL_BG_NAME As String = "SET_DATE_FORMAT_TRACK"
Private Const DATE_CONTROL_KNOB_NAME As String = "SET_DATE_FORMAT_KNOB"

Private Const GANTT_REFERENCE_TRACK_WIDTH As Double = 28
Private Const GANTT_THREE_POSITION_TRACK_WIDTH As Double = 46
Private Const GANTT_REFERENCE_TRACK_HEIGHT As Double = 12
Private Const GANTT_REFERENCE_KNOB_SIZE As Double = 9
Private Const GANTT_REFERENCE_LABEL_GAP As Double = 7
Private Const SETTINGS_TRACK_WIDTH As Double = 44
Private Const SETTINGS_TRACK_HEIGHT As Double = 16
Private Const SETTINGS_KNOB_SIZE As Double = 12
Private Const DATE_CONTROL_LABEL_WIDTH As Double = 104

Private Const MODULE_DASHBOARD As String = "DASHBOARD"
Private Const MODULE_GANTT As String = "GANTT"
Private Const MODULE_SCURVE As String = "SCURVE"
Private Const MODULE_WBS As String = "WBS"
Private Const MODULE_CONSTRAINTS As String = "CONSTRAINTS"
Private Const MODULE_EVENT As String = "EVENT"

'------------------------------------------------------------------------------
' FR: Met a jour Settings Initialize dans le contexte settings and language.
' EN: Updates Settings Initialize in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_Initialize()

    Dim ws As Worksheet
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo SafeExit

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set ws = Settings_EnsureSheet()
    Settings_BuildLayout ws

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Hydrate Runtime State dans le contexte settings and language.
' EN: Updates Settings Hydrate Runtime State in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_HydrateRuntimeState()

    EventHistory_SetShowInfo Settings_GetInfoEnabled()
    EventHistory_SetLanguage Settings_GetOwnerLanguage(MODULE_EVENT)
    WBS_SetLanguage Settings_GetOwnerLanguage(MODULE_WBS)
    Gantt_SetLanguage Settings_GetOwnerLanguage(MODULE_GANTT)
    SCurve_SetLanguage Settings_GetOwnerLanguage(MODULE_SCURVE)
    Dashboard_SetLanguage Settings_GetOwnerLanguage(MODULE_DASHBOARD)
    Constraints_SetLanguage Settings_GetOwnerLanguage(MODULE_CONSTRAINTS)

End Sub
'------------------------------------------------------------------------------
' FR: Retourne la langue persistante d'un owner sans modifier le workbook.
' EN: Returns an owner's persisted language without mutating the workbook.
'------------------------------------------------------------------------------
Public Function Settings_GetOwnerLanguage(ByVal ownerKey As String) As String

    Dim ws As Worksheet
    Dim globalLanguage As String
    Dim storageCell As String

    On Error GoTo UseFallback
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    globalLanguage = Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).Value2), "EN")
    storageCell = Settings_ModuleStorageCell(ownerKey)
    If storageCell = "" Then
        Settings_GetOwnerLanguage = globalLanguage
    Else
        Settings_GetOwnerLanguage = _
            Settings_NormalizeLanguage(CStr(ws.Range(storageCell).Value2), globalLanguage)
    End If
    Exit Function

UseFallback:
    Settings_GetOwnerLanguage = "EN"

End Function

'------------------------------------------------------------------------------
' FR: Retourne le reglage Info persiste sans modifier le workbook.
' EN: Returns the persisted Info setting without mutating the workbook.
'------------------------------------------------------------------------------
Public Function Settings_GetInfoEnabled() As Boolean

    Dim ws As Worksheet

    On Error GoTo UseFallback
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    Settings_GetInfoEnabled = Settings_InfoIsEnabled(ws)
    Exit Function

UseFallback:
    Settings_GetInfoEnabled = True

End Function

'------------------------------------------------------------------------------
' FR: Migre explicitement le stockage de langue vers le schema V2.
' EN: Explicitly migrates language storage to the V2 schema.
'------------------------------------------------------------------------------
Public Sub Settings_MigrateLanguageStorageV2()

    Dim ws As Worksheet
    Dim globalLanguage As String

    Set ws = Settings_EnsureSheet()
    globalLanguage = Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).Value2), "EN")

    If Not Settings_IsRecognizedLanguage(CStr(ws.Range(CELL_CONSTRAINTS_LANGUAGE).Value2)) Then
        ws.Range(CELL_CONSTRAINTS_LANGUAGE).Value2 = globalLanguage
    End If
    ws.Range("X1").Value2 = SETTINGS_STORAGE_VERSION
    ws.Columns("X:Y").Hidden = True

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Info Messages dans le contexte settings and language.
' EN: Updates Settings Toggle Info Messages in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleInfoMessages()

    Dim ws As Worksheet
    Dim newValue As Boolean

    Set ws = Settings_EnsureSheet()

    newValue = Not Settings_InfoIsEnabled(ws)
    ws.Range(CELL_INFO_ENABLED).value = newValue

    EventHistory_SetShowInfo newValue
    Refresh_EventHistory_View
    Settings_RefreshVisuals ws

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Global Language dans le contexte settings and language.
' EN: Updates Settings Toggle Global Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleGlobalLanguage()

    Dim ws As Worksheet
    Dim languageCode As String

    Set ws = Settings_EnsureSheet()

    languageCode = Settings_OppositeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).value))

    If Settings_GlobalIsActivated(ws) Then
        Settings_ApplyGlobalLanguageTransaction ws, languageCode, True
    Else
        ws.Range(CELL_GLOBAL_LANGUAGE).value = languageCode
        Settings_RefreshVisuals ws
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Global Activated dans le contexte settings and language.
' EN: Updates Settings Toggle Global Activated in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleGlobalActivated()

    Dim ws As Worksheet
    Dim newValue As Boolean

    Set ws = Settings_EnsureSheet()

    newValue = Not Settings_GlobalIsActivated(ws)
    If newValue Then
        Settings_ApplyGlobalLanguageTransaction ws, _
            Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).value), "EN"), False
    Else
        ws.Range(CELL_GLOBAL_ACTIVATED).value = False
        Settings_RefreshVisuals ws
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Dashboard Language dans le contexte settings and language.
' EN: Updates Settings Toggle Dashboard Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleDashboardLanguage()
    Settings_ToggleModuleLanguage MODULE_DASHBOARD
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Gantt Language dans le contexte settings and language.
' EN: Updates Settings Toggle Gantt Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleGanttLanguage()
    Settings_ToggleModuleLanguage MODULE_GANTT
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle SCurve Language dans le contexte settings and language.
' EN: Updates Settings Toggle SCurve Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleSCurveLanguage()
    Settings_ToggleModuleLanguage MODULE_SCURVE
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle WBSLanguage dans le contexte settings and language.
' EN: Updates Settings Toggle WBSLanguage in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleWBSLanguage()
    Settings_ToggleModuleLanguage MODULE_WBS
End Sub

'------------------------------------------------------------------------------
' FR: Bascule la langue de l'owner Constraints depuis SETTINGS.
' EN: Toggles the Constraints owner language from SETTINGS.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleConstraintsLanguage()
    Settings_ToggleModuleLanguage MODULE_CONSTRAINTS
End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Event History Language dans le contexte settings and language.
' EN: Updates Settings Toggle Event History Language in the settings and language context.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleEventHistoryLanguage()
    Settings_ToggleModuleLanguage MODULE_EVENT
End Sub

'------------------------------------------------------------------------------
' FR: Retourne le mode d'affichage des dates sans modifier le stockage.
' EN: Returns the date display mode without mutating persisted storage.
'------------------------------------------------------------------------------
Public Function Settings_GetDateDisplayMode() As String

    Dim ws As Worksheet

    On Error GoTo UseFallback
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    Settings_GetDateDisplayMode = _
        Settings_NormalizeDateDisplayMode(CStr(ws.Range(CELL_DATE_DISPLAY_MODE).Value2))
    Exit Function

UseFallback:
    Settings_GetDateDisplayMode = DATE_MODE_DMY

End Function

'------------------------------------------------------------------------------
' FR: Retourne le format invariant de date visible selectionne.
' EN: Returns the selected invariant visible date format.
'------------------------------------------------------------------------------
Public Function Settings_GetDateNumberFormat() As String
    Settings_GetDateNumberFormat = _
        Settings_DateNumberFormatForMode(Settings_GetDateDisplayMode())
End Function

'------------------------------------------------------------------------------
' FR: Retourne le format invariant de date et heure visible selectionne.
' EN: Returns the selected invariant visible date-time format.
'------------------------------------------------------------------------------
Public Function Settings_GetDateTimeNumberFormat() As String
    Settings_GetDateTimeNumberFormat = _
        Settings_DateTimeNumberFormatForMode(Settings_GetDateDisplayMode())
End Function

'------------------------------------------------------------------------------
' FR: Retourne le format invariant compact utilise par les axes visibles.
' EN: Returns the invariant compact format used by visible date axes.
'------------------------------------------------------------------------------
Public Function Settings_GetCompactDateNumberFormat() As String
    Settings_GetCompactDateNumberFormat = _
        Settings_CompactDateNumberFormatForMode(Settings_GetDateDisplayMode())
End Function

'------------------------------------------------------------------------------
' FR: Fait avancer le selecteur partage DMY, MDY, ISO.
' EN: Advances the shared DMY, MDY, ISO selector.
'------------------------------------------------------------------------------
Public Sub Settings_ToggleDateDisplayMode()

    Dim requestedMode As String

    Select Case Settings_GetDateDisplayMode()
        Case DATE_MODE_DMY
            requestedMode = DATE_MODE_MDY
        Case DATE_MODE_MDY
            requestedMode = DATE_MODE_ISO
        Case Else
            requestedMode = DATE_MODE_DMY
    End Select

    Settings_ApplyDateDisplayMode requestedMode

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Toggle Module Language dans le contexte settings and language.
' EN: Updates Settings Toggle Module Language in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_ToggleModuleLanguage(ByVal moduleKey As String)

    Dim ws As Worksheet
    Dim languageCode As String
    Dim storageCell As String
    Dim oldLanguage As String
    Dim oldStoredValue As Variant
    Dim oldGlobalActivated As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim headerSnapshot As Object
    Dim stateCaptured As Boolean
    Dim applyStarted As Boolean
    Dim headersApplied As Boolean
    Dim errorDescription As String

    Set ws = Settings_EnsureSheet()

    storageCell = Settings_ModuleStorageCell(moduleKey)
    If storageCell = "" Then Exit Sub

    oldLanguage = Settings_ModuleLanguage(ws, moduleKey)
    oldStoredValue = ws.Range(storageCell).Value2
    oldGlobalActivated = ws.Range(CELL_GLOBAL_ACTIVATED).Value2
    languageCode = Settings_OppositeLanguage(oldLanguage)
    If languageCode = oldLanguage Then Exit Sub

    On Error GoTo ApplyFailed
    Settings_PreflightLanguageOwner moduleKey
    VisibleHeaders_PreflightOwner moduleKey, languageCode
    Set headerSnapshot = VisibleHeaders_CaptureOwnerSnapshot(moduleKey)

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    stateCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    VisibleHeaders_ApplyOwnerFromSnapshot headerSnapshot, moduleKey, languageCode
    headersApplied = True
    applyStarted = True
    Settings_ApplySingleModule moduleKey, languageCode
    ws.Range(storageCell).Value2 = languageCode
    ws.Range(CELL_GLOBAL_ACTIVATED).Value2 = False
    Settings_RefreshVisuals ws

CleanExit:
    Schema_ClearPhysicalHeaderLanguageOverrides
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    Exit Sub

ApplyFailed:
    errorDescription = Err.Source & ": " & Err.Description
    On Error Resume Next
    If headersApplied Then VisibleHeaders_RestoreSnapshot headerSnapshot
    If applyStarted Then Settings_ApplySingleModule moduleKey, oldLanguage
    ws.Range(storageCell).Value2 = oldStoredValue
    ws.Range(CELL_GLOBAL_ACTIVATED).Value2 = oldGlobalActivated
    Settings_RefreshVisuals ws
    Schema_ClearPhysicalHeaderLanguageOverrides
    On Error GoTo 0
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    CalcBridge_ShowSingleConsoleMessage "STOP", _
        "La langue " & moduleKey & " n'a pas ete modifiee. " & errorDescription, _
        "The " & moduleKey & " language was not changed. " & errorDescription

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Apply Single Module dans le contexte settings and language.
' EN: Updates Settings Apply Single Module in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_ApplySingleModule( _
    ByVal moduleKey As String, _
    ByVal languageCode As String)

    Select Case UCase$(Trim$(moduleKey))
        Case MODULE_DASHBOARD
            Dashboard_ApplyLanguage languageCode
        Case MODULE_GANTT
            Gantt_ApplyLanguage languageCode
        Case MODULE_SCURVE
            SCurve_ApplyLanguage languageCode
        Case MODULE_WBS
            WBS_ApplyLanguage languageCode
        Case MODULE_CONSTRAINTS
            Constraints_ApplyLanguage languageCode
        Case MODULE_EVENT
            EventHistory_ApplyLanguage languageCode
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Applique les six owners comme une transaction unique depuis GLOBAL.
' EN: Applies all six owners as one transaction from GLOBAL.
'------------------------------------------------------------------------------
Private Sub Settings_ApplyGlobalLanguageTransaction( _
    ByVal ws As Worksheet, _
    ByVal requestedLanguage As String, _
    ByVal persistGlobalLanguage As Boolean)

    Dim ownerKeys As Variant
    Dim oldLanguages(0 To 5) As String
    Dim oldStoredValues(0 To 5) As Variant
    Dim oldGlobalLanguage As Variant
    Dim oldGlobalActivated As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim headerSnapshot As Object
    Dim targetLanguage As String
    Dim stateCaptured As Boolean
    Dim applyStarted As Boolean
    Dim headersApplied As Boolean
    Dim failedOwner As String
    Dim errorDescription As String
    Dim i As Long

    targetLanguage = Settings_NormalizeLanguage(requestedLanguage, "EN")
    ownerKeys = Settings_LanguageOwnerKeys()

    On Error GoTo ApplyFailed

    oldGlobalLanguage = ws.Range(CELL_GLOBAL_LANGUAGE).Value2
    oldGlobalActivated = ws.Range(CELL_GLOBAL_ACTIVATED).Value2

    For i = LBound(ownerKeys) To UBound(ownerKeys)
        oldLanguages(i) = Settings_ModuleLanguage(ws, CStr(ownerKeys(i)))
        oldStoredValues(i) = ws.Range(Settings_ModuleStorageCell(CStr(ownerKeys(i)))).Value2
    Next i

    For i = LBound(ownerKeys) To UBound(ownerKeys)
        Settings_PreflightLanguageOwner CStr(ownerKeys(i))
    Next i
    VisibleHeaders_PreflightAll targetLanguage
    Set headerSnapshot = VisibleHeaders_CaptureAllSnapshot()

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    stateCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    VisibleHeaders_ApplyAllFromSnapshot headerSnapshot, targetLanguage
    headersApplied = True
    applyStarted = True
    For i = LBound(ownerKeys) To UBound(ownerKeys)
        failedOwner = CStr(ownerKeys(i))
        Settings_ApplySingleModule failedOwner, targetLanguage
    Next i

    For i = LBound(ownerKeys) To UBound(ownerKeys)
        ws.Range(Settings_ModuleStorageCell(CStr(ownerKeys(i)))).Value2 = targetLanguage
    Next i
    If persistGlobalLanguage Then ws.Range(CELL_GLOBAL_LANGUAGE).Value2 = targetLanguage
    ws.Range(CELL_GLOBAL_ACTIVATED).Value2 = True
    Settings_RefreshVisuals ws

CleanExit:
    Schema_ClearPhysicalHeaderLanguageOverrides
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    Exit Sub

ApplyFailed:
    errorDescription = Err.Source & ": " & Err.Description
    On Error Resume Next
    If headersApplied Then VisibleHeaders_RestoreSnapshot headerSnapshot
    If applyStarted Then
        For i = UBound(ownerKeys) To LBound(ownerKeys) Step -1
            Settings_ApplySingleModule CStr(ownerKeys(i)), oldLanguages(i)
        Next i
    End If
    For i = LBound(ownerKeys) To UBound(ownerKeys)
        ws.Range(Settings_ModuleStorageCell(CStr(ownerKeys(i)))).Value2 = oldStoredValues(i)
    Next i
    ws.Range(CELL_GLOBAL_LANGUAGE).Value2 = oldGlobalLanguage
    ws.Range(CELL_GLOBAL_ACTIVATED).Value2 = oldGlobalActivated
    Settings_RefreshVisuals ws
    Schema_ClearPhysicalHeaderLanguageOverrides
    On Error GoTo 0
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    CalcBridge_ShowSingleConsoleMessage "STOP", _
        "Le changement GLOBAL a echoue sur " & failedOwner & ". " & errorDescription, _
        "The GLOBAL language change failed on " & failedOwner & ". " & errorDescription

End Sub

'------------------------------------------------------------------------------
' FR: Retourne les owners dans l'ordre transactionnel canonique.
' EN: Returns owners in canonical transaction order.
'------------------------------------------------------------------------------
Private Function Settings_LanguageOwnerKeys() As Variant
    Settings_LanguageOwnerKeys = Array( _
        MODULE_DASHBOARD, MODULE_GANTT, MODULE_SCURVE, _
        MODULE_WBS, MODULE_CONSTRAINTS, MODULE_EVENT)
End Function

'------------------------------------------------------------------------------
' FR: Valide les ressources minimales d'un owner avant toute mutation.
' EN: Validates an owner's minimum resources before any mutation.
'------------------------------------------------------------------------------
Private Sub Settings_PreflightLanguageOwner(ByVal moduleKey As String)

    Select Case UCase$(Trim$(moduleKey))
        Case MODULE_DASHBOARD
            Settings_RequireWorksheet "DASHBOARD"
        Case MODULE_GANTT
            Settings_RequireWorksheet "GANTT"
        Case MODULE_SCURVE
            Settings_RequireWorksheet "SCURVE"
        Case MODULE_WBS
            Settings_RequireTable "WBS", "tbl_WBS"
        Case MODULE_CONSTRAINTS
            Settings_RequireTable "CONSTRAINTS", "tbl_CONSTRAINTS"
        Case MODULE_EVENT
            Settings_RequireTable "EVENT_HISTORY", "tbl_EVENT_HISTORY"
            Settings_RequireTable "EVENT_ACK", "tbl_EVENT_ACK"
        Case Else
            Err.Raise vbObjectError + 5210, "Settings_PreflightLanguageOwner", _
                "Unknown language owner: " & moduleKey
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Ensure Sheet dans le contexte settings and language.
' EN: Updates Settings Ensure Sheet in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_EnsureSheet() As Worksheet

    Dim ws As Worksheet
    Dim wsAfter As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set wsAfter = ThisWorkbook.Worksheets("DASHBOARD")
        On Error GoTo 0

        If wsAfter Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=wsAfter)
        End If

        ws.Name = SETTINGS_SHEET
    End If

    ws.Visible = xlSheetVisible
    Set Settings_EnsureSheet = ws

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Info Is Enabled dans le contexte settings and language.
' EN: Updates Settings Info Is Enabled in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_InfoIsEnabled(ByVal ws As Worksheet) As Boolean

    Dim rawValue As Variant
    Dim textValue As String

    If ws Is Nothing Then Exit Function

    rawValue = ws.Range(CELL_INFO_ENABLED).value
    If VarType(rawValue) = vbBoolean Then
        Settings_InfoIsEnabled = CBool(rawValue)
        Exit Function
    End If

    textValue = UCase$(Trim$(CStr(rawValue)))
    Settings_InfoIsEnabled = _
        (textValue = "TRUE" Or textValue = "VRAI" Or textValue = "YES" Or _
         textValue = "Y" Or textValue = "1" Or textValue = "ON")

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Global Is Activated dans le contexte settings and language.
' EN: Updates Settings Global Is Activated in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_GlobalIsActivated(ByVal ws As Worksheet) As Boolean

    Dim rawValue As Variant
    Dim textValue As String

    If ws Is Nothing Then Exit Function

    rawValue = ws.Range(CELL_GLOBAL_ACTIVATED).value

    If VarType(rawValue) = vbBoolean Then
        Settings_GlobalIsActivated = CBool(rawValue)
        Exit Function
    End If

    textValue = UCase$(Trim$(CStr(rawValue)))
    Settings_GlobalIsActivated = _
        (textValue = "TRUE" Or textValue = "VRAI" Or textValue = "YES" Or textValue = "Y" Or textValue = "1")

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Normalize Language dans le contexte settings and language.
' EN: Updates Settings Normalize Language in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_NormalizeLanguage( _
    ByVal languageCode As String, _
    ByVal fallbackLanguage As String) As String

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            Settings_NormalizeLanguage = "FR"
        Case "EN"
            Settings_NormalizeLanguage = "EN"
        Case Else
            If UCase$(Trim$(fallbackLanguage)) = "FR" Then
                Settings_NormalizeLanguage = "FR"
            Else
                Settings_NormalizeLanguage = "EN"
            End If
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Indique si un code de langue persiste est reconnu.
' EN: Returns whether a persisted language code is recognized.
'------------------------------------------------------------------------------
Private Function Settings_IsRecognizedLanguage(ByVal languageCode As String) As Boolean
    Select Case UCase$(Trim$(languageCode))
        Case "FR", "EN"
            Settings_IsRecognizedLanguage = True
    End Select
End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Opposite Language dans le contexte settings and language.
' EN: Updates Settings Opposite Language in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_OppositeLanguage(ByVal languageCode As String) As String

    If Settings_NormalizeLanguage(languageCode, "EN") = "FR" Then
        Settings_OppositeLanguage = "EN"
    Else
        Settings_OppositeLanguage = "FR"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Module Storage Cell dans le contexte settings and language.
' EN: Updates Settings Module Storage Cell in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ModuleStorageCell(ByVal moduleKey As String) As String

    Select Case UCase$(Trim$(moduleKey))
        Case MODULE_DASHBOARD
            Settings_ModuleStorageCell = CELL_DASHBOARD_LANGUAGE
        Case MODULE_GANTT
            Settings_ModuleStorageCell = CELL_GANTT_LANGUAGE
        Case MODULE_SCURVE
            Settings_ModuleStorageCell = CELL_SCURVE_LANGUAGE
        Case MODULE_WBS
            Settings_ModuleStorageCell = CELL_WBS_LANGUAGE
        Case MODULE_CONSTRAINTS
            Settings_ModuleStorageCell = CELL_CONSTRAINTS_LANGUAGE
        Case MODULE_EVENT
            Settings_ModuleStorageCell = CELL_EVENT_LANGUAGE
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Module Language dans le contexte settings and language.
' EN: Updates Settings Module Language in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ModuleLanguage( _
    ByVal ws As Worksheet, _
    ByVal moduleKey As String) As String

    Dim cellAddress As String

    cellAddress = Settings_ModuleStorageCell(moduleKey)
    If cellAddress = "" Then
        Settings_ModuleLanguage = "EN"
    Else
        Settings_ModuleLanguage = _
            Settings_NormalizeLanguage( _
                CStr(ws.Range(cellAddress).value), _
                Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).Value2), "EN"))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Build Layout dans le contexte settings and language.
' EN: Updates Settings Build Layout in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_BuildLayout(ByVal ws As Worksheet)

    Dim previousSheet As Object
    Dim dateTitle As Shape
    Dim panelLeft As Double
    Dim panelTop As Double
    Dim settingsToggleCenterX As Double
    Dim horizontalScale As Double
    Dim dateTrackWidth As Double
    Dim dateTrackHeight As Double
    Dim dateKnobSize As Double
    Dim dateLabelGap As Double
    Dim dateTrackLeft As Double
    Dim dateTrackTop As Double
    Dim dateLabelLeft As Double

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set previousSheet = Application.ActiveSheet
    On Error GoTo 0

    Settings_DeleteGeneratedShapes ws
    Settings_PrepareCanvas ws

    panelLeft = ws.Range("B4").Left
    panelTop = ws.Range("B4").Top
    settingsToggleCenterX = panelLeft + 28 + 310 + (SETTINGS_TRACK_WIDTH / 2)
    horizontalScale = SETTINGS_TRACK_WIDTH / GANTT_REFERENCE_TRACK_WIDTH
    dateTrackWidth = GANTT_THREE_POSITION_TRACK_WIDTH * horizontalScale
    dateTrackHeight = GANTT_REFERENCE_TRACK_HEIGHT * _
        (SETTINGS_TRACK_HEIGHT / GANTT_REFERENCE_TRACK_HEIGHT)
    dateKnobSize = GANTT_REFERENCE_KNOB_SIZE * _
        (SETTINGS_KNOB_SIZE / GANTT_REFERENCE_KNOB_SIZE)
    dateLabelGap = GANTT_REFERENCE_LABEL_GAP * horizontalScale
    dateTrackLeft = settingsToggleCenterX - (dateTrackWidth / 2)
    dateTrackTop = panelTop + 388 + ((20 - dateTrackHeight) / 2)
    dateLabelLeft = dateTrackLeft - dateLabelGap - DATE_CONTROL_LABEL_WIDTH

    Settings_AddPanel ws, "SET_PANEL_LANGUAGE", panelLeft, panelTop, 610, 444, RGB(255, 255, 255), RGB(214, 220, 228)
    Settings_AddCenteredTitle ws, "SET_TITLE_LANGUAGE", Settings_L(ws, "Langue", "Language"), panelLeft + 20, panelTop + 16, 570, 28, 15

    Settings_AddLanguageSwitch ws, "GLOBAL", "GLOBAL", panelLeft + 28, panelTop + 64, _
        Settings_ModuleDisplayLanguage(ws, "GLOBAL"), "Settings_ToggleGlobalLanguage", True
    Settings_AddActivatedControl ws, panelLeft + 430, panelTop + 63, Settings_GlobalIsActivated(ws)

    Settings_AddLanguageSwitch ws, MODULE_DASHBOARD, "Dashboard", panelLeft + 28, panelTop + 118, _
        Settings_ModuleLanguage(ws, MODULE_DASHBOARD), "Settings_ToggleDashboardLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_GANTT, "Gantt", panelLeft + 28, panelTop + 162, _
        Settings_ModuleLanguage(ws, MODULE_GANTT), "Settings_ToggleGanttLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_SCURVE, "S-Curve", panelLeft + 28, panelTop + 206, _
        Settings_ModuleLanguage(ws, MODULE_SCURVE), "Settings_ToggleSCurveLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_WBS, "WBS", panelLeft + 28, panelTop + 250, _
        Settings_ModuleLanguage(ws, MODULE_WBS), "Settings_ToggleWBSLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_CONSTRAINTS, _
        Settings_L(ws, "Contraintes", "Constraints"), panelLeft + 28, panelTop + 294, _
        Settings_ModuleLanguage(ws, MODULE_CONSTRAINTS), "Settings_ToggleConstraintsLanguage", False
    Settings_AddLanguageSwitch ws, MODULE_EVENT, "Messages & Event History", panelLeft + 28, panelTop + 338, _
        Settings_ModuleLanguage(ws, MODULE_EVENT), "Settings_ToggleEventHistoryLanguage", False
    Settings_AddInfoSwitch ws, panelLeft + 430, panelTop + 338, Settings_InfoIsEnabled(ws)

    Set dateTitle = Settings_AddTextShape(ws, "SET_DATE_FORMAT_TITLE", _
        Settings_L(ws, "Format des dates", "Date format"), _
        panelLeft + 28, panelTop + 388, 150, 20, msoAlignLeft, 10.5, True)
    ThreePositionControl_Ensure ws, _
        DATE_CONTROL_LABEL_NAME, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
        dateLabelLeft, panelTop + 388, DATE_CONTROL_LABEL_WIDTH, "DMY / MDY / ISO", _
        dateTrackLeft, dateTrackTop, dateTrackWidth, dateTrackHeight, dateKnobSize, _
        "Settings_ToggleDateDisplayMode", xlFreeFloating
    ThreePositionControl_Refresh ws, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
        Settings_DateDisplayPosition(Settings_GetDateDisplayMode())

    Settings_AddPanel ws, "SET_PANEL_RESET", panelLeft + 634, panelTop, 300, 226, RGB(255, 255, 255), RGB(214, 220, 228)
    Settings_AddCenteredTitle ws, "SET_TITLE_RESET", Settings_L(ws, "R" & ChrW$(&HE9) & "initialisation", "Reset"), panelLeft + 654, panelTop + 16, 260, 28, 14
    Settings_AddCommandButton ws, "SET_BTN_CLEAR_HISTORY", Settings_L(ws, "Nettoyer historique", "Clear History"), "ClearPlanningEventHistory", panelLeft + 674, panelTop + 52, 220, 34, RGB(68, 114, 196), RGB(255, 255, 255)
    Settings_AddCommandButton ws, "SET_BTN_CLEAR_ACK", Settings_L(ws, "Nettoyer les messages acquités", "Clear Acknowledged"), "ClearPlanningWarningAcknowledgements", panelLeft + 674, panelTop + 92, 220, 34, RGB(68, 114, 196), RGB(255, 255, 255)
    Settings_AddCommandButton ws, "SET_BTN_CLEAN_DASHBOARD", Settings_L(ws, "Nettoyer Dashboard", "Clean Dashboard"), "Reset_Dashboard", panelLeft + 674, panelTop + 132, 220, 34, RGB(68, 114, 196), RGB(255, 255, 255)
    Settings_AddCommandButton ws, "SET_BTN_RESET_PLANNING", Settings_L(ws, "R" & ChrW$(&HE9) & "initialiser planning", "Reset Planning"), "Reset_Planning", panelLeft + 674, panelTop + 172, 220, 34, RGB(192, 120, 0), RGB(255, 255, 255)

    Settings_AddPanel ws, "SET_PANEL_DANGER", panelLeft + 634, panelTop + 250, 300, 150, RGB(255, 247, 247), RGB(220, 80, 80)
    Settings_AddCenteredTitle ws, "SET_TITLE_DANGER", Settings_L(ws, "Zone de danger", "Danger Zone"), panelLeft + 654, panelTop + 266, 260, 28, 14
    Settings_AddCommandButton ws, "SET_BTN_FULL_RESET", Settings_L(ws, "R" & ChrW$(&HE9) & "initialisation compl" & ChrW$(&HE8) & "te", "Full Reset"), "Armageddon", panelLeft + 674, panelTop + 316, 220, 42, RGB(192, 0, 0), RGB(255, 255, 255)

    Settings_RefreshVisuals ws

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    If Not previousSheet Is Nothing Then previousSheet.Activate
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Info Switch dans le contexte settings and language.
' EN: Updates Settings Add Info Switch in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddInfoSwitch( _
    ByVal ws As Worksheet, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal isEnabled As Boolean)

    Dim labelShape As Shape
    Dim offShape As Shape
    Dim trackShape As Shape
    Dim knobShape As Shape
    Dim onShape As Shape
    Dim trackLeft As Double
    Dim knobLeft As Double

    trackLeft = leftPos + 58

    Set labelShape = Settings_AddTextShape(ws, "SET_INFO_LABEL", "Info", leftPos, topPos, 50, 24, msoAlignLeft, 9.5, True)
    Set offShape = Settings_AddTextShape(ws, "SET_INFO_OFF", "OFF", trackLeft - 34, topPos, 28, 24, msoAlignCenter, 9.5, True)

    Set trackShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, topPos + 4, 44, 16)
    trackShape.Name = "SET_INFO_TRACK"
    trackShape.Adjustments.item(1) = 0.5
    trackShape.Placement = xlFreeFloating
    trackShape.OnAction = "Settings_ToggleInfoMessages"

    knobLeft = trackLeft + 2
    If isEnabled Then knobLeft = trackLeft + 28
    Set knobShape = ws.Shapes.AddShape(msoShapeOval, knobLeft, topPos + 6, 12, 12)
    knobShape.Name = "SET_INFO_KNOB"
    knobShape.Placement = xlFreeFloating
    knobShape.OnAction = "Settings_ToggleInfoMessages"
    knobShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    knobShape.Line.Visible = msoFalse

    Set onShape = Settings_AddTextShape(ws, "SET_INFO_ON", "ON", trackLeft + 50, topPos, 30, 24, msoAlignCenter, 9.5, True)

    labelShape.OnAction = "Settings_ToggleInfoMessages"
    offShape.OnAction = "Settings_ToggleInfoMessages"
    onShape.OnAction = "Settings_ToggleInfoMessages"
    Settings_FormatToggle trackShape, knobShape, isEnabled

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Module Display Language dans le contexte settings and language.
' EN: Updates Settings Module Display Language in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ModuleDisplayLanguage( _
    ByVal ws As Worksheet, _
    ByVal moduleKey As String) As String

    If UCase$(moduleKey) = "GLOBAL" Then
        Settings_ModuleDisplayLanguage = _
            Settings_NormalizeLanguage(CStr(ws.Range(CELL_GLOBAL_LANGUAGE).value), "EN")
    Else
        Settings_ModuleDisplayLanguage = Settings_ModuleLanguage(ws, moduleKey)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings L dans le contexte settings and language.
' EN: Updates Settings L in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_L( _
    ByVal ws As Worksheet, _
    ByVal frText As String, _
    ByVal enText As String) As String

    If Settings_ModuleDisplayLanguage(ws, "GLOBAL") = "FR" Then
        Settings_L = frText
    Else
        Settings_L = enText
    End If

End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Prepare Canvas dans le contexte settings and language.
' EN: Updates Settings Prepare Canvas in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_PrepareCanvas(ByVal ws As Worksheet)

    ws.cells.Interior.Color = RGB(242, 244, 247)
    ws.cells.Font.Name = "Segoe UI"
    ws.cells.Font.Color = RGB(35, 45, 58)

    With ws.Range("A1:W30")
        .UnMerge
        .ClearContents
        .Interior.Color = RGB(242, 244, 247)
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(35, 45, 58)
    End With

    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B:L").ColumnWidth = 12
    ws.rows("1:30").rowHeight = 22
    ws.rows("1:2").rowHeight = 28

    With ws.Range("B1:K2")
        .Merge
        .value = Settings_L(ws, "Options", "Settings")
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(28, 42, 58)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ws.Tab.Color = RGB(96, 111, 128)

End Sub


'------------------------------------------------------------------------------
' FR: Met a jour Settings Delete Generated Shapes dans le contexte settings and language.
' EN: Updates Settings Delete Generated Shapes in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_DeleteGeneratedShapes(ByVal ws As Worksheet)

    Dim i As Long

    For i = ws.Shapes.Count To 1 Step -1
        If Left$(CStr(ws.Shapes(i).Name), Len(SETTINGS_PREFIX)) = SETTINGS_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Panel dans le contexte settings and language.
' EN: Updates Settings Add Panel in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddPanel( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal fillColor As Long, _
    ByVal lineColor As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
    shp.Name = shapeName
    shp.Placement = xlFreeFloating
    shp.Adjustments.item(1) = 0.08
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = lineColor
    shp.Line.Weight = 0.8
    shp.Shadow.Visible = msoFalse
    shp.ZOrder msoSendToBack

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Centered Title dans le contexte settings and language.
' EN: Updates Settings Add Centered Title in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddCenteredTitle( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal titleText As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal fontSize As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, widthVal, heightVal)
    shp.Name = shapeName
    shp.Placement = xlFreeFloating
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = titleText
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI Semibold"
    shp.TextFrame2.TextRange.Font.Size = fontSize
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(35, 45, 58)

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Language Switch dans le contexte settings and language.
' EN: Updates Settings Add Language Switch in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddLanguageSwitch( _
    ByVal ws As Worksheet, _
    ByVal keyName As String, _
    ByVal labelText As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal languageCode As String, _
    ByVal macroName As String, _
    ByVal emphasize As Boolean)

    Dim labelShape As Shape
    Dim frShape As Shape
    Dim trackShape As Shape
    Dim knobShape As Shape
    Dim enShape As Shape
    Dim trackLeft As Double
    Dim knobLeft As Double
    Dim isEnglish As Boolean

    isEnglish = (Settings_NormalizeLanguage(languageCode, "EN") = "EN")
    trackLeft = leftPos + 310

    Set labelShape = Settings_AddTextShape(ws, "SET_LANG_" & keyName & "_LABEL", labelText, leftPos, topPos, 250, 24, msoAlignLeft, 10.5, emphasize)
    Set frShape = Settings_AddTextShape(ws, "SET_LANG_" & keyName & "_FR", "FR", trackLeft - 34, topPos, 28, 24, msoAlignCenter, 9.5, True)

    Set trackShape = ws.Shapes.AddShape(msoShapeRoundedRectangle, trackLeft, topPos + 4, 44, 16)
    trackShape.Name = "SET_LANG_" & keyName & "_TRACK"
    trackShape.Adjustments.item(1) = 0.5
    trackShape.Placement = xlFreeFloating
    trackShape.OnAction = macroName

    knobLeft = trackLeft + 2
    If isEnglish Then knobLeft = trackLeft + 28
    Set knobShape = ws.Shapes.AddShape(msoShapeOval, knobLeft, topPos + 6, 12, 12)
    knobShape.Name = "SET_LANG_" & keyName & "_KNOB"
    knobShape.Placement = xlFreeFloating
    knobShape.OnAction = macroName
    knobShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    knobShape.Line.Visible = msoFalse

    Set enShape = Settings_AddTextShape(ws, "SET_LANG_" & keyName & "_EN", "EN", trackLeft + 50, topPos, 30, 24, msoAlignCenter, 9.5, True)

    labelShape.OnAction = macroName
    frShape.OnAction = macroName
    enShape.OnAction = macroName
    Settings_FormatToggle trackShape, knobShape, isEnglish

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Text Shape dans le contexte settings and language.
' EN: Updates Settings Add Text Shape in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_AddTextShape( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal textValue As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal alignment As MsoParagraphAlignment, _
    ByVal fontSize As Double, _
    ByVal isBold As Boolean) As Shape

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPos, topPos, widthVal, heightVal)
    shp.Name = shapeName
    shp.Placement = xlFreeFloating
    shp.Line.Visible = msoFalse
    shp.Fill.Visible = msoFalse
    shp.TextFrame2.MarginLeft = 0
    shp.TextFrame2.MarginRight = 0
    shp.TextFrame2.MarginTop = 0
    shp.TextFrame2.MarginBottom = 0
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.Text = textValue
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = alignment
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shp.TextFrame2.TextRange.Font.Size = fontSize
    shp.TextFrame2.TextRange.Font.Bold = isBold
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(55, 67, 82)

    Set Settings_AddTextShape = shp

End Function

'------------------------------------------------------------------------------
' FR: Normalise la geometrie commune des toggles Settings, y compris les shapes persistantes.
' EN: Normalizes shared Settings toggle geometry, including persisted shapes.
'------------------------------------------------------------------------------
Private Sub Settings_NormalizeToggleGeometry( _
    ByVal trackShape As Shape, _
    ByVal knobShape As Shape)

    If trackShape Is Nothing Or knobShape Is Nothing Then Exit Sub

    trackShape.LockAspectRatio = msoFalse
    trackShape.Width = 44
    trackShape.Height = 16
    trackShape.Adjustments.item(1) = 0.5

    knobShape.LockAspectRatio = msoFalse
    knobShape.Width = 12
    knobShape.Height = 12

End Sub

'------------------------------------------------------------------------------
' FR: Aligne les labels gauche/droite sur la geometrie commune des toggles Settings.
' EN: Aligns left/right labels to the shared Settings toggle geometry.
'------------------------------------------------------------------------------
Private Sub Settings_NormalizeToggleLabels( _
    ByVal leftShape As Shape, _
    ByVal trackShape As Shape, _
    ByVal rightShape As Shape)

    If trackShape Is Nothing Then Exit Sub

    If Not leftShape Is Nothing Then
        leftShape.Left = trackShape.Left - 34
        leftShape.Top = trackShape.Top - 4
        leftShape.Width = 28
        leftShape.Height = 24
        leftShape.TextFrame2.TextRange.Font.Size = 9.5
    End If

    If Not rightShape Is Nothing Then
        rightShape.Left = trackShape.Left + 50
        rightShape.Top = trackShape.Top - 4
        rightShape.Width = 30
        rightShape.Height = 24
        rightShape.TextFrame2.TextRange.Font.Size = 9.5
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Normalise les labels de tous les toggles Settings existants.
' EN: Normalizes labels for all existing Settings toggles.
'------------------------------------------------------------------------------
Private Sub Settings_NormalizeAllToggleLabels(ByVal ws As Worksheet)

    Dim keys As Variant
    Dim keyName As Variant
    Dim leftShape As Shape
    Dim trackShape As Shape
    Dim rightShape As Shape

    If ws Is Nothing Then Exit Sub

    keys = Array( _
        "GLOBAL", MODULE_DASHBOARD, MODULE_GANTT, MODULE_SCURVE, _
        MODULE_WBS, MODULE_CONSTRAINTS, MODULE_EVENT)

    For Each keyName In keys
        Set leftShape = Nothing
        Set trackShape = Nothing
        Set rightShape = Nothing
        On Error Resume Next
        Set leftShape = ws.Shapes("SET_LANG_" & CStr(keyName) & "_FR")
        Set trackShape = ws.Shapes("SET_LANG_" & CStr(keyName) & "_TRACK")
        Set rightShape = ws.Shapes("SET_LANG_" & CStr(keyName) & "_EN")
        On Error GoTo 0
        Settings_NormalizeToggleLabels leftShape, trackShape, rightShape
    Next keyName

    Set leftShape = Nothing
    Set trackShape = Nothing
    Set rightShape = Nothing
    On Error Resume Next
    Set leftShape = ws.Shapes("SET_INFO_OFF")
    Set trackShape = ws.Shapes("SET_INFO_TRACK")
    Set rightShape = ws.Shapes("SET_INFO_ON")
    On Error GoTo 0
    Settings_NormalizeToggleLabels leftShape, trackShape, rightShape

End Sub

'------------------------------------------------------------------------------
' FR: Normalise ou formate Settings Format Toggle selon le contrat canonique du composant.
' EN: Normalizes or formats Settings Format Toggle according to the component contract.
'------------------------------------------------------------------------------

Private Sub Settings_FormatToggle( _
    ByVal trackShape As Shape, _
    ByVal knobShape As Shape, _
    ByVal isRightActive As Boolean)

    If trackShape Is Nothing Or knobShape Is Nothing Then Exit Sub
    Settings_NormalizeToggleGeometry trackShape, knobShape

    If isRightActive Then
        trackShape.Fill.ForeColor.RGB = RGB(68, 114, 196)
        trackShape.Line.ForeColor.RGB = RGB(68, 114, 196)
        knobShape.Left = trackShape.Left + trackShape.Width - knobShape.Width - 2
    Else
        trackShape.Fill.ForeColor.RGB = RGB(150, 160, 172)
        trackShape.Line.ForeColor.RGB = RGB(150, 160, 172)
        knobShape.Left = trackShape.Left + 2
    End If

    trackShape.Line.Weight = 0.75
    knobShape.Top = trackShape.Top + ((trackShape.Height - knobShape.Height) / 2)

End Sub
'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Activated Control dans le contexte settings and language.
' EN: Updates Settings Add Activated Control in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddActivatedControl( _
    ByVal ws As Worksheet, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal isActivated As Boolean)

    Dim boxShape As Shape
    Dim labelShape As Shape

    Set boxShape = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos + 2, 18, 18)
    boxShape.Name = "SET_GLOBAL_ACTIVATED_BOX"
    boxShape.Placement = xlFreeFloating
    boxShape.OnAction = "Settings_ToggleGlobalActivated"
    boxShape.TextFrame2.TextRange.Text = IIf(isActivated, ChrW$(&H2713), "")
    boxShape.TextFrame2.TextRange.Font.Name = "Segoe UI Symbol"
    boxShape.TextFrame2.TextRange.Font.Size = 11
    boxShape.TextFrame2.TextRange.Font.Bold = msoTrue
    boxShape.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    boxShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    boxShape.TextFrame2.MarginLeft = 0
    boxShape.TextFrame2.MarginRight = 0
    boxShape.TextFrame2.MarginTop = 0
    boxShape.TextFrame2.MarginBottom = 0

    Set labelShape = Settings_AddTextShape(ws, "SET_GLOBAL_ACTIVATED_LABEL", Settings_ActivatedLabel(), leftPos + 26, topPos, 130, 24, msoAlignLeft, 9.5, True)
    labelShape.OnAction = "Settings_ToggleGlobalActivated"

    Settings_FormatActivatedControl boxShape, isActivated

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Format Activated Control dans le contexte settings and language.
' EN: Updates Settings Format Activated Control in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_FormatActivatedControl( _
    ByVal boxShape As Shape, _
    ByVal isActivated As Boolean)

    If boxShape Is Nothing Then Exit Sub

    If isActivated Then
        boxShape.Fill.ForeColor.RGB = RGB(68, 114, 196)
        boxShape.Line.ForeColor.RGB = RGB(68, 114, 196)
        boxShape.TextFrame2.TextRange.Text = ChrW$(&H2713)
        boxShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        boxShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
        boxShape.Line.ForeColor.RGB = RGB(150, 160, 172)
        boxShape.TextFrame2.TextRange.Text = ""
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Add Command Button dans le contexte settings and language.
' EN: Updates Settings Add Command Button in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_AddCommandButton( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal captionText As String, _
    ByVal macroName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal fillColor As Long, _
    ByVal textColor As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, widthVal, heightVal)
    shp.Name = shapeName
    shp.Placement = xlFreeFloating
    shp.Adjustments.item(1) = 0.16
    shp.OnAction = macroName
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoFalse
    shp.Shadow.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = captionText
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.alignment = msoAlignCenter
    shp.TextFrame2.TextRange.Font.Name = "Segoe UI Semibold"
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textColor

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Visuals dans le contexte settings and language.
' EN: Updates Settings Refresh Visuals in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshVisuals(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    Settings_RefreshSwitch ws, "GLOBAL", Settings_ModuleDisplayLanguage(ws, "GLOBAL")
    Settings_RefreshSwitch ws, MODULE_DASHBOARD, Settings_ModuleLanguage(ws, MODULE_DASHBOARD)
    Settings_RefreshSwitch ws, MODULE_GANTT, Settings_ModuleLanguage(ws, MODULE_GANTT)
    Settings_RefreshSwitch ws, MODULE_SCURVE, Settings_ModuleLanguage(ws, MODULE_SCURVE)
    Settings_RefreshSwitch ws, MODULE_WBS, Settings_ModuleLanguage(ws, MODULE_WBS)
    Settings_RefreshSwitch ws, MODULE_CONSTRAINTS, Settings_ModuleLanguage(ws, MODULE_CONSTRAINTS)
    Settings_RefreshSwitch ws, MODULE_EVENT, Settings_ModuleLanguage(ws, MODULE_EVENT)
    Settings_RefreshInfoSwitch ws, Settings_InfoIsEnabled(ws)
    ThreePositionControl_Refresh ws, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
        Settings_DateDisplayPosition(Settings_GetDateDisplayMode())
    Settings_NormalizeAllToggleLabels ws
    Settings_RefreshTitles ws
    Settings_RefreshCommandCaptions ws

    On Error Resume Next
    Settings_FormatActivatedControl ws.Shapes("SET_GLOBAL_ACTIVATED_BOX"), Settings_GlobalIsActivated(ws)
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Titles dans le contexte settings and language.
' EN: Updates Settings Refresh Titles in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshTitles(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    ws.Range("B1").value = Settings_L(ws, "Options", "Settings")
    Settings_SetShapeText ws, "SET_TITLE_LANGUAGE", Settings_L(ws, "Langue", "Language")
    Settings_SetShapeText ws, "SET_LANG_CONSTRAINTS_LABEL", _
        Settings_L(ws, "Contraintes", "Constraints")
    Settings_SetShapeText ws, "SET_DATE_FORMAT_TITLE", Settings_L(ws, "Format des dates", "Date format")
    Settings_SetShapeText ws, "SET_TITLE_RESET", Settings_L(ws, "R" & ChrW$(&HE9) & "initialisation", "Reset")
    Settings_SetShapeText ws, "SET_TITLE_DANGER", Settings_L(ws, "Zone de danger", "Danger Zone")

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Command Captions dans le contexte settings and language.
' EN: Updates Settings Refresh Command Captions in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshCommandCaptions(ByVal ws As Worksheet)

    If ws Is Nothing Then Exit Sub

    Settings_SetShapeText ws, "SET_BTN_CLEAR_HISTORY", Settings_L(ws, "Nettoyer historique", "Clear History")
    Settings_SetShapeText ws, "SET_BTN_CLEAR_ACK", Settings_L(ws, "Nettoyer les messages acquités", "Clear Acknowledged")
    Settings_SetShapeText ws, "SET_BTN_CLEAN_DASHBOARD", Settings_L(ws, "Nettoyer Dashboard", "Clean Dashboard")
    Settings_SetShapeText ws, "SET_BTN_RESET_PLANNING", Settings_L(ws, "R" & ChrW$(&HE9) & "initialiser planning", "Reset Planning")
    Settings_SetShapeText ws, "SET_BTN_FULL_RESET", Settings_L(ws, "R" & ChrW$(&HE9) & "initialisation compl" & ChrW$(&HE8) & "te", "Full Reset")

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Set Shape Text dans le contexte settings and language.
' EN: Updates Settings Set Shape Text in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_SetShapeText( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal captionText As String)

    On Error Resume Next
    ws.Shapes(shapeName).TextFrame2.TextRange.Text = captionText
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Info Switch dans le contexte settings and language.
' EN: Updates Settings Refresh Info Switch in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshInfoSwitch( _
    ByVal ws As Worksheet, _
    ByVal isEnabled As Boolean)

    Dim trackShape As Shape
    Dim knobShape As Shape

    On Error Resume Next
    Set trackShape = ws.Shapes("SET_INFO_TRACK")
    Set knobShape = ws.Shapes("SET_INFO_KNOB")
    On Error GoTo 0

    Settings_FormatToggle trackShape, knobShape, isEnabled

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Refresh Switch dans le contexte settings and language.
' EN: Updates Settings Refresh Switch in the settings and language context.
'------------------------------------------------------------------------------
Private Sub Settings_RefreshSwitch( _
    ByVal ws As Worksheet, _
    ByVal keyName As String, _
    ByVal languageCode As String)

    Dim trackShape As Shape
    Dim knobShape As Shape

    On Error Resume Next
    Set trackShape = ws.Shapes("SET_LANG_" & keyName & "_TRACK")
    Set knobShape = ws.Shapes("SET_LANG_" & keyName & "_KNOB")
    On Error GoTo 0

    Settings_FormatToggle trackShape, knobShape, _
        (Settings_NormalizeLanguage(languageCode, "EN") = "EN")

End Sub

'------------------------------------------------------------------------------
' FR: Met a jour Settings Reset Title dans le contexte settings and language.
' EN: Updates Settings Reset Title in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ResetTitle() As String
    Settings_ResetTitle = "Reset / R" & ChrW$(&HE9) & "initialisation"
End Function

'------------------------------------------------------------------------------
' FR: Met a jour Settings Activated Label dans le contexte settings and language.
' EN: Updates Settings Activated Label in the settings and language context.
'------------------------------------------------------------------------------
Private Function Settings_ActivatedLabel() As String
    Settings_ActivatedLabel = "Activated / Activ" & ChrW$(&HE9)
End Function

'------------------------------------------------------------------------------
' FR: Applique transactionnellement un nouveau format visible aux surfaces produit.
' EN: Transactionally applies a new visible format to product-facing surfaces.
'------------------------------------------------------------------------------
Public Sub Settings_ApplyDateDisplayMode(ByVal requestedMode As String)

    Dim ws As Worksheet
    Dim oldMode As String
    Dim normalizedRequested As String
    Dim oldStoredValue As Variant
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim stateCaptured As Boolean
    Dim formattingStarted As Boolean
    Dim errorDescription As String

    If Not Settings_IsValidDateDisplayMode(requestedMode) Then
        CalcBridge_ShowSingleConsoleMessage "STOP", _
            "Mode de format de date inconnu : " & requestedMode, _
            "Unknown date display mode: " & requestedMode
        Exit Sub
    End If

    normalizedRequested = Settings_NormalizeDateDisplayMode(requestedMode)
    oldMode = Settings_GetDateDisplayMode()
    If normalizedRequested = oldMode Then Exit Sub

    On Error GoTo ApplyFailed

    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    oldStoredValue = ws.Range(CELL_DATE_DISPLAY_MODE).Value2
    Settings_PreflightDateDisplaySurfaces

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    stateCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    formattingStarted = True
    Settings_ApplyDateDisplayFormatsForMode normalizedRequested
    ws.Range(CELL_DATE_DISPLAY_MODE).Value2 = normalizedRequested
    ThreePositionControl_Refresh ws, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
        Settings_DateDisplayPosition(normalizedRequested)

CleanExit:
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    Exit Sub

ApplyFailed:
    errorDescription = Err.Description
    On Error Resume Next
    If formattingStarted Then Settings_ApplyDateDisplayFormatsForMode oldMode
    If Not ws Is Nothing Then
        ws.Range(CELL_DATE_DISPLAY_MODE).Value2 = oldStoredValue
        ThreePositionControl_Refresh ws, DATE_CONTROL_BG_NAME, DATE_CONTROL_KNOB_NAME, _
            Settings_DateDisplayPosition(oldMode)
    End If
    On Error GoTo 0
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
    End If
    CalcBridge_ShowSingleConsoleMessage "STOP", _
        "Le format des dates n'a pas ete modifie. " & errorDescription, _
        "The date format was not changed. " & errorDescription

End Sub

'------------------------------------------------------------------------------
' FR: Valide toutes les surfaces obligatoires avant la premiere ecriture de format.
' EN: Validates every required surface before the first format write.
'------------------------------------------------------------------------------
Private Sub Settings_PreflightDateDisplaySurfaces()

    Dim tbl As ListObject
    Dim scurveTbl As ListObject
    Dim ch As Chart

    Set tbl = Settings_RequireTable("WBS", "tbl_WBS")
    Settings_RequireColumns tbl, VTS_TABLE_WBS, Array( _
        VTS_COL_BASELINE_START, VTS_COL_BASELINE_FINISH, VTS_COL_ACTUAL_START, VTS_COL_ACTUAL_FINISH, _
        VTS_COL_FORECAST_START, VTS_COL_FORECAST_FINISH, _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH)

    Settings_RequireWorksheet "GANTT"

    Set scurveTbl = Settings_RequireTable("SCURVE", "tbl_SCURVE")
    Settings_RequireColumns scurveTbl, VTS_TABLE_SCURVE, Array(VTS_COL_DATE)
    Set ch = Settings_RequireVisibleDateChart( _
        "SCURVE", "cht_SCurve", scurveTbl, VTS_TABLE_SCURVE, VTS_COL_DATE)

    Set tbl = Settings_RequireTable("CONSTRAINTS", "tbl_CONSTRAINTS")
    Settings_RequireColumns tbl, VTS_TABLE_CONSTRAINTS, Array( _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH, VTS_COL_START_CONSTRAINT_DATE, _
        VTS_COL_FINISH_CONSTRAINT_DATE, VTS_COL_DEADLINE)

    Set tbl = Settings_RequireTable("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    Settings_RequireColumns tbl, VTS_TABLE_EVENT_HISTORY, Array(VTS_COL_DATE)
    Set tbl = Settings_RequireTable("EVENT_ACK", "tbl_EVENT_ACK")
    Settings_RequireColumns tbl, VTS_TABLE_EVENT_ACK, Array(VTS_COL_ACKNOWLEDGED_AT)

    Set ch = Settings_RequireVisibleDateChart( _
        "DASHBOARD", "cht_Dashboard_SCurve", scurveTbl, VTS_TABLE_SCURVE, VTS_COL_DATE)

End Sub

'------------------------------------------------------------------------------
' FR: Applique uniquement NumberFormat aux surfaces visibles contractuelles.
' EN: Applies only NumberFormat to the contracted visible surfaces.
'------------------------------------------------------------------------------
Private Sub Settings_ApplyDateDisplayFormatsForMode(ByVal modeValue As String)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim scurveTbl As ListObject
    Dim ch As Chart
    Dim dateFormat As String
    Dim dateTimeFormat As String
    Dim compactFormat As String
    Dim lastRow As Long

    dateFormat = Settings_DateNumberFormatForMode(modeValue)
    dateTimeFormat = Settings_DateTimeNumberFormatForMode(modeValue)
    compactFormat = Settings_CompactDateNumberFormatForMode(modeValue)

    Set tbl = Settings_RequireTable("WBS", "tbl_WBS")
    Settings_ApplyTableColumnFormats tbl, VTS_TABLE_WBS, Array( _
        VTS_COL_BASELINE_START, VTS_COL_BASELINE_FINISH, VTS_COL_ACTUAL_START, VTS_COL_ACTUAL_FINISH, _
        VTS_COL_FORECAST_START, VTS_COL_FORECAST_FINISH, _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH), _
        dateFormat

    Set ws = Settings_RequireWorksheet("GANTT")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then lastRow = 5
    ws.Range(ws.Cells(5, 3), ws.Cells(lastRow, 6)).NumberFormat = dateFormat

    Set scurveTbl = Settings_RequireTable("SCURVE", "tbl_SCURVE")
    Settings_ApplyTableColumnFormats scurveTbl, VTS_TABLE_SCURVE, Array(VTS_COL_DATE), dateFormat
    Set ch = Settings_RequireVisibleDateChart( _
        "SCURVE", "cht_SCurve", scurveTbl, VTS_TABLE_SCURVE, VTS_COL_DATE)
    If Not ch Is Nothing Then ch.Axes(xlCategory).TickLabels.NumberFormat = dateFormat

    Set tbl = Settings_RequireTable("CONSTRAINTS", "tbl_CONSTRAINTS")
    Settings_ApplyTableColumnFormats tbl, VTS_TABLE_CONSTRAINTS, Array( _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH, VTS_COL_START_CONSTRAINT_DATE, _
        VTS_COL_FINISH_CONSTRAINT_DATE, VTS_COL_DEADLINE), dateFormat

    Set tbl = Settings_RequireTable("EVENT_HISTORY", "tbl_EVENT_HISTORY")
    Settings_ApplyTableColumnFormats tbl, VTS_TABLE_EVENT_HISTORY, Array(VTS_COL_DATE), dateFormat
    Set tbl = Settings_RequireTable("EVENT_ACK", "tbl_EVENT_ACK")
    Settings_ApplyTableColumnFormats tbl, VTS_TABLE_EVENT_ACK, Array(VTS_COL_ACKNOWLEDGED_AT), dateTimeFormat

    Set ch = Settings_RequireVisibleDateChart( _
        "DASHBOARD", "cht_Dashboard_SCurve", scurveTbl, VTS_TABLE_SCURVE, VTS_COL_DATE)
    If Not ch Is Nothing Then ch.Axes(xlCategory).TickLabels.NumberFormat = compactFormat

End Sub

Private Sub Settings_ApplyTableColumnFormats( _
    ByVal tbl As ListObject, _
    ByVal tableKey As String, _
    ByVal columnKeys As Variant, _
    ByVal numberFormat As String)

    Dim columnKey As Variant

    For Each columnKey In columnKeys
        SchemaListColumn(tbl, tableKey, CStr(columnKey)).Range.NumberFormat = numberFormat
    Next columnKey

End Sub

Private Function Settings_RequireWorksheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set Settings_RequireWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Settings_RequireWorksheet Is Nothing Then
        Err.Raise vbObjectError + 5281, "Settings_RequireWorksheet", _
            "Required worksheet is missing: " & sheetName
    End If

End Function

Private Function Settings_RequireTable( _
    ByVal sheetName As String, _
    ByVal tableName As String) As ListObject

    Dim ws As Worksheet

    Set ws = Settings_RequireWorksheet(sheetName)
    On Error Resume Next
    Set Settings_RequireTable = ws.ListObjects(tableName)
    On Error GoTo 0
    If Settings_RequireTable Is Nothing Then
        Err.Raise vbObjectError + 5282, "Settings_RequireTable", _
            "Required table is missing: " & sheetName & "!" & tableName
    End If

End Function

Private Function Settings_RequireVisibleDateChart( _
    ByVal sheetName As String, _
    ByVal chartName As String, _
    ByVal dateTable As ListObject, _
    ByVal tableKey As String, _
    ByVal dateColumnKey As String) As Chart

    Dim ws As Worksheet
    Dim dateRange As Range

    Set ws = Settings_RequireWorksheet(sheetName)
    On Error Resume Next
    Set Settings_RequireVisibleDateChart = ws.ChartObjects(chartName).Chart
    On Error GoTo 0

    If Settings_RequireVisibleDateChart Is Nothing Then
        Set dateRange = SchemaListColumn(dateTable, tableKey, dateColumnKey).DataBodyRange
        If Not dateRange Is Nothing Then
            If Application.Count(dateRange) > 0 Then
                Err.Raise vbObjectError + 5283, "Settings_RequireVisibleDateChart", _
                    "Required chart is missing: " & sheetName & "!" & chartName
            End If
        End If
        Exit Function
    End If

    If Not Settings_RequireVisibleDateChart.HasAxis(xlCategory) Then
        Err.Raise vbObjectError + 5284, "Settings_RequireVisibleDateChart", _
            "Required date axis is missing: " & sheetName & "!" & chartName
    End If

End Function

Private Sub Settings_RequireColumns( _
    ByVal tbl As ListObject, _
    ByVal tableKey As String, _
    ByVal columnKeys As Variant)

    Dim columnKey As Variant
    Dim listColumn As ListColumn

    For Each columnKey In columnKeys
        Set listColumn = Nothing
        On Error Resume Next
        Set listColumn = SchemaListColumn(tbl, tableKey, CStr(columnKey))
        On Error GoTo 0
        If listColumn Is Nothing Then
            Err.Raise vbObjectError + 5285, "Settings_RequireColumns", _
                "Required column is missing: " & tbl.Name & "[" & CStr(columnKey) & "]"
        End If
    Next columnKey

End Sub

Private Function Settings_IsValidDateDisplayMode(ByVal modeValue As String) As Boolean

    Select Case UCase$(Trim$(modeValue))
        Case DATE_MODE_DMY, DATE_MODE_MDY, DATE_MODE_ISO
            Settings_IsValidDateDisplayMode = True
    End Select

End Function

Private Function Settings_NormalizeDateDisplayMode(ByVal modeValue As String) As String

    Select Case UCase$(Trim$(modeValue))
        Case DATE_MODE_MDY
            Settings_NormalizeDateDisplayMode = DATE_MODE_MDY
        Case DATE_MODE_ISO
            Settings_NormalizeDateDisplayMode = DATE_MODE_ISO
        Case Else
            Settings_NormalizeDateDisplayMode = DATE_MODE_DMY
    End Select

End Function

Private Function Settings_DateNumberFormatForMode(ByVal modeValue As String) As String

    Select Case Settings_NormalizeDateDisplayMode(modeValue)
        Case DATE_MODE_MDY
            Settings_DateNumberFormatForMode = DATE_FORMAT_MDY
        Case DATE_MODE_ISO
            Settings_DateNumberFormatForMode = DATE_FORMAT_ISO
        Case Else
            Settings_DateNumberFormatForMode = DATE_FORMAT_DMY
    End Select

End Function

Private Function Settings_DateTimeNumberFormatForMode(ByVal modeValue As String) As String

    Select Case Settings_NormalizeDateDisplayMode(modeValue)
        Case DATE_MODE_MDY
            Settings_DateTimeNumberFormatForMode = DATETIME_FORMAT_MDY
        Case DATE_MODE_ISO
            Settings_DateTimeNumberFormatForMode = DATETIME_FORMAT_ISO
        Case Else
            Settings_DateTimeNumberFormatForMode = DATETIME_FORMAT_DMY
    End Select

End Function

Private Function Settings_CompactDateNumberFormatForMode(ByVal modeValue As String) As String

    Select Case Settings_NormalizeDateDisplayMode(modeValue)
        Case DATE_MODE_MDY
            Settings_CompactDateNumberFormatForMode = DATE_COMPACT_FORMAT_MDY
        Case DATE_MODE_ISO
            Settings_CompactDateNumberFormatForMode = DATE_COMPACT_FORMAT_ISO
        Case Else
            Settings_CompactDateNumberFormatForMode = DATE_COMPACT_FORMAT_DMY
    End Select

End Function

Private Function Settings_DateDisplayPosition(ByVal modeValue As String) As Long

    Select Case Settings_NormalizeDateDisplayMode(modeValue)
        Case DATE_MODE_MDY
            Settings_DateDisplayPosition = 1
        Case DATE_MODE_ISO
            Settings_DateDisplayPosition = 2
        Case Else
            Settings_DateDisplayPosition = 0
    End Select

End Function
