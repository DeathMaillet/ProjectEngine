Attribute VB_Name = "mod_VisibleTableSchemaHarness"
Option Explicit

' Test-only proof harness. It is invoked explicitly on isolated workbook copies.

Public Function VisibleTableSchemaHarness_Run() As String
    Dim tables As Variant
    Dim tableKey As Variant
    Dim totalColumns As Long

    On Error GoTo Fail

    tables = Array( _
        VTS_TABLE_WBS, VTS_TABLE_SCURVE, VTS_TABLE_CONSTRAINTS, _
        VTS_TABLE_EVENT_HISTORY, VTS_TABLE_EVENT_ACK)

    For Each tableKey In tables
        totalColumns = totalColumns + VisibleTableSchemaHarness_CheckTable(CStr(tableKey))
    Next tableKey

    VisibleTableSchemaHarness_Assert totalColumns = 81, _
        "Expected 81 visible-table columns but checked " & CStr(totalColumns) & "."
    VisibleTableSchemaHarness_CheckErrors

    VisibleTableSchemaHarness_Run = "PASS|tables=5|columns=" & CStr(totalColumns)
    Exit Function

Fail:
    VisibleTableSchemaHarness_Run = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function VisibleTableSchemaHarness_CaptureListColumnError( _
    ByVal tableKey As String, _
    ByVal columnKey As String) As String

    Dim tableObject As ListObject
    Dim resolvedColumn As ListColumn

    On Error GoTo Captured
    Set tableObject = VisibleTableSchemaHarness_FindTable(tableKey)
    Set resolvedColumn = SchemaListColumn(tableObject, tableKey, columnKey)
    VisibleTableSchemaHarness_CaptureListColumnError = "NO_ERROR|" & resolvedColumn.Name
    Exit Function

Captured:
    VisibleTableSchemaHarness_CaptureListColumnError = _
        "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function VisibleTableSchemaHarness_RunHeaderEngineRoundtrip() As String
    Dim tables As Variant
    Dim tableKey As Variant
    Dim snapshot As Object

    On Error GoTo Fail

    tables = Array( _
        VTS_TABLE_WBS, VTS_TABLE_SCURVE, VTS_TABLE_CONSTRAINTS, _
        VTS_TABLE_EVENT_HISTORY, VTS_TABLE_EVENT_ACK)

    For Each tableKey In tables
        VisibleHeaders_PreflightOwner SchemaOwnerForTable(CStr(tableKey)), VTS_LANG_FR
    Next tableKey

    Set snapshot = VisibleHeaders_CaptureAllSnapshot()
    VisibleHeaders_ApplyAllFromSnapshot snapshot, VTS_LANG_FR
    For Each tableKey In tables
        VisibleHeaders_ValidateTable CStr(tableKey), VTS_LANG_FR
    Next tableKey

    Set snapshot = VisibleHeaders_CaptureAllSnapshot()
    VisibleHeaders_ApplyAllFromSnapshot snapshot, VTS_LANG_EN
    For Each tableKey In tables
        VisibleHeaders_ValidateTable CStr(tableKey), VTS_LANG_EN
    Next tableKey

    Schema_ClearPhysicalHeaderLanguageOverrides
    VisibleTableSchemaHarness_RunHeaderEngineRoundtrip = "PASS|tables=5|columns=81"
    Exit Function

Fail:
    Schema_ClearPhysicalHeaderLanguageOverrides
    VisibleTableSchemaHarness_RunHeaderEngineRoundtrip = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Public Function VisibleTableSchemaHarness_CaptureOwnerPreflight( _
    ByVal ownerKey As String, _
    ByVal targetLanguage As String) As String

    On Error GoTo Captured
    VisibleHeaders_PreflightOwner ownerKey, targetLanguage
    VisibleTableSchemaHarness_CaptureOwnerPreflight = "NO_ERROR"
    Exit Function

Captured:
    VisibleTableSchemaHarness_CaptureOwnerPreflight = _
        "ERROR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Public Function VisibleTableSchemaHarness_CaptureAllPreflight( _
    ByVal targetLanguage As String) As String

    On Error GoTo Captured
    VisibleHeaders_PreflightAll targetLanguage
    VisibleTableSchemaHarness_CaptureAllPreflight = "NO_ERROR"
    Exit Function

Captured:
    VisibleTableSchemaHarness_CaptureAllPreflight = _
        "ERROR|" & CStr(Err.Number) & "|" & Err.Source & "|" & Err.Description
End Function

Private Function VisibleTableSchemaHarness_CheckTable(ByVal tableKey As String) As Long
    Dim tableObject As ListObject
    Dim columnKeys As Variant
    Dim columnKey As Variant
    Dim englishTitles As Object
    Dim frenchTitles As Object
    Dim physicalTitle As String
    Dim englishTitle As String
    Dim frenchTitle As String
    Dim physicalLanguage As String
    Dim resolvedColumn As ListColumn

    Set tableObject = VisibleTableSchemaHarness_FindTable(tableKey)
    columnKeys = SchemaColumnKeys(tableKey)
    physicalLanguage = SchemaPhysicalHeaderLanguageForTable(tableKey)
    Set englishTitles = CreateObject("Scripting.Dictionary")
    Set frenchTitles = CreateObject("Scripting.Dictionary")
    englishTitles.CompareMode = vbBinaryCompare
    frenchTitles.CompareMode = vbBinaryCompare

    For Each columnKey In columnKeys
        englishTitle = SchemaColumnTitle(tableKey, CStr(columnKey), VTS_LANG_EN)
        frenchTitle = SchemaColumnTitle(tableKey, CStr(columnKey), VTS_LANG_FR)
        physicalTitle = SchemaCurrentColumnTitle(tableKey, CStr(columnKey))

        VisibleTableSchemaHarness_Assert Len(englishTitle) > 0, _
            "Empty English title for " & tableKey & "/" & CStr(columnKey) & "."
        VisibleTableSchemaHarness_Assert Len(frenchTitle) > 0, _
            "Empty French title for " & tableKey & "/" & CStr(columnKey) & "."
        VisibleTableSchemaHarness_Assert Not englishTitles.Exists(englishTitle), _
            "Duplicate English title '" & englishTitle & "' in " & tableKey & "."
        VisibleTableSchemaHarness_Assert Not frenchTitles.Exists(frenchTitle), _
            "Duplicate French title '" & frenchTitle & "' in " & tableKey & "."
        englishTitles.Add englishTitle, CStr(columnKey)
        frenchTitles.Add frenchTitle, CStr(columnKey)

        VisibleTableSchemaHarness_Assert _
            StrComp(physicalTitle, SchemaColumnTitle(tableKey, CStr(columnKey), physicalLanguage), vbBinaryCompare) = 0, _
            "Physical header language is not " & physicalLanguage & _
            " for " & tableKey & "/" & CStr(columnKey) & "."
        Set resolvedColumn = SchemaListColumn(tableObject, tableKey, CStr(columnKey))
        VisibleTableSchemaHarness_Assert _
            StrComp(resolvedColumn.Name, physicalTitle, vbBinaryCompare) = 0, _
            "Resolved title mismatch for " & tableKey & "/" & CStr(columnKey) & "."
        VisibleTableSchemaHarness_Assert _
            SchemaColumnIndex(tableObject, tableKey, CStr(columnKey)) = resolvedColumn.Index, _
            "Resolved index mismatch for " & tableKey & "/" & CStr(columnKey) & "."
        Set resolvedColumn = Nothing
    Next columnKey

    VisibleTableSchemaHarness_Assert _
        SchemaBuildColumnKeyMap(tableObject, tableKey).Count = UBound(columnKeys) - LBound(columnKeys) + 1, _
        "Column-key map count mismatch for " & tableKey & "."

    VisibleTableSchemaHarness_CheckTable = UBound(columnKeys) - LBound(columnKeys) + 1
End Function

Private Function VisibleTableSchemaHarness_FindTable(ByVal tableKey As String) As ListObject
    Dim worksheetObject As Worksheet
    Dim tableObject As ListObject

    For Each worksheetObject In ThisWorkbook.Worksheets
        For Each tableObject In worksheetObject.ListObjects
            If StrComp(tableObject.Name, tableKey, vbBinaryCompare) = 0 Then
                Set VisibleTableSchemaHarness_FindTable = tableObject
                Exit Function
            End If
        Next tableObject
    Next worksheetObject

    Err.Raise vbObjectError + 5390, "VisibleTableSchemaHarness_FindTable", _
        "Visible table '" & tableKey & "' was not found in the workbook."
End Function

Private Sub VisibleTableSchemaHarness_CheckErrors()
    VisibleTableSchemaHarness_ExpectError "UNKNOWN_TABLE", "ID", VTS_LANG_EN, _
        "UNKNOWN_TABLE", "ID", VTS_LANG_EN
    VisibleTableSchemaHarness_ExpectError VTS_TABLE_WBS, "UNKNOWN_COLUMN", VTS_LANG_EN, _
        VTS_TABLE_WBS, "UNKNOWN_COLUMN", VTS_LANG_EN
    VisibleTableSchemaHarness_ExpectError VTS_TABLE_WBS, VTS_COL_ID, "DE", _
        VTS_TABLE_WBS, VTS_COL_ID, "DE"
End Sub

Private Sub VisibleTableSchemaHarness_ExpectError( _
    ByVal tableKey As String, _
    ByVal columnKey As String, _
    ByVal languageCode As String, _
    ByVal expectedTable As String, _
    ByVal expectedColumn As String, _
    ByVal expectedLanguage As String)

    Dim ignored As String
    Dim capturedDescription As String

    On Error Resume Next
    ignored = SchemaColumnTitle(tableKey, columnKey, languageCode)
    capturedDescription = Err.Description
    Err.Clear
    On Error GoTo 0

    VisibleTableSchemaHarness_Assert Len(capturedDescription) > 0, _
        "Expected schema error was not raised."
    VisibleTableSchemaHarness_Assert InStr(1, capturedDescription, expectedTable, vbBinaryCompare) > 0, _
        "Schema error does not identify table '" & expectedTable & "'."
    VisibleTableSchemaHarness_Assert InStr(1, capturedDescription, expectedColumn, vbBinaryCompare) > 0, _
        "Schema error does not identify column '" & expectedColumn & "'."
    VisibleTableSchemaHarness_Assert InStr(1, capturedDescription, expectedLanguage, vbBinaryCompare) > 0, _
        "Schema error does not identify language '" & expectedLanguage & "'."
End Sub

Private Sub VisibleTableSchemaHarness_Assert( _
    ByVal condition As Boolean, _
    ByVal message As String)

    If Not condition Then
        Err.Raise vbObjectError + 5391, "VisibleTableSchemaHarness_Assert", message
    End If
End Sub
