Attribute VB_Name = "mod_ConstrDiagService"
Option Explicit

'===============================================================================
' MODULE : mod_ConstrDiagService
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Possede le workflow specialise indique par son nom et expose ses contrats stables.
' Ne possede pas les domaines appeles en dependance.
'
' EN
' Owns the named specialized workflow and exposes its stable contracts.
' Does not own the domains it calls as dependencies.
'
' CONTRATS / CONTRACTS : BuildConstraintValidationMessage, AddConstraintWarning
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Construit le message bilingue d'un diagnostic Constraints pour une ligne source.
' EN: Builds the bilingual Constraints diagnostic message for a source row.
'------------------------------------------------------------------------------
Public Function BuildConstraintValidationMessage( _
    ByRef arrConstraints As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapConstraints As Object, _
    ByVal frPrefix As String, _
    ByVal enPrefix As String, _
    Optional ByVal frExplanation As String = "", _
    Optional ByVal enExplanation As String = "") As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskName As String
    Dim frText As String
    Dim enText As String

    idVal = Trim$(CStr(arrConstraints(rowIdx, mapConstraints("ID"))))
    wbsVal = Trim$(CStr(arrConstraints(rowIdx, mapConstraints("WBS"))))
    taskName = Trim$(CStr(arrConstraints(rowIdx, mapConstraints("Task Name"))))

    frText = "FR:" & vbCrLf & frPrefix
    If Trim$(frExplanation) <> "" Then
        frText = frText & vbCrLf & vbCrLf & "-> " & frExplanation
    End If
    frText = frText & vbCrLf & vbCrLf & _
        "ID : " & idVal & vbCrLf & _
        "WBS : " & wbsVal & vbCrLf & _
        "Task : " & taskName

    enText = "EN:" & vbCrLf & enPrefix
    If Trim$(enExplanation) <> "" Then
        enText = enText & vbCrLf & vbCrLf & "-> " & enExplanation
    End If
    enText = enText & vbCrLf & vbCrLf & _
        "ID: " & idVal & vbCrLf & _
        "WBS: " & wbsVal & vbCrLf & _
        "Task: " & taskName

    BuildConstraintValidationMessage = frText & vbCrLf & vbCrLf & enText

End Function

'------------------------------------------------------------------------------
' FR: Ajoute un warning Constraints au flux console sans relire les regles metier.
' EN: Adds a Constraints warning to the console stream without re-reading business rules.
'------------------------------------------------------------------------------
Public Sub AddConstraintWarning( _
    ByVal consoleMessages As Collection, _
    ByVal messageText As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "")

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, "WARNING", messageText, historyHandled, eventType, eventHash

End Sub

