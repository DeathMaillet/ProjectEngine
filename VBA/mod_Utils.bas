Attribute VB_Name = "mod_Utils"
Option Explicit

'===============================================================================
' MODULE : mod_Utils
' DOMAINE / DOMAIN : Shared Utilities
'
' FR
' Fournit les primitives transversales de dates, WBS, valeurs et messages bilingues.
' Ne possede aucun workflow metier ni table Excel.
'
' EN
' Provides shared date, WBS, value and bilingual-message primitives.
' Owns no business workflow or Excel table.
'
' CONTRATS / CONTRACTS : HasValue, GetCellValue, maxDate, GetParentWBS, BiMsg, NormalizePercentInput, ValuesDiffer, NormalizeWBS
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' Return True only when value is usable (not Error / Empty / blank).
'------------------------------------------------------------------------------
' FR: Indique si Value est vrai pour le contexte courant.
' EN: Returns whether Value is true for the current context.
'------------------------------------------------------------------------------
Public Function HasValue(ByVal v As Variant) As Boolean

    If isError(v) Then
        HasValue = False
    ElseIf IsEmpty(v) Then
        HasValue = False
    ElseIf Trim(CStr(v)) = "" Then
        HasValue = False
    Else
        HasValue = True
    End If

End Function

' Return value if usable, otherwise return Empty.
'------------------------------------------------------------------------------
' FR: Retourne Cell Value depuis le contexte utilities.
' EN: Returns Cell Value from the utilities context.
'------------------------------------------------------------------------------
Public Function GetCellValue(ByVal v As Variant) As Variant

    If HasValue(v) Then
        GetCellValue = v
    Else
        GetCellValue = Empty
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la valeur max Date sans modifier les donnees d'entree.
' EN: Returns the max Date value without mutating input data.
'------------------------------------------------------------------------------

Public Function maxDate(ByVal d1 As Variant, ByVal d2 As Variant) As Double

    If CDbl(d1) >= CDbl(d2) Then
        maxDate = CDbl(d1)
    Else
        maxDate = CDbl(d2)
    End If

End Function


' Return parent WBS code from a given WBS string.
'------------------------------------------------------------------------------
' FR: Retourne Parent WBS depuis le contexte utilities.
' EN: Returns Parent WBS from the utilities context.
'------------------------------------------------------------------------------
Public Function GetParentWBS(ByVal wbs As String) As String

    Dim lastDotPos As Long

    wbs = Replace(Trim$(wbs), ",", ".")
    lastDotPos = InStrRev(wbs, ".")

    If lastDotPos > 0 Then
        GetParentWBS = Left$(wbs, lastDotPos - 1)
    Else
        GetParentWBS = ""
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la valeur Bi Msg sans modifier les donnees d'entree.
' EN: Returns the Bi Msg value without mutating input data.
'------------------------------------------------------------------------------

Public Function BiMsg(ByVal frText As String, ByVal enText As String) As String

    BiMsg = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

End Function

' Normalize percent input (accept 0-1 or 0-100 scale).
'------------------------------------------------------------------------------
' FR: Normalise Percent Input dans un format exploitable.
' EN: Normalizes Percent Input into a usable format.
'------------------------------------------------------------------------------
Public Function NormalizePercentInput(ByVal v As Variant) As Variant

    Dim x As Double

    If Not HasValue(v) Then
        NormalizePercentInput = Empty
        Exit Function
    End If

    If Not IsNumeric(v) Then
        NormalizePercentInput = Empty
        Exit Function
    End If

    x = CDbl(v)

    If x < 0 Then
        NormalizePercentInput = Empty
    ElseIf x <= 1 Then
        NormalizePercentInput = x
    Else
        NormalizePercentInput = x / 100#
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Values Differ sans modifier les donnees d'entree.
' EN: Returns the Values Differ value without mutating input data.
'------------------------------------------------------------------------------

Public Function ValuesDiffer(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean

    If Not HasValue(v1) And Not HasValue(v2) Then Exit Function

    If HasValue(v1) Xor HasValue(v2) Then
        ValuesDiffer = True
        Exit Function
    End If

    If IsNumeric(v1) And IsNumeric(v2) Then
        ValuesDiffer = (Abs(CDbl(v1) - CDbl(v2)) > 0.0000001)
    Else
        ValuesDiffer = (CStr(v1) <> CStr(v2))
    End If

End Function

' Normalize WBS string format (trim + replace comma with dot).
'------------------------------------------------------------------------------
' FR: Normalise WBS dans un format exploitable.
' EN: Normalizes WBS into a usable format.
'------------------------------------------------------------------------------
Public Function NormalizeWBS(ByVal rawValue As Variant) As String

    Dim s As String

    If isError(rawValue) Then
        NormalizeWBS = ""
        Exit Function
    End If

    If IsEmpty(rawValue) Then
        NormalizeWBS = ""
        Exit Function
    End If

    s = Trim$(CStr(rawValue))

    If s = "" Then
        NormalizeWBS = ""
        Exit Function
    End If

    ' Canonical separator for WBS = dot
    s = Replace$(s, ",", ".")

    ' Remove accidental spaces
    s = Replace$(s, " ", "")

    ' Defensive cleanup for repeated dots
    Do While InStr(1, s, "..", vbBinaryCompare) > 0
        s = Replace$(s, "..", ".")
    Loop

    ' Remove leading/trailing dot if ever present
    If Left$(s, 1) = "." Then s = Mid$(s, 2)
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)

    NormalizeWBS = s

End Function

'=================================================
' Parse a single predecessor token.
'
' INPUT
' - tokenText examples:
'   1
'   1+3
'   1FS+3
'   1SS-2
'   1FF
'
' OUTPUT
' - predWBS  : predecessor WBS
' - linkType : FS / SS / FF
' - lagVal   : integer lag
' - rawToken : normalized raw token
'
' RETURN
' - True if token parsed successfully
' - False if token is invalid
'=================================================
'------------------------------------------------------------------------------
' FR: Analyse Predecessor Token et extrait les informations utiles.
' EN: Parses Predecessor Token and extracts useful information.
'------------------------------------------------------------------------------
Public Function ParsePredecessorToken( _
    ByVal token As String, _
    ByRef predWbs As String, _
    ByRef linkType As String, _
    ByRef lagVal As Long, _
    ByRef rawToken As String, _
    ByRef errText As String) As Boolean

    Dim t As String
    Dim posType As Long
    Dim posPlus As Long
    Dim posMinus As Long
    Dim posLag As Long
    Dim suffix As String
    Dim lagText As String

    ParsePredecessorToken = False
    predWbs = ""
    linkType = "FS"
    lagVal = 0
    rawToken = ""
    errText = ""

    t = Trim$(token)

    If ContainsForbiddenWhitespace(t) Then
        errText = "Spaces are not allowed in predecessor tokens."
        Exit Function
    End If

    t = Replace$(t, ",", ".")

    If t = "" Then
        errText = "Token vide."
        Exit Function
    End If

    '----------------------------------------
    ' 1) Detect explicit link type if present
    '----------------------------------------
    posType = 0

    If InStr(1, t, "SS", vbTextCompare) > 0 Then
        linkType = "SS"
        posType = InStr(1, t, "SS", vbTextCompare)
    ElseIf InStr(1, t, "FF", vbTextCompare) > 0 Then
        linkType = "FF"
        posType = InStr(1, t, "FF", vbTextCompare)
    ElseIf InStr(1, t, "FS", vbTextCompare) > 0 Then
        linkType = "FS"
        posType = InStr(1, t, "FS", vbTextCompare)
    End If

    '----------------------------------------
    ' 2) Split WBS / lag
    '    Accepted forms:
    '    - 1.2.3
    '    - 1.2.3+4
    '    - 1.2.3-2
    '    - 1.2.3FS
    '    - 1.2.3FS+4
    '    - 1.2.3SS-2
    '----------------------------------------
    If posType > 0 Then

        predWbs = Left$(t, posType - 1)
        suffix = Mid$(t, posType + 2)

    Else

        posPlus = InStrRev(t, "+")
        posMinus = InStrRev(t, "-")

        posLag = 0
        If posPlus > 0 And posMinus > 0 Then
            If posPlus > posMinus Then
                posLag = posPlus
            Else
                posLag = posMinus
            End If
        ElseIf posPlus > 0 Then
            posLag = posPlus
        ElseIf posMinus > 0 Then
            posLag = posMinus
        End If

        If posLag > 0 Then
            predWbs = Left$(t, posLag - 1)
            suffix = Mid$(t, posLag)
        Else
            predWbs = t
            suffix = ""
        End If

    End If

    predWbs = Trim$(predWbs)

    If predWbs = "" Then
        errText = "Prédécesseur WBS vide."
        Exit Function
    End If

    If Not IsValidPureWBS(predWbs) Then
        errText = "WBS prédécesseur invalide : " & predWbs
        Exit Function
    End If

    '----------------------------------------
    ' 3) Parse lag if present
    '----------------------------------------
    If suffix <> "" Then
        If Left$(suffix, 1) <> "+" And Left$(suffix, 1) <> "-" Then
            errText = "Suffix invalide : " & suffix
            Exit Function
        End If

        lagText = suffix

        If Not IsNumeric(lagText) Then
            errText = "Lag invalide : " & lagText
            Exit Function
        End If

        lagVal = CLng(lagText)
    End If

    rawToken = predWbs & linkType
    If lagVal > 0 Then
        rawToken = rawToken & "+" & CStr(lagVal)
    ElseIf lagVal < 0 Then
        rawToken = rawToken & CStr(lagVal)
    End If

    ParsePredecessorToken = True

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Contains Forbidden Whitespace sans modifier les donnees d'entree.
' EN: Returns the Contains Forbidden Whitespace value without mutating input data.
'------------------------------------------------------------------------------

Private Function ContainsForbiddenWhitespace(ByVal textValue As String) As Boolean

    ContainsForbiddenWhitespace = _
        (InStr(1, textValue, " ", vbBinaryCompare) > 0) Or _
        (InStr(1, textValue, vbTab, vbBinaryCompare) > 0) Or _
        (InStr(1, textValue, Chr$(160), vbBinaryCompare) > 0)

End Function
'------------------------------------------------------------------------------
' FR: Indique si Valid Pure WBS est vrai pour le contexte courant.
' EN: Returns whether Valid Pure WBS is true for the current context.
'------------------------------------------------------------------------------
Private Function IsValidPureWBS(ByVal wbsText As String) As Boolean

    Dim reWBS As Object

    Set reWBS = CreateObject("VBScript.RegExp")
    reWBS.Pattern = "^\d+(\.\d+)*$"

    IsValidPureWBS = reWBS.Test(Trim$(wbsText))

End Function

'=================================================
' Parse a full Predecessors WBS cell into a collection
' of dictionaries ready for later use.
'
' EACH ITEM CONTAINS
' - Succ ID
' - Succ WBS
' - Pred ID
' - Pred WBS
' - Link Type
' - Lag
' - Raw Token
'
' RETURN
' - True if all tokens were parsed successfully
' - False if at least one token failed
'
' NOTE
' - empty cell = valid, zero link
'=================================================
'------------------------------------------------------------------------------
' FR: Analyse Predecessors Text et extrait les informations utiles.
' EN: Parses Predecessors Text and extracts useful information.
'------------------------------------------------------------------------------
Public Function ParsePredecessorsText( _
    ByVal succId As String, _
    ByVal succWBS As String, _
    ByVal predText As String, _
    ByVal wbsToId As Object, _
    ByRef linksOut As Collection, _
    ByRef errText As String) As Boolean

    Dim perfScope As clsPerfScope

    Dim parts As Variant
    Dim i As Long
    Dim tokenText As String

    Dim predWbs As String
    Dim predId As String
    Dim linkType As String
    Dim lagVal As Long
    Dim rawToken As String

    Dim linkRow As Object

    Set perfScope = Profiler_BeginScope("ParsePredecessorsText", "Parsing")

    Set linksOut = New Collection
    errText = ""

    predText = Replace$(Trim$(predText), ",", ".")

    If predText = "" Then
        ParsePredecessorsText = True
        Exit Function
    End If

    parts = Split(predText, ";")

    For i = LBound(parts) To UBound(parts)

        tokenText = Trim$(CStr(parts(i)))

        If tokenText = "" Then
            errText = "Empty predecessor token in: " & predText
            Exit Function
        End If

        If Not ParsePredecessorToken(tokenText, predWbs, linkType, lagVal, rawToken, errText) Then
            errText = "Successor WBS " & succWBS & " -> " & errText
            Exit Function
        End If

        If wbsToId.Exists(predWbs) Then
            predId = CStr(wbsToId(predWbs))
        Else
            predId = ""
        End If

        Set linkRow = CreateObject("Scripting.Dictionary")
        linkRow("Succ ID") = Trim$(succId)
        linkRow("Succ WBS") = NormalizeWBS(succWBS)
        linkRow("Pred ID") = predId
        linkRow("Pred WBS") = predWbs
        linkRow("Link Type") = linkType
        linkRow("Lag") = lagVal
        linkRow("Raw Token") = rawToken
        linkRow("Entered Token") = tokenText

        linksOut.Add linkRow

    Next i

    ParsePredecessorsText = True

End Function
