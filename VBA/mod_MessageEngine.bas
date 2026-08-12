Attribute VB_Name = "mod_MessageEngine"
Option Explicit

'===============================================================================
' MODULE : mod_MessageEngine
' DOMAINE / DOMAIN : Message Engine
'
' FR
' Prepare, filtre et regroupe les diagnostics destines a la console et a l'historique.
' Ne produit pas les diagnostics metier et ne possede pas leur UI.
'
' EN
' Prepares, filters and groups diagnostics for console and event history.
' Does not produce business diagnostics or own their UI.
'
' CONTRATS / CONTRACTS : MessageEngine_PrepareConsoleMessages, MessageEngine_PrepareDisplayMessages, MessageEngine_ShouldShowConsole, MessageEngine_BuildCategoryProgressCaption
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Retourne la collection Prepare Console Messages sans modifier les donnees d'entree.
' EN: Returns the Prepare Console Messages collection without mutating input data.
'------------------------------------------------------------------------------

Public Function MessageEngine_PrepareConsoleMessages(ByVal messages As Collection) As Collection

    Dim result As Collection
    Dim item As Variant

    Set result = New Collection

    If messages Is Nothing Then
        Set MessageEngine_PrepareConsoleMessages = result
        Exit Function
    End If

    For Each item In messages
        Select Case MessageEngine_NormalizeSeverity(CStr(item("Type")))
            Case "STOP"
                result.Add item
        End Select
    Next item

    For Each item In messages
        If MessageEngine_NormalizeSeverity(CStr(item("Type"))) = "WARNING" Then result.Add item
    Next item

    For Each item In messages
        If MessageEngine_NormalizeSeverity(CStr(item("Type"))) = "INFO" Then result.Add item
    Next item

    Set MessageEngine_PrepareConsoleMessages = result

End Function

'------------------------------------------------------------------------------
' FR: Retourne la collection Prepare Display Messages sans modifier les donnees d'entree.
' EN: Returns the Prepare Display Messages collection without mutating input data.
'------------------------------------------------------------------------------

Public Function MessageEngine_PrepareDisplayMessages(ByVal messages As Collection) As Collection

    Dim included As Collection
    Dim item As Variant

    Set included = New Collection

    If messages Is Nothing Then
        Set MessageEngine_PrepareDisplayMessages = included
        Exit Function
    End If

    For Each item In messages
        If MessageEngine_ShouldIncludeInDisplay(item) Then
            MessageEngine_AnnotateAcknowledgement item
            included.Add item
        End If
    Next item

    Set MessageEngine_PrepareDisplayMessages = MessageEngine_GroupDisplayStops(included)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la collection Group Display Stops sans modifier les donnees d'entree.
' EN: Returns the Group Display Stops collection without mutating input data.
'------------------------------------------------------------------------------

Private Function MessageEngine_GroupDisplayStops(ByVal messages As Collection) As Collection

    Dim result As Collection
    Dim item As Variant
    Dim consolidatedStop As Object
    Dim stopCount As Long
    Dim stopBlock As String
    Dim stopInserted As Boolean
    Dim severity As String

    Set result = New Collection

    If messages Is Nothing Then
        Set MessageEngine_GroupDisplayStops = result
        Exit Function
    End If

    For Each item In messages
        If MessageEngine_NormalizeSeverity(CStr(item("Type"))) = "STOP" Then
            stopCount = stopCount + 1
        End If
    Next item

    If stopCount <= 1 Then
        Set MessageEngine_GroupDisplayStops = messages
        Exit Function
    End If

    For Each item In messages
        severity = MessageEngine_NormalizeSeverity(CStr(item("Type")))

        If severity = "STOP" Then
            If stopBlock <> "" Then
                stopBlock = stopBlock & vbCrLf & vbCrLf & String$(36, "-") & vbCrLf & vbCrLf
            End If
            stopBlock = stopBlock & FormatPlanningConsoleMessageForCurrentLanguage(CStr(item("Message")))

            If Not stopInserted Then
                Set consolidatedStop = CreateObject("Scripting.Dictionary")
                consolidatedStop("Type") = "STOP"
                consolidatedStop("Message") = ""
                consolidatedStop("HistoryHandled") = True
                consolidatedStop("Acknowledged") = False
                result.Add consolidatedStop
                stopInserted = True
            End If
        Else
            result.Add item
        End If
    Next item

    If Not consolidatedStop Is Nothing Then
        consolidatedStop("Message") = stopBlock
    End If

    Set MessageEngine_GroupDisplayStops = result

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Include In Display satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Include In Display value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function MessageEngine_ShouldIncludeInDisplay(ByVal item As Variant) As Boolean

    Dim severity As String

    On Error GoTo SafeExit

    severity = MessageEngine_NormalizeSeverity(CStr(item("Type")))

    If severity = "INFO" Then
        MessageEngine_ShouldIncludeInDisplay = ShouldShowInfoOnlyPlanningConsole()
        Exit Function
    End If

    If severity = "WARNING" Then
        MessageEngine_ShouldIncludeInDisplay = Not PlanningMessage_IsAcknowledged(item)
        Exit Function
    End If

    MessageEngine_ShouldIncludeInDisplay = True
    Exit Function

SafeExit:
    MessageEngine_ShouldIncludeInDisplay = True

End Function

'------------------------------------------------------------------------------
' FR: Ajoute au message les metadonnees ACK calculees par EventHistory sans modifier son texte ni sa severite.
' EN: Adds ACK metadata computed by EventHistory without changing message text or severity.
'------------------------------------------------------------------------------

Private Sub MessageEngine_AnnotateAcknowledgement(ByVal item As Variant)

    On Error GoTo SafeExit

    If MessageEngine_NormalizeSeverity(CStr(item("Type"))) <> "WARNING" Then
        item("Acknowledged") = False
        Exit Sub
    End If

    item("Acknowledged") = PlanningMessage_IsAcknowledged(item)

SafeExit:
End Sub

'------------------------------------------------------------------------------
' FR: Indique si la collection Show Console satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Show Console collection satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Public Function MessageEngine_ShouldShowConsole(ByVal messages As Collection) As Boolean

    If messages Is Nothing Then Exit Function
    If messages.Count = 0 Then Exit Function

    MessageEngine_ShouldShowConsole = True

End Function

'------------------------------------------------------------------------------
' FR: Construit la collection Category Progress Caption a partir des donnees fournies par l'appelant.
' EN: Builds the Category Progress Caption collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function MessageEngine_BuildCategoryProgressCaption( _
    ByVal messages As Collection, _
    ByVal currentIndex As Long, _
    ByVal currentType As String) As String

    Dim stopTotal As Long
    Dim warnTotal As Long
    Dim infoTotal As Long
    Dim stopIndex As Long
    Dim warnIndex As Long
    Dim infoIndex As Long
    Dim infoCaption As String
    Dim i As Long
    Dim item As Object
    Dim sev As String
    Dim normalizedCurrent As String

    If messages Is Nothing Then
        MessageEngine_BuildCategoryProgressCaption = "STOP 0/0 | WARNING 0/0 | INFO 0/0"
        Exit Function
    End If

    normalizedCurrent = MessageEngine_NormalizeSeverity(currentType)

    For i = 1 To messages.Count
        Set item = messages(i)
        sev = MessageEngine_NormalizeSeverity(CStr(item("Type")))

        Select Case sev
            Case "STOP"
                stopTotal = stopTotal + 1
                If i <= currentIndex Then stopIndex = stopIndex + 1
            Case "WARNING"
                warnTotal = warnTotal + 1
                If i <= currentIndex Then warnIndex = warnIndex + 1
            Case "INFO"
                infoTotal = infoTotal + 1
                If i <= currentIndex Then infoIndex = infoIndex + 1
        End Select
    Next i

    If normalizedCurrent <> "STOP" Then stopIndex = 0
    If normalizedCurrent <> "WARNING" Then warnIndex = 0
    If normalizedCurrent <> "INFO" Then infoIndex = 0

    If ShouldShowInfoOnlyPlanningConsole() Then
        infoCaption = "INFO " & CStr(infoIndex) & "/" & CStr(infoTotal)
    Else
        infoCaption = "INFO Muted"
    End If

    MessageEngine_BuildCategoryProgressCaption = _
        "STOP " & CStr(stopIndex) & "/" & CStr(stopTotal) & _
        " | WARNING " & CStr(warnIndex) & "/" & CStr(warnTotal) & _
        " | " & infoCaption

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Severity selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Severity according to the component contract.
'------------------------------------------------------------------------------

Private Function MessageEngine_NormalizeSeverity(ByVal msgType As String) As String

    Select Case UCase$(Trim$(msgType))
        Case "ERROR", "STOP"
            MessageEngine_NormalizeSeverity = "STOP"
        Case "WARNING"
            MessageEngine_NormalizeSeverity = "WARNING"
        Case "INFO"
            MessageEngine_NormalizeSeverity = "INFO"
        Case Else
            MessageEngine_NormalizeSeverity = UCase$(Trim$(msgType))
    End Select

End Function
