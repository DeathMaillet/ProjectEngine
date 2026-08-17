Attribute VB_Name = "mod_PlanningConsolePolicy"
Option Explicit

'===============================================================================
' MODULE : mod_PlanningConsolePolicy
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Centralise la decision d'affichage modal ou de capture non interactive de la console planning.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Centralizes the decision between modal display and noninteractive capture of the planning console.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : PlanningConsolePolicy_EnableNonInteractive, PlanningConsolePolicy_DisableNonInteractive, PlanningConsolePolicy_IsNonInteractive, PlanningConsolePolicy_GetCapturedMessageCount, PlanningConsolePolicy_CaptureDisplayMessages
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gPlanningConsoleNonInteractive As Boolean
Private gPlanningConsoleCapturePath As String
Private gPlanningConsoleContext As String
Private gPlanningConsoleCapturedCount As Long

'------------------------------------------------------------------------------
' FR: Active la capture non interactive de la console Planning pour les harnais.
' EN: Enables non-interactive Planning console capture for harnesses.
'------------------------------------------------------------------------------
Public Sub PlanningConsolePolicy_EnableNonInteractive( _
    ByVal capturePath As String, _
    Optional ByVal contextName As String = "")

    gPlanningConsoleCapturePath = capturePath
    gPlanningConsoleContext = contextName
    gPlanningConsoleCapturedCount = 0
    gPlanningConsoleNonInteractive = (Trim$(capturePath) <> "")

    If gPlanningConsoleNonInteractive Then
        PlanningConsolePolicy_WriteHeader
        RunButtonsTrace_Checkpoint "ConsolePolicy", "Noninteractive console capture enabled"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Desactive le mode non interactif de la console Planning.
' EN: Disables non-interactive Planning console mode.
'------------------------------------------------------------------------------
Public Sub PlanningConsolePolicy_DisableNonInteractive()

    If gPlanningConsoleNonInteractive Then
        RunButtonsTrace_Checkpoint "ConsolePolicy", "Noninteractive console capture disabled"
    End If

    gPlanningConsoleNonInteractive = False
    gPlanningConsoleCapturePath = vbNullString
    gPlanningConsoleContext = vbNullString
    gPlanningConsoleCapturedCount = 0

End Sub

'------------------------------------------------------------------------------
' FR: Indique si la console Planning doit capturer sans afficher de fenetre modale.
' EN: Returns whether the Planning console should capture without showing a modal window.
'------------------------------------------------------------------------------
Public Function PlanningConsolePolicy_IsNonInteractive() As Boolean

    PlanningConsolePolicy_IsNonInteractive = gPlanningConsoleNonInteractive

End Function

'------------------------------------------------------------------------------
' FR: Retourne le nombre de messages captures pendant le contexte courant.
' EN: Returns the number of messages captured during the current context.
'------------------------------------------------------------------------------
Public Function PlanningConsolePolicy_GetCapturedMessageCount() As Long

    PlanningConsolePolicy_GetCapturedMessageCount = gPlanningConsoleCapturedCount

End Function

'------------------------------------------------------------------------------
' FR: Capture les messages qui auraient ete affiches dans frmPlanningMessages.
' EN: Captures messages that would have been displayed in frmPlanningMessages.
'------------------------------------------------------------------------------
Public Sub PlanningConsolePolicy_CaptureDisplayMessages( _
    ByVal messages As Collection, _
    Optional ByVal windowTitle As String = "Planning console")

    Dim i As Long
    Dim item As Object
    Dim msgType As String
    Dim msgText As String
    Dim ackText As String
    Dim eventType As String
    Dim eventHash As String
    Dim fileNo As Integer

    If Not gPlanningConsoleNonInteractive Then Exit Sub
    If Trim$(gPlanningConsoleCapturePath) = "" Then Exit Sub
    If messages Is Nothing Then Exit Sub

    On Error GoTo SafeExit

    fileNo = FreeFile
    Open gPlanningConsoleCapturePath For Append As #fileNo

    For i = 1 To messages.Count
        Set item = messages(i)
        msgType = PlanningConsolePolicy_SafeString(item, "Type")
        msgText = PlanningConsolePolicy_SafeString(item, "Message")
        ackText = PlanningConsolePolicy_SafeString(item, "Acknowledged")
        eventType = PlanningConsolePolicy_SafeString(item, "EventType")
        eventHash = PlanningConsolePolicy_SafeString(item, "EventHash")

        gPlanningConsoleCapturedCount = gPlanningConsoleCapturedCount + 1

        Print #fileNo, _
            PlanningConsolePolicy_Tsv(CStr(gPlanningConsoleCapturedCount)) & vbTab & _
            PlanningConsolePolicy_Tsv(gPlanningConsoleContext) & vbTab & _
            PlanningConsolePolicy_Tsv(windowTitle) & vbTab & _
            PlanningConsolePolicy_Tsv(msgType) & vbTab & _
            PlanningConsolePolicy_Tsv(ackText) & vbTab & _
            PlanningConsolePolicy_Tsv(eventType) & vbTab & _
            PlanningConsolePolicy_Tsv(eventHash) & vbTab & _
            PlanningConsolePolicy_Tsv(msgText)
    Next i

SafeExit:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Planning Console Policy Write Header dans le stockage possede par le domaine.
' EN: Writes or synchronizes Planning Console Policy Write Header in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub PlanningConsolePolicy_WriteHeader()

    Dim fileNo As Integer

    On Error GoTo SafeExit
    fileNo = FreeFile
    Open gPlanningConsoleCapturePath For Output As #fileNo
    Print #fileNo, "Index" & vbTab & "Context" & vbTab & "WindowTitle" & vbTab & _
        "Type" & vbTab & "Acknowledged" & vbTab & "EventType" & vbTab & _
        "EventHash" & vbTab & "Message"

SafeExit:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Planning Console Policy Safe String sans modifier les donnees d'entree.
' EN: Returns the Planning Console Policy Safe String map without mutating input data.
'------------------------------------------------------------------------------

Private Function PlanningConsolePolicy_SafeString( _
    ByVal item As Object, _
    ByVal keyName As String) As String

    On Error GoTo SafeExit
    If item Is Nothing Then Exit Function
    If item.Exists(keyName) Then PlanningConsolePolicy_SafeString = CStr(item(keyName))

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Planning Console Policy Tsv sans modifier les donnees d'entree.
' EN: Returns the Planning Console Policy Tsv value without mutating input data.
'------------------------------------------------------------------------------

Private Function PlanningConsolePolicy_Tsv(ByVal value As String) As String

    PlanningConsolePolicy_Tsv = Replace(Replace(Replace(value, vbTab, " "), vbCr, "\r"), vbLf, "\n")

End Function
