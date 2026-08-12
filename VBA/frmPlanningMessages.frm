VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlanningMessages 
   Caption         =   "Planning warnings"
   ClientHeight    =   12840
   ClientLeft      =   930
   ClientTop       =   3705
   ClientWidth     =   3.06870e5
   OleObjectBlob   =   "frmPlanningMessages.frx":0000
End
Attribute VB_Name = "frmPlanningMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' MODULE : frmPlanningMessages
' DOMAINE / DOMAIN : Planning Console UI
'
' FR
' Affiche les messages planning, la navigation et les actions ACK dans la console modale.
' Ne prepare pas, ne filtre pas et ne journalise pas les diagnostics.
'
' EN
' Displays planning messages, navigation and ACK actions in the modal console.
' Does not prepare, filter or log diagnostics.
'
' CONTRATS / CONTRACTS : LoadMessages
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : UserForm_Initialize
'===============================================================================

Private mItems As Collection
Private mTitle As String
Private mIndex As Long
Private mLoadingMessage As Boolean

'------------------------------------------------------------------------------
' FR: Charge la collection Messages depuis sa source proprietaire sans appliquer de politique aval.
' EN: Loads the Messages collection from its owning source without applying downstream policy.
'------------------------------------------------------------------------------

Public Sub LoadMessages( _
    ByVal messages As Collection, _
    Optional ByVal windowTitle As String = "Planning console")

    Set mItems = MessageEngine_PrepareDisplayMessages(messages)
    mTitle = windowTitle
    mIndex = 1

    ApplyFormLayout
    CenterFormOnExcel
    RenderCurrentMessage

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Previous Click dans planning message form.
' EN: Handles the cmd Previous Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdPrevious_Click()

    If mItems Is Nothing Then Exit Sub
    If mItems.Count = 0 Then Exit Sub

    If mIndex > 1 Then
        mIndex = mIndex - 1
        RenderCurrentMessage
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Next Click dans planning message form.
' EN: Handles the cmd Next Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdNext_Click()

    If mItems Is Nothing Then Exit Sub
    If mItems.Count = 0 Then Exit Sub

    If mIndex < mItems.Count Then
        mIndex = mIndex + 1
        RenderCurrentMessage
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Close Click dans planning message form.
' EN: Handles the cmd Close Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdClose_Click()

    Unload Me

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Render Current Message sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Render Current Message without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub RenderCurrentMessage()

    Dim item As Object
    Dim msgType As String

    If mItems Is Nothing Then Exit Sub

    If mItems.Count = 0 Then
        lblTitle.caption = "Messages"
        lblCounter.caption = "STOP 0/0 | WARNING 0/0 | INFO 0/0"
        lblMessageType.caption = ""
        txtMessage.Text = ""
        cmdPrevious.enabled = False
        cmdNext.enabled = False
        chkWarningAck.Visible = False
        Exit Sub
    End If

    If mIndex < 1 Then mIndex = 1
    If mIndex > mItems.Count Then mIndex = mItems.Count

    Set item = mItems(mIndex)
    msgType = UCase$(Trim$(CStr(item("Type"))))
    If MessageIsWarning(msgType) Then item("Acknowledged") = PlanningMessage_IsAcknowledged(item)

    lblTitle.caption = "Messages"
    lblCounter.caption = BuildCategoryProgressCaption(msgType)

    mLoadingMessage = True
    ApplyMessageTypeVisual msgType
    ApplyAckControlVisual item, msgType
    mLoadingMessage = False

    txtMessage.Text = FormatPlanningConsoleMessageForCurrentLanguage(CStr(item("Message")))

    cmdPrevious.enabled = (mIndex > 1)
    cmdNext.enabled = (mIndex < mItems.Count)

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Message Type Visual sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Message Type Visual without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyMessageTypeVisual(ByVal msgType As String)

    Select Case UCase$(Trim$(msgType))

        Case "STOP", "ERROR"
            lblMessageType.caption = "STOP"
            lblMessageType.ForeColor = RGB(156, 0, 6)
            lblMessageType.BackColor = RGB(255, 235, 238)

        Case "WARNING"
            If CurrentMessageAcknowledged() Then
                lblMessageType.caption = "WARNING"
                lblMessageType.ForeColor = RGB(92, 74, 0)
                lblMessageType.BackColor = RGB(236, 232, 214)
            Else
                lblMessageType.caption = "WARNING"
                lblMessageType.ForeColor = RGB(156, 101, 0)
                lblMessageType.BackColor = RGB(255, 248, 225)
            End If

        Case "INFO"
            lblMessageType.caption = "INFO"
            lblMessageType.ForeColor = RGB(0, 97, 0)
            lblMessageType.BackColor = RGB(232, 245, 233)

        Case Else
            lblMessageType.caption = msgType
            ApplyMessageTypeBadgeLayout
            lblMessageType.ForeColor = RGB(60, 60, 60)
            lblMessageType.BackColor = RGB(240, 240, 240)

    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Current Message Acknowledged sans exposer de mutateur sur l'etat source.
' EN: Returns the Current Message Acknowledged map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function CurrentMessageAcknowledged() As Boolean

    Dim item As Object

    On Error GoTo SafeExit

    If mItems Is Nothing Then Exit Function
    If mIndex < 1 Or mIndex > mItems.Count Then Exit Function
    If TypeName(mItems(mIndex)) <> "Dictionary" Then Exit Function

    Set item = mItems(mIndex)
    CurrentMessageAcknowledged = CBool(item("Acknowledged"))

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Message Is Warning sans modifier les donnees d'entree.
' EN: Returns the Message Is Warning value without mutating input data.
'------------------------------------------------------------------------------

Private Function MessageIsWarning(ByVal msgType As String) As Boolean

    MessageIsWarning = (UCase$(Trim$(msgType)) = "WARNING")

End Function

'------------------------------------------------------------------------------
' FR: Actualise Apply Ack Control Visual sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Ack Control Visual without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyAckControlVisual( _
    ByVal item As Object, _
    ByVal msgType As String)

    If Not MessageIsWarning(msgType) Then
        chkWarningAck.Visible = False
        Exit Sub
    End If

    chkWarningAck.Visible = PlanningMessage_CanAcknowledge(item)
    chkWarningAck.value = PlanningMessage_IsAcknowledged(item)
    chkWarningAck.caption = "Cacher / Hide"

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur chk Warning Ack Click dans planning message form.
' EN: Handles the chk Warning Ack Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub chkWarningAck_Click()

    Dim item As Object

    If mLoadingMessage Then Exit Sub
    If mItems Is Nothing Then Exit Sub
    If mIndex < 1 Or mIndex > mItems.Count Then Exit Sub
    If TypeName(mItems(mIndex)) <> "Dictionary" Then Exit Sub

    Set item = mItems(mIndex)
    If Not PlanningMessage_CanAcknowledge(item) Then Exit Sub

    SetPlanningWarningAckState item, CBool(chkWarningAck.value)
    item("Acknowledged") = PlanningMessage_IsAcknowledged(item)

    mLoadingMessage = True
    chkWarningAck.value = CBool(item("Acknowledged"))
    mLoadingMessage = False

    ApplyMessageTypeVisual CStr(item("Type"))

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Clear Ack Click dans planning message form.
' EN: Handles the cmd Clear Ack Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdClearAck_Click()

    ClearPlanningWarningAcknowledgements
    RenderCurrentMessage

End Sub

'------------------------------------------------------------------------------
' FR: Traite le clic utilisateur cmd Clear History Click dans planning message form.
' EN: Handles the cmd Clear History Click user click in planning message form.
'------------------------------------------------------------------------------
Private Sub cmdClearHistory_Click()

    ClearPlanningEventHistory
    RenderCurrentMessage

End Sub
'------------------------------------------------------------------------------
' FR: Construit la valeur Category Progress Caption a partir des donnees fournies par l'appelant.
' EN: Builds the Category Progress Caption value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildCategoryProgressCaption(ByVal currentType As String) As String

    BuildCategoryProgressCaption = MessageEngine_BuildCategoryProgressCaption(mItems, mIndex, currentType)

End Function
'------------------------------------------------------------------------------
' FR: Actualise Apply Form Layout sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Form Layout without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyFormLayout()

    Dim formW As Single
    Dim formH As Single
    Dim margin As Single
    Dim headerTop As Single
    Dim headerH As Single
    Dim textTop As Single
    Dim textH As Single
    Dim footerH As Single
    Dim buttonTop As Single
    Dim buttonW As Single
    Dim buttonH As Single
    Dim gap As Single

    margin = 24
    headerTop = 18
    headerH = 42
    footerH = 78
    buttonW = 138
    buttonH = 30
    gap = 14

    formW = CalcDynamicFormWidth()
    formH = CalcDynamicFormHeight()

    Me.caption = mTitle
    Me.StartUpPosition = 0
    Me.Width = formW
    Me.Height = formH
    Me.BackColor = RGB(250, 250, 250)

    lblTitle.Left = margin
    lblTitle.Top = headerTop
    lblTitle.Width = 130
    lblTitle.Height = 18
    lblTitle.caption = "Messages"
    lblTitle.Font.Name = "Segoe UI"
    lblTitle.Font.Size = 10
    lblTitle.Font.Bold = True
    lblTitle.ForeColor = RGB(35, 35, 35)
    lblTitle.BackStyle = fmBackStyleTransparent

    lblMessageType.Left = margin
    lblMessageType.Top = headerTop + 22
    lblMessageType.Width = 82
    lblMessageType.Height = 14
    lblMessageType.TextAlign = fmTextAlignCenter
    lblMessageType.Font.Name = "Segoe UI"
    lblMessageType.Font.Size = 8
    lblMessageType.Font.Bold = True
    lblMessageType.BorderStyle = fmBorderStyleSingle
    lblMessageType.SpecialEffect = fmSpecialEffectFlat
    lblMessageType.WordWrap = False
    lblMessageType.AutoSize = False

    lblCounter.Left = margin + 150
    lblCounter.Top = headerTop
    lblCounter.Width = formW - (2 * margin) - 150
    lblCounter.Height = 18
    lblCounter.TextAlign = fmTextAlignRight
    lblCounter.Font.Name = "Segoe UI"
    lblCounter.Font.Size = 10
    lblCounter.Font.Bold = True
    lblCounter.ForeColor = RGB(35, 35, 35)
    lblCounter.BackStyle = fmBackStyleTransparent

    textTop = headerTop + headerH + 10
    textH = formH - textTop - footerH

    If textH < 150 Then textH = 150

    txtMessage.Left = margin
    txtMessage.Top = textTop
    txtMessage.Width = formW - (2 * margin)
    txtMessage.Height = textH
    txtMessage.Multiline = True
    txtMessage.ScrollBars = fmScrollBarsVertical
    txtMessage.Locked = True
    txtMessage.WordWrap = True
    txtMessage.BorderStyle = fmBorderStyleSingle
    txtMessage.SpecialEffect = fmSpecialEffectFlat
    txtMessage.BackColor = RGB(255, 255, 255)
    txtMessage.ForeColor = RGB(25, 25, 25)
    txtMessage.Font.Name = "Segoe UI"
    txtMessage.Font.Size = 9

    buttonTop = txtMessage.Top + txtMessage.Height + 18

    cmdPrevious.Left = margin
    cmdPrevious.Top = buttonTop
    cmdPrevious.Width = buttonW
    cmdPrevious.Height = buttonH
    cmdPrevious.caption = "Précédent / Previous"
    cmdPrevious.Font.Name = "Segoe UI"
    cmdPrevious.Font.Size = 9

    cmdNext.Left = cmdPrevious.Left + buttonW + gap
    cmdNext.Top = buttonTop
    cmdNext.Width = buttonW
    cmdNext.Height = buttonH
    cmdNext.caption = "Suivant / Next"
    cmdNext.Font.Name = "Segoe UI"
    cmdNext.Font.Size = 9

    cmdClose.Left = formW - margin - buttonW
    cmdClose.Top = buttonTop
    cmdClose.Width = buttonW
    cmdClose.Height = buttonH
    cmdClose.caption = "Fermer / Close"
    cmdClose.Font.Name = "Segoe UI"
    cmdClose.Font.Size = 9

    chkWarningAck.Left = lblMessageType.Left + lblMessageType.Width + 10
    chkWarningAck.Top = lblMessageType.Top + 1
    chkWarningAck.Width = 92
    chkWarningAck.Height = 18
    chkWarningAck.caption = "Cacher / Hide"
    chkWarningAck.Font.Name = "Segoe UI"
    chkWarningAck.Font.Size = 9
    chkWarningAck.BackStyle = fmBackStyleTransparent
    chkWarningAck.Visible = False

    cmdClearAck.Visible = False
    cmdClearHistory.Visible = False

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Calc Dynamic Form Width sans modifier les donnees d'entree.
' EN: Returns the Calc Dynamic Form Width map without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcDynamicFormWidth() As Single

    Dim maxLineLen As Long
    Dim msg As Variant
    Dim lines As Variant
    Dim oneLine As Variant
    Dim estimatedW As Single
    Dim item As Object
    Dim msgText As String

    maxLineLen = 0

    If Not mItems Is Nothing Then
        For Each msg In mItems

            If TypeName(msg) = "Dictionary" Then
                Set item = msg
                msgText = FormatPlanningConsoleMessageForCurrentLanguage(CStr(item("Message")))
            Else
                msgText = FormatPlanningConsoleMessageForCurrentLanguage(CStr(msg))
            End If

            lines = Split(msgText, vbCrLf)

            For Each oneLine In lines
                If Len(CStr(oneLine)) > maxLineLen Then
                    maxLineLen = Len(CStr(oneLine))
                End If
            Next oneLine
        Next msg
    End If

    estimatedW = 470 + (maxLineLen * 1.4)

    If estimatedW < 600 Then estimatedW = 600
    If estimatedW > 760 Then estimatedW = 760

    CalcDynamicFormWidth = estimatedW

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Calc Dynamic Form Height sans modifier les donnees d'entree.
' EN: Returns the Calc Dynamic Form Height map without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcDynamicFormHeight() As Single

    Dim maxLineCount As Long
    Dim msg As Variant
    Dim lines As Variant
    Dim estimatedH As Single
    Dim item As Object
    Dim msgText As String

    maxLineCount = 0

    If Not mItems Is Nothing Then
        For Each msg In mItems

            If TypeName(msg) = "Dictionary" Then
                Set item = msg
                msgText = FormatPlanningConsoleMessageForCurrentLanguage(CStr(item("Message")))
            Else
                msgText = FormatPlanningConsoleMessageForCurrentLanguage(CStr(msg))
            End If

            lines = Split(msgText, vbCrLf)

            If UBound(lines) - LBound(lines) + 1 > maxLineCount Then
                maxLineCount = UBound(lines) - LBound(lines) + 1
            End If
        Next msg
    End If

    estimatedH = 250 + (maxLineCount * 13)

    If estimatedH < 370 Then estimatedH = 370
    If estimatedH > 520 Then estimatedH = 520

    CalcDynamicFormHeight = estimatedH

End Function

'------------------------------------------------------------------------------
' FR: Positionne la valeur Form On Excel relativement a la fenetre Excel active.
' EN: Positions the Form On Excel value relative to the active Excel window.
'------------------------------------------------------------------------------

Private Sub CenterFormOnExcel()

    On Error GoTo SafeExit

    Me.StartUpPosition = 0

    Me.Left = Application.Left + ((Application.Width - Me.Width) / 2)
    Me.Top = Application.Top + ((Application.Height - Me.Height) / 2)

    If Me.Left < 0 Then Me.Left = 20
    If Me.Top < 0 Then Me.Top = 20

SafeExit:
End Sub

'------------------------------------------------------------------------------
' FR: Initialise l'etat visuel du formulaire de messages planning.
' EN: Initializes the planning message form visual state.
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

    ApplyMessageTypeBadgeLayout

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Message Type Badge Layout sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Message Type Badge Layout without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyMessageTypeBadgeLayout()

    Const BADGE_HEIGHT As Single = 20

    Dim oldTop As Single
    Dim oldHeight As Single

    On Error Resume Next

    With lblMessageType

        oldTop = .Top
        oldHeight = .Height

        .AutoSize = False
        .WordWrap = False
        .TextAlign = fmTextAlignCenter

        .Height = BADGE_HEIGHT

        If oldHeight > BADGE_HEIGHT Then
            .Top = oldTop + ((oldHeight - BADGE_HEIGHT) / 2)
        End If

        .Font.Bold = True
        .Font.Size = 9

        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat

    End With

    On Error GoTo 0

End Sub






