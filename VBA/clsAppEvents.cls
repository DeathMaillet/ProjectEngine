VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit

'===============================================================================
' MODULE : clsAppEvents
' DOMAINE / DOMAIN : Workbook / Sheet Events
'
' FR
' Encapsule l'etat et le contrat App Events sans exposer ses structures mutables.
' Ne porte aucune orchestration de workflow.
'
' EN
' Encapsulates App Events state and contract without exposing mutable internals.
' Owns no workflow orchestration.
'
' CONTRATS / CONTRACTS : Aucun / None
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : App_SheetChange
'===============================================================================

    
    Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
    
'------------------------------------------------------------------------------
' FR: Relaye les changements de feuilles du classeur vers le dispatcher central.
' EN: Routes workbook sheet changes to the central dispatcher.
'------------------------------------------------------------------------------
    Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
        If Sh Is Nothing Then Exit Sub
        If Target Is Nothing Then Exit Sub
        If Not (Sh.Parent Is ThisWorkbook) Then Exit Sub

        Handle_Worksheet_Change Sh, Target
    
    End Sub
    
    

