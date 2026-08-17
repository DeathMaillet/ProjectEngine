Attribute VB_Name = "mod_AppEventsService"
Option Explicit

'===============================================================================
' MODULE : mod_AppEventsService
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
' CONTRATS / CONTRACTS : AppEvents_EnsureInitialized
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'==============================================================
' Stable boundary for application-event initialization.
' ThisWorkbook remains the owner of the clsAppEvents instance.
'==============================================================

'------------------------------------------------------------------------------
' FR: Garantit que les evenements Application du classeur sont initialises.
' EN: Ensures that the workbook Application events are initialized.
'------------------------------------------------------------------------------
Public Sub AppEvents_EnsureInitialized()

    ThisWorkbook.Init_AppEvents

End Sub
