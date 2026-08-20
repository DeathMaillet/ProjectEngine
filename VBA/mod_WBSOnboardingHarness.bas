Attribute VB_Name = "mod_WBSOnboardingHarness"
Option Explicit

'===============================================================================
' MODULE : mod_WBSOnboardingHarness
' DOMAINE / DOMAIN : WBS Proof Harness
'
' FR
' Expose les preuves VBA necessaires au harnais non interactif du guide WBS.
' Ne contient aucune logique produit et ne doit etre appele que sur une copie.
'
' EN
' Exposes VBA proofs required by the noninteractive WBS onboarding harness.
' Contains no product logic and must only be called on a copy.
'
' CONTRATS / CONTRACTS : WBSOnboardingHarness_CaptureMissingColumnDiagnostic
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

'------------------------------------------------------------------------------
' FR: Renomme temporairement Task Name et retourne le diagnostic produit par le contrat WBS.
' EN: Temporarily renames Task Name and returns the diagnostic produced by the WBS contract.
' FR - Effet de bord : restaure toujours le nom de colonne avant de rendre la main.
' EN - Side effect: always restores the column name before returning.
'------------------------------------------------------------------------------
Public Function WBSOnboardingHarness_CaptureMissingColumnDiagnostic() As String

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim capturedDescription As String
    Dim renamedColumn As Boolean

    On Error GoTo HarnessError

    Set ws = ThisWorkbook.Worksheets("WBS")
    Set tbl = ws.ListObjects("tbl_WBS")

    tbl.ListColumns("Task Name").Name = "Task Name Missing"
    renamedColumn = True

    On Error Resume Next
    WBS_EnsureOnboardingGuide
    capturedDescription = Err.Description
    Err.Clear
    On Error GoTo HarnessError

SafeExit:
    On Error Resume Next
    If renamedColumn Then tbl.ListColumns("Task Name Missing").Name = "Task Name"
    On Error GoTo 0

    WBSOnboardingHarness_CaptureMissingColumnDiagnostic = capturedDescription
    Exit Function

HarnessError:
    capturedDescription = "HARNESS_ERROR: " & Err.Description
    Resume SafeExit

End Function
