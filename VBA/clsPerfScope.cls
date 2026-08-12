VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerfScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' MODULE : clsPerfScope
' DOMAINE / DOMAIN : Performance / Profiler
'
' FR
' Encapsule l'etat et le contrat Perf Scope sans exposer ses structures mutables.
' Ne porte aucune orchestration de workflow.
'
' EN
' Encapsulates Perf Scope state and contract without exposing mutable internals.
' Owns no workflow orchestration.
'
' CONTRATS / CONTRACTS : Init
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Class_Terminate
'===============================================================================


Private mToken As Long

'------------------------------------------------------------------------------
' FR: Initialise la valeur Init et en prend une copie conforme au contrat du composant.
' EN: Initializes the Init value and takes a contract-compliant copy of its input.
'------------------------------------------------------------------------------

Public Sub Init(ByVal token As Long)
    mToken = token
End Sub

'------------------------------------------------------------------------------
' FR: Libere automatiquement un scope de profiling a la destruction de l'objet.
' EN: Automatically closes a profiling scope when the object is destroyed.
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
    If mToken <> 0 Then Profiler_EndScope mToken
    mToken = 0
End Sub
