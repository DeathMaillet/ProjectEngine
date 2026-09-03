VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDashboardReadContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' MODULE : clsDashboardReadContext
' DOMAINE / DOMAIN : Dashboard
'
' FR
' Encapsule l'etat et le contrat Dashboard Read Context sans exposer ses structures mutables.
' Ne porte aucune orchestration de workflow.
'
' EN
' Encapsulates Dashboard Read Context state and contract without exposing mutable internals.
' Owns no workflow orchestration.
'
' CONTRATS / CONTRACTS : WbsTable, CalcTable, WbsColumns, CalcColumns, SCurveProjection
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private mWbsTable As ListObject
Private mCalcTable As ListObject
Private mWbsColumns As Object
Private mCalcColumns As Object
Private mSCurveProjection As Object

'------------------------------------------------------------------------------
' FR: Initialise la map Dashboard Read Context et en prend une copie conforme au contrat du composant.
' EN: Initializes the Dashboard Read Context map and takes a contract-compliant copy of its input.
'------------------------------------------------------------------------------

Friend Sub InitializeDashboardReadContext( _
    ByVal wbsTable As ListObject, _
    ByVal calcTable As ListObject, _
    ByVal wbsColumns As Object, _
    ByVal calcColumns As Object, _
    ByVal scurveProjection As Object)

    Set mWbsTable = wbsTable
    Set mCalcTable = calcTable
    Set mWbsColumns = wbsColumns
    Set mCalcColumns = calcColumns
    Set mSCurveProjection = scurveProjection

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Wbs Table sans modifier les donnees d'entree.
' EN: Returns the Wbs Table reference without mutating input data.
'------------------------------------------------------------------------------

Public Property Get WbsTable() As ListObject
    Set WbsTable = mWbsTable
End Property

'------------------------------------------------------------------------------
' FR: Retourne la reference Calc Table sans modifier les donnees d'entree.
' EN: Returns the Calc Table reference without mutating input data.
'------------------------------------------------------------------------------

Public Property Get CalcTable() As ListObject
    Set CalcTable = mCalcTable
End Property

'------------------------------------------------------------------------------
' FR: Retourne la map Wbs Columns sans modifier les donnees d'entree.
' EN: Returns the Wbs Columns map without mutating input data.
'------------------------------------------------------------------------------

Public Property Get WbsColumns() As Object
    Set WbsColumns = mWbsColumns
End Property

'------------------------------------------------------------------------------
' FR: Retourne la map Calc Columns sans modifier les donnees d'entree.
' EN: Returns the Calc Columns map without mutating input data.
'------------------------------------------------------------------------------

Public Property Get CalcColumns() As Object
    Set CalcColumns = mCalcColumns
End Property

'------------------------------------------------------------------------------
' FR: Retourne la map S Curve Projection sans modifier les donnees d'entree.
' EN: Returns the S Curve Projection map without mutating input data.
'------------------------------------------------------------------------------

Public Property Get SCurveProjection() As Object
    Set SCurveProjection = mSCurveProjection
End Property
