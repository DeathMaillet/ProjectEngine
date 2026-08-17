VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParsedPlanningLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' MODULE : clsParsedPlanningLink
' DOMAINE / DOMAIN : Parsed Planning Network
'
' FR
' Encapsule l'etat et le contrat Parsed Planning Link sans exposer ses structures mutables.
' Ne porte aucune orchestration de workflow.
'
' EN
' Encapsulates Parsed Planning Link state and contract without exposing mutable internals.
' Owns no workflow orchestration.
'
' CONTRATS / CONTRACTS : RowIndex, SuccId, SuccWbs, PredId, PredWbs, RawLinkType, LinkType, Lag
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private mRowIndex As Long
Private mSuccId As String
Private mSuccWbs As String
Private mPredId As String
Private mPredWbs As String
Private mRawLinkType As String
Private mLinkType As String
Private mLag As Double
Private mRawToken As String
Private mExpandedFrom As String

'------------------------------------------------------------------------------
' FR: Copie les champs structurels d'un lien parse dans ce record immutable.
' EN: Copies parsed-link structural fields into this immutable record.
'------------------------------------------------------------------------------

Friend Sub Initialize( _
    ByVal rowIndex As Long, _
    ByVal succId As String, _
    ByVal succWbs As String, _
    ByVal predId As String, _
    ByVal predWbs As String, _
    ByVal rawLinkType As String, _
    ByVal linkType As String, _
    ByVal lag As Double, _
    ByVal rawToken As String, _
    ByVal expandedFrom As String)

    mRowIndex = rowIndex
    mSuccId = succId
    mSuccWbs = succWbs
    mPredId = predId
    mPredWbs = predWbs
    mRawLinkType = rawLinkType
    mLinkType = linkType
    mLag = lag
    mRawToken = rawToken
    mExpandedFrom = expandedFrom

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Row Index sans modifier les donnees d'entree.
' EN: Returns the Row Index value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get RowIndex() As Long
    RowIndex = mRowIndex
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Succ ID sans modifier les donnees d'entree.
' EN: Returns the Succ ID value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get SuccId() As String
    SuccId = mSuccId
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Succ WBS sans modifier les donnees d'entree.
' EN: Returns the Succ WBS value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get SuccWbs() As String
    SuccWbs = mSuccWbs
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Pred ID sans modifier les donnees d'entree.
' EN: Returns the Pred ID value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get PredId() As String
    PredId = mPredId
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Pred WBS sans modifier les donnees d'entree.
' EN: Returns the Pred WBS value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get PredWbs() As String
    PredWbs = mPredWbs
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Raw Link Type sans modifier les donnees d'entree.
' EN: Returns the Raw Link Type value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get RawLinkType() As String
    RawLinkType = mRawLinkType
End Property

'------------------------------------------------------------------------------
' FR: Retourne le type FS/SS/FF du record de lien sans modifier les donnees d'entree.
' EN: Returns the link record FS/SS/FF type without mutating input data.
'------------------------------------------------------------------------------

Public Property Get LinkType() As String
    LinkType = mLinkType
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Lag sans modifier les donnees d'entree.
' EN: Returns the Lag value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get Lag() As Double
    Lag = mLag
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Raw Token sans modifier les donnees d'entree.
' EN: Returns the Raw Token value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get RawToken() As String
    RawToken = mRawToken
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Expanded From sans modifier les donnees d'entree.
' EN: Returns the Expanded From value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get ExpandedFrom() As String
    ExpandedFrom = mExpandedFrom
End Property

