VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParsedPlanningNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' MODULE : clsParsedPlanningNetwork
' DOMAINE / DOMAIN : Parsed Planning Network
'
' FR
' Proprietaire canonique du reseau de planning parse et de ses liens immuables.
' Ne calcule ni ordre topologique, ni politique Core, ni routing Gantt.
'
' EN
' Canonical owner of the parsed planning network and its immutable links.
' Does not calculate topological order, Core policy or Gantt routing.
'
' CONTRATS / CONTRACTS : Count, Item, HasColumn
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private mLinks As Collection
Private mColumns As Object

'------------------------------------------------------------------------------
' FR: Initialise la collection Network From et en prend une copie conforme au contrat du composant.
' EN: Initializes the Network From collection and takes a contract-compliant copy of its input.
'------------------------------------------------------------------------------

Friend Sub InitializeNetworkFrom(ByVal links As Collection, ByVal columns As Object)

    Dim link As clsParsedPlanningLink

    Set mLinks = New Collection
    For Each link In links
        mLinks.Add link
    Next link

    Set mColumns = columns

End Sub

'------------------------------------------------------------------------------
' FR: Retourne le nombre d'elements exposes par le composant sans exposer de mutateur sur l'etat source.
' EN: Returns the number of items exposed by the component without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Property Get Count() As Long
    Count = mLinks.Count
End Property

'------------------------------------------------------------------------------
' FR: Retourne la valeur Item sans modifier les donnees d'entree.
' EN: Returns the Item value without mutating input data.
'------------------------------------------------------------------------------

Public Property Get Item(ByVal index As Long) As clsParsedPlanningLink
Attribute Item.VB_UserMemId = 0
    Set Item = mLinks(index)
End Property

'------------------------------------------------------------------------------
' FR: Indique si la condition Has Column est satisfaite sans modifier les donnees source.
' EN: Returns whether the Has Column condition is satisfied without mutating source data.
'------------------------------------------------------------------------------

Public Function HasColumn(ByVal columnName As String) As Boolean
    HasColumn = mColumns.Exists(columnName)
End Function
