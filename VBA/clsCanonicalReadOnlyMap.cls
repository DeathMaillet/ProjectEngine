VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCanonicalReadOnlyMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' MODULE : clsCanonicalReadOnlyMap
' DOMAINE / DOMAIN : Canonical Read Models
'
' FR
' Encapsule l'etat et le contrat Canonical Read Only Map sans exposer ses structures mutables.
' Ne porte aucune orchestration de workflow.
'
' EN
' Encapsulates Canonical Read Only Map state and contract without exposing mutable internals.
' Owns no workflow orchestration.
'
' CONTRATS / CONTRACTS : Exists, Item, Count, Keys, Items, CompareMode
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private mValues As Object

'------------------------------------------------------------------------------
' FR: Copie une map source dans le stockage prive immuable du Read Model.
' EN: Copies a source map into the Read Model's private immutable storage.
'------------------------------------------------------------------------------
Friend Sub InitializeFrom(ByVal source As Object)

    Dim key As Variant

    Set mValues = CreateObject("Scripting.Dictionary")
    mValues.CompareMode = source.CompareMode

    For Each key In source.Keys
        mValues(key) = source(key)
    Next key

End Sub

'------------------------------------------------------------------------------
' FR: Indique si la cle demandee existe dans l'index immuable.
' EN: Returns whether the requested key exists in the immutable index.
'------------------------------------------------------------------------------
Public Function Exists(ByVal key As Variant) As Boolean

    Exists = mValues.Exists(key)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur associee a une cle sans exposer de mutateur.
' EN: Returns the value associated with a key without exposing a mutator.
'------------------------------------------------------------------------------
Public Property Get Item(ByVal key As Variant) As Variant
Attribute Item.VB_UserMemId = 0

    If mValues.Exists(key) Then
        Item = mValues(key)
    Else
        Item = Empty
    End If

End Property

'------------------------------------------------------------------------------
' FR: Retourne le nombre d'entrees de l'index immuable.
' EN: Returns the number of entries in the immutable index.
'------------------------------------------------------------------------------
Public Property Get Count() As Long

    Count = mValues.Count

End Property

'------------------------------------------------------------------------------
' FR: Retourne une copie du tableau des cles de l'index.
' EN: Returns a copy of the index key array.
'------------------------------------------------------------------------------
Public Property Get Keys() As Variant

    Keys = mValues.Keys

End Property

'------------------------------------------------------------------------------
' FR: Retourne une copie du tableau des valeurs de l'index.
' EN: Returns a copy of the index value array.
'------------------------------------------------------------------------------
Public Property Get Items() As Variant

    Items = mValues.Items

End Property

'------------------------------------------------------------------------------
' FR: Retourne le mode de comparaison conserve depuis la map source.
' EN: Returns the comparison mode preserved from the source map.
'------------------------------------------------------------------------------
Public Property Get CompareMode() As VbCompareMethod

    CompareMode = mValues.CompareMode

End Property
