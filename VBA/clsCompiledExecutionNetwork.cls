VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompiledExecutionNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' MODULE : clsCompiledExecutionNetwork
' DOMAINE / DOMAIN : Core Calculation / Analytics
'
' FR
' Snapshot transactionnel immuable de la structure d'execution d'un dataset.
' Partage les index, adjacences et ordres entre Core et analytics.
' Ne possede ni Excel, ni writers, ni cache inter-transaction.
'
' EN
' Immutable transaction snapshot of a dataset execution structure.
' Shares indexes, adjacency and orders between Core and analytics.
' Owns neither Excel, writers nor an inter-transaction cache.
'
' CONTRATS / CONTRACTS : InitializeStructures, InitializeIndexedArrays, structural property getters
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private mRowById As Object
Private mParentIds As Object
Private mDirectChildrenById As Object
Private mChildrenByPred As Object
Private mValidLeafIds As Object
Private mLoeIds As Object
Private mTopoOrder As Collection
Private mAnalyticsPredsById As Object
Private mAnalyticsChildrenById As Object
Private mPredLagBySuccPred As Object
Private mPredTypeBySuccPred As Object
Private mIdToWbs As Object
Private mNodeIndexById As Object

Private mNodeIds As Variant
Private mRowsByNode As Variant
Private mParentByNode As Variant
Private mPredOffsets As Variant
Private mPredNodes As Variant
Private mPredTypes As Variant
Private mPredLags As Variant
Private mCorePredOffsets As Variant
Private mCorePredNodes As Variant
Private mCorePredIds As Variant
Private mCorePredTypes As Variant
Private mCorePredLags As Variant
Private mCorePredSummarySourceIds As Variant
Private mTopoNodes As Variant
Private mComponentByNode As Variant

Private mNodeCount As Long
Private mLeafCount As Long
Private mEdgeCount As Long
Private mComponentCount As Long
Private mIsValid As Boolean

' FR: Initialise une fois le snapshot compile; l'appelant conserve le dataset.
' EN: Initializes the compiled snapshot once; the caller retains the dataset.
Friend Sub InitializeStructures( _
    ByVal rowById As Object, _
    ByVal parentIds As Object, _
    ByVal directChildrenById As Object, _
    ByVal childrenByPred As Object, _
    ByVal validLeafIds As Object, _
    ByVal loeIds As Object, _
    ByVal topoOrder As Collection, _
    ByVal analyticsPredsById As Object, _
    ByVal analyticsChildrenById As Object, _
    ByVal predLagBySuccPred As Object, _
    ByVal predTypeBySuccPred As Object, _
    ByVal idToWbs As Object, _
    ByVal nodeIndexById As Object)

    Set mRowById = rowById
    Set mParentIds = parentIds
    Set mDirectChildrenById = directChildrenById
    Set mChildrenByPred = childrenByPred
    Set mValidLeafIds = validLeafIds
    Set mLoeIds = loeIds
    Set mTopoOrder = topoOrder
    Set mAnalyticsPredsById = analyticsPredsById
    Set mAnalyticsChildrenById = analyticsChildrenById
    Set mPredLagBySuccPred = predLagBySuccPred
    Set mPredTypeBySuccPred = predTypeBySuccPred
    Set mIdToWbs = idToWbs
    Set mNodeIndexById = nodeIndexById

End Sub

' FR: Initialise les aretes Core completes, y compris diagnostics et expansion parent.
' EN: Initializes complete Core edges, including diagnostics and parent expansion.
Friend Sub InitializeCoreIndexedArrays( _
    ByRef predOffsets As Variant, _
    ByRef predNodes As Variant, _
    ByRef predIds As Variant, _
    ByRef predTypes As Variant, _
    ByRef predLags As Variant, _
    ByRef predSummarySourceIds As Variant)

    mCorePredOffsets = predOffsets
    mCorePredNodes = predNodes
    mCorePredIds = predIds
    mCorePredTypes = predTypes
    mCorePredLags = predLags
    mCorePredSummarySourceIds = predSummarySourceIds

End Sub

' FR: Initialise les arrays plats et compteurs de la representation compilee.
' EN: Initializes the compiled representation's flat arrays and counters.
Friend Sub InitializeIndexedArrays( _
    ByRef nodeIds As Variant, _
    ByRef rowsByNode As Variant, _
    ByRef parentByNode As Variant, _
    ByRef predOffsets As Variant, _
    ByRef predNodes As Variant, _
    ByRef predTypes As Variant, _
    ByRef predLags As Variant, _
    ByRef topoNodes As Variant, _
    ByRef componentByNode As Variant, _
    ByVal edgeCount As Long, _
    ByVal componentCount As Long, _
    ByVal isValid As Boolean)

    mNodeIds = nodeIds
    mRowsByNode = rowsByNode
    mParentByNode = parentByNode
    mPredOffsets = predOffsets
    mPredNodes = predNodes
    mPredTypes = predTypes
    mPredLags = predLags
    mTopoNodes = topoNodes
    mComponentByNode = componentByNode

    mNodeCount = mRowById.Count
    mLeafCount = mValidLeafIds.Count
    mEdgeCount = edgeCount
    mComponentCount = componentCount
    mIsValid = isValid

End Sub

' FR: Expose l'index ID -> ligne compile sans le reconstruire.
' EN: Exposes the compiled ID-to-row index without rebuilding it.
Public Property Get RowById() As Object
    Set RowById = mRowById
End Property

' FR: Expose l'ensemble des IDs parents compile.
' EN: Exposes the compiled parent-ID set.
Public Property Get ParentIds() As Object
    Set ParentIds = mParentIds
End Property

' FR: Expose la parenté directe compilee.
' EN: Exposes compiled direct parentage.
Public Property Get DirectChildrenById() As Object
    Set DirectChildrenById = mDirectChildrenById
End Property

' FR: Expose l'adjacence successeurs utilisee par le Core.
' EN: Exposes the successor adjacency used by Core.
Public Property Get ChildrenByPred() As Object
    Set ChildrenByPred = mChildrenByPred
End Property

' FR: Expose les feuilles non-LOE participant au reseau.
' EN: Exposes non-LOE leaves participating in the network.
Public Property Get ValidLeafIds() As Object
    Set ValidLeafIds = mValidLeafIds
End Property

' FR: Expose les IDs LOE exclus de la topologie standard.
' EN: Exposes LOE IDs excluded from standard topology.
Public Property Get LoeIds() As Object
    Set LoeIds = mLoeIds
End Property

' FR: Expose l'ordre topologique canonique unique.
' EN: Exposes the single canonical topological order.
Public Property Get TopoOrder() As Collection
    Set TopoOrder = mTopoOrder
End Property

' FR: Expose les predecesseurs analytics compiles.
' EN: Exposes compiled analytics predecessors.
Public Property Get AnalyticsPredsById() As Object
    Set AnalyticsPredsById = mAnalyticsPredsById
End Property

' FR: Expose les successeurs analytics compiles.
' EN: Exposes compiled analytics successors.
Public Property Get AnalyticsChildrenById() As Object
    Set AnalyticsChildrenById = mAnalyticsChildrenById
End Property

' FR: Expose les lags des aretes compilees.
' EN: Exposes compiled edge lags.
Public Property Get PredLagBySuccPred() As Object
    Set PredLagBySuccPred = mPredLagBySuccPred
End Property

' FR: Expose les types FS/SS/FF des aretes compilees.
' EN: Exposes compiled FS/SS/FF edge types.
Public Property Get PredTypeBySuccPred() As Object
    Set PredTypeBySuccPred = mPredTypeBySuccPred
End Property

' FR: Expose l'index ID -> WBS compile pour les diagnostics.
' EN: Exposes the compiled ID-to-WBS diagnostic index.
Public Property Get IdToWbs() As Object
    Set IdToWbs = mIdToWbs
End Property

' FR: Expose l'index entier d'un ID a la frontiere du chemin chaud.
' EN: Exposes an ID integer index at the hot-path boundary.
Public Property Get NodeIndexById() As Object
    Set NodeIndexById = mNodeIndexById
End Property

' FR: Retourne les IDs ordonnes par index entier.
' EN: Returns IDs ordered by integer index.
Public Property Get NodeIds() As Variant
    NodeIds = mNodeIds
End Property

' FR: Retourne les lignes CALC ordonnees par index entier.
' EN: Returns CALC rows ordered by integer index.
Public Property Get RowsByNode() As Variant
    RowsByNode = mRowsByNode
End Property

' FR: Retourne les parents indexes par noeud.
' EN: Returns parent indexes by node.
Public Property Get ParentByNode() As Variant
    ParentByNode = mParentByNode
End Property

' FR: Retourne les offsets des predecesseurs dans les arrays plats.
' EN: Returns predecessor offsets into flat arrays.
Public Property Get PredOffsets() As Variant
    PredOffsets = mPredOffsets
End Property

' FR: Retourne les indices predecesseurs plats.
' EN: Returns flat predecessor indexes.
Public Property Get PredNodes() As Variant
    PredNodes = mPredNodes
End Property

' FR: Retourne les types d'aretes plats.
' EN: Returns flat edge types.
Public Property Get PredTypes() As Variant
    PredTypes = mPredTypes
End Property

' FR: Retourne les lags plats.
' EN: Returns flat edge lags.
Public Property Get PredLags() As Variant
    PredLags = mPredLags
End Property

' FR: Retourne les offsets des liens complets consommes par le Core.
' EN: Returns complete-link offsets consumed by Core.
Public Property Get CorePredOffsets() As Variant
    CorePredOffsets = mCorePredOffsets
End Property

' FR: Retourne les indices predecesseurs Core; zero signale un ID absent.
' EN: Returns Core predecessor indices; zero identifies a missing ID.
Public Property Get CorePredNodes() As Variant
    CorePredNodes = mCorePredNodes
End Property

' FR: Retourne les IDs predecesseurs Core pour diagnostics et frontieres.
' EN: Returns Core predecessor IDs for diagnostics and boundaries.
Public Property Get CorePredIds() As Variant
    CorePredIds = mCorePredIds
End Property

' FR: Retourne les types de liens Core sans normalisation supplementaire.
' EN: Returns Core link types without additional normalization.
Public Property Get CorePredTypes() As Variant
    CorePredTypes = mCorePredTypes
End Property

' FR: Retourne les lags Core.
' EN: Returns Core lags.
Public Property Get CorePredLags() As Variant
    CorePredLags = mCorePredLags
End Property

' FR: Retourne la source summary de chaque arete issue d'une expansion parent.
' EN: Returns each parent-expanded edge's summary source.
Public Property Get CorePredSummarySourceIds() As Variant
    CorePredSummarySourceIds = mCorePredSummarySourceIds
End Property

' FR: Retourne l'ordre topologique sous forme d'indices entiers.
' EN: Returns topological order as integer indexes.
Public Property Get TopoNodes() As Variant
    TopoNodes = mTopoNodes
End Property

' FR: Retourne la composante reseau de chaque noeud.
' EN: Returns the network component of each node.
Public Property Get ComponentByNode() As Variant
    ComponentByNode = mComponentByNode
End Property

' FR: Retourne le nombre de noeuds compiles.
' EN: Returns the compiled node count.
Public Property Get NodeCount() As Long
    NodeCount = mNodeCount
End Property

' FR: Retourne le nombre de feuilles du reseau.
' EN: Returns the network leaf count.
Public Property Get LeafCount() As Long
    LeafCount = mLeafCount
End Property

' FR: Retourne le nombre d'aretes atomiques compilees.
' EN: Returns the compiled atomic edge count.
Public Property Get EdgeCount() As Long
    EdgeCount = mEdgeCount
End Property

' FR: Retourne le nombre de composantes feuille compilees.
' EN: Returns the compiled leaf-component count.
Public Property Get ComponentCount() As Long
    ComponentCount = mComponentCount
End Property

' FR: Indique si la topologie compilee couvre toutes les feuilles.
' EN: Indicates whether compiled topology covers every leaf.
Public Property Get IsValid() As Boolean
    IsValid = mIsValid
End Property




