Attribute VB_Name = "mod_GanttTestAnalyticsSnapshot"
Option Explicit

'===============================================================================
' MODULE : mod_GanttTestAnalyticsSnapshot
' DOMAINE / DOMAIN : Gantt TEST
'
' FR
' Possede la projection runtime coherente du planning TEST, indexee par ID.
' Reunit les resultats temporels du Core et les analytics calculees sur le meme
' dataset. Ne persiste aucune valeur dans WBS ou CALC.
'
' EN
' Owns the coherent runtime TEST planning projection indexed by ID.
' Combines Core temporal results with analytics computed from the same dataset.
' Persists no value to WBS or CALC.
'===============================================================================

Private gProjectionById As Object

' FR: Remplace atomiquement le snapshot TEST par une copie des resultats fournis.
' EN: Atomically replaces the TEST snapshot with a copy of the supplied results.
Public Sub GanttTestAnalyticsSnapshot_Set( _
    ByVal resultById As Object, _
    ByVal analyticsById As Object)

    Dim perfScope As clsPerfScope
    Dim projection As Object
    Dim idKey As Variant
    Dim resultData As Variant
    Dim analyticsData As Variant

    Set perfScope = Profiler_BeginScope("GanttTestAnalyticsSnapshot_Build", "Gantt TEST")
    Set projection = CreateObject("Scripting.Dictionary")

    If resultById Is Nothing Then
        Set gProjectionById = projection
        Exit Sub
    End If

    For Each idKey In resultById.Keys
        resultData = resultById(CStr(idKey))
        analyticsData = Array(Empty, Empty, vbNullString, vbNullString)
        If Not analyticsById Is Nothing Then
            If analyticsById.Exists(CStr(idKey)) Then analyticsData = analyticsById(CStr(idKey))
        End If

        projection(CStr(idKey)) = Array( _
            resultData(0), resultData(1), resultData(2), resultData(3), _
            resultData(4), resultData(5), resultData(6), _
            analyticsData(0), analyticsData(1), analyticsData(2), analyticsData(3))
    Next idKey

    Set gProjectionById = projection

End Sub

' FR: Efface le snapshot runtime lorsque le mode TEST prend fin.
' EN: Clears the runtime snapshot when TEST mode ends.
Public Sub GanttTestAnalyticsSnapshot_Clear()
    Set gProjectionById = Nothing
End Sub

' FR: Retourne le snapshot courant en lecture seule par convention interne.
' EN: Returns the current snapshot, treated as read-only by internal convention.
Public Function GanttTestAnalyticsSnapshot_GetProjectionById() As Object
    If gProjectionById Is Nothing Then
        Set GanttTestAnalyticsSnapshot_GetProjectionById = CreateObject("Scripting.Dictionary")
    Else
        Set GanttTestAnalyticsSnapshot_GetProjectionById = gProjectionById
    End If
End Function

' FR: Indique si un snapshot analytique TEST coherent est disponible.
' EN: Returns whether a coherent TEST analytics snapshot is available.
Public Function GanttTestAnalyticsSnapshot_IsCurrent() As Boolean
    GanttTestAnalyticsSnapshot_IsCurrent = Not gProjectionById Is Nothing
End Function

' FR: Retourne le Total Float simule d'une tache, ou Empty si absent.
' EN: Returns a task's simulated Total Float, or Empty when unavailable.
Public Function GanttTestAnalyticsSnapshot_GetTotalFloat(ByVal taskId As String) As Variant
    If GanttTestAnalyticsSnapshot_HasTask(taskId) Then _
        GanttTestAnalyticsSnapshot_GetTotalFloat = gProjectionById(taskId)(7)
End Function

' FR: Retourne le Free Float simule d'une tache, ou Empty si absent.
' EN: Returns a task's simulated Free Float, or Empty when unavailable.
Public Function GanttTestAnalyticsSnapshot_GetFreeFloat(ByVal taskId As String) As Variant
    If GanttTestAnalyticsSnapshot_HasTask(taskId) Then _
        GanttTestAnalyticsSnapshot_GetFreeFloat = gProjectionById(taskId)(8)
End Function

' FR: Retourne la classification Critical Path ou Longest Path du snapshot TEST.
' EN: Returns the TEST snapshot's Critical Path or Longest Path classification.
Public Function GanttTestAnalyticsSnapshot_GetPathValue( _
    ByVal taskId As String, _
    ByVal pathColumnName As String) As String

    If Not GanttTestAnalyticsSnapshot_HasTask(taskId) Then Exit Function

    Select Case pathColumnName
        Case "Critical Path"
            GanttTestAnalyticsSnapshot_GetPathValue = CStr(gProjectionById(taskId)(9))
        Case "Longest Path"
            GanttTestAnalyticsSnapshot_GetPathValue = CStr(gProjectionById(taskId)(10))
    End Select

End Function

' FR: Verifie en lecture seule la presence d'une tache dans le snapshot courant.
' EN: Read-only check for a task in the current snapshot.
Private Function GanttTestAnalyticsSnapshot_HasTask(ByVal taskId As String) As Boolean
    If gProjectionById Is Nothing Then Exit Function
    GanttTestAnalyticsSnapshot_HasTask = gProjectionById.Exists(taskId)
End Function
