Attribute VB_Name = "mod_TaskTypeRules"
Option Explicit

'===============================================================================
' MODULE : mod_TaskTypeRules
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Possede la normalisation et les classifications partagees LOE, Milestone et Task Type exclu.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Owns shared LOE, Milestone and excluded Task Type normalization and classification.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : TaskTypeRules_IsMilestoneValue, TaskTypeRules_IsLevelOfEffortValue, TaskTypeRules_IsMilestoneRow, TaskTypeRules_IsLevelOfEffortRow
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'==============================================================
' Neutral task-type classification rules shared by calculation,
' analytics and rendering domains. This module has no worksheet,
' shape, message or mutation responsibility.
'==============================================================

'------------------------------------------------------------------------------
' FR: Indique si une valeur de type de tache designe une Milestone.
' EN: Returns whether a task-type value identifies a Milestone.
'------------------------------------------------------------------------------
Public Function TaskTypeRules_IsMilestoneValue(ByVal rawValue As Variant) As Boolean

    Dim normalizedValue As String

    normalizedValue = TaskTypeRules_NormalizeComparisonValue(rawValue)

    TaskTypeRules_IsMilestoneValue = _
        (normalizedValue = "MILESTONE") Or _
        (normalizedValue = "MS") Or _
        (normalizedValue = "JALON")

End Function

'------------------------------------------------------------------------------
' FR: Indique si une valeur de type de tache designe une Level of Effort.
' EN: Returns whether a task-type value identifies a Level of Effort task.
'------------------------------------------------------------------------------
Public Function TaskTypeRules_IsLevelOfEffortValue(ByVal rawValue As Variant) As Boolean

    Dim normalizedValue As String

    normalizedValue = TaskTypeRules_NormalizeComparisonValue(rawValue)

    TaskTypeRules_IsLevelOfEffortValue = _
        (normalizedValue = "LOE") Or _
        (normalizedValue = "LEVEL OF EFFORT") Or _
        (normalizedValue = "LEVEL OF EFFORT TASK")

End Function

'------------------------------------------------------------------------------
' FR: Classe une ligne de dataset comme Milestone a partir de la colonne Task Type.
' EN: Classifies a dataset row as a Milestone from its Task Type column.
'------------------------------------------------------------------------------
Public Function TaskTypeRules_IsMilestoneRow( _
    ByRef dataArr As Variant, _
    ByVal columnMap As Object, _
    ByVal dataRow As Long) As Boolean

    On Error GoTo SafeExit

    If columnMap Is Nothing Then Exit Function
    If Not columnMap.Exists("Task Type") Then Exit Function
    If dataRow < LBound(dataArr, 1) Then Exit Function
    If dataRow > UBound(dataArr, 1) Then Exit Function

    TaskTypeRules_IsMilestoneRow = _
        TaskTypeRules_IsMilestoneValue(dataArr(dataRow, columnMap("Task Type")))

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Classe une ligne de dataset comme Level of Effort a partir de Task Type.
' EN: Classifies a dataset row as Level of Effort from its Task Type column.
'------------------------------------------------------------------------------
Public Function TaskTypeRules_IsLevelOfEffortRow( _
    ByRef dataArr As Variant, _
    ByVal columnMap As Object, _
    ByVal dataRow As Long) As Boolean

    On Error GoTo SafeExit

    If columnMap Is Nothing Then Exit Function
    If Not columnMap.Exists("Task Type") Then Exit Function
    If dataRow < LBound(dataArr, 1) Then Exit Function
    If dataRow > UBound(dataArr, 1) Then Exit Function

    TaskTypeRules_IsLevelOfEffortRow = _
        TaskTypeRules_IsLevelOfEffortValue(dataArr(dataRow, columnMap("Task Type")))

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Normalise uniquement la syntaxe necessaire aux comparaisons de type.
' EN: Normalizes only the syntax required for task-type comparisons.
'------------------------------------------------------------------------------
Private Function TaskTypeRules_NormalizeComparisonValue(ByVal rawValue As Variant) As String

    Dim normalizedValue As String

    normalizedValue = UCase$(Trim$(CStr(rawValue)))
    normalizedValue = Replace$(normalizedValue, "-", " ")
    normalizedValue = Replace$(normalizedValue, "_", " ")

    Do While InStr(1, normalizedValue, "  ", vbBinaryCompare) > 0
        normalizedValue = Replace$(normalizedValue, "  ", " ")
    Loop

    TaskTypeRules_NormalizeComparisonValue = normalizedValue

End Function
