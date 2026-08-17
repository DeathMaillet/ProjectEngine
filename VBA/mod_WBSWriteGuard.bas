Attribute VB_Name = "mod_WBSWriteGuard"
Option Explicit

'===============================================================================
' MODULE : mod_WBSWriteGuard
' DOMAINE / DOMAIN : WBS Mutation Guard
'
' FR
' Possede les scopes tokenises qui autorisent temporairement les ecritures moteur dans WBS.
' Ne realise aucune ecriture et ne ferme jamais un scope non possede.
'
' EN
' Owns tokenized scopes that temporarily authorize engine writes to WBS.
' Performs no write and never closes an unowned scope.
'
' CONTRATS / CONTRACTS : OpenAuthorizedWBSWriteScope, CloseAuthorizedWBSWriteScope, BeginAuthorizedWBSWrite, EndAuthorizedWBSWrite, IsAuthorizedWBSWriteActive, GetAuthorizedWBSWriteSource, IsAuthorizedWBSColumn
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gWBSWriteScopes As Collection
Private gNextWBSWriteScopeToken As Long

'------------------------------------------------------------------------------
' FR: Ouvre un scope WBS possede par l'appelant et retourne son token unique.
' EN: Opens a caller-owned WBS scope and returns its unique token.
'------------------------------------------------------------------------------
Public Function OpenAuthorizedWBSWriteScope( _
    ByVal sourceName As String, _
    ByVal allowedColumnNames As Variant) As Long

    Dim allowedColumns As Object
    Dim scopeFrame As Object
    Dim i As Long

    EnsureWBSWriteGuardScopes

    gNextWBSWriteScopeToken = gNextWBSWriteScopeToken + 1
    If gNextWBSWriteScopeToken <= 0 Then gNextWBSWriteScopeToken = 1

    Set allowedColumns = CreateObject("Scripting.Dictionary")
    For i = LBound(allowedColumnNames) To UBound(allowedColumnNames)
        allowedColumns(CStr(allowedColumnNames(i))) = True
    Next i

    Set scopeFrame = CreateObject("Scripting.Dictionary")
    scopeFrame("Token") = gNextWBSWriteScopeToken
    scopeFrame("Source") = CStr(sourceName)
    Set scopeFrame("AllowedColumns") = allowedColumns
    gWBSWriteScopes.Add scopeFrame

    OpenAuthorizedWBSWriteScope = gNextWBSWriteScopeToken

End Function

'------------------------------------------------------------------------------
' FR: Ferme uniquement le scope WBS identifie, dans l'ordre LIFO attendu.
' EN: Closes only the identified WBS scope in the expected LIFO order.
'------------------------------------------------------------------------------
Public Sub CloseAuthorizedWBSWriteScope(ByVal scopeToken As Long)

    Dim scopeFrame As Object
    Dim activeToken As Long

    If scopeToken = 0 Then Exit Sub

    EnsureWBSWriteGuardScopes

    If gWBSWriteScopes.Count = 0 Then
        Err.Raise vbObjectError + 9620, "CloseAuthorizedWBSWriteScope", _
            "No active WBS write scope matches token " & CStr(scopeToken) & "."
    End If

    Set scopeFrame = gWBSWriteScopes(gWBSWriteScopes.Count)
    activeToken = CLng(scopeFrame("Token"))

    If activeToken <> scopeToken Then
        Err.Raise vbObjectError + 9621, "CloseAuthorizedWBSWriteScope", _
            "WBS write scopes must close in LIFO order. Active token=" & _
            CStr(activeToken) & ", requested token=" & CStr(scopeToken) & "."
    End If

    gWBSWriteScopes.Remove gWBSWriteScopes.Count

End Sub

'------------------------------------------------------------------------------
' FR: Ouvre le cycle historique Authorized WBSWrite sur la pile tokenisee.
' EN: Opens the legacy Authorized WBSWrite cycle on the tokenized stack.
'------------------------------------------------------------------------------
Public Sub BeginAuthorizedWBSWrite( _
    ByVal sourceName As String, _
    ByVal allowedColumnNames As Variant)

    Dim ignoredToken As Long

    ignoredToken = OpenAuthorizedWBSWriteScope(sourceName, allowedColumnNames)

End Sub

'------------------------------------------------------------------------------
' FR: Ferme le scope historique courant pour compatibilite avec les appelants existants.
' EN: Closes the current legacy scope for compatibility with existing callers.
'------------------------------------------------------------------------------
Public Sub EndAuthorizedWBSWrite()

    Dim scopeFrame As Object

    EnsureWBSWriteGuardScopes
    If gWBSWriteScopes.Count = 0 Then Exit Sub

    Set scopeFrame = gWBSWriteScopes(gWBSWriteScopes.Count)
    CloseAuthorizedWBSWriteScope CLng(scopeFrame("Token"))

End Sub

'------------------------------------------------------------------------------
' FR: Indique si Authorized WBSWrite Active est vrai pour le contexte courant.
' EN: Returns whether Authorized WBSWrite Active is true for the current context.
'------------------------------------------------------------------------------
Public Function IsAuthorizedWBSWriteActive() As Boolean

    EnsureWBSWriteGuardScopes
    IsAuthorizedWBSWriteActive = (gWBSWriteScopes.Count > 0)

End Function

'------------------------------------------------------------------------------
' FR: Retourne Authorized WBSWrite Source depuis le contexte WBS write guard.
' EN: Returns Authorized WBSWrite Source from the WBS write guard context.
'------------------------------------------------------------------------------
Public Function GetAuthorizedWBSWriteSource() As String

    Dim scopeFrame As Object

    EnsureWBSWriteGuardScopes
    If gWBSWriteScopes.Count = 0 Then Exit Function

    Set scopeFrame = gWBSWriteScopes(gWBSWriteScopes.Count)
    GetAuthorizedWBSWriteSource = CStr(scopeFrame("Source"))

End Function

'------------------------------------------------------------------------------
' FR: Indique si Authorized WBSColumn est vrai pour le contexte courant.
' EN: Returns whether Authorized WBSColumn is true for the current context.
'------------------------------------------------------------------------------
Public Function IsAuthorizedWBSColumn(ByVal columnName As String) As Boolean

    Dim scopeFrame As Object
    Dim allowedColumns As Object

    EnsureWBSWriteGuardScopes
    If gWBSWriteScopes.Count = 0 Then Exit Function

    Set scopeFrame = gWBSWriteScopes(gWBSWriteScopes.Count)
    Set allowedColumns = scopeFrame("AllowedColumns")
    IsAuthorizedWBSColumn = allowedColumns.Exists(CStr(columnName))

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree la pile de scopes WBS tokenises.
' EN: Ensures or creates the tokenized WBS scope stack.
'------------------------------------------------------------------------------
Private Sub EnsureWBSWriteGuardScopes()

    If gWBSWriteScopes Is Nothing Then Set gWBSWriteScopes = New Collection

End Sub
