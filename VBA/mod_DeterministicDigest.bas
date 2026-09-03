Attribute VB_Name = "mod_DeterministicDigest"
Option Explicit

'===============================================================================
' MODULE : mod_DeterministicDigest
' DOMAINE / DOMAIN : Deterministic hashing
'
' FR
' Produit des digests SHA-256 stables a partir du texte UTF-16LE natif de VBA.
' L'implementation utilise Windows CNG et ne depend ni de la locale ni du bitness
' d'Office.
'
' EN
' Produces stable SHA-256 digests from VBA native UTF-16LE text. The
' implementation uses Windows CNG and is independent from locale and Office
' bitness.
'
' CONTRATS / CONTRACTS : DeterministicDigest_SHA256Hex
'===============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
        ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, _
        ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _
        ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, _
        ByVal pbOutput As LongPtr, ByVal cbOutput As Long, _
        ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, _
        ByVal pbHashObject As LongPtr, ByVal cbHashObject As Long, _
        ByVal pbSecret As LongPtr, ByVal cbSecret As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr, ByVal pbInput As LongPtr, _
        ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr, ByVal pbOutput As LongPtr, _
        ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr) As Long
    Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
        ByRef phAlgorithm As Long, ByVal pszAlgId As Long, _
        ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptGetProperty Lib "bcrypt.dll" ( _
        ByVal hObject As Long, ByVal pszProperty As Long, _
        ByVal pbOutput As Long, ByVal cbOutput As Long, _
        ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptCreateHash Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, ByRef phHash As Long, _
        ByVal pbHashObject As Long, ByVal cbHashObject As Long, _
        ByVal pbSecret As Long, ByVal cbSecret As Long, _
        ByVal dwFlags As Long) As Long
    Private Declare Function BCryptHashData Lib "bcrypt.dll" ( _
        ByVal hHash As Long, ByVal pbInput As Long, _
        ByVal cbInput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptFinishHash Lib "bcrypt.dll" ( _
        ByVal hHash As Long, ByVal pbOutput As Long, _
        ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
    Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" ( _
        ByVal hHash As Long) As Long
    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
#End If

Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"
Private Const BCRYPT_OBJECT_LENGTH As String = "ObjectLength"
Private Const BCRYPT_HASH_LENGTH As String = "HashDigestLength"

Public Function DeterministicDigest_SHA256Hex(ByVal textValue As String) As String

#If VBA7 Then
    Dim algorithmHandle As LongPtr
    Dim hashHandle As LongPtr
#Else
    Dim algorithmHandle As Long
    Dim hashHandle As Long
#End If
    Dim algorithmName As String
    Dim objectLengthProperty As String
    Dim hashLengthProperty As String
    Dim hashObjectLength As Long
    Dim hashLength As Long
    Dim resultLength As Long
    Dim textByteLength As Long
    Dim hashObject() As Byte
    Dim digest() As Byte
    Dim status As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim i As Long
    Dim hexValue As String

    On Error GoTo ErrHandler

    algorithmName = BCRYPT_SHA256_ALGORITHM
    objectLengthProperty = BCRYPT_OBJECT_LENGTH
    hashLengthProperty = BCRYPT_HASH_LENGTH

    status = BCryptOpenAlgorithmProvider(algorithmHandle, StrPtr(algorithmName), 0, 0)
    DeterministicDigest_RequireSuccess status, "BCryptOpenAlgorithmProvider"

    status = BCryptGetProperty(algorithmHandle, StrPtr(objectLengthProperty), _
        VarPtr(hashObjectLength), 4, resultLength, 0)
    DeterministicDigest_RequireSuccess status, "BCryptGetProperty(ObjectLength)"

    status = BCryptGetProperty(algorithmHandle, StrPtr(hashLengthProperty), _
        VarPtr(hashLength), 4, resultLength, 0)
    DeterministicDigest_RequireSuccess status, "BCryptGetProperty(HashDigestLength)"

    If hashObjectLength <= 0 Or hashLength <> 32 Then
        Err.Raise vbObjectError + 2711, "DeterministicDigest_SHA256Hex", _
            "Windows CNG returned an invalid SHA-256 contract."
    End If

    ReDim hashObject(0 To hashObjectLength - 1)
    ReDim digest(0 To hashLength - 1)

    status = BCryptCreateHash(algorithmHandle, hashHandle, VarPtr(hashObject(0)), _
        hashObjectLength, 0, 0, 0)
    DeterministicDigest_RequireSuccess status, "BCryptCreateHash"

    textByteLength = LenB(textValue)
    If textByteLength > 0 Then
        status = BCryptHashData(hashHandle, StrPtr(textValue), textByteLength, 0)
        DeterministicDigest_RequireSuccess status, "BCryptHashData"
    End If

    status = BCryptFinishHash(hashHandle, VarPtr(digest(0)), hashLength, 0)
    DeterministicDigest_RequireSuccess status, "BCryptFinishHash"

    For i = 0 To hashLength - 1
        hexValue = hexValue & Right$("0" & Hex$(digest(i)), 2)
    Next i

    DeterministicDigest_SHA256Hex = UCase$(hexValue)

CleanExit:
    If hashHandle <> 0 Then BCryptDestroyHash hashHandle
    If algorithmHandle <> 0 Then BCryptCloseAlgorithmProvider algorithmHandle, 0
    If errorNumber <> 0 Then
        Err.Raise errorNumber, "DeterministicDigest_SHA256Hex", errorDescription
    End If
    Exit Function

ErrHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Resume CleanExit

End Function

Private Sub DeterministicDigest_RequireSuccess( _
    ByVal status As Long, _
    ByVal operationName As String)

    If status <> 0 Then
        Err.Raise vbObjectError + 2712, "DeterministicDigest_SHA256Hex", _
            operationName & " failed with NTSTATUS " & CStr(status) & "."
    End If

End Sub
