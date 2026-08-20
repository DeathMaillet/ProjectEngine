VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGanttDependencyRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' MODULE : clsGanttDependencyRoute
' DOMAIN : Gantt dependency logical model
'
' Owns one canonical dependency route independently from Excel Shapes. Segment
' records use Gantt coordinates in points and retain terminal-arrow ownership.
'===============================================================================

Private mLinkId As String
Private mPredId As String
Private mSuccId As String
Private mLinkType As String
Private mLag As Double
Private mSegments As Collection
Private mMinX As Double
Private mMinY As Double
Private mMaxX As Double
Private mMaxY As Double
Private mHasBounds As Boolean
Private mVisible As Boolean

Public Sub Initialize( _
    ByVal linkId As String, _
    ByVal predId As String, _
    ByVal succId As String, _
    ByVal linkType As String, _
    ByVal lag As Double)

    mLinkId = linkId
    mPredId = predId
    mSuccId = succId
    mLinkType = UCase$(Trim$(linkType))
    mLag = lag
    mVisible = True
    Set mSegments = New Collection

End Sub

Public Sub AddSegment( _
    ByVal x1 As Double, _
    ByVal y1 As Double, _
    ByVal x2 As Double, _
    ByVal y2 As Double, _
    ByVal withArrow As Boolean)

    Dim segment As Object

    If mSegments Is Nothing Then Set mSegments = New Collection

    If Abs(x2 - x1) < 0.1 And Abs(y2 - y1) < 0.1 Then
        If withArrow And mSegments.Count > 0 Then
            Set segment = mSegments(mSegments.Count)
            segment("Arrow") = True
        End If
        Exit Sub
    End If

    Set segment = CreateObject("Scripting.Dictionary")
    segment("X1") = x1
    segment("Y1") = y1
    segment("X2") = x2
    segment("Y2") = y2
    segment("Arrow") = withArrow
    mSegments.Add segment

    IncludePoint x1, y1
    IncludePoint x2, y2

End Sub

Private Sub IncludePoint(ByVal xValue As Double, ByVal yValue As Double)

    If Not mHasBounds Then
        mMinX = xValue
        mMaxX = xValue
        mMinY = yValue
        mMaxY = yValue
        mHasBounds = True
        Exit Sub
    End If

    If xValue < mMinX Then mMinX = xValue
    If xValue > mMaxX Then mMaxX = xValue
    If yValue < mMinY Then mMinY = yValue
    If yValue > mMaxY Then mMaxY = yValue

End Sub

Public Property Get LinkId() As String
    LinkId = mLinkId
End Property

Public Property Get PredId() As String
    PredId = mPredId
End Property

Public Property Get SuccId() As String
    SuccId = mSuccId
End Property

Public Property Get LinkType() As String
    LinkType = mLinkType
End Property

Public Property Get Lag() As Double
    Lag = mLag
End Property

Public Property Get Segments() As Collection
    Set Segments = mSegments
End Property

Public Property Get SegmentCount() As Long
    If Not mSegments Is Nothing Then SegmentCount = mSegments.Count
End Property

Public Property Get MinX() As Double
    MinX = mMinX
End Property

Public Property Get MinY() As Double
    MinY = mMinY
End Property

Public Property Get MaxX() As Double
    MaxX = mMaxX
End Property

Public Property Get MaxY() As Double
    MaxY = mMaxY
End Property

Public Property Get HasBounds() As Boolean
    HasBounds = mHasBounds
End Property

Public Property Get Visible() As Boolean
    Visible = mVisible
End Property

Public Property Let Visible(ByVal value As Boolean)
    mVisible = value
End Property
