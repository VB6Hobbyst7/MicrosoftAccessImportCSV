VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Used to navigate through the database. See lecture on data-structures and graphs

Private pcGraph As cGraph
Private pvEdge As Collection 'of cNodes
Private psName As String
Private psParentName As String
Private pbImported As Boolean
Private poParentTable As TableDef
Private poTableDef As TableDef
Private piIndexNo As Integer 'The index within the graph class
Private poField2 As Field2
Private piCSVcolumnNo As Integer
Private pvBin As Collection 'of integers
Private piWalk As Integer
Private piCopyNo As Integer
Private psData As String

Public Property Get iWalk() As Integer
    iWalk = piWalk
End Property

Public Property Let iWalk(iValue As Integer)
    piWalk = iValue
End Property

Public Property Get sData() As String
    sData = psData
End Property

Public Property Let sData(sValue As String)
    psData = sValue
End Property

Public Property Get iCopyNo() As Integer
    iCopyNo = piCopyNo
End Property

Public Property Let iCopyNo(iValue As Integer)
    piCopyNo = iValue
End Property

Public Property Get bImported() As Boolean
    bImported = pbImported
End Property

Public Property Let bImported(bValue As Boolean)
    pbImported = bValue
End Property

Public Property Get vEdge() As Collection
    Set vEdge = pvEdge
End Property

Public Property Get oTableDef() As Variant
    Set oTableDef = poTableDef
End Property

Public Property Set oTableDef(oValue As TableDef)
    Set poTableDef = oValue
End Property

Public Property Get oParentTable() As Variant
    Set oParentTable = poParentTable
End Property

Public Property Set oParentTable(oValue As TableDef)
    Set poParentTable = oValue
End Property

Public Property Get oField2() As Variant
    Set oField2 = poField2
End Property

Public Property Set oField2(oValue As Field2)
    Set poField2 = oValue
End Property

Public Property Get sName() As String
    sName = psName
End Property

Public Property Let sName(sValue As String)
    psName = sValue
End Property

Public Property Get sParentName() As String
    sParentName = psParentName
End Property

Public Property Let sParentName(sValue As String)
    psParentName = sValue
End Property

Public Property Get iIndexNo() As Integer
    iIndexNo = piIndexNo
End Property

Public Property Let iIndexNo(sValue As Integer)
    piIndexNo = sValue
End Property

Public Property Get iCSVcolumnNo() As Integer
    iCSVcolumnNo = piCSVcolumnNo
End Property

Public Property Let iCSVcolumnNo(iValue As Integer)
    piCSVcolumnNo = iValue
End Property

Function GetEdge(iValue As Integer) As cNode
    GetEdge = pvEdge(iValue)
End Function

Function GetvBin() As Collection
    Set GetvBin = pvBin
End Function

Sub AddvBin(iValue As Integer)
    If pvBin Is Nothing Then
        Set pvBin = New Collection
    End If
    pvBin.Add iValue
End Sub

Sub ErasevBin()
    Dim i As Integer
    
    For i = 1 To pvBin.count
        pvBin.Remove 1
    Next i
End Sub

Function CountvBin() As Integer
    CountvBin = pvBin.count
End Function



Sub AddEdge(oValue As cNode)
    
    'Error handling
    If pvEdge Is Nothing Then
        Set pvEdge = New Collection
    End If
    
    If oValue Is Nothing Then Exit Sub
    
    pvEdge.Add oValue
End Sub

Sub RemoveEdge(iValue As Integer)
'iValue is index starting at 0 (not 1)
'For some reason collections say "item 1" but its index is actually 0

    'Error handling
    If pvEdge Is Nothing Then
        Set pvEdge = New Collection
    End If
    
    If iValue < (pvEdge.count - 1) Then Exit Sub
    If iValue > -1 Then Exit Sub
    If apEdge.count = 0 Then Exit Sub
    
    pvEdge.Remove (iValue)
End Sub

Sub InitializecNode(Optional oTable As TableDef, Optional oField As Field2)
    
    
    
    Stop
    Set pvBin = New Collection
    
    If Not oTable Is Nothing Then
        Set poTableDef = oTable
    End If
    
    If Not oField Is Nothing Then
        Set poField2 = oField
    End If
    
End Sub
