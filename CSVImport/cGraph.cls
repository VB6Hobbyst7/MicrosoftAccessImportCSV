VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'Architecture used to navigate through the tables

Private pvcNodes As Collection 'of cNodes
Private pDatabase As Database
Private piNodeCt As Integer

Public Property Get iNodeCt() As Integer
    iNodeCt = piNodeCt
End Property

Public Property Get vcNodes() As Collection
    Set vcNodes = pvcNodes
End Property


Function GetNode(iValue As Integer) As cNode
    Set GetNode = pvcNodes(iValue)
End Function

Sub ConstructGraph(MyDatabase As Database)
'Constructs a graph model from the Microsoft Access database data. Used to navigate between fields

    Dim aRelations As Relations
    Dim aTableDefs As TableDefs
    Dim aFields As Fields
    Dim oTableDefcNode As cNode
    Dim oFieldcNode As cNode
    Dim oTableDef As TableDef
    Dim sName As String
    Dim oField As Field2
    Dim sName1 As String
    Dim sName2 As String
    Dim ocNode As cNode
    Dim ocNode1 As cNode
    Dim ocNode2 As cNode
    Dim oRelation As Relation
    Dim oRelFields As Fields
    Dim sTableName1 As String
    Dim sTableName2 As String
    
    'Construct all the nodes and connect the edges
    Set pDatabase = Application.CurrentDb
    Set pvcNodes = New Collection


    Set aTableDefs = MyDatabase.TableDefs
    For Each oTableDef In aTableDefs
        
        'Create the node for oTableDef
        Set oTableDefcNode = New cNode
        oTableDefcNode.iWalk = 1
        Set oTableDefcNode.oTableDef = oTableDef
        sName = oTableDef.Name
        oTableDefcNode.sName = sName
        
        'Skip over the "msys" tables
        If (InStr(sName, "MSys") = 0) Or (InStr(sName, "MSys") > 4) Then
        
            'Create the cNodes for the table children
            For Each oField In oTableDef.Fields
                
                'Create the oField cNode
                Set oFieldcNode = New cNode
                oFieldcNode.iWalk = 1
                Set oFieldcNode.oField2 = oField
                Set oFieldcNode.oParentTable = oTableDef
                sName = oField.Name
                oFieldcNode.sName = sName
                oFieldcNode.sParentName = oTableDef.Name
                
                'Cross-reference the table and field nodes
                oFieldcNode.AddEdge oTableDefcNode
                oTableDefcNode.AddEdge oFieldcNode
                
                'Add the fieldnode node to the collection
                pvcNodes.Add oFieldcNode
                oFieldcNode.iIndexNo = pvcNodes.count
                
            Next oField
            
            'add the table to the collection
            pvcNodes.Add oTableDefcNode
            oTableDefcNode.iIndexNo = pvcNodes.count
        End If
    Next oTableDef

    'Add the relations to the graph
    Set aRelations = MyDatabase.Relations
    For Each oRelation In aRelations
        'Skip over any internal "Msys" relations
        If (InStr(sName, "MSys") = 0) Or (InStr(sName, "MSys") > 4) Then
            
            'Get the names of the two related fields within the relation
            sTableName1 = oRelation.Table
            sTableName2 = oRelation.ForeignTable
            Set oRelFields = oRelation.Fields
            For Each oField In oRelFields
                sName1 = oField.Name
                sName2 = oField.ForeignName
            Next oField
            
            'Search your existing graph for both nodes
            For Each ocNode In pvcNodes
                If (ocNode.sName = sName1) And (ocNode.sParentName = sTableName1) Then
                    Set ocNode1 = ocNode
                End If
            
                If (ocNode.sName = sName2) And (ocNode.sParentName = sTableName2) Then
                    Set ocNode2 = ocNode
                End If
            
                'Abort once nodes are found
                'If (Not ocNode1 Is Nothing) And (Not ocNode2 Is Nothing) Then
                    'Exit For
                'End If
            Next ocNode
            
            'Cross reference both fields in graph
            ocNode1.AddEdge ocNode2
            ocNode2.AddEdge ocNode1
        End If
    Next oRelation
    
    'populate the number of nodes
    piNodeCt = pvcNodes.count
End Sub
