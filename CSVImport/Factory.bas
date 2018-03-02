Attribute VB_Name = "Factory"
Option Compare Database

Function NewcDataBase() As cDataBase
'Creates a NewImport object
    Set NewcDataBase = New cDataBase
    NewcDataBase.InitializecDatabase
End Function

Function NewcNode(oData As Object) As cNode
'Creates a NewcTable object
    If oData Is Nothing Then Exit Function
    If TypeOf oData Is Field2 Then
        Set NewcNode = New cNode
        NewcNode.InitializecNode Nothing, oData
        Exit Function
    End If
    If TypeOf oData Is TableDef Then
        Set NewcNode = New cNode
        NewcNode.InitializecNode oData, Nothing
        Exit Function
    Else
        'YOU HAVE AN OBJECT THAT IS NOT ON THIS LIST!!!!
        Stop
        'YOU HAVE AN OBJECT THAT IS NOT ON THIS LIST!!!!
        'YOU HAVE AN OBJECT THAT IS NOT ON THIS LIST!!!!
        'YOU HAVE AN OBJECT THAT IS NOT ON THIS LIST!!!!
    End If
End Function

Function NewcGraph(MyDatabase As Database) As cGraph
    Set NewcGraph = New cGraph
    NewcGraph.ConstructGraph MyDatabase
End Function
    
