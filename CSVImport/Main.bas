Attribute VB_Name = "Main"
Option Compare Database
Option Explicit
'Main code to run your CSV import

Sub Main()
'Imports data into database using a CSV file
    Dim MyDatabase As Database
    Dim MyGraph As cGraph
    Dim Path As String
    Dim FSO As Object 'File Scripting Object
    Dim oMyCSVFile As Object
    Dim sMyCSVContent As String
    Dim vMyCSVRows As Variant
    Dim x As Integer
    Dim sTemp As String
    Dim vCurrentRow As Variant
    Dim Headers As Variant
    Dim Col As String
    Dim y As Integer
    Dim vCSVColNodes As Collection
    Dim z As Integer
    Dim Temp As Object
    Dim CurrentNode As cNode
    Dim a As Integer
    
    Set MyDatabase = Application.CurrentDb
    Set MyGraph = NewcGraph(MyDatabase)
    Set vCSVColNodes = New Collection
      
'NOTE: 'Rework:' is used to redo the program as you add changes to your database
Rework:
    
    'Open up the CSV file
    Path = "" 'add your csv filepath here
    
    'Error Handling
    If IsEmpty(Path) Or Path = "" Then
        MsgBox ("no path found")
        Exit Sub
    End If
    
    If MyGraph Is Nothing Then
        MsgBox ("no database")
        Exit Sub
    End If
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Const ForReading = 1
    Set oMyCSVFile = FSO.OpenTextFile(Path, ForReading, True)
    sMyCSVContent = oMyCSVFile.readall
    vMyCSVRows = Split(sMyCSVContent, vbCrLf)
    
    For a = 0 To UBound(vMyCSVRows)
        'Clean the data, split row into columns
        sTemp = Replace(vMyCSVRows(a), "﻿", "")
        vCurrentRow = Split(sTemp, ",")
        
        'Construct what to do from using the headers
        If a = 0 Then
            
            'Find the nodes in graph
            For y = 0 To UBound(vCurrentRow)
                Col = vCurrentRow(y)
                'MyGraph.vcNodes (z)
                
                For z = 1 To MyGraph.vcNodes.count
                    Set CurrentNode = MyGraph.vcNodes(z)
                    If Col = CurrentNode.sName Then
                        CurrentNode.iCSVcolumnNo = y
                        'if you have a field or a table
                        vCSVColNodes.Add CurrentNode
                        Exit For
                    End If
                    
                    'if you can't find the node
                    If z = MyGraph.vcNodes.count And Col <> CurrentNode.sName Then
                        MsgBox "Node: " & Col & " is missing. Please add to database and rerun macro"
                        Stop
                        GoTo Rework
                    End If
                Next z
            Next y
            
            'Determine which fields belong to the same table
            Dim BinCt As Integer
            Dim Group As Collection
            Dim cNode As cNode
            Dim Parent As cNode
            Dim oField As Field2
            Dim sName As String
            Dim PrimaryEdge As cNode
            Dim SecondaryEdge As cNode
            Dim que As Collection
            Dim i As Integer
            Dim vArr As Collection
            Dim Bin As Collection
            Dim j As Integer
            Dim k As Integer
            Dim CopyNo As Integer
            Dim sFirst As String
            Dim sSecond As String
            Dim sThird As String
            Dim sFour As String
            Dim m As Integer
        
            
            Set Bin = New Collection
            
            'Construct the bins to execute the data
            'Go through headers, and mark which ones belong in the same bin
            For i = 1 To vCSVColNodes.count
                If i = 1 Then
                    Set Group = New Collection
                    Group.Add vCSVColNodes(1)
                    Bin.Add Group
                    'Record what bin the data goes into, record the total bins
                    vCSVColNodes(i).AddvBin 1
                    BinCt = BinCt + 1
                End If
                
                'Check your existing bins, if you find a matching table
                'then add it to the group. if they match, you need to mark that it was a match
                If i > 1 Then
                    sSecond = LCase(vCSVColNodes(i).sParentName)
                    sFour = LCase(vCSVColNodes(i).sName)
                    
                    For j = 1 To Bin.count
                    'Debug.Print Bin(j)(1).sParentName
                    'Debug.Print vCSVColNodes(i).sParentName
                    'Debug.Print "i=" & i
                    'Debug.Print "j=" & j
                    'Debug.Print "---"
                        sFirst = LCase((Bin(j)(1).sParentName))
                        
                        'If you have matching tables
                        If StrComp(sFirst, sSecond) = 0 Then
                            
                            'Check for duplicates
                            For k = 1 To Bin(j).count
                                sThird = LCase(Bin(j)(k).sName)
                                
                                If StrComp(sThird, sFour) = 0 Then
                                    'copy your collection with new data
                                    If CopyNo = 0 Then
                                        CopyNo = CopyNo + 1
                                        vCSVColNodes(i).iCopyNo = CopyNo
                                        'Bin(j)(k).iCopyNo = CopyNo
                                    Else
                                       vCSVColNodes(i).iCopyNo = vArr(k).iCopyNo
                                    End If
                                End If
                            Next k
                            
                            'Add item to bin
                            vCSVColNodes(i).AddvBin j
                            Bin(j).Add vCSVColNodes(i)
                            Exit For
                        End If
                        
                        'if you cant find a bin, you need to create one
                        If j = Bin.count And StrComp(LCase((Bin(j)(1).sParentName)), LCase(vCSVColNodes(i).sParentName)) <> 0 Then
                            Set Group = New Collection
                            Group.Add vCSVColNodes(i)
                            Bin.Add Group
                            vCSVColNodes(i).AddvBin (j + 1)
                            BinCt = BinCt + 1
                        End If
                    Next j
                End If
            Next i
            
            Dim ToCopy2 As Collection
            Dim Duplicate As Collection
            Dim NonDup As Collection
            Dim SubNonDup As Collection
            Dim Temp2 As Collection
            Dim Dup As Collection
            Dim Noder2 As cNode
            Dim OldBinCt As Integer
            Dim n As Integer
            Dim BinsRemoved As Integer
            
            Set ToCopy2 = New Collection
            Set Duplicate = New Collection
            Set NonDup = New Collection
            Set SubNonDup = New Collection
            
            'Handle Duplicates
            For z = 1 To Bin.count
                
                'Separate duplicates from non-duplicates
                For m = 1 To Bin(z).count
                    
                    'Bin either duplicate or non duplicate
                    If Bin(z)(m).iCopyNo <> 0 Then
                        OldBinCt = z
                        
                        'If duplicate, check whether it already exists in the duplicate bin
                        'If you have multiple duplicates, the duplicates collection would have more
                        'than one item. If you have repeating duplicates (repeating fields in the CVS file)
                        'then you will have two items within the duplicate collection
                        If Duplicate.count = 0 Then
                            Set Temp2 = New Collection
                            Temp2.Add Bin(z)(m)
                            Duplicate.Add Temp2
                        Else
                            For Each Dup In Duplicate
                                If Dup(1).iCopyNo = Bin(z)(m).iCopyNo Then
                                    Dup.Add Bin(z)(m)
                                Else
                                    Set Temp2 = New Collection
                                    Temp2.Add Bin(z)(m)
                                    Duplicate.Add Temp
                                End If
                            Next Dup
                        End If
                        'strip the index number
                        Bin(z)(m).ErasevBin
                    Else
                        NonDup.Add Bin(z)(m)
                    End If
                Next m
                
                Dim q As Integer
                
                'If Duplicate.count = 1 Then Stop
                'Rebuild vBin to account for duplicates
                For n = 1 To Duplicate.count
                    For q = 1 To Duplicate(n).count
                        If q = 1 Then
                            'One duplicate should be put back into the old list
                            Duplicate(n)(q).AddvBin OldBinCt
                        Else
                            BinCt = BinCt + 1
                            Duplicate(n)(q).AddvBin BinCt
                            For Each Noder2 In NonDup
                                Noder2.AddvBin BinCt
                            Next Noder2
                        End If
                    Next q
                Next n
                
                Set Duplicate = New Collection
                Set NonDup = New Collection
                Set SubNonDup = New Collection
                
                'Stop
            Next z
            
        End If
        
        Dim sField As String
        Dim sData As String
        Dim sSQL As String
        Dim b As Integer
        Dim CurrentBin As Collection
        
        'First data entry stuff, fix duplicate problem
        'Populate the database
        If a > 0 Then
            'Redo what you did with headers, but use the saved information
            
            'build empty bins
            Set Bin = New Collection
            For y = 1 To BinCt
                Set Group = New Collection
                Bin.Add Group
            Next y
            
            'Load the data into the vCSV array
            For y = 1 To vCSVColNodes.count
                vCSVColNodes(y).sData = vCurrentRow(y - 1)
                Set CurrentBin = vCSVColNodes(y).GetvBin
                If vCSVColNodes(y).bImported = False Then
                    For b = 1 To CurrentBin.count
                        Bin(CurrentBin(b)).Add vCSVColNodes(y)
                    Next b
                    vCSVColNodes(y).bImported = True
                End If
            Next y
            
            'Build the SQL string from the bin and execute
            For z = 1 To Bin.count
                'string field and data information together
                For m = 1 To Bin(z).count
                    If sField = "" Then
                        sField = Bin(z)(m).sName
                    Else
                        sField = sField & "," & Bin(z)(m).sName
                    End If
                
                    If sData = "" Then
                        sData = "'" & Bin(z)(m).sData & "'"
                    Else
                        sData = sData & "," & "'" & Bin(z)(m).sData & "'"
                    End If
                Next m
                'execute the SQL statement
                
                sSQL = "INSERT INTO " & Bin(z)(1).sParentName & " (" & sField & ") " & "VALUES (" & sData & ");"
                MyDatabase.Execute sSQL
                sData = ""
                sField = ""
            Next z
        End If
    Next a
Stop
End Sub
