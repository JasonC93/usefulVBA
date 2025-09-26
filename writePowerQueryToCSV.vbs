'Your query must be added to the data model for this to work

Option Explicit

Public Sub ConnectToQuery()
    On Error GoTo ErrHandlerConnectToQuery
    
    Call AddToLog("Exporting queries to CSV..", "", "")
    
    'Loop our combined queries which have been added to the data model and export as CSV
    Dim arrQueriesToExport As Variant: arrQueriesToExport = Array("combinedPayout", "combinedLineItem")
    
    Dim i As Long 'use this to loop
    
    For i = 0 To UBound(arrQueriesToExport)
    
        Dim TABLE_NAME As String
        TABLE_NAME = arrQueriesToExport(i)
    
        Dim rs As Variant
        Set rs = GetRecordSetFromConnection(TABLE_NAME)
    
        Dim dataArray As Variant
        dataArray = RecordsetToArray(rs)
    
        ' Now dataArray contains all the data from the recordset
        ' Wrap the array in double quoutes for inline commas
        Call WriteToCSV(dataArray, arrQueriesToExport(i))
    
    Next i
    
    On Error GoTo 0
    Exit Sub

ErrHandlerConnectToQuery:
    'MsgBox "An error occurred in ConnectToQuery: " & Err.Description, vbCritical
    Call AddToLog("An error occurred when exporting queries to CSV..", "ConnectToQuery", Err.Description)
    
End Sub

Private Function GetModelADOConnection()
'We just need the ADOConnection; the rest is for perusal

    Dim wbConnections, Model, ModelDMC, ModelDMCMC As Object
    
    Set wbConnections = ThisWorkbook.Connections
    Set Model = ThisWorkbook.Model
    Set ModelDMC = Model.DataModelConnection
    Set ModelDMCMC = ModelDMC.ModelConnection
    Set GetModelADOConnection = ModelDMCMC.ADOConnection
    
End Function

Private Function GetRecordSetFromConnection(ByVal TABLE_NAME As String) As Variant
'Requires that connection is added to data model.

    Dim conn, rs As Object
    Set conn = GetModelADOConnection

    Set rs = CreateObject("ADODB.RecordSet")
    rs.Open "SELECT * From $" & TABLE_NAME & ".$" & TABLE_NAME, conn
    
    
    Set GetRecordSetFromConnection = rs
       
End Function

Private Function RecordsetToArray(ByVal rs As Object) As Variant
    If rs.EOF Then
        RecordsetToArray = Array() ' Return an empty array if the recordset is empty
        Exit Function
    End If

    ' Move to the first record
    rs.MoveFirst

    ' Define a dynamic array
    Dim dataArray() As Variant
    ReDim dataArray(rs.RecordCount, rs.Fields.Count - 2) '-2 omits the row number automatically added to the query

    Dim row As Long, col As Long

    ' Add headers to the first row of the array
    For col = 0 To rs.Fields.Count - 2
        dataArray(0, col) = rs.Fields(col).Name
        dataArray(0, col) = Split(dataArray(0, col), ".")(1) 'Remove the table name
        dataArray(0, col) = Replace(dataArray(0, col), "[", "")
        dataArray(0, col) = Replace(dataArray(0, col), "]", "")
    Next col

    ' Start adding data from the second row of the array
    row = 1

    ' Loop through the recordset
    Do Until rs.EOF
        For col = 0 To rs.Fields.Count - 2
            dataArray(row, col) = rs.Fields(col).Value
        Next col
        row = row + 1
        rs.MoveNext
    Loop

    RecordsetToArray = dataArray
End Function


Public Sub WriteToCSV(ByVal arrSource As Variant, Optional ByVal fileType As String)

    On Error GoTo ErrWriteToCSV
    
    Dim outputFile As String: outputFile = GetOutputFile(fileType)
    
    Dim FSO, TSWrite As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    arrSource = Get1dArray(arrSource) 'We will write the 1d array to a  csv
    
    Set TSWrite = FSO.OpenTextFile(outputFile, 2, True)
    
        TSWrite.Write Join(arrSource, vbCrLf)
        TSWrite.Close
        
     Set TSWrite = Nothing
     
    Call AddToLog("Exporting query", fileType, "Success")
    
    On Error GoTo 0
    Exit Sub
    
ErrWriteToCSV:
    Call AddToLog("Error when exporting query", fileType, "Fail", Err.Description)
        
End Sub

Private Function Get1dArray(ByVal arrSource As Variant) As Variant

    ReDim arrTempOutput(0 To UBound(arrSource, 1)) As Variant 'Changing back to 1d array
    
    Dim i, j As Long
    
    For i = LBound(arrSource, 1) To UBound(arrSource, 1)
        For j = 0 To UBound(arrSource, 2)
        
            If j < UBound(arrSource, 2) Then
                arrTempOutput(i) = arrTempOutput(i) & Chr(34) & arrSource(i, j) & Chr(34) & ","
            Else
                arrTempOutput(i) = arrTempOutput(i) & Chr(34) & arrSource(i, j) & Chr(34)
            End If

        Next j
    Next i
    
    Get1dArray = arrTempOutput

End Function

Private Function GetOutputFile(Optional ByVal fileType As String) As String
    
    Dim strTemp As String

    'folderPath is a variable, user can decide how to implement
    folderPath = "Your path here"

    strTemp = folderPath & "\" & "Macro_output"
    
    If CheckSubfolder(strTemp) = False Then MkDir (strTemp)
    
    strTemp = strTemp & "\" & fileType & Format(Now(), "YYYYMMDDhhmmss") & ".csv"
    
    GetOutputFile = strTemp

End Function
