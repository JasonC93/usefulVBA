Option Explicit

Dim blnRefreshQuery As Boolean

Public Function RefreshQueries() As Boolean
    
    Call AddToLog("Entering query refresh", "", "")
    blnRefreshQuery = True 'Set to false if an error occurs
    
    On Error GoTo ErrRefreshQueries
    
    'Gather list of queries that need to be run & refreshed
    Dim arrQueryList As Variant: arrQueryList = Array("SizeRanges", "ProductData", "Product_Upload_Array_Size", "tempBarcode", "staticBarcode", "Product_Upload", "PO_Upload")
    
    'Add these to a dictionary and then we can assign Yes/No to their item values if they are found in our process table
    Dim dictQueryList As Object: Set dictQueryList = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 0 To UBound(arrQueryList)
        dictQueryList.Add key:=CStr(arrQueryList(i)), Item:=""
    Next i
    
    Dim key As Variant
    For Each key In dictQueryList.Keys
        Call UpdateQuery(key, "Refreshing Query", "Refreshing " & key & " , please wait.")
    Next key
    
    On Error GoTo 0
    
    If blnRefreshQuery Then
        RefreshQueries = True 'Successful if we enter this condition
        Call AddToLog("Exiting query refresh", "", "Success")
    Else
        RefreshQueries = False
    End If
    
    Exit Function

ErrRefreshQueries:
    RefreshQueries = False
    Call AddToLog("Unexpected error updating query", "", "Fail", Err.Description)

End Function

Public Sub UpdateQuery(ByVal queryName As String, ByVal strCaption As String, ByVal strStatus As String, Optional ByVal formula As String = "")
    
    On Error GoTo ErrUpdateQuery
    
    'ProgressIndicator.UpdateProgress strCaption, strStatus
    
'    With ThisWorkbook.Queries(queryName)
'        If formula <> "" Then .formula = formula
'        .
'        .Refresh
'        Call AddToLog("Updating query", queryName, "Success")
'    End With

    queryName = "Query - " & queryName 'Different when using this method. Needs the query prefix.

    With ThisWorkbook.Connections(queryName).OLEDBConnection
        'If formula <> "" Then .formula = formula
        
        .BackgroundQuery = False
        .Refresh
    End With
    
    On Error GoTo 0
    
    Exit Sub

ErrUpdateQuery:
    blnRefreshQuery = False
    Call AddToLog("Error when updating query", queryName, "Fail", Err.Description)
    
End Sub
