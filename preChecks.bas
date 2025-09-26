Option Explicit

Dim blnCheckNotOk As Boolean 'Set this to true if any failures occur

Public Function CheckInputRanges() As Boolean

    'Use for both validation & run macro checks
    'Reason: it is possible for the validations to be true, but a user could delete info before running the macro

    On Error GoTo ErrHandlerCheckInputRanges

     ' Check if all cells in the checkRange are populated
    'Dim checkRange As Range: Set checkRange = wsControl.Range("rngInputFolder", "rngInputFile", "rngSaveFolder")
    Dim checkRange As Range: Set checkRange = Union(wsControl.Range("rngInputFolder"), wsControl.Range("rngInputFile"), wsControl.Range("rngSaveFolder"))
    
    If Not Application.CountA(checkRange) = checkRange.Cells.Count Then
        MsgBox "Empty input fields on control panel identified. Please try again", vbCritical, "User Input"
        CheckInputRanges = False
        Call AddToLog("Check input range", "", "Fail", Err.Description)
        MsgBox "Please populate each of the input ranges before trying again.", vbInformation, "Pre Checks"
    Else
        CheckInputRanges = True
        Call AddToLog("Check input range", "", "Success")
    End If

    On Error GoTo 0 'Reset error handling and exit the function

    Exit Function

ErrHandlerCheckInputRanges:
    CheckInputRanges = False
    Call AddToLog("Unexpected error on check input range function", "modPreChecks", "Fail", Err.Description)

End Function


'*********************************************************************
'Main pre checks validation below | Checked input ranges & file is accessible before this

Public Function PreChecks() As Boolean
    
    Call AddToLog("Initiating pre checks validation...", "", "")
    
    On Error GoTo ErrorHandler
    blnCheckNotOk = False 'reset in case macro crashes without running
    
    'If CheckNamedRanges = False Then blnCheckNotOk = True 'Check named ranges
    'If CheckInputRanges = False Then blnCheckNotOk = True
    'If CheckFolderLocations = False Then blnCheckNotOk = True
    'If CheckFilesExist(wsControl.Range("rngInputFolder").Value & "\" & wsControl.Range("rngInputFile").Value) = False Then blnCheckNotOk = True 'Check the files saved exist before starting checks
    'If CheckFileFormat(wsControl.Range("rngInputFolder").Value & "\" & wsControl.Range("rngInputFile").Value) = False Then blnCheckNotOk = True 'Check headers and data begins in cell A1
    If CheckSheetsExist = False Then blnCheckNotOk = True
    
        'Maybe move the check queries until after the tables - the SIZE VALUES columns will be dynamic.
    If CheckQueriesExist = False Then blnCheckNotOk = True
    If CheckTablesExist = False Then blnCheckNotOk = True
    

    If blnCheckNotOk Then
        MsgBox "Inconsistent data detected during pre-checks validation. Refer to the failed logfile items for more information before trying again.", vbCritical, "PreChecks"
        Call AddToLog("Pre checks validation", "PreCheck", "Fail", "Refer to the failed log items above, amend, and try again.")
    Else
        PreChecks = True
        Call AddToLog("Pre checks validation", "PreCheck", "Success")
    End If
    
    On Error GoTo 0 'Reset the err handling
    
    Exit Function

ErrorHandler:
    PreChecks = False
    Call AddToLog("Pre checks validation", "", "Fail", Err.Description)
    
End Function

Private Function CheckSheetsExist() As Boolean

    Dim arrExpectedSheets As Variant: arrExpectedSheets = Array("wsControl", "wsOrderDetail", "wsProductUpload", "wsPoUpload", "wsBarcode", "wsLog")
    
    Dim ws As Worksheet
    Dim dictWorksheetCodeNames As Object: Set dictWorksheetCodeNames = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrHandlerCheckSheetsExist
    For Each ws In ThisWorkbook.Worksheets
        dictWorksheetCodeNames(ws.CodeName) = Empty
    Next ws
    
    Dim i As Long
    Dim allExist As Boolean: allExist = True
    
    For i = 0 To UBound(arrExpectedSheets, 1)
        If Not dictWorksheetCodeNames.Exists(arrExpectedSheets(i)) Then
            allExist = False
            Call AddToLog("Checking worksheet exists", arrExpectedSheets(i), "Fail", "Worksheet may have been deleted and re-added. The worksheet codename is important.")
        Else
            Call AddToLog("Checking worksheet exists", arrExpectedSheets(i), "Success")
        End If
    Next i
    
    ' If code reaches here, the codename was found
    If allExist Then
        CheckSheetsExist = True
        Call AddToLog("Checking worksheet exists", "All expected worksheets", "Success")
    End If
    
    On Error GoTo 0
    
    Exit Function
    
ErrHandlerCheckSheetsExist:
    
    CheckSheetsExist = False
    Call AddToLog("Checking worksheet exists", "", "Fail", Err.Description)

End Function

Private Function CheckQueriesExist() As Boolean
    
    'Amend query names when all known
    'sizeRangeExists is a validation and will be invoked in the check headers. Therefore move the queries exist to before the tableExist
    Dim arrQueries As Variant: arrQueries = Array("SizeRanges", "ProductData", "Product_Upload", "Product_Upload_Array_Size", "PO_Upload", "tempBarcode", "staticBarcode", "sizeRangeExists")
    
    On Error GoTo ErrCheckQueriesExist
    
    Dim qry As WorkbookQuery
    Dim dictQueryNames As Object: Set dictQueryNames = CreateObject("Scripting.Dictionary")
    
    For Each qry In ThisWorkbook.Queries
        dictQueryNames(qry.name) = Empty
    Next qry
    
    Dim i As Long
    Dim allExist As Boolean: allExist = True
    
    
    For i = 0 To UBound(arrQueries, 1)
        If Not dictQueryNames.Exists(arrQueries(i)) Then
            Call AddToLog("Checking queries exists", arrQueries(i), "Fail", "Please download original file containing the queries and try again.")
            allExist = False
        Else
            Call AddToLog("Checking queries exists", arrQueries(i), "Success")
        End If
    Next i
    
    On Error GoTo 0
    
    If allExist Then
        CheckQueriesExist = True
        Call AddToLog("Checking queries exists", "All queries found", "Success")
        Exit Function
    End If
    

ErrCheckQueriesExist:
    CheckQueriesExist = False
    Call AddToLog("Checking queries exists", "", "Fail", Err.Description)

End Function

Private Function CheckTablesExist() As Boolean
    
    Dim arrTablesToCheck As Variant: arrTablesToCheck = Array("SizeRanges", "ProductData")  ', "MissingFunds", "tblProfiles")
    Dim dictTableNames As Object: Set dictTableNames = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet, tbl As ListObject
    
    ' Populate the dictionary with the names of existing tables
    
    On Error GoTo ErrCheckTablesExist
    
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            dictTableNames(tbl.name) = Empty
        Next tbl
    Next ws

    ' Check if each required table exists
    Dim allExist As Boolean: allExist = True
    Dim i As Long
    
    For i = LBound(arrTablesToCheck) To UBound(arrTablesToCheck)
        If Not dictTableNames.Exists(arrTablesToCheck(i)) Then
            allExist = False
            Call AddToLog("Checking tables exists", arrTablesToCheck(i), "Fail", "Table may have been deleted.")
        
        Else 'Check the necessary headers are available for each table
            Call AddToLog("Checking tables exists", arrTablesToCheck(i), "Success")
            
            If Not CheckHeadersExist(arrTemp:="", tableName:=arrTablesToCheck(i)) Then
                allExist = False 'Log entries have already been handled in the function being called
            End If
        End If
    Next i
    
    On Error GoTo 0 'reset err handling

    If allExist Then
        CheckTablesExist = True
    End If
    
    Exit Function

ErrCheckTablesExist:
    
    CheckTablesExist = False
    Call AddToLog("Checking tables exists", "", "Fail", Err.Description)

End Function

Private Function CheckFilesExist(ByVal fileName As String) As Boolean

    CheckFilesExist = True 'Return false if something cannot be found
    
    On Error GoTo ErrCheckFileExists

    'Check if fileName exists
    If Dir(fileName) = "" Then
        CheckFilesExist = False
        Call AddToLog("Checking file exists", "", "Fail", "Check the input folder & file names are correct before trying again. Ensure there are no trailing or leading slashes in both given strings.")
    Else
        Call AddToLog("Checking file exists", "fileName", "Success")
    End If
    
    On Error GoTo 0
    
    Exit Function
    
ErrCheckFileExists:

    CheckFilesExist = False
    Call AddToLog("Checking file exists", "", "Fail", Err.Description)

End Function

Private Function CheckFileFormat(ByVal fileName As String) As Boolean

    CheckFileFormat = True 'Return false if something cannot be found
    
    On Error GoTo ErrCheckFileFormat
    
    Dim FSO, TSReadFile As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TSReadFile = FSO.OpenTextFile(fileName, 1, False)
    
    'For the CSV only
    Dim arrTemp As Variant: arrTemp = TSReadFile.ReadAll
    arrTemp = Split(arrTemp, vbCrLf)
    arrTemp = Split(arrTemp(0), ",") 'This will take the first row of data only
    
    TSReadFile.Close
    Set TSReadFile = Nothing
    Set FSO = Nothing
    
    
    'Pass the array to the check headers function
    If Not CheckHeadersExist(arrTemp) Then CheckFileFormat = False 'Logging is already in the function being called
    
    On Error GoTo 0
    
    Exit Function
    
ErrCheckFileFormat:

    CheckFileFormat = False
    Call AddToLog("Checking file format", "", "Fail", Err.Description)

End Function

'Private Function CheckHeadersExist(ByVal tableName As String) As Boolean
Private Function CheckHeadersExist(ByVal arrTemp As Variant, Optional ByVal tableName As String) As Boolean
    
    CheckHeadersExist = True 'Return false if something cannot be found
    
    Dim dictExpectedHeaders, dictActualHeaders As Object
    
    On Error GoTo ErrCheckHeadersExist
    
    'Use the get
    If tableName = "" Then
        Set dictExpectedHeaders = GetExpectedHeaders '(tableName)
        Set dictActualHeaders = GetActualHeaders(arrTemp)  '(tableName)
    Else
        Set dictExpectedHeaders = GetExpectedHeaders(tableName)
        Set dictActualHeaders = GetActualHeaders(arrTemp:="", tableName:=tableName)
    End If
    
    'Loop the expected dictionaries
    Dim key As Variant, blnNotExists As Boolean 'set this boolean to true if something is not found

    For Each key In dictExpectedHeaders.Keys
        If Not dictActualHeaders.Exists(key) Then
            CheckHeadersExist = False
            If tableName = "" Then
                Call AddToLog("Checking headers exists", key, "Fail", "Header not found. Check input file & try again.")
            Else
                Call AddToLog("Checking headers exists", "Table: " & tableName & ": Header: " & key, "Fail", "Header not found. Check table and amend if needed.")
            End If
        End If
    Next key
    
    'Extra step to check the dynamic SIZE VALUES are appearing in the dataset
    Call UpdateQuery("sizeRangeExists", "Refreshing Query", "Refreshing " & key & " , please wait.")
    
    'Get the data from the resulting table, sizeRangeExists and if it is TRUE, then write a Success result, Fail if not.
    Dim tblSizeRangeExists As ListObject: Set tblSizeRangeExists = wsDataValidations.ListObjects("sizeRangeExists")
    Dim sizeRangeValue As Boolean: sizeRangeValue = tblSizeRangeExists.DataBodyRange.Cells(1, 1).Value 'Get the first data value from the table
    
    If sizeRangeValue = True Then
        Call AddToLog("Checking SIZE VALUE dynamic headers exists", "sizeRangeExists", "Success")
    Else
        CheckHeadersExist = False
        Call AddToLog("Checking SIZE VALUE dynamic headers exists", "sizeRangeExists", "Fail", "Mismatch on SIZE VALUE columsn between SizeRanges and ProductData tables. Amend table data & try again.")
    End If
    
    On Error GoTo 0
    
    If CheckHeadersExist Then Call AddToLog("Checking headers exists", "", "Success")
    
    Exit Function

ErrCheckHeadersExist:
    
    Debug.Print Err.Description
    
    CheckHeadersExist = False
    Call AddToLog("Checking headers exists", "", "Fail", Err.Description)

End Function

Sub CheckSizeRangeExists()
    Dim tblSizeRangeExists As ListObject
    Dim ws As Worksheet
    Dim sizeRangeValue As Boolean

    'Set the worksheet containing the table
    Set ws = ThisWorkbook.Worksheets("Sheet1") ' <-- Change to your sheet name

    'Reference the table
    Set tblSizeRangeExists = ws.ListObjects("tempDataValidations")

    'Get the first data value from the table (assuming only one row, one column)
    sizeRangeValue = tblSizeRangeExists.DataBodyRange.Cells(1, 1).Value

    'Check if TRUE or FALSE
    If sizeRangeValue = True Then
        MsgBox "Success"
    Else
        MsgBox "Fail"
    End If
End Sub

Private Function GetExpectedHeaders(Optional ByVal tableName As String) As Object
'Private Function GetExpectedHeaders() As Object

    'Use this to hold the headers of each table
    Dim arrTemp As Variant

    If tableName = "ProductData" Then
        arrTemp = Array("PRODUCTTYPE", "WEB DEPARTMENT FOR DESCRIPTION", "STYLE NAME", "BRAND", "MAJORITY FABRIC", "SUB TYPE", "SUBTYPE TARGETS HELPER", "FABRIC TARGETS HELPER", "SUPPLIERCODE", "SUPPLIER", "MANSKU", "COLOUR", "SIZE RANGES", "UNITS", "SUPPLIERBUY", "POS PRICE", "ITEM GP$", "TOTAL SELL")
    
    ElseIf tableName = "SizeRanges" Then
        arrTemp = Array("SIZE RANGES")
        
    'Else
    '    arrTemp = Array("Name", "Financial Account Name", "Account Description", "Investment Portfolio Profile", "MV Account", "MV Holding", "Financial Holding Name", "% of Holdings")
    End If
    
    Dim dictTemp As Object: Set dictTemp = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    
    For i = 0 To UBound(arrTemp)
        dictTemp(arrTemp(i)) = Empty 'Assign each array element to our dictionary
    Next i
    
    Set GetExpectedHeaders = dictTemp
    
End Function

Private Function GetActualHeaders(ByVal arrTemp As Variant, Optional ByVal tableName As String) As Object
    
    'Using the table name, load the first row into a dictionary
    Dim dictHeaders As Object: Set dictHeaders = CreateObject("Scripting.Dictionary")
    
    If tableName = "" Then
    
        'Add the first row of the array to the dictionary
        Dim j As Long
        
        For j = 0 To UBound(arrTemp, 1)
            arrTemp(j) = UCase(Trim(Replace(arrTemp(j), Chr(34), "")))    'Eliminate double quote text qualifiers from the headers | And uppercase and trim - will be assesed in PQ also
            dictHeaders(arrTemp(j)) = Empty
        Next j
        
        Set GetActualHeaders = dictHeaders ' Assign the dictionary to the function's return value
        
        Exit Function
    
    Else
    
'******************* Version for managing multiple tables, not being used
        Dim ws As Worksheet, tbl As ListObject, headerRowRange As Range, cell As Range
        Dim found As Boolean
    
        ' Populate the dictionary with the names of existing tables
        'The reason for no error handling is because a table does not allow duplicate headers
        For Each ws In ThisWorkbook.Worksheets
            For Each tbl In ws.ListObjects
                If tbl.name = tableName Then 'If the tableName parameter matches then load the table and get the first row for our headers
                    Set headerRowRange = tbl.headerRowRange
    
                    Dim columnIndex As Integer: columnIndex = 1 ' Initialize the column index
    
                    For Each cell In headerRowRange.Cells
                        ' Add to dictionary with the column index as the key
                        ' and the actual header value as the item
                        dictHeaders(cell.Value) = Empty
                        columnIndex = columnIndex + 1
                    Next cell
                    found = True ' Set found flag
                    Exit For
                End If
            Next tbl
    
            If found Then Exit For ' Exit the worksheet loop if table is found
        Next ws
    
    
        Set GetActualHeaders = dictHeaders ' Assign the dictionary to the function's return value
    
    End If

End Function

Private Function CheckFolderLocations() As Boolean
    
    CheckFolderLocations = True 'Set to false if something is not found
    
'    '1. folder location that is saved in the "rngSaveFolder' named range
'
'    If Not CheckSubfolder(wsControl.Range("rngSaveFolder").Value) Then
'        CheckFolderLocations = False
'        Call AddToLog("Checking folder location", wsControl.Range("rngSaveFolder").Value, "Fail", "Please ensure the given folder path exists and try again.")
'    Else
'        Call AddToLog("Checking folder location", wsControl.Range("rngSaveFolder").Value, "Success")
'    End If
    
    Dim arrFolders As Variant: arrFolders = Array(wsControl.Range("rngInputFolder").Value, wsControl.Range("rngSaveFolder").Value)

    Dim i As Long
    For i = 0 To UBound(arrFolders)
        If Not CheckSubfolder(arrFolders(i)) Then
            CheckFolderLocations = False
            Call AddToLog("Checking folder location", arrFolders(i), "Fail", "Please ensure the given folder path exists and try again.")
        Else
            Call AddToLog("Checking folder location", arrFolders(i), "Success")
        End If
    Next i

End Function

Public Function CheckSubfolder(ByVal folderPath As String) As Boolean

    CheckSubfolder = (Len(Dir(folderPath, vbDirectory)) > 0)

End Function

Private Function CheckNamedRanges() As Boolean
    
    'Start with true and set to false if we cannot locate one
    CheckNamedRanges = True
    
    ' List of expected named ranges
    Dim expectedNames As Variant: expectedNames = Array("rngInputFolder", "rngInputFile", "rngSaveFolder", "rngLastRuntime")
    
    On Error GoTo ErrCheckNamedRanges
    
    Dim name As Variant
    
    ' Check each name to see if it exists in the Names collection
    For Each name In expectedNames
        If Not NameExists(name) Then
            CheckNamedRanges = False
            Call AddToLog("Checking dependent named ranges", name, "Fail", "Named range not found. Please ensure no data has been deleted from the original template.")
        Else
            Call AddToLog("Checking dependent named ranges", name, "Success")
        End If
    Next name
    
    On Error GoTo 0
    
    Exit Function

ErrCheckNamedRanges:

    CheckNamedRanges = False
    Call AddToLog("Checking dependent named ranges", "", "Fail", Err.Description)
    
End Function

' Helper function to check if a given name exists in the Names collection of the workbook
Private Function NameExists(ByVal name As String) As Boolean
    On Error Resume Next  ' In case the name does not exist, avoid an error
    Dim testRange As Range
    Set testRange = ThisWorkbook.Names(name).RefersToRange
    If Not testRange Is Nothing Then
        NameExists = True
    Else
        NameExists = False
    End If
    On Error GoTo 0  ' Turn off error handling
End Function

