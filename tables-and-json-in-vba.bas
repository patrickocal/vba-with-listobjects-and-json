Option Explicit
' Patrick O'Callaghan March 2025
' the new import sub: now only imports the data; uses tables; still uses wasp.Get_Good_Date.
Public Sub ImportData2025()

    ' integers
    Dim i As Integer
    ' dates
    Dim revalDate As Date
    Dim candidate As Date
    Dim fileDate As Date
    ' strings
    Dim str As String
    Dim inputFilePath As String
    Dim inputDirectory As String
    ' variants
    Dim fieldNames As Variant
    Dim tempDate As Variant
    ' tables
    Dim tbl As ListObject
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
On Error GoTo errHandler:

    'admin sheet
    With Sheets("admin")
        Set tbl = .ListObjects("adminDatesPaths")
        With tbl
            ' store inputfilepath for the import step
            inputFilePath = .ListColumns("inputfilepath").DataBodyRange.value
            inputDirectory = .ListColumns("InputDirectory").DataBodyRange.value
            'application.statusbar = "inputFilePath is" & inputFilePath
            '
            ' KEY STEP: import the landing file data
            ' check if we can import a specific file
            If inputFilePath = "" Then
                ' work with inputdirectory instead
                If inputDirectory = "" Then inputDirectory = ThisWorkbook.path & "\input\"
                ' select the input workbook manually
                inputFilePath = selectInputFile(inputDirectory)
                ' check if the inputfilepath was successfully selected
                If inputFilePath = "" Then
                    Application.StatusBar = "File not selected, so no new Apollo Data imported this time."
                    Exit Sub
                End If
            End If
            If Right(inputFilePath, 5) = ".xlsx" Or Right(inputFilePath, 4) = ".csv" Then
                ' try to populate the input sheet
                If populateInputSheet(inputFilePath) Then
                    ' clear the inputfilepath to avoid inadvertant import of stale data
                    .ListColumns("inputfilepath").DataBodyRange = ""
                Else
                    MsgBox "The input file cannot be found at " & inputFilePath
                    Exit Sub
                End If
            End If
            '
        End With
        Set tbl = .ListObjects("adminCheck")
        With tbl
            .ListColumns("Input Data Last Updated On").DataBodyRange = Format(Date + Time, "dd-mmm-YYYY")
        End With
    End With
    
  
    'Remove extra dates from input data
        'Call sb_FormatInputData
        
        Call workbookStatus("ImportData2025")
        
    Exit Sub
    
errHandler:
    MsgBox Err.Description
    'Resume
    
    Exit Sub

End Sub
' the new import opics rates and curves: preserves dates as text (strings) not dates for more stable and
' robust data; no longer uses wasp
Public Sub opicsDiscountCurves2025()
    Dim tbl As ListObject
    Dim sqlStr As String
    Dim a_CFRS() As Variant
    Dim strLastRow As String
    'dates
    Dim revalDate As String
    Dim todayDate As Date
    Dim candidate As String
    Dim i As Integer
    
    
On Error GoTo errHandler:

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    With Sheets("admin")
        'revalDate = .Range("revalDate").value
        Set tbl = .ListObjects("adminDatesPaths")
        With tbl
            ' get candidate
            candidate = .ListColumns("candidate").DataBodyRange
            todayDate = Date
            If candidate = "" Then
                revalDate = Format(CorrectRevalDate(todayDate), "dd-mmm-yyyy")
            Else
                revalDate = Format(CorrectRevalDate(CDate(candidate)), "dd-mmm-yyyy")
            End If
            ' populate the table with the revalDate
            .ListColumns("revalDate").DataBodyRange = revalDate
        End With
        
    End With
    With Sheets("OPICS")
        'clear and then repopulate the opicsDiscountFactor table with the revalDate's AUDSWAP1 curve data
        Set tbl = .ListObjects("opicsDiscountFactor")
        Call getDataFromOpics(tbl, "discountFactorQry-ordered.sql", revalDate, "AUDSWAP1")
        'clear and then repopulate the opicsRatesHistory table with the revalDate's BBSW30 data
        Set tbl = .ListObjects("opicsRatesHistory")
        Call getDataFromOpics(tbl, "ratesHistoryQry-2025.sql", revalDate, "BBSW30")
    End With
    'rinse and repeat for the OPICS_AUDSWAP3 sheet:
    With Sheets("OPICS_AUDSWAP3")
        'clear and then repopulate the opicsAudSwap3DiscountFactor table with the revalDate's AUDSWAP3 curve data
        Set tbl = .ListObjects("opics_AudSwap3DiscountFactor")
        Call getDataFromOpics(tbl, "discountFactorQry-ordered.sql", revalDate, "AUDSWAP3")
        'clear and then repopulate the opics_audswap3RatesHistory table with the revalDate's BBSW90 data
        Set tbl = .ListObjects("opics_audswap3RatesHistory")
        Call getDataFromOpics(tbl, "ratesHistoryQry-2025.sql", revalDate, "BBSW90")
    End With
    'rinse and repeat for the OPICS_AUDSWAP3 sheet:
    With Sheets("OPICS_OIS")
        'clear and then repopulate the opics_OISDiscountFactor table with the revalDate's OIS curve data
        Set tbl = .ListObjects("opics_OISDiscountFactor")
        Call getDataFromOpics(tbl, "discountFactorQry-ordered.sql", revalDate, "OIS")
        'clear and then repopulate the opics_OISRatesHistory table with the revalDate's ONCASH data
        Set tbl = .ListObjects("opics_OISRatesHistory")
        Call getDataFromOpics(tbl, "ratesHistoryQry-2025.sql", revalDate, "ONCASH")
    End With
            
    Call workbookStatus("opicsDiscountCurves2025")
    Exit Sub

errHandler:
    MsgBox Err.Description
    Exit Sub
    
End Sub
Public Sub genSwaps2025()
    ' doubles
    Dim runTime As Double
    Dim frac As Integer
    ' tables
    Dim tbl As ListObject
    ' strings
    Dim s As String
    Dim sec As String
    
    'record start time
    runTime = Timer

 
    ' ------------------------------------------------------------------------------------
    ' KEY STEP: completes all calculations and stores the results to json-output\parentDict.json
    ' once it has run successfully, there is no need to re-run.
    Call CreateDictsFromTemplate
    ' ------------------------------------------------------------------------------------

    With Sheets("admin")
        ' lastly record time taken
        Set tbl = .ListObjects("adminDatesPaths")
        With tbl
            ' Number of seconds it took to run the model
            runTime = Round(Timer - runTime, 0)
            ' save to table
            .ListColumns("Time taken to run model (seconds)").DataBodyRange.value = runTime
        End With
        Set tbl = .ListObjects("adminCheck")
        With tbl
            ' save current user to table
            .ListColumns("Last User").DataBodyRange.value = Environ("USERNAME")
            ' save time stamp to table
            .ListColumns("Last Time Stamp").DataBodyRange.value = Format(Date, "dd-mmm-yyyy") & " " & Format(Time, "hh:mm")
        End With
    End With
End Sub

' the new genswaps subroutine: it pulls data in from the dicts
Public Sub populateTablesAndExportLocally2025()

'On Error GoTo ErrHandler:
    ' integers
    Dim i As Integer: i = 1
    Dim j As Integer: j = 1
    Dim k As Integer: k = 1
    Dim n As Integer
    Dim nTranches As Integer
    Dim numSheets As Integer: numSheets = 3
    ' strings
    Dim strSearch As String
    Dim pathToParentDict As String
    Dim parentTemplatePath As String
    Dim colName As String
    Dim filename As String
    Dim saveFilePath As String
    Dim staticDataPath As String
    Dim outputCSVPath As String
    Dim localOutputDirectoryPath As String
    Dim revalDateStr As String
    ' variants
    Dim arrMortgRate As Variant
    Dim arrDates As Variant
    Dim apolloItem As Variant  ' Could be a dictionary or another type
    Dim key As Variant
    Dim path As Variant
    Dim name As Variant
    ' dicts
    Dim apolloDict As Object
    Dim adminDict As Object
    Dim localDict As Object
    Dim calcDict As Object
    Dim fixedSwapsDict As Object
    Dim otcDict As Object
    Dim fso As Object
    Dim file As Object
    ' ranges
    Dim headRng As Range
    Dim productRng As Range
    ' booleans
    Dim bool As Boolean
    ' dates
    Dim revalDate As Date
    ' workbooks
    Dim wb As Workbook
    Dim newWB As Workbook
    ' tables
    Dim tbl As ListObject
    Dim tbl2 As ListObject
    ' arrays
    Dim paths() As String: ReDim paths(1 To 2)
    Dim sheetNames() As String: ReDim sheetNames(1 To numSheets)
    ' collections
    Dim excludeCols As Collection
    Dim coll As Collection
    
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    
    ' get the path to the json file with the populated parent dict
    pathToParentDict = ThisWorkbook.path & localPathToOutputDict
    
    ' Check if parentDict is Nothing.
    If parentDict Is Nothing Then
        Set parentDict = LoadParentDictFromFile(pathToParentDict)
    End If

    ' all that remains is to populate the tables
    '
    ' the admin sheet
    With ThisWorkbook.Sheets("admin")
        ' save the parentdict's revalDate to the adminCheck table
        Set tbl = .ListObjects("adminCheck")
        With tbl
            .ListColumns("json reval date").DataBodyRange = parentDict("revalDate")
        End With
        ' check the parentdict matches the revaldate
        Set tbl = .ListObjects("adminDatesPaths")
        revalDateStr = tbl.ListColumns("revalDate").DataBodyRange
        ' reset tbl to adminCheck which is where we will store the date if all is well
        If parentDict("revalDate") = revalDateStr Then
            ' Do nothing
        Else
            MsgBox "Mismatch of reval dates on the admin sheet and json file." _
                & " Try re-running the GenSwaps... subroutine again (press 'Generate Swap Data and save to JSON')"
                Exit Sub
        End If
        '
        ' assign new and expired trust names to the admin sheet table adminExpNewTrusts
        Set tbl = .ListObjects("adminExpNewTrusts")
        With tbl
            Application.StatusBar = "Populating the " & .name & " table"
            .DataBodyRange.ClearContents
            ' set k to be the length of the list
            k = parentDict("expApolloNames").Count
            'application.statusbar = k & ", " & parentDict("expApolloNames")(1)
            If k > 0 Then
                For i = 1 To k Step 1
                .ListColumns("Expired Trusts").DataBodyRange(i, 1) _
                    = parentDict("expApolloNames")(i)
                Next i
            End If
            ' set k to be the length of the list
            k = parentDict("newApolloNames").Count
            If k > 0 Then
                For i = 1 To k Step 1
                .ListColumns("New Trusts").DataBodyRange(i, 1) = parentDict("newApolloNames")(i)
                Next i
            End If
        End With
        '
        ' the new names and "input data" to the adminInput table
        Set tbl = .ListObjects("adminInput")
        With tbl
            Application.StatusBar = "Populating the " & .name & " table"
            For i = 1 To .ListColumns.Count
                With .ListColumns(i)
                    If .DataBodyRange.Cells(1, 1).HasFormula = False Then
                        .DataBodyRange.ClearContents
                        colName = .name
                        ' fill the main (top) table on the ADMIN sheet
                        j = 1
                        parentDict.CompareMode = vbBinaryCompare ' ensure case-sensitivity
                        For Each key In parentDict.Keys
                            If Left(key, 6) = "APOLLO" Then
                                'Application.StatusBar = dict("APOLLO 28")("ApolloName")
                                Set apolloDict = parentDict(key)
                                ' for each sub-dict, clone create a new handle
                                Set adminDict = apolloDict("adminInput")
                                'application.statusbar = j & ", " & colName & ", " & adminDict(colName)
                                .DataBodyRange.Cells(j, 1) = adminDict(colName)
                                j = j + 1
                            End If
                        Next key
                    End If
                End With
            Next i
            With .sort
                .SortFields.Clear
                .SortFields.Add key:=tbl.ListColumns("Product").Range, Order:=xlAscending
                .Header = xlYes
                .Apply
            End With
        End With
        '
        ' assign the new names to the adminDiff table (other columns are formulae and untouched)
        Set tbl = .ListObjects("adminDiff")
        Call assignDictToTable(tbl, parentDict, True)
    End With
    ' the main calculations:
    ' commonData
    With ThisWorkbook.Sheets("commonData")
        Set tbl = .ListObjects("commonDataMain")
        Call assignDictToTable(tbl, parentDict)
    End With
    ' Basis
    With ThisWorkbook.Sheets("BASIS_SWAPS")
        Set tbl = .ListObjects("basisSwapsMain")
        Call assignDictToTable(tbl, parentDict)
    End With
    '
    ' Fixed
    With ThisWorkbook.Sheets("FIXED_SWAPS")
        Application.StatusBar = "Populating the fixedSwapsMain table"
        Set tbl = .ListObjects("fixedSwapsMain")
        Call assignDictToTable(tbl, parentDict)
    End With
    '
    ' the all important report
    With ThisWorkbook.Sheets("otcreport")
        Application.StatusBar = "Populating the otcReportMain table"
        Set tbl = .ListObjects("otcreportmain")
        Call assignDictToTable(tbl, parentDict, True)
    End With
    '
    'If parentDict("newApolloNames").Count > 0 Then
    '    With ThisWorkbook.Sheets("accounting")
    '        Application.StatusBar = "Populating the accountingMain table"
    '        Set tbl = .ListObjects("accountingMain")
    '        With tbl
    '
    '        End With
    '    End With
    'End If
    
    '
    
    ' ------------------------------------------------------------------------------------
    ' Exporting the data locally
    ' ------------------------------------------------------------------------------------
    ' 0. set the local output directory (create it if one doesn't exist and clear contents otherwise)
    localOutputDirectoryPath = ThisWorkbook.path & "\output\"
    Call checkClearLocalDir(localOutputDirectoryPath)
        
    ' ---------------------------------------------------------------------------
    ' 1. set path to the OTC reval workbook (for liquidity): stores the main calculations (not otcReport)
    '
    ' first get reval date and config/static file
    With ThisWorkbook.Sheets("admin")
        Set tbl = .ListObjects("adminDatesPaths")
        ' get reval date
        revalDate = tbl.ListColumns("revalDate").DataBodyRange
        ' get STATIC/CONFIG file paths: typically local
        staticDataPath = tbl.ListColumns("staticDataPath").DataBodyRange
        If staticDataPath = "" Then
            staticDataPath = ThisWorkbook.path & "\config\APOLLO Fixed and Basis Swaps - New XML - LEIS.xlsm"
        End If
    End With
    
    ' set file name to the current date
    filename = Format(revalDate, "yyyymmdd") & "_OTC reval.xlsx"
    saveFilePath = localOutputDirectoryPath & filename
    '
    ' EXPORT main calculation sheets
    ' select sheet names for export
    sheetNames(1) = "output"
    sheetNames(2) = "FIXED_SWAPS"
    sheetNames(3) = "BASIS_SWAPS"
    ' create and save a copy of the sheets in sheetNames
    Set newWB = CreateCopyWithSheets(ThisWorkbook, saveFilePath, sheetNames)
    ' and close it
    newWB.Close SaveChanges:=True
    
    ' ---------------------------------------------------------------------------
    ' ---------------------------------------------------------------------------
    ' exports 2 and 3 both relate to the otcReport table
    With ThisWorkbook.Sheets("otcreport")
        Set tbl = .ListObjects("otcreportmain")
    End With
    ' ---------------------------------------------------------------------------
    ' 2. save the otcReportMain table to the locations in following paths:
    paths(1) = localOutputDirectoryPath & "\" & Format(revalDate, "yyyymmdd") & "_APOLLO_OTCREPORT.csv"
    paths(2) = localOutputDirectoryPath & "\" & Format(revalDate, "yyyymmdd") & "_ApolloCheck.csv"
    
    Set excludeCols = New Collection
    For Each path In paths
        bool = ExportTableToCSV_NoWorkbook(tbl, path, False, excludeCols)
        If bool Then Application.StatusBar = "Saved OTCReportMain to " & path
    Next path

    ' ---------------------------------------------------------------------------
    ' 3. STATIC/config workbook update
    
    outputCSVPath = localOutputDirectoryPath & "\" & Format(revalDate, "yyyymmdd") & "_ApolloSwaps.csv"
    ' take the file name from the current config/statics file
    filename = Mid(staticDataPath, InStrRev(staticDataPath, "\") + 1)
    '
    ' open the config/static workbook (if it is not already open)
    If IsWorkbookOpen(staticDataPath) = False Then Workbooks.Open filename:=staticDataPath
    ' set a handle
    Set wb = Workbooks(filename)
    ' get the destination table
    With wb.Sheets("Summary")
        Set tbl2 = .ListObjects("summaryOTCReport")
        With tbl2
            ' clear the decks
            .DataBodyRange.ClearContents
            ' load the otcreport data into the statics/config workbook
            For i = 1 To .ListColumns.Count
                With .ListColumns(i)
                    name = .name
                    .DataBodyRange = tbl.ListColumns(name).DataBodyRange.value
                End With
            Next i
        End With
        '
        Set tbl = .ListObjects("summaryExport")
        excludeCols.Add "Field Description"
        bool = ExportTableToCSV_NoWorkbook(tbl, outputCSVPath, True, excludeCols)
        If bool Then Application.StatusBar = "Saved OTCReportMain to " & outputCSVPath
    End With
    ' and close it
    wb.Close SaveChanges:=True
    ' ---------------------------------------------------------------------------
    ' 4. accountant's email sheet
    outputCSVPath = localOutputDirectoryPath & "\" & Format(revalDate, "yyyymmdd") & "_ApolloCheck_ACCOUNTING.csv"
    With Sheets("accounting")
        Set tbl = .ListObjects("accountingMain")
        Set excludeCols = New Collection
        bool = ExportTableToCSV_NoWorkbook(tbl, outputCSVPath, True, excludeCols)
        If bool Then Application.StatusBar = "Saved accountingMain to " & outputCSVPath
    End With
    
    ' ---------------------------------------------------------------------------
    ' finally

    
    'Save Historic Values within spreadsheet
    'Call sb_RollSwapHistory

'    Worksheets("ADMIN").Select
'    Range("A1").Select
    Call workbookStatus("PopulateTablesAndExportDataLocally2025")

End Sub
' the production step
Public Sub copyOutputFilesToProduction2025()

    ' strings
    Dim filename As String
    Dim saveFilePath As String
    Dim staticDataPath As String
    Dim prd1OpicsOTCDir As String
    Dim localOutputDir As String
    Dim MarketRiskOTCClearingDir As String
    Dim ApolloSwapsOutputDir As String
    Dim localTestDir As String
    Dim localFilePath As String
    Dim productionFilePath As String
    ' tables
    Dim tbl As ListObject
    ' collections
    Dim productionPaths As Collection
    Dim localPaths As Collection
    ' dates
    Dim revalDate As Date
    
    
    ' the local output directory that contains the exports from the ExportDataLocally macro
    localOutputDir = ThisWorkbook.path & "\output\"
    If Dir(localOutputDir, vbDirectory) = "" Then MsgBox "There is no local output directory! " _
        & "Rerun `Export Results To Local Output` macro."
    '
    ' the local test directory (make it if it doesn't exist and clear it otherwise)
    localTestDir = ThisWorkbook.path & "\test\"
    Call checkClearLocalDir(localTestDir)
    
    ' get file times and paths
    With Sheets("admin")
        Set tbl = .ListObjects("adminDatesPaths")
        ' get reval date
        revalDate = tbl.ListColumns("revalDate").DataBodyRange
        ' get output directory
        ApolloSwapsOutputDir = tbl.ListColumns("ApolloSwapsOutputDir").DataBodyRange
        ' set it to the local output directory if the cell is empty
        If ApolloSwapsOutputDir = "" Then
            ApolloSwapsOutputDir = localTestDir
        End If
        ' get apollo check path
        MarketRiskOTCClearingDir = tbl.ListColumns("MarketRiskOTCClearingDir").DataBodyRange
        If MarketRiskOTCClearingDir = "" Then
            MarketRiskOTCClearingDir = localTestDir
        End If
        ' get prd1Opics path
        prd1OpicsOTCDir = tbl.ListColumns("PRD1OpicsOTCDir").DataBodyRange
        If prd1OpicsOTCDir = "" Then
            prd1OpicsOTCDir = localTestDir
        End If
    End With

    ' ---------------------------------------------------------------------------
    ' 1. Apollo revaluations (for Liquidity team)
    filename = Format(revalDate, "yyyymmdd") & "_OTC reval.xlsx"
    ' set source path
    localFilePath = localOutputDir & "\" & filename
    ' set destination path
    productionFilePath = ApolloSwapsOutputDir & "\" & filename
    ' checking if files are recent, if so copy across to production
    Call copyFileUsingFileCopy(localFilePath, productionFilePath)
    
    ' ---------------------------------------------------------------------------
    ' 2. Apollo check
    filename = "apolloCheck.csv"
    ' set source path
    localFilePath = localOutputDir & "\" & Format(revalDate, "yyyymmdd") & "_" & filename
    ' set destination path
    productionFilePath = MarketRiskOTCClearingDir & "\" & filename
    ' checking if files are recent, if so copy across to production
    If copyFileUsingFileCopy(localFilePath, productionFilePath) Then
        ' send email
        Call createEmail2025("rng_emails", productionFilePath)
    End If

    ' ---------------------------------------------------------------------------
    ' 3. Accounting Apollo Check
    filename = "apolloCheck_accounting.csv"
    ' set source path
    localFilePath = localOutputDir & "\" & Format(revalDate, "yyyymmdd") & "_" & filename
    ' set destination path
    productionFilePath = MarketRiskOTCClearingDir & "\" & filename
    ' checking if files are recent, if so copy across to production
    If copyFileUsingFileCopy(localFilePath, productionFilePath) Then
        ' send email
        Call createEmail2025("rng_Accounting_Emails", productionFilePath)
    End If

    ' ---------------------------------------------------------------------------
    ' 4. copy to PRD1OPICS
    filename = "apolloSwaps.csv"
    ' set source path
    localFilePath = localOutputDir & "\" & Format(revalDate, "yyyymmdd") & "_" & filename
    ' set destination path
    productionFilePath = prd1OpicsOTCDir & "\" & filename
    ' checking if files are recent, if so copy across to production
    Call copyFileUsingFileCopy(localFilePath, productionFilePath)

    Call workbookStatus("copyOutputFilesToProduction2025")
End Sub

