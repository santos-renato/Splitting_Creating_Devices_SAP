Attribute VB_Name = "Process"
Dim sap_system As String

Sub import_report()

    Dim FileOpen As String
    Dim SelectedBook As Workbook, MacroBook As Workbook
    Dim ReportSheet As Worksheet
    Dim ShCnt As Byte
    Dim FirstVisibleRow As Long
    Dim VisibleRange As Range
    Dim ArrCountries() As String, ArrCountriesP05() As String, ArrCountriesPA3() As String
    Dim TbCountries As ListObject, TbCountriesP05 As ListObject, TbCountriesPA3 As ListObject
    Dim i As Long
    Dim ReportLastRow As Long, tempLastRow As Long
    
    Set TbCountries = ShConfig.ListObjects("TbCountries")
    Set TbCountriesP05 = ShConfig.ListObjects("TbCountriesP05")
    Set TbCountriesPA3 = ShConfig.ListObjects("TbCountriesPA3")
    
    ' defining array for list of countries in PE5, PA3 & P05
    ReDim ArrCountries(2 To TbCountries.Range.Rows.Count)
    ReDim ArrCountriesP05(2 To TbCountriesP05.Range.Rows.Count)
    ReDim ArrCountriesPA3(2 To TbCountriesPA3.Range.Rows.Count)
    For i = LBound(ArrCountries) To UBound(ArrCountries)
    ArrCountries(i) = ShConfig.Cells(i, 3).Value
    Next i
    For i = LBound(ArrCountriesP05) To UBound(ArrCountriesP05)
    ArrCountriesP05(i) = ShConfig.Cells(i, 8).Value
    Next i
    For i = LBound(ArrCountriesPA3) To UBound(ArrCountriesPA3)
    ArrCountriesPA3(i) = ShConfig.Cells(i, 13).Value
    Next i
    
    Call EntryPoint
    
    Set MacroBook = ActiveWorkbook
    
    Shtemp.Visible = xlSheetVisible
    ShEqui.Visible = xlSheetVisible
    Shzset.Visible = xlSheetVisible
    
    FileOpen = Application.GetOpenFilename(Filefilter:="Excel Files(*.xls*),*xls*", Title:="Select Workbook to import", MultiSelect:=False)
    If FileOpen = "False" Then
        Shtemp.Visible = xlSheetVeryHidden
        ShEqui.Visible = xlSheetVeryHidden
        Shzset.Visible = xlSheetVeryHidden
        MsgBox "No file selected.", vbCritical
        Call ExitPoint
        Exit Sub
    Else
        Set SelectedBook = Workbooks.Open(FileOpen)
        SelectedBook.Activate
    End If
    
    ' loop to check from which sheet we retrieve data
    For ShCnt = 1 To SelectedBook.Sheets.Count
        If Sheets(ShCnt).Range("A1").Value = "Serialnumber" Then
            Set ReportSheet = Sheets(ShCnt)
        End If
    Next ShCnt
    
    ReportSheet.Activate
    
    ' check if it has a filter, if not then set it
    If ReportSheet.AutoFilterMode Then
        'do nothing
    Else
        ReportSheet.Range("A1").AutoFilter
    End If
    
    ReportLastRow = ReportSheet.Range("A" & Rows.Count).End(xlUp).row
    
    'sorts Country ASC -> will help to save zeuequi based on country
'    Selection.AutoFilter
'    Shtemp.AutoFilter.Sort.SortFields.Clear
    ReportSheet.AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("D1:D" & ReportLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ReportSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    For i = 1 To 3 ' (loop for PE5,PA3 & P05)
        If i = 1 Then ' PE5
            sap_system = "PE5"
            Shtemp.Range("A2:J100000").Clear
            ReportSheet.Range("G1").AutoFilter field:=7, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("H1").AutoFilter field:=8, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("I1").AutoFilter field:=9, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("J1").AutoFilter field:=10, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("D1").AutoFilter field:=4, Criteria1:=ArrCountries, Operator:=xlFilterValues
            FirstVisibleRow = ReportSheet.Range("A2:A" & Rows.Count).SpecialCells(xlCellTypeVisible)(1).row
            Set myrange = ReportSheet.Range("A" & FirstVisibleRow & ":" & "J" & ReportLastRow).SpecialCells(xlCellTypeVisible)
            myrange.Copy
            Shtemp.Range("A2").PasteSpecial (xlPasteValuesAndNumberFormats)
            If Shtemp.Range("A2").Value <> "" Then
            Call ZeuEqui
            Else
                MsgBox "There are no countries to be processed in PE5", vbInformation
            End If
        End If
        If i = 2 Then ' PA3
            sap_system = "PA3"
            Shtemp.Range("A2:J100000").Clear
            ReportSheet.Range("G1").AutoFilter field:=7, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("H1").AutoFilter field:=8, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("I1").AutoFilter field:=9, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("J1").AutoFilter field:=10, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("D1").AutoFilter field:=4, Criteria1:=ArrCountriesPA3, Operator:=xlFilterValues
            FirstVisibleRow = ReportSheet.Range("A2:A" & Rows.Count).SpecialCells(xlCellTypeVisible)(1).row
            Set myrange = ReportSheet.Range("A" & FirstVisibleRow & ":" & "J" & ReportLastRow).SpecialCells(xlCellTypeVisible)
            myrange.Copy
            Shtemp.Range("A2").PasteSpecial (xlPasteValuesAndNumberFormats)
            If Shtemp.Range("A2").Value <> "" Then
            Call ZeuEqui
            Else
                MsgBox "There are no countries to be processed in PA3", vbInformation
            End If
        End If
        If i = 3 Then ' P05
            sap_system = "P05"
            Shtemp.Range("A2:J100000").Clear
            ReportSheet.Range("G1").AutoFilter field:=7, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("H1").AutoFilter field:=8, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("I1").AutoFilter field:=9, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("J1").AutoFilter field:=10, Criteria1:="", Operator:=xlFilterValues
            ReportSheet.Range("D1").AutoFilter field:=4, Criteria1:=ArrCountriesP05, Operator:=xlFilterValues
            FirstVisibleRow = ReportSheet.Range("A2:A" & Rows.Count).SpecialCells(xlCellTypeVisible)(1).row
            Set myrange = ReportSheet.Range("A" & FirstVisibleRow & ":" & "J" & ReportLastRow).SpecialCells(xlCellTypeVisible)
            myrange.Copy
            Shtemp.Range("A2").PasteSpecial (xlPasteValuesAndNumberFormats)
            If Shtemp.Range("A2").Value <> "" Then
            Call Zset
            Else
                MsgBox "There are no countries to be processed in P05", vbInformation
            End If
        End If
    Next i
    
    SelectedBook.Close False
    Shtemp.Visible = xlSheetVeryHidden
    ShEqui.Visible = xlSheetVeryHidden
    Shzset.Visible = xlSheetVeryHidden
    ShMain.Activate
    ShMain.Range("A1").Select
    MsgBox "Files saved in same directory of this macro.", vbInformation
    
    Call ExitPoint
    
End Sub

Sub ZeuEqui()

    Dim tempLastRow As Long
    Dim i As Long, j As Long
    Dim FilePath As String
    Dim NewBook As Workbook
    Dim row As Long
    
    tempLastRow = Shtemp.Range("A" & Rows.Count).End(xlUp).row
    ShEqui.Range("A2:AM100000").Clear
    
    row = 2
    
    For i = 2 To tempLastRow
        If Shtemp.Cells(i, 4).Value = Shtemp.Cells(i + 1, 4).Value Then 'checks if next line is same country
            Select Case Trim(Shtemp.Cells(i, 5).Value) 'case for material number and description
                Case "HYC_300"
                    ShEqui.Cells(row, 1).Value = "Model 1"
                    ShEqui.Cells(row, 2).Value = "Material 1"
                Case "HYC_150"
                    ShEqui.Cells(row, 1).Value = "Model 2"
                    ShEqui.Cells(row, 2).Value = "Material 2"
                Case "HYC_50"
                    ShEqui.Cells(row, 1).Value = "Model 3"
                    ShEqui.Cells(row, 2).Value = "Material 3"
            End Select
            ShEqui.Cells(row, 3).Value = Trim(Shtemp.Cells(i, 1).Value)
            If sap_system = "PE5" Then
                ShEqui.Cells(row, 4).Value = Application.WorksheetFunction.VLookup(Trim(Shtemp.Cells(i, 4).Value), ShConfig.Range("C:D"), 2, False)
            Else
                ShEqui.Cells(row, 4).Value = Application.WorksheetFunction.VLookup(Trim(Shtemp.Cells(i, 4).Value), ShConfig.Range("M:N"), 2, False)
            End If
            ShEqui.Cells(row, 5).Value = Mid(ShEqui.Cells(row, 4).Value, 4, 7)
            row = row + 1
        Else 'if next line is not same country
            Select Case Trim(Shtemp.Cells(i, 5).Value) 'case for material number and description
                Case "HYC_300"
                    ShEqui.Cells(row, 1).Value = "Model 1"
                    ShEqui.Cells(row, 2).Value = "Material 1"
                Case "HYC_150"
                    ShEqui.Cells(row, 1).Value = "Model 2"
                    ShEqui.Cells(row, 2).Value = "Material 2"
                Case "HYC_50"
                    ShEqui.Cells(row, 1).Value = "Model 3"
                    ShEqui.Cells(row, 2).Value = "Material 3"
            End Select
            ShEqui.Cells(row, 3).Value = Trim(Shtemp.Cells(i, 1).Value)
            If sap_system = "PE5" Then
                ShEqui.Cells(row, 4).Value = Application.WorksheetFunction.VLookup(Trim(Shtemp.Cells(i, 4).Value), ShConfig.Range("C:D"), 2, False)
            Else
                ShEqui.Cells(row, 4).Value = Application.WorksheetFunction.VLookup(Trim(Shtemp.Cells(i, 4).Value), ShConfig.Range("M:N"), 2, False)
            End If
            ShEqui.Cells(row, 5).Value = Mid(ShEqui.Cells(row, 4).Value, 4, 7)
            row = row + 1
            FilePath = ThisWorkbook.Path & "\EVC_ZEUEQUI_" & Shtemp.Cells(i, 4).Value & "_" & sap_system & ".txt"
            ShEqui.Copy
            Set NewBook = ActiveWorkbook
            NewBook.SaveAs FilePath, xlUnicodeText
            NewBook.Close False
            ShEqui.Range("A2:AM100000").Clear
            row = 2
            If i = tempLastRow + 1 Then
                Exit Sub
            End If
        End If
    Next i
    
End Sub

Sub Zset()
    
    Dim tempLastRow As Long
    Dim i As Long, j As Long
    Dim FilePath As String
    Dim NewBook As Workbook
    Dim row As Byte
    
    tempLastRow = Shtemp.Range("A" & Rows.Count).End(xlUp).row
    Shzset.Range("B17:AR100000").Clear
    
    row = 17
    
    For i = 2 To tempLastRow
        Shzset.Cells(row, 3) = "11014059"
        Select Case Trim(Shtemp.Cells(i, 5).Value)
            Case "HYC_150"
                Shzset.Cells(row, 4) = "Model 1"
            Case "HYC_300"
                Shzset.Cells(row, 4) = "Model 2"
            Case "HYC_50"
                Shzset.Cells(row, 4) = "Model 3"
        End Select
        Shzset.Cells(row, 5) = Application.WorksheetFunction.XLookup(Shzset.Cells(row, 4), ShConfig.Range("T:T"), ShConfig.Range("S:S"))
        Shzset.Cells(row, 6) = Trim(Shtemp.Cells(i, 1).Value)
        Shzset.Cells(row, 30) = "01.01.2020"
        Shzset.Cells(row, 31) = "31.12.2020"
        Shzset.Cells(row, 32) = "Please ask Customer"
        If Shtemp.Cells(i, 4).Value = "" Then
            Shzset.Cells(row, 33) = "Unknown"
            Shzset.Cells(row, 35) = "Unknown"
        Else
            Shzset.Cells(row, 33) = Shtemp.Cells(i, 4).Value
            Shzset.Cells(row, 35) = Shtemp.Cells(i, 4).Value
        End If
        Shzset.Cells(row, 34) = "99999"
        Shzset.Cells(row, 43) = "ja"
        Shzset.Cells(row, 44) = "ja"
        row = row + 1
    Next i
    
    FilePath = ThisWorkbook.Path & "\EVC_ZSET_" & sap_system & ".txt"
    Shzset.Copy
    Set NewBook = ActiveWorkbook
    NewBook.SaveAs FilePath, xlUnicodeText
    NewBook.Close False

End Sub
