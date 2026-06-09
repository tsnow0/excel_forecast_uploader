Attribute VB_Name = "DataDownload"
Sub DownloadData()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
ActiveWorkbook.Sheets("Raw Data").Visible = True
ActiveWorkbook.Sheets("Data Validation").Visible = True

'Declare variables
    Dim frow As Integer, lrow As Long, fcol As Integer, lcol As Integer, NumSKURows As Integer, sales_Period As Integer
    Dim prod_IDs As String, aproducts_type As String, oprice_list_id As String, category As String, sku As String, channel As String, family As String, manual_sku_list As String, tier As String
    Dim CustomResult As VbMsgBoxResult, AllResult As VbMsgBoxResult
    Dim full_ouput As Range, backup_rng As Range
    Set full_output = Range("B29:AJ100000")
    Set backup_rng = Range("AN30:BS100000")
    
'Define variables
    frow = 30
    fcol = 6
    lcol = 36
    salesPeriod = 18
    
   ' Call Refresh
    
    Sheets("Uploader").Select
    
    entity = Sheets("Uploader").Cells(2, 2).Value
    sku = Sheets("Uploader").Cells(2, 8).Value
    family = Sheets("Uploader").Cells(2, 13).Value
    category = Sheets("Uploader").Cells(2, 18).Value
    channel = Sheets("Uploader").Cells(2, 4).Value
    tier = Sheets("Uploader").Cells(2, 23).Value
    manual_sku_list = Sheets("Uploader").Cells(30, 1).Value
    
    'make sure there are SKUs in the 'table' below or that there is a selection
    If entity <> "" And channel <> "" Then
        If sku = "" Then
            If family = "" Then
                If category = "" Then
                    If tier = "" Then
                        If Cells(30, 2) <> "" Then
                            CustomResult = MsgBox("Do you want to download the custom list of SKUs below for " & channel & "?", vbYesNo + vbQuestion)
                            If CustomResult = vbNo Then
                               'clear previous data
                                full_output.ClearContents
                                full_output.ClearFormats
                                backup_rng.ClearContents
                                End
                            End If
                        ElseIf manual_sku_list = "" Then
                            AllResult = MsgBox("Do you want to download all SKUs for " & channel & "? Note: This will include any new SKUs as well since they don't have any sales channels to filter by.", vbYesNo + vbQuestion)
                            If AllResult = vbNo Then
                                full_output.ClearContents
                                full_output.ClearFormats
                                backup_rng.ClearContents
                                End
                            End If
                        Else: GoTo SkipToHere
                        End If
                    End If
                End If
            End If
        End If
    Else:
        MsgBox ("Please make sure an entity and a channel are selected. Then try again.")
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        End
    End If
    
    
    ActiveWorkbook.Sheets("Data Validation").Activate
    ActiveWorkbook.Sheets("Uploader").Activate
    ActiveWorkbook.Sheets("Data Validation").Visible = False
    
    'clear color
    With Range("B30", Range("AB30").End(xlDown)).Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
    End With
    
'Get data
SkipToHere:
    Call SalesDataDisplay(frow, fcol, lcol, lrow, category, sku, channel, family, manual_sku_list, tier, CustomResult, AllResult)
    Call Format_Data(frow, lrow, lcol)
    
    Range("B29").Select
    
    
ActiveWorkbook.Sheets("Data Validation").Visible = False
ActiveWorkbook.Sheets("Raw Data").Visible = False
ActiveWorkbook.Sheets("Products Query").Visible = False
ActiveWorkbook.Sheets("Uploader").Activate
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

Sub SalesDataDisplay(ByRef frow As Integer, ByRef fcol As Integer, ByRef lcol As Integer, ByRef lrow As Long, ByRef category As String, ByRef sku As String, ByRef channel As String, ByRef family As String, ByRef manual_sku_list As String, ByRef tier As String, ByRef CustomResult As VbMsgBoxResult, ByRef AllResult As VbMsgBoxResult)
Application.ScreenUpdating = False
    
    Dim strSQL As String, prod_IDs As String
    Dim Count As Integer
    Dim table As ListObject
    Dim found_row As Variant
    
    'Dim skuNumbers As Object
    'Set skuNumbers = CreateObject("Scripting.Dictionary")

    Sheets("Uploader").Select
    entity_filter = "(Table_Query_from_hq_1[Entity]='Uploader'!R3C2)"
    channel_filter = "(Table_Query_from_hq_1[Channel]='Uploader'!R2C4)"
    family_filter = "(Table_Query_from_hq_1[Family]='Uploader'!R2C13)"
    category_filter = "(Table_Query_from_hq_1[Category]='Uploader'!R2C18)"
    tier_filter = "(Table_Query_from_hq_1[Tier]='Uploader'!R2C23)"
    
    blank_channel = "+(Table_Query_from_hq_1[Channel] = """")"
    
    If category = "" Then
        If tier = "" Then
            If family = "" Then
                If sku = "" And CustomResult = vbNo Then
                    Sheets("Uploader").Cells(frow, 2).Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(Table_Query_from_hq_1[oproduct_sku]," & channel_filter & "*" & entity_filter & ")),"""")"
                ElseIf Result = vbYes Then
                    GoTo AddSkuInfo
                ElseIf sku = "" And AllResult = vbYes Then
                    Sheets("Uploader").Cells(frow, 2).Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(Table_Query_from_hq_1[oproduct_sku]," & channel_filter & ")),"""")"
                Else: Cells(frow, 2) = sku
                End If
            Else
                Sheets("Uploader").Select
                'pull unique SKUs based on family selection
                Sheets("Uploader").Cells(frow, 2).Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(Table_Query_from_hq_1[oproduct_sku]," & family_filter & "*" & channel_filter & "*" & entity_filter & ")),"""")"
            End If
        Else
            Sheets("Uploader").Select
                'pull unique SKUs based on tier selection
                Sheets("Uploader").Cells(frow, 2).Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(Table_Query_from_hq_1[oproduct_sku]," & tier_filter & "*" & channel_filter & "*" & entity_filter & ")),"""")"
        End If
    Else
        If tier = "" Then
            Sheets("Uploader").Select
            'pull unique SKUs based on category selection
            Sheets("Uploader").Cells(frow, 2).Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(Table_Query_from_hq_1[oproduct_sku]," & category_filter & "*" & channel_filter & "*" & entity_filter & ")),"""")"
        Else: Sheets("Uploader").Cells(frow, 2).Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(Table_Query_from_hq_1[oproduct_sku]," & category_filter & "*" & tier_filter & "*" & channel_filter & "*" & entity_filter & ")),"""")"
        End If
    End If

    
AddSkuInfo:
    'No SKUs: msgbox and macro is done
    'only one SKU: lrow = frow
    'else: count number of SKUs
    If Cells(frow, 2).Value = "" Then
        MsgBox ("No SKU's were pulled for that channel. Double check your selections. Reach out to Analytics if you have further problems.")
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        End
    ElseIf Cells(frow + 1, 2).Value = "" Then
        lrow = frow
    Else
        lrow = Cells(frow, 2).End(xlDown).Row
    End If


    'add product tier
    Range(Cells(frow, 3), Cells(lrow, 3)).FormulaR1C1 = "=INDEX(Products_Query[Tier],MATCH(Uploader!RC[-1],Products_Query[oproduct_sku],0))"
    
    'add product description
    Range(Cells(frow, 4), Cells(lrow, 4)).FormulaR1C1 = "=INDEX(Products_Query[Description],MATCH(Uploader!RC[-2],Products_Query[oproduct_sku],0))"
    
    
    'Populate sales and forecast
    sku_formula = "Table_Query_from_hq_1[oproduct_sku],'Uploader'!RC2,"
    month_formula = "Table_Query_from_hq_1[Month],'Uploader'!R6C,"
    year_formula = "Table_Query_from_hq_1[Year], 'Uploader'!R5C,"
    channel_formula = "Table_Query_from_hq_1[Channel],'Uploader'!R2C4,"
    entity_formula = "Table_Query_from_hq_1[Entity],'Uploader'!R2C2,"
    type_formula = "Table_Query_from_hq_1[Type],'Uploader'!R7C"
    complete_formula = "=SUMIFS(Table_Query_from_hq_1[Sales_Forecast]," & sku_formula & month_formula & year_formula & channel_formula & entity_formula & type_formula & ")"
    Range(Cells(frow, 6), Cells(lrow, lcol)).FormulaR1C1 = complete_formula
    
    'copy over formulas
    Sheets("Uploader").Select
    Range(Cells(frow, 2), Cells(lrow, lcol)).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'copy original forecast data to seperate area
    Sheets("Uploader").Select
    Range(Cells(frow, 25), Cells(lrow, lcol)).Select
    Selection.Copy
    Range("BH30").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'copy original sales data to seperate area
    Sheets("Uploader").Select
    Range(Cells(frow, 6), Cells(lrow, 24)).Select
    Selection.Copy
    Range("AN30").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
End Sub

Sub Format_Data(ByRef frow As Integer, ByRef lrow As Long, ByRef lcol As Integer)
Application.ScreenUpdating = False

    Dim fcol As Integer
    fcol = 1


Sheets("Uploader").Select


'add total category row
    Range("B29").FormulaR1C1 = "Total"
    Application.CutCopyMode = False
    Range("F29").FormulaR1C1 = "=SUM(R[1]C:R" & lrow & "C)"
    Range("F29").AutoFill Destination:=Range("F29:AJ29"), Type:=xlFillDefault
    Range("B29:AJ29").Select
    Selection.Font.Bold = True

'add color to category row
    Range("B29:AJ29").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

'update formulas
    Range("F23").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R29C:R1000C,MATCH(R14C2,R29C2:R1000C2,0))"
    Range("F23").Select
    Selection.AutoFill Destination:=Range("F23:X23"), Type:=xlFillDefault
    Range("F23:X23").Select
    Range("Y23").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]"
    Range("Y24").Select
    ActiveCell.FormulaR1C1 = "=INDEX(R29C:R10000C,MATCH(R14C2,R29C2:R10000C2,0))"
    Range("Y24").Select
    Selection.AutoFill Destination:=Range("Y24:AJ24"), Type:=xlFillDefault
    
'center
    Range("B29:C" & lrow).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("E29:E" & lrow).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("F29:AJ" & lrow).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

'conditional formatting
    Sheets("Uploader").Select
    
    'forecast
    formatRng = Range("Y30", Range("AJ30").End(xlDown)).Select
       
    OffsetValue = Cells(30, Range("BH30").Column).Address(RowAbsolute = True, ColumnAbsolute = True)
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=Y30<>" & OffsetValue
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Color = 8834210
    End With
    
    'sales history
    formatRng = Range("F30", Range("X30").End(xlDown)).Select
       
    OffsetValue = Cells(30, Range("AN30").Column).Address(RowAbsolute = True, ColumnAbsolute = True)
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=F30<>" & OffsetValue
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Color = 255
    End With

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub


