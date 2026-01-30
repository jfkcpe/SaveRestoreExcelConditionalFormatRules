Attribute VB_Name = "SaveRestore"
' Save and Restore Conditional Formatting Rules - Version 0.4 - Jan 30 2026
' Copyright (c) 2026 jfkcpe.  Governed by an MIT license.  See https://github.com/jfkcpe/SaveRestoreExcelConditionalFormatRules
' Use at your own risk - see license.

Sub SaveConditionalFormattingToString()
    Dim ws As Worksheet
    Dim cf As Object
    Dim allRules As String
    Dim ruleCount As Long, intType As Long, i As Long
    Dim allColumns As String, rngHdrName As String, intCol As Variant, stTmp As String
    
    Set ws = ActiveSheet ' We're interested in rules associated with the active worksheet.
    
    Dim rngName As String: rngName = ws.Name & "_CF_RULES"
    
    ' Check if any rules exist in the selection
    If ws.Cells.FormatConditions.Count = 0 Then
        MsgBox "No conditional formatting rules found for worksheet """ & ws.Name & """"
        Exit Sub
    End If
    
    ' Make sure the range with the correct Name has been defined.
    If Not RangeExists(rngName) Then
        MsgBox "You do not have a cell (range) named """ & rngName & """ in which to save the rules."
        Exit Sub
    End If
    
    rngHdrName = ws.Name & "_hdrRow"
    If RangeExists(rngHdrName) Then
        intHdrRow = Range(rngHdrName).Row
    Else
        intHdrRow = 0
    End If
    
    ruleCount = 0

    ' Loop through all format conditions in the worksheet
    For Each cf In ws.Cells.FormatConditions
        ' For this conditional format, render out ever conceivably relevant non-null parameter we think will be needed to re-establish the rule
        ruleCount = ruleCount + 1
        allRules = allRules & "Rule #" & ruleCount & ":" & vbCrLf
        If intHdrRow > 0 Then
            allColumns = "Column(s) Affected:"
            For Each intCol In cf.AppliesTo.Columns
                stTmp = Cells(intHdrRow, intCol.Column).Value
                If InStr(1, allColumns, " " & stTmp & ",") = 0 Then
                    allColumns = allColumns & " " & Cells(intHdrRow, intCol.Column).Value & ","
                End If
            Next intCol
            allRules = allRules & Left(allColumns, Len(allColumns) - 1) & vbCrLf
        End If
        allRules = allRules & "  Applies To: " & cf.AppliesTo.Address & vbCrLf
        allRules = allRules & "  Type ID: " & TypeIntToString(cf.Type) & vbCrLf
        intType = cf.Type
              
        On Error Resume Next ' There WILL be "errors" and there's no elegant way to pre-determine if a condition index (like cf.dupeUnique) exists before accessing it.  Sigh.
       
        If intType = xlColorScale Then ' I believe ColorScale only has these variables
            If Not IsEmpty(cf.ColorScaleCriteria) Then
                allRules = allRules & "  ColorScaleType: " & cf.ColorScaleCriteria.Count & vbCrLf
                For i = 1 To cf.ColorScaleCriteria.Count
                    With cf.ColorScaleCriteria(i)
                        allRules = allRules & "  ColorScaleCriteriaType " & i & ": " & CritTypeIntToString(.Type) & vbCrLf
                        allRules = allRules & "  ScaleColor " & i & ": " & .FormatColor.Color & vbCrLf
                        'If .FormatColor.TintAndShade <> 0 Then allRules = allRules & "  ScaleTintShade " & i & ": " & .FormatColor.TintAndShade & vbCrLf
                        allRules = allRules & "  ColorScaleCriteriaValue " & i & ": " & .Value & vbCrLf
                    End With
                Next i
            End If
        ElseIf intType = xlIconSets Then ' I believe IconSet only has these variables
            allRules = allRules & "  IconSet_ID: " & IconSetIntToString(cf.IconSet.ID) & vbCrLf
            If cf.ReverseOrder <> "" And cf.ReverseOrder <> False Then allRules = allRules & "  ReverseOrder: " & cf.ReverseOrder & vbCrLf
            If cf.ShowIconOnly <> "" And cf.ShowIconOnly <> False Then allRules = allRules & "  ShowIconOnly: " & cf.ShowIconOnly & vbCrLf
            For i = 1 To cf.IconCriteria.Count
                allRules = allRules & "  IconCriteria(" & i & ")_Type: " & CondValTypeIntToString(cf.IconCriteria(i).Type) & vbCrLf
                allRules = allRules & "  IconCriteria(" & i & ")_Value: " & cf.IconCriteria(i).Value & vbCrLf
                allRules = allRules & "  IconCriteria(" & i & ")_Operator: " & OpIntToString(cf.IconCriteria(i).Operator) & vbCrLf
            Next i
        Else ' Not colorScale, not IconSet - everything else
            If cf.Type = xlUniqueValues Then allRules = allRules & "  DupeUnique: " & cf.dupeUnique & vbCrLf
            If cf.Type = xlTop10 Then allRules = allRules & "  TopBottom: " & cf.TopBottom & vbCrLf
            If cf.Type = xlTop10 Then allRules = allRules & "  Rank: " & cf.Rank & vbCrLf
            
            If Not IsEmpty(cf.StopIfTrue) And cf.StopIfTrue Then allRules = allRules & "  StopIfTrue: " & cf.StopIfTrue & vbCrLf
            If Not IsEmpty(cf.Percent) Then allRules = allRules & "  Percent: " & cf.Percent & vbCrLf
            If Not IsEmpty(cf.Operator) Then allRules = allRules & "  Operator: " & OpIntToString(cf.Operator) & vbCrLf
            If Not IsEmpty(cf.TextOperator) Then allRules = allRules & "  TextOperator: " & TxtOpIntToString(cf.TextOperator) & vbCrLf
            If Not IsEmpty(cf.DateOperator) Then allRules = allRules & "  DateOperator: " & DateOpIntToString(cf.DateOperator) & vbCrLf
            If Not IsEmpty(cf.Text) Then allRules = allRules & "  Text: " & cf.Text & vbCrLf
            If Not IsEmpty(cf.formula1) Then allRules = allRules & "  Formula 1: " & cf.formula1 & vbCrLf
            If Not IsEmpty(cf.Formula2) Then allRules = allRules & "  Formula 2: " & cf.Formula2 & vbCrLf
            If cf.Interior.ColorIndex <> xlNone Then allRules = allRules & "  FillColor: " & cf.Interior.Color & vbCrLf
            If cf.Font.ColorIndex <> xlNone Then allRules = allRules & "  FontColor: " & cf.Font.Color & vbCrLf
            If cf.Font.Bold <> xlNone Then allRules = allRules & "  FontBold: " & cf.Font.Bold & vbCrLf
            If cf.Font.Italic <> xlNone Then allRules = allRules & "  FontItalic: " & cf.Font.Italic & vbCrLf
            
            For i = 1 To cf.Borders.Count
                With cf.Borders(i)
                    If Not IsEmpty(.LineStyle) And .LineStyle <> xlNone Then
                        If Not IsEmpty(.Color) Then allRules = allRules & "  Borders.Color " & i & ": " & .Color & vbCrLf
                        If Not IsEmpty(.LineSyle) Then allRules = allRules & "  Borders.LineStyle " & i & ": " & BorderStyleIntToString(.LineStyle) & vbCrLf
                        'If Not IsEmpty(.Weight) Then allRules = allRules & "  Borders.Weight " & i & ": " & BorderWeightIntToString(.Weight) & vbCrLf ' No longer allowed to set border weight
                    End If
                End With
            Next i

        'Else ' IS xlColorScale

        End If
        If Not IsEmpty(cf.StopIfTrue) And cf.StopIfTrue Then allRules = allRules & "  StopIfTrue: " & cf.StopIfTrue & vbCrLf
        ' Excel default with manually created rules is to leave StopIfTrue unchecked (False).
        ' Based on an assumption about programmer's committment to efficiency, Excel's default for most programmatically created rules is to set StopIfTrue to True
        ' We, here, are setting it to False by default, unless the user manually sets it to True before saving.
        On Error GoTo 0
        
        allRules = allRules & vbCrLf
    Next cf
    
    allRules = ruleCount & " Conditional Formatting Rules for tab: """ & ws.Name & """ Saved on " & Format(Now, "mm/dd/yyyy H:mm:ss") & vbCrLf & vbCrLf & allRules

    If ruleCount = 0 Then
        MsgBox "No conditional formatting rules found."
    Else
        Range(rngName).Value = allRules
        MsgBox "Saved " & ruleCount & " rules to string. View Named Range """ & rngName & """ for details."
    End If
End Sub

Sub RecreateConditionalFormattingFromString()
    Dim strTmp As String
    Dim ruleData As String
    Dim ruleArray() As String
    Dim iRule As Variant
    Dim stTrgRange As String
    Dim ruleCount As Long: ruleCount = 0
    Dim stType As String, intType As Long
    Dim stOperator As String, intOperator As Long
    Dim stFormula1 As String, stFormula2 As String
    Dim stFillColor As String, stFontColor As String
    Dim fontBold As String, fontItalic As String
    Dim stColScaleType As String
    
    Dim stDupeUnique As String, stText As String
    Dim stTextOperator As String
    Dim stDateOperator As String, intDateOperator As Long
    Dim stTopBottom As String
    Dim stRank As String
    Dim stPercent As String
    Dim stIconSetID As String, stShowIconOnly As String, stReverseOrder As String
    Dim stStopIfTrue As String, boolStopIfTrue As Boolean
    Dim doFontsBorders As Boolean

    
    Dim ws As Worksheet
    Dim cf As Object
       
    Set ws = ActiveSheet ' We're interested in rules associated with the active worksheet.
    
    Dim rngName As String: rngName = ws.Name & "_CF_RULES"
    
    ' Make sure the range with the correct Name has been defined.
    If Not RangeExists(rngName) Then
        MsgBox "You do not have a cell (range) named """ & rngName & """"
        Exit Sub
    End If
    
    ruleData = Range(rngName).Value
    
    ActiveSheet.Cells.FormatConditions.Delete
    
    ' Use "Rule #" as the Rule Block delimiter based on the Save macro's output
    ruleArray = Split(ruleData, "Rule #")
    
    ' Iterate through the array (skip index 0)
    Dim stR As String
    Dim i As Integer, j As Integer
    For i = 1 To UBound(ruleArray)
        
        'Re-initialize all the arrays - some rules set fewer indexed values that others; don't want left-over schmutz
        Dim arrstrScaleCol(3) As String
        'Dim arrstrScaleTintShade(3) As String ' Do not do TintAndShade: It will be saved and will continue to modify colors with each Save/Restore cycle!
        Dim arrstrColScaleCritType(3) As String: Dim arrlngColScaleCritType(3) As Long
        Dim arrvarColScaleCritValue(3) As String
        
        Dim arBorders_LineStyle(32) As String ' Not sure how high border indices can theoretically go.
        Dim arBorders_Color(32) As String
        'Dim arBorders_Weight(32) As String
        
        Dim arrstIconSetType(32) As String
        Dim arrstIconSetValue(32) As String
        Dim arrstIconSetOperator(32) As String
        
        iRule = ruleArray(i)
        iRule = Replace(iRule, vbCrLf, "|") ' Line feed characters are handled inconsistently, could be CrLf, Cr, or Lf.
        iRule = Replace(iRule, vbLf, "|")
        iRule = Replace(iRule, vbCr, "|")
        iRule = iRule & "|"
        
        stTrgRange = ParseMyParm(iRule, "Applies To") ' Always present
        stType = ParseMyParm(iRule, "Type ID"): intType = TypeStringToInt(stType) ' Always present
        
        stOperator = ParseMyParm(iRule, "Operator"): If stOperator <> "" Then intOperator = OpStringToInt(stOperator)
        stTextOperator = ParseMyParm(iRule, "TextOperator")
        stDateOperator = ParseMyParm(iRule, "DateOperator"): If stDateOperator <> "" Then intDateOperator = DateOpStringToInt(stDateOperator) ' -9999
        stText = ParseMyParm(iRule, "Text")
        stFormula1 = ParseMyParm(iRule, "Formula 1")
        stFormula2 = ParseMyParm(iRule, "Formula 2")
        stFillColor = ParseMyParm(iRule, "FillColor")
        stFontColor = ParseMyParm(iRule, "FontColor")
        stFontBold = ParseMyParm(iRule, "FontBold")
        stFontItalic = ParseMyParm(iRule, "FontItalic")
        stColScaleType = ParseMyParm(iRule, "ColorScaleType")
        For j = 1 To 3
            arrstrColScaleCritType(j) = ParseMyParm(iRule, "ColorScaleCriteriaType " & j): arrlngColScaleCritType(j) = CritTypeStringToInt(arrstrColScaleCritType(j)) '-9999'
            arrstrScaleCol(j) = ParseMyParm(iRule, "ScaleColor " & j)
            'arrstrScaleTintShade(j) = ParseMyParm(iRule, "ScaleTintShade " & j) ' Do not do TintAndShade: It will be saved and will continue to modify colors with each Save/Restore cycle!
            strTmp = ParseMyParm(iRule, "ColorScaleCriteriaValue " & j)
            If strTmp <> "" Then arrvarColScaleCritValue(j) = ParseMyParm(iRule, "ColorScaleCriteriaValue " & j)
        Next j
        stDupeUnique = ParseMyParm(iRule, "DupeUnique")
        stTopBottom = ParseMyParm(iRule, "TopBottom")
        stRank = ParseMyParm(iRule, "Rank")
        stPercent = ParseMyParm(iRule, "Percent")
        
        For j = 1 To 4
            arBorders_LineStyle(j) = ParseMyParm(iRule, "Borders.LineStyle " & j) ' There is a known gap in documentation on border values & constants.  e.g. xlTop vs xlEdgeTop
            arBorders_Color(j) = ParseMyParm(iRule, "Borders.Color " & j)
            'arBorders_Weight(j) = ParseMyParm(iRule, "Borders.Weight " & j) ' No longer allowed
        Next j
        
        stIconSetID = ParseMyParm(iRule, "IconSet_ID")
        stShowIconOnly = ParseMyParm(iRule, "ShowIconOnly")
        
        For j = 1 To 5 ' I think 5 is the max number of iconcriteria
            arrstIconSetType(j) = ParseMyParm(iRule, "IconCriteria(" & j & ")_Type")
            arrstIconSetValue(j) = ParseMyParm(iRule, "IconCriteria(" & j & ")_Value")
            arrstIconSetOperator(j) = ParseMyParm(iRule, "IconCriteria(" & j & ")_Operator")
        Next j
        
        stStopIfTrue = ParseMyParm(iRule, "StopIfTrue"): boolStopIfTrue = IIf(stStopIfTrue = "True", True, False) ' IIf, a nice compact alternative to If Then Else!
        
        On Error Resume Next ' Comment this line out if you are debugging.
        
        doFontsBorders = True
        
        'Different conditions and types require different sets of arguments at Add time.
        'The order of this sequence was derived empirically and is a bit fussy.
        If intType = xlTop10 Then
            Set cf = ws.Range(stTrgRange).FormatConditions.AddTop10
            With cf
                .TopBottom = stTopBottom
                .Rank = stRank
                .Percent = stPercent
            End With
        ElseIf intType = xlTimePeriod And intDateOperator <> -9999 Then
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=intType, DateOperator:=intDateOperator)
        ElseIf intType = xlColorScale Then ' these parms are just for xlColorScale and other parms are not relevant
            Set cf = ws.Range(stTrgRange).FormatConditions.AddColorScale(colorScaleType:=CLng(stColScaleType))
            For j = 1 To cf.ColorScaleCriteria.Count
                With cf.ColorScaleCriteria(j)
                    .Type = arrlngColScaleCritType(j)
                    If arrvarColScaleCritValue(j) <> 0 And .Type <> xlConditionValueLowestValue And .Type <> xlConditionValueHighestValue Then
                        .Value = CDbl(arrvarColScaleCritValue(j))
                    End If
                    .FormatColor.Color = CLng(arrstrScaleCol(j))
                    'If arrstrScaleTintShade(j) <> "" Then .FormatColor.TintAndShade = CDbl(arrstrScaleTintShade(j)) ' Do not do TintAndShade: It will be saved and will continue to modify colors with each Save/Restore cycle!
                End With
            Next j
            doFontsBorders = False
        ElseIf intType = 6 Then
            Set cf = ws.Range(stTrgRange).FormatConditions.AddIconSetCondition
            cf.IconSet = ActiveWorkbook.IconSets(IconSetStringToInt(stIconSetID))
            cf.ReverseOrder = CBool(stReverseOrder)
            cf.ShowIconOnly = CBool(stShowIconOnly)
            For j = 2 To 5
                cf.IconCriteria.item(j).Type = CDbl(arrstIconSetType(j))
                cf.IconCriteria.item(j).Value = CDbl(arrstIconSetValue(j))
                cf.IconCriteria.item(j).Operator = CDbl(arrstIconSetOperator(j))
            Next j
            doFontsBorders = False
        
        ElseIf stTextOperator <> "" And stText <> "" Then
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=xlTextString, String:=stText, TextOperator:=TxtOpStringToInt(stTextOperator))
        ElseIf stText <> "" And intType = xlTextString Then
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=xlTextString, String:=stText)
        ElseIf stOperator <> "" Then
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=intType, Operator:=intOperator, formula1:=stFormula1, Formula2:=stFormula2)


        Else
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=intType, formula1:=stFormula1, Formula2:=stFormula2)
        End If
        
        If doFontsBorders Then
            With cf
                If intType <> xlColorScale And intType <> xlIconSets Then
                    If stFillColor <> "" And stFillColor <> "None" Then .Interior.Color = CLng(stFillColor)
                    If stFontColor <> "" And stFontColor <> "None" Then .Font.Color = CLng(stFontColor)
                    If fontBold <> "" And fontBold <> "None" Then .Font.Bold = CBool(fontBold)
                    If fontItalic <> "" And fontItalic <> "None" Then .Font.Italic = CBool(fontItalic)
                    If stDupeUnique <> "" Then
                        .dupeUnique = CLng(stDupeUnique)
                    End If
    
                    For j = 1 To 4
                        If arBorders_LineStyle(j) <> "" Then .Borders(j).LineStyle = BorderStyleStringToInt(arBorders_LineStyle(j))
                        If arBorders_Color(j) <> "" Then .Borders(j).Color = CLng(arBorders_Color(j))
                        'If arBorders_Weight(j) <> "" Then .Borders(j).Weight = BorderWeightStringToInt(arBorders_Weight(j)) ' Excel no longer allows Border Weight other than the default
                    Next j
                    .StopIfTrue = boolStopIfTrue ' which WE will set to False by default, unlike Excel's default for programmatically created rules.
                End If
            End With
        End If
        
        On Error GoTo 0
        ruleCount = ruleCount + 1
    Next i
    
    MsgBox ws.Cells.FormatConditions.Count & " rules successfully created out of " & UBound(ruleArray) & " specified in cell with Name """ & rngName & """", vbInformation
End Sub

Function ParseMyParm(singleRule As Variant, stV As String) As String
    Dim stR As String
    stV = stV & ":"
    If InStr(singleRule, stV) > 0 Then
        stR = Mid(singleRule, InStr(singleRule, stV) + Len(stV))
        stR = Trim(Left(stR, InStr(stR, "|") - 1))
    Else
        stR = ""
    End If
    ParseMyParm = stR
End Function

Function RangeExists(rngName As String) As Boolean
    ' This function returns True if the named range exists, False otherwise.
    On Error Resume Next ' Ignore errors if the name doesn't exist - not elegant, but it works.
    RangeExists = False
    RangeExists = Len(ThisWorkbook.Names(rngName).Name) <> 0
End Function

Function TypeIntToString(intType As Long) As String
' https://www.google.com/search?q=excel+macro+Conditional+Formatting+type+integer+equivalents+for+xl+constants
    Select Case intType
        Case 1: TypeIntToString = "xlCellValue"
        Case 2: TypeIntToString = "xlExpression"
        Case 3: TypeIntToString = "xlColorScale"
        Case 4: TypeIntToString = "xlDataBar"
        Case 5: TypeIntToString = "xlTop10"
        Case 6: TypeIntToString = "xlIconSet"
        Case 8: TypeIntToString = "xlUniqueValues"
        Case 9: TypeIntToString = "xlTextString"
        Case 10: TypeIntToString = "xlBlanksCondition"
        Case 11: TypeIntToString = "xlTimePeriod"
        Case 12: TypeIntToString = "xlAboveAverageCondition"
        Case 13: TypeIntToString = "xlNoBlanksCondition"
        Case 16: TypeIntToString = "xlErrorsCondition"
        Case 17: TypeIntToString = "xlNoErrorsCondition"
        Case Else: TypeIntToString = ""
    End Select
End Function

Function TypeStringToInt(stType As String) As Long
    Select Case stType
        Case "xlCellValue": TypeStringToInt = 1
        Case "xlExpression": TypeStringToInt = 2
        Case "xlColorScale": TypeStringToInt = 3
        Case "xlDataBar": TypeStringToInt = 4
        Case "xlTop10": TypeStringToInt = 5
        Case "xlIconSet": TypeStringToInt = 6
        Case "xlUniqueValues": TypeStringToInt = 8
        Case "xlTextString": TypeStringToInt = 9
        Case "xlBlanksCondition": TypeStringToInt = 10
        Case "xlTimePeriod": TypeStringToInt = 11
        Case "xlAboveAverageCondition": TypeStringToInt = 12
        Case "xlNoBlanksCondition": TypeStringToInt = 13
        Case "xlErrorsCondition": TypeStringToInt = 16
        Case "xlNoErrorsCondition": TypeStringToInt = 17
        Case Else: TypeStringToInt = -9999 ' zero is used in some Excel constants (like Conditional Formatting xlEqual operator)
    End Select
End Function

Function OpIntToString(intOp As Long) As String
    Select Case intOp
        Case 1: OpIntToString = "xlBetween"
        Case 2: OpIntToString = "xlNotBetween"
        Case 3: OpIntToString = "xlEqual"
        Case 4: OpIntToString = "xlNotEqual"
        Case 5: OpIntToString = "xlGreater"
        Case 7: OpIntToString = "xlGreaterEqual"
        Case 6: OpIntToString = "xlLess"
        Case 8: OpIntToString = "xlLessEqual"
        Case Else: OpIntToString = ""
    End Select
End Function

Function OpStringToInt(strOp As String) As Long
    Select Case strOp
        Case "xlBetween": OpStringToInt = 1
        Case "xlNotBetween": OpStringToInt = 2
        Case "xlEqual": OpStringToInt = 3
        Case "xlNotEqual": OpStringToInt = 4
        Case "xlGreater": OpStringToInt = 5
        Case "xlGreaterEqual": OpStringToInt = 7
        Case "xlLess": OpStringToInt = 6
        Case "xlLessEqual": OpStringToInt = 8
        Case Else: OpStringToInt = -9999
    End Select
End Function

Function CritTypeIntToString(intType As Long) As String
    Select Case intType
        Case 1: CritTypeIntToString = "xlConditionalValueLowestValue"
        Case 2: CritTypeIntToString = "xlConditionalValueHighestValue"
        Case 4: CritTypeIntToString = "xlCondtionalValueFormula"
        Case 5: CritTypeIntToString = "xlConditionalValuePercentile"
        Case Else: CritTypeIntToString = ""
    End Select
End Function

Function CritTypeStringToInt(stType As String) As Long
    Select Case stType
        Case "xlConditionalValueLowestValue": CritTypeStringToInt = 1
        Case "xlConditionalValueHighestValue": CritTypeStringToInt = 2
        Case "xlCondtionalValueFormula": CritTypeStringToInt = 4
        Case "xlConditionalValuePercentile": CritTypeStringToInt = 5
        Case Else: CritTypeStringToInt = -9999
    End Select
End Function

Function TxtOpIntToString(intTxtOp As Long) As String
    Select Case intTxtOp
        Case 0: TxtOpIntToString = "xlContains"
        Case 1: TxtOpIntToString = "xlDoesNotContain"
        Case 2: TxtOpIntToString = "xlBeginsWith"
        Case 3: TxtOpIntToString = "xlEndsWith"
        Case Else: TxtOpIntToString = ""
    End Select
End Function

Function TxtOpStringToInt(stTxtOp As String) As Long
    Select Case stTxtOp
        Case "xlContains": TxtOpStringToInt = 0
        Case "xlDoesNotContain": TxtOpStringToInt = 1
        Case "xlBeginsWith": TxtOpStringToInt = 2
        Case "xlEndsWith": TxtOpStringToInt = 3
        Case Else: TxtOpStringToInt = -9999
    End Select
End Function

Function DateOpIntToString(intDateOp As Long) As String
    Select Case intDateOp
        Case 0: DateOpIntToString = "xlToday"
        Case 1: DateOpIntToString = "xlYesterday"
        Case 2: DateOpIntToString = "xlLastWeek"
        Case 3: DateOpIntToString = "xlThisWeek"
        Case 4: DateOpIntToString = "xlNextWeek"
        Case 5: DateOpIntToString = "xlLastMonth"
        Case 6: DateOpIntToString = "xlThisMonth"
        Case 7: DateOpIntToString = "xlNextMonth"
        Case 8: DateOpIntToString = "xlTomorrow"
        Case 9: DateOpIntToString = "xlNextMonth"
        Case Else: DateOpIntToString = ""
    End Select
End Function

Function DateOpStringToInt(stDateOp As String) As Long
    Select Case stDateOp
        Case "xlToday": DateOpStringToInt = 0
        Case "xlYesterday": DateOpStringToInt = 1
        Case "xlLastWeek": DateOpStringToInt = 2
        Case "xlThisWeek": DateOpStringToInt = 3
        Case "xlNextWeek": DateOpStringToInt = 4
        Case "xlLastMonth": DateOpStringToInt = 5
        Case "xlThisMonth": DateOpStringToInt = 6
        Case "xlNextMonth": DateOpStringToInt = 7
        Case "xlTomorrow": DateOpStringToInt = 8
        Case "xlNextMonth": DateOpStringToInt = 9
        Case Else: DateOpStringToInt = -9999
    End Select
End Function

'Function TopBotIntToString(intTopBot As Long) As String ' Not used?
'    Select Case intTopBot
'        Case 0: TopBotIntToString = "xlTop10Top"
'        Case 1: TopBotIntToString = "xlTop10Bottom"
'        Case -9999: TopBotIntToString = ""
'        Case Else: TopBotIntToString = ""
'    End Select
'End Function

'Function TopBotStringToInt(stTopBot As String) As Long ' Not used?
'    Select Case stTopBot
'        Case "xlTop10Top": TopBotStringToInt = 0
'        Case "xlTop10Bottom": TopBotStringToInt = 1
'        Case Else: TopBotStringToInt = -9999
'    End Select
'End Function

'Function BorderTypeStringToInt(stTest As String) As Long ' Not Used?
'  Select Case stTest
'    Case "xlDiagonalDown": BorderTypeStringToInt = 5
'    Case "xlDiagonalUp": BorderTypeStringToInt = 6
'    Case "xlEdgeLeft": BorderTypeStringToInt = 7
'    Case "xlEdgeTop": BorderTypeStringToInt = 8
'    Case "xlEdgeBottom": BorderTypeStringToInt = 9
'    Case "xlEdgeRight": BorderTypeStringToInt = 10
'    Case "xlInsideVertical": BorderTypeStringToInt = 11
'    Case "xlInsideHorizontal": BorderTypeStringToInt = 12
'    Case Else: BorderTypeStringToInt = -9999
'  End Select
'End Function

'Function BorderTypeIntToString(intTest As Long) As String ' Not used?
'  Select Case intTest
'    Case 5: BorderTypeIntToString = "xlDiagonalDown"
'    Case 6: BorderTypeIntToString = "xlDiagonalUp"
'    Case 7: BorderTypeIntToString = "xlEdgeLeft"
'    Case 8: BorderTypeIntToString = "xlEdgeTop"
'    Case 9: BorderTypeIntToString = "xlEdgeBottom"
'    Case 10: BorderTypeIntToString = "xlEdgeRight"
'    Case 11: BorderTypeIntToString = "xlInsideVertical"
'    Case 12: BorderTypeIntToString = "xlInsideHorizontal"
'    Case Else: BorderTypeIntToString = ""
'  End Select
'End Function
 
Function BorderStyleIntToString(intTest As Long) As String
  Select Case intTest
    Case 1: BorderStyleIntToString = "xlContinuous"
    Case -4115: BorderStyleIntToString = "xlDash"
    Case 4: BorderStyleIntToString = "xlDashDot"
    Case 5: BorderStyleIntToString = "xlDashDotDot"
    Case -4118: BorderStyleIntToString = "xlDot"
    Case -4119: BorderStyleIntToString = "xlDouble"
    Case -4142: BorderStyleIntToString = "xlLineStyleNone"
    Case 13: BorderStyleIntToString = "xlSlantDashDot"
    Case Else: BorderStyleIntToString = ""
  End Select
End Function
Function BorderStyleStringToInt(stTest As String) As Long
  Select Case stTest
    Case "xlContinuous": BorderStyleStringToInt = 1
    Case "xlDash": BorderStyleStringToInt = -4115
    Case "xlDashDot": BorderStyleStringToInt = 4
    Case "xlDashDotDot": BorderStyleStringToInt = 5
    Case "xlDot": BorderStyleStringToInt = -4118
    Case "xlDouble": BorderStyleStringToInt = -4119
    Case "xlLineStyleNone": BorderStyleStringToInt = -4142
    Case "xlSlantDashDot": BorderStyleStringToInt = 13
    Case Else: BorderStyleStringToInt = -9999
  End Select
End Function



Function CondValTypeIntToString(intTest As Long) As String
  Select Case intTest
    Case 7: CondValTypeIntToString = "xlConditionValueAutomaticMax"
    Case 6: CondValTypeIntToString = "xlConditionValueAutomaticMin"
    Case 4: CondValTypeIntToString = "xlConditionValueFormula"
    Case 2: CondValTypeIntToString = "xlConditionValueHighestValue"
    Case 1: CondValTypeIntToString = "xlConditionValueLowestValue"
    Case -1: CondValTypeIntToString = "xlConditionValueNone"
    Case 0: CondValTypeIntToString = "xlConditionValueNumber"
    Case 3: CondValTypeIntToString = "xlConditionValuePercent"
    Case 5: CondValTypeIntToString = "xlConditionValuePercentile"
    Case Else: CondValTypeIntToString = ""
  End Select
End Function

'Function CondValTypeStringToInt(stTest As String) As Long ' Not used
'  Select Case stTest
'    Case "xlConditionValueAutomaticMax": CondValTypeStringToInt = 7
'    Case "xlConditionValueAutomaticMin": CondValTypeStringToInt = 6
'    Case "xlConditionValueFormula": CondValTypeStringToInt = 4
'    Case "xlConditionValueHighestValue": CondValTypeStringToInt = 2
'    Case "xlConditionValueLowestValue": CondValTypeStringToInt = 1
'    Case "xlConditionValueNone": CondValTypeStringToInt = -1
'    Case "xlConditionValueNumber": CondValTypeStringToInt = 0
'    Case "xlConditionValuePercent": CondValTypeStringToInt = 3
'    Case "xlConditionValuePercentile": CondValTypeStringToInt = 5
'    Case Else: CondValTypeStringToInt = -9999
'  End Select
'End Function

Function IconSetIntToString(intTest As Long) As String
  Select Case intTest
    Case 1: IconSetIntToString = "xl3Arrows"
    Case 2: IconSetIntToString = "xl3ArrowsGray"
    Case 3: IconSetIntToString = "xl3Flags"
    Case 4: IconSetIntToString = "xl3TrafficLights1"
    Case 5: IconSetIntToString = "xl3TrafficLights2"
    Case 6: IconSetIntToString = "xl3Signs"
    Case 7: IconSetIntToString = "xl3Symbols"
    Case 9: IconSetIntToString = "xl4Arrows"
    Case 10: IconSetIntToString = "xl4ArrowsGray"
    Case 11: IconSetIntToString = "xl4RedToBlack"
    Case 12: IconSetIntToString = "xl4CRV"
    Case 13: IconSetIntToString = "xl4TrafficLights"
    Case 14: IconSetIntToString = "xl5Arrows"
    Case 15: IconSetIntToString = "xl5ArrowsGray"
    Case 16: IconSetIntToString = "xl5CRV"
    Case 17: IconSetIntToString = "xl5Quarters"

    Case Else: IconSetIntToString = ""
  End Select
End Function

Function IconSetStringToInt(stTest As String) As Long
  Select Case stTest
    Case "xl3Arrows": IconSetStringToInt = 1
    Case "xl3ArrowsGray": IconSetStringToInt = 2
    Case "xl3Flags": IconSetStringToInt = 3
    Case "xl3TrafficLights1": IconSetStringToInt = 4
    Case "xl3TrafficLights2": IconSetStringToInt = 5
    Case "xl3Signs": IconSetStringToInt = 6
    Case "xl3Symbols": IconSetStringToInt = 7
    Case "xl4Arrows": IconSetStringToInt = 9
    Case "xl4ArrowsGray": IconSetStringToInt = 10
    Case "xl4RedToBlack": IconSetStringToInt = 11
    Case "xl4CRV": IconSetStringToInt = 12
    Case "xl4TrafficLights": IconSetStringToInt = 13
    Case "xl5Arrows": IconSetStringToInt = 14
    Case "xl5ArrowsGray": IconSetStringToInt = 15
    Case "xl5CRV": IconSetStringToInt = 16
    Case "xl5Quarters": IconSetStringToInt = 17

    Case Else: IconSetStringToInt = -9999
  End Select
End Function

