Attribute VB_Name = "SaveRestore"
' Save and Restore Conditional Formatting Rules - Version 0.2
' Copyright (c) 2026 jfkcpe (GitHub ID).  Governed by an MIT license.  See https://github.com/jfkcpe/SaveRestoreExcelConditionalFormatRules
' Use at your own risk - see license.

Sub SaveConditionalFormattingToString()
    Dim ws As Worksheet
    Dim cf As Object
    Dim allRules As String
    Dim ruleCount As Long, intType As Long, i As Long
    
    Set ws = ActiveSheet ' We're interested in rules associated with the active worksheet.
    
    Dim rngName As String: rngName = ws.Name & "_CF_RULES"
    
    ' Check if any rules exist in the selection
    If ws.Cells.FormatConditions.Count = 0 Then
        MsgBox "No conditional formatting rules found for worksheet """ & ws.Name & """"
        Exit Sub
    End If
    
    ' Make sure the range with the correct Name has been defined.
    If Not RangeExists(rngName) Then
        MsgBox "You do not have a cell (range) named """ & rngName & """"
        Exit Sub
    End If
    
    ruleCount = 0

    ' Loop through all format conditions in the worksheet
    For Each cf In ws.Cells.FormatConditions
        ' For this conditional format, render out ever conceivably relevant non-null parameter we think will be needed to re-establish the rule
        ruleCount = ruleCount + 1
        allRules = allRules & "Rule #" & ruleCount & ":" & vbCrLf
        allRules = allRules & "  Applies To: " & cf.AppliesTo.Address & vbCrLf
        allRules = allRules & "  Type ID: " & TypeIntToString(cf.Type) & vbCrLf
        intType = cf.Type
              
        On Error Resume Next
        
        If intType <> xlColorScale Then
            If Not IsEmpty(cf.dupeUnique) Then allRules = allRules & "  DupeUnique: " & cf.dupeUnique & vbCrLf
            If Not IsEmpty(cf.TopBottom) Then allRules = allRules & "  TopBottom: " & cf.TopBottom & vbCrLf
            If Not IsEmpty(cf.Rank) Then allRules = allRules & "  Rank: " & cf.Rank & vbCrLf
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
            
            'If cf.Borders.LineStyle <> xlNone Then
            For i = 1 To cf.Borders.Count
                With cf.Borders(i)
                    If Not IsEmpty(.LineStyle) And .LineStyle <> xlNone Then
                        If Not IsEmpty(.Color) Then allRules = allRules & "  Borders.Color " & i & ": " & .Color & vbCrLf
                        If Not IsEmpty(.LineSyle) Then allRules = allRules & "  Borders.LineStyle " & i & ": " & .LineStyle & vbCrLf
                        If Not IsEmpty(.Weight) Then allRules = allRules & "  Borders.Weight " & i & ": " & .Weight & vbCrLf
                    End If
                End With
            Next i
            'End If

        Else ' IS xlColorScale
            If Not IsEmpty(cf.ColorScaleCriteria) Then
                allRules = allRules & "  ColorScaleType: " & cf.ColorScaleCriteria.Count & vbCrLf
                For i = 1 To cf.ColorScaleCriteria.Count
                    allRules = allRules & "  ColorScaleCriteriaType " & i & ": " & CritTypeIntToString(cf.ColorScaleCriteria(i).Type) & vbCrLf
                    allRules = allRules & "  ScaleColor " & i & ": " & cf.ColorScaleCriteria(i).FormatColor.Color & vbCrLf
                    allRules = allRules & "  ScaleTintShade " & i & ": " & cf.ColorScaleCriteria(i).FormatColor.TintAndShade & vbCrLf
                Next i
            End If
        End If
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
    Dim stScaleCol1 As String, stScaleCol2 As String, stScaleCol3 As String
    Dim stScaleTintShare1 As String, stScaleTintShare2 As String
    Dim stColScaleCritType1 As String, intColScaleCritType1 As Long
    Dim stColScaleCritType2 As String, intColScaleCritType2 As Long
    Dim stColScaleCritType3 As String, intColScaleCritType3 As Long
    Dim stDupeUnique As String, stText As String
    Dim stTextOperator As String
    Dim stDateOperator As String, intDateOperator As Long
    Dim stTopBottom As String
    Dim stRank As String
    Dim stPercent As String
    
    Dim arBorders_LineStyle(32) As String ' Not sure how high border indices can theoretically go.
    Dim arBorders_Color(32) As String
    Dim arBorders_Weight(32) As String
    
    Dim stBorders_LineStyle As String
    Dim stBorders_Weight As String
    
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
        stColScaleCritType1 = ParseMyParm(iRule, "ColorScaleCriteriaType 1"): intColScaleCritType1 = CritTypeStringToInt(stColScaleCritType1) ' -9999'
        stColScaleCritType2 = ParseMyParm(iRule, "ColorScaleCriteriaType 2"): intColScaleCritType2 = CritTypeStringToInt(stColScaleCritType2) ' -9999'
        stColScaleCritType3 = ParseMyParm(iRule, "ColorScaleCriteriaType 2"): intColScaleCritType3 = CritTypeStringToInt(stColScaleCritType3) ' -9999
        stColScaleType = ParseMyParm(iRule, "ColorScaleType")
        stScaleCol1 = ParseMyParm(iRule, "ScaleColor 1")
        stScaleCol2 = ParseMyParm(iRule, "ScaleColor 2")
        stScaleCol3 = ParseMyParm(iRule, "ScaleColor 3")
        stScaleTintShade1 = ParseMyParm(iRule, "ScaleTintShade 1")
        stScaleTintShade2 = ParseMyParm(iRule, "ScaleTintShade 2")
        stScaleTintShade3 = ParseMyParm(iRule, "ScaleTintShade 3")
        stDupeUnique = ParseMyParm(iRule, "DupeUnique")
        stTopBottom = ParseMyParm(iRule, "TopBottom")
        stRank = ParseMyParm(iRule, "Rank")
        stPercent = ParseMyParm(iRule, "Percent")
        
        For j = 1 To 4
            arBorders_LineStyle(j) = ParseMyParm(iRule, "Borders.Linestyle " & j) ' There is a known gap in documentation on border values & constants.  e.g. xlTop vs xlEdgeTop
            arBorders_Color(j) = ParseMyParm(iRule, "Borders.Color " & j)
            arBorders_Weight(j) = ParseMyParm(iRule, "Borders.Weight " & j)
        Next j
        
        On Error Resume Next ' Comment this line out if you are debugging.
        
        'Different conditions and types require different sets of arguments at Add time.
        'The order of this sequence was derived empirically and is a bit fussy.
        If intType = xlTop10 Then
            Set cf = ws.Range(stTrgRange).FormatConditions.AddTop10
            With cf
                .TopBottom = stTopBottom
                .Rank = stRank
                .Percent = stPercent
            End With
        ElseIf stTextOperator <> "" And stText <> "" Then
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=xlTextString, String:=stText, TextOperator:=TxtOpStringToInt(stTextOperator))
        ElseIf stText <> "" And intType = xlTextString Then
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=xlTextString, String:=stText)
        ElseIf intType = xlTimePeriod And intDateOperator <> -9999 Then
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=intType, DateOperator:=intDateOperator)
        ElseIf stOperator <> "" Then
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=intType, Operator:=intOperator, formula1:=stFormula1, Formula2:=stFormula2)
        ElseIf intType = xlColorScale Then ' these parms are just for xlColorScale and other parms are not relevant
            Set cf = ws.Range(stTrgRange).FormatConditions.AddColorScale(colorScaleType:=CLng(stColScaleType))
            With cf
                .ColorScaleCriteria(1).FormatColor.Color = CLng(stScaleCol1)
                .ColorScaleCriteria(1).FormatColor.TintAndShade = CLng(stScaleTintShade1)
                If intColScaleCrit1 <> -9999 Then .ColorScaleCriteria(1).Type = intColScaleCritType1
                .ColorScaleCriteria(2).FormatColor.Color = CLng(stScaleCol2)
                .ColorScaleCriteria(2).FormatColor.TintAndShade = CLng(stScaleTintShade2)
                If intColScaleCrit2 <> -9999 Then .ColorScaleCriteria(2).Type = intColScaleCritType2
                If stColScaleType = "3" Then
                    .ColorScaleCriteria(3).FormatColor.Color = CLng(stScaleCol3)
                    .ColorScaleCriteria(3).FormatColor.TintAndShade = CLng(stScaleTintShade3)
                    If intColScaleCrit3 <> -9999 Then .ColorScaleCriteria(3).Type = intColScaleCritType3
                End If
            End With
        Else
            Set cf = ws.Range(stTrgRange).FormatConditions.Add(Type:=intType, formula1:=stFormula1, Formula2:=stFormula2)
        End If
        
        With cf
            If intType <> xlColorScale Then
                If stFillColor <> "" And stFillColor <> "None" Then .Interior.Color = CLng(stFillColor)
                If stFontColor <> "" And stFontColor <> "None" Then .Font.Color = CLng(stFontColor)
                If fontBold <> "" And fontBold <> "None" Then .Font.Bold = CBool(fontBold)
                If fontItalic <> "" And fontItalic <> "None" Then .Font.Italic = CBool(fontItalic)
                If stDupeUnique <> "" Then
                    .dupeUnique = CLng(stDupeUnique)
                End If

                For j = 1 To 4
                    If arBorders_LineStyle(j) <> "" Then .Borders(j).LineStyle = CLng(arBorders_LineStyle(j))
                    If arBorders_Color(j) <> "" Then .Borders(j).Color = CLng(arBorders_Color(j))
                    If arBorders_Weight(j) <> "" Then .Borders(j).Weight = CLng(arBorders_Weight(j))
                Next j
            End If
        End With
        On Error GoTo 0
        ruleCount = ruleCount + 1
    Next i
    
    'Some day, I'd like to do better error handling in case the "Resume Next" block above fails silently.
    MsgBox ws.Cells.FormatConditions.Count & " rules successfully created out of " & UBound(ruleArray) & " specified in cell with Name """ & rngName & """", vbInformation
End Sub


