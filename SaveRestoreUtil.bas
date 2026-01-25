Attribute VB_Name = "SaveRestoreUtil"
' Save and Restore Conditional Formatting Rules - Utility/Helper Functions - Version 0.2
' Copyright (c) 2026 jfkcpe (GitHub ID).  Governed by an MIT license.  See https://github.com/jfkcpe/SaveRestoreExcelConditionalFormatRules
' Use at your own risk - see license.

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

        Case -9999: TypeIntToString = ""
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
        Case -9999: OpIntToString = ""
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
        Case -9999: CritTypeIntToString = ""
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
        Case -9999: TxtOpIntToString = ""
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

        Case -9999: DateOpIntToString = ""
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

Function TopBotIntToString(intTopBot As Long) As String
    Select Case intTopBot
        Case 0: TopBotIntToString = "xlTop10Top"
        Case 1: TopBotIntToString = "xlTop10Bottom"
        Case -9999: TopBotIntToString = ""
        Case Else: TopBotIntToString = ""
    End Select
End Function

Function TopBotStringToInt(stTopBot As String) As Long
    Select Case stTopBot
        Case "xlTop10Top": TopBotStringToInt = 0
        Case "xlTop10Bottom": TopBotStringToInt = 1
        Case Else: TopBotStringToInt = -9999
    End Select
End Function

Function BorderTypeStringToInt(stTest As String) As Long
  Select Case stTest
    Case "xlDiagonalDown": BorderTypeStringToInt = 5
    Case "xlDiagonalUp": BorderTypeStringToInt = 6
    Case "xlEdgeLeft": BorderTypeStringToInt = 7
    Case "xlEdgeTop": BorderTypeStringToInt = 8
    Case "xlEdgeBottom": BorderTypeStringToInt = 9
    Case "xlEdgeRight": BorderTypeStringToInt = 10
    Case "xlInsideVertical": BorderTypeStringToInt = 11
    Case "xlInsideHorizontal": BorderTypeStringToInt = 12
  End Select
End Function

Function BorderTypeIntToString(intTest As Long) As String
  Select Case intTest
    Case 5: BorderTypeIntToString = "xlDiagonalDown"
    Case 6: BorderTypeIntToString = "xlDiagonalUp"
    Case 7: BorderTypeIntToString = "xlEdgeLeft"
    Case 8: BorderTypeIntToString = "xlEdgeTop"
    Case 9: BorderTypeIntToString = "xlEdgeBottom"
    Case 10: BorderTypeIntToString = "xlEdgeRight"
    Case 11: BorderTypeIntToString = "xlInsideVertical"
    Case 12: BorderTypeIntToString = "xlInsideHorizontal"
  End Select
End Function

