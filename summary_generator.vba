' ============================================================================================
' Title: Route Group Summary with Grade Distribution (Configurable VBA Macro)
' by Fayyaz Minhas (version 16-06-2025)
' WHAT THIS MACRO DOES:
' ------------------------
' 1. Summarizes coursework, exam, and total marks per route group (e.g., CS, CSE, etc.).
' 2. Only includes valid numeric entries in average calculations:
'    - If coursework is missing ? excluded from coursework average.
'    - If exam or total is missing ? excluded from that respective average.
'    - A row is never entirely skipped unless route code is empty.
' 3. Outputs:
'    - Group-wise student count, average marks, and how many were included per metric.
'    - One overall grade distribution (1st, 2.1, 2.2, 3rd, Fail) across all students.
'
' CONFIGURE BEFORE RUNNING:
' ----------------------------
' - Ensure your worksheet is named "Marks"
' - Update these variables if your layout differs:
'     - routeCol = 4     ' Route column (e.g., "G503 Computer Science MEng" in col D)
'     - examCol = 7      ' Exam marks column (e.g., col G)
'     - totalCol = 9     ' Total marks column (e.g., col I)
'     - courseworkCols = Array(5, 6)   ' Coursework columns (e.g., col E and F)
'     - weightRow = 2    ' Row from which coursework weights are read (e.g., "25%", "25%")
'     - groupDefinitions = Array(...) ' Define your route groups here
'
' HOW TO RUN:
' --------------
' 1. Press Alt + F11 to open the VBA editor
' 2. Insert a new module, paste this full code
' 3. Press F5 (or Alt + F8 > "SummarizeRouteGroups" > Run)
' 4. A new "Summary" sheet will be created with all outputs
' ============================================================================================


Sub SummarizeRouteGroups()

    ' =============================================================================
    ' Configurable Summary Macro with Grade Distribution and Valid Entry Counts
    ' =============================================================================
    Const weightRow As Long = 2
    Const routeCol As Long = 4
    Const examCol As Long = 7
    Const totalCol As Long = 9

    Dim courseworkCols As Variant
    courseworkCols = Array(5, 6)

    Dim groupDefinitions As Variant
    groupDefinitions = Array( _
        "CS:G500,G502,G503,G504", _
        "CSE:G406,GG407,G408,G409", _
        "DM:G4G1,G4G2,G4G3,G4G4", _
        "CSBS:I1N1,I1NA", _
        "DS:G302,G303,G304" _
    )

    Dim wsData As Worksheet, wsOut As Worksheet
    Set wsData = ThisWorkbook.Sheets("Marks")

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = ThisWorkbook.Sheets.Add
    wsOut.Name = "Summary"

    Dim weights() As Double, weightSum As Double
    ReDim weights(LBound(courseworkCols) To UBound(courseworkCols))
    Dim i As Long, j As Long
    weightSum = 0
    For i = LBound(courseworkCols) To UBound(courseworkCols)
        weights(i) = CDbl(Replace(wsData.Cells(weightRow, courseworkCols(i)).Text, "%", "")) / 100
        weightSum = weightSum + weights(i)
    Next i

    Dim dictGroups As Object, dictStats As Object
    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set dictStats = CreateObject("Scripting.Dictionary")

    Dim groupKey, groupName As String, groupCodes As Variant, code
    For Each groupKey In groupDefinitions
        groupName = Split(groupKey, ":")(0)
        groupCodes = Split(Split(groupKey, ":")(1), ",")
        For Each code In groupCodes
            dictGroups(UCase(Trim(code))) = groupName
        Next code
    Next

    Dim gradeCounts(0 To 4) As Long ' 1st, 2.1, 2.2, 3rd, Fail

    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.Count, routeCol).End(xlUp).Row
    Dim routeCode As String, totalMark As Double, courseworkMark As Double, examMark As Double
    Dim stats As Variant

    Dim allStats(0 To 6) As Variant ' count, cw_sum, cw_n, exam_sum, exam_n, total_sum, total_n

    For i = 3 To lastRow
        routeCode = UCase(Trim(Split(WorksheetFunction.Clean(wsData.Cells(i, routeCol).Text), " ")(0)))
        If dictGroups.exists(routeCode) Then
            groupName = dictGroups(routeCode)
        Else
            groupName = "Other"
        End If

        If Not dictStats.exists(groupName) Then
            dictStats(groupName) = Array(0, 0#, 0, 0#, 0, 0#, 0)
        End If

        stats = dictStats(groupName)
        stats(0) = stats(0) + 1  ' total students
        allStats(0) = allStats(0) + 1

        ' Coursework
        courseworkMark = 0: Dim cwN As Long: cwN = 0
        For j = LBound(courseworkCols) To UBound(courseworkCols)
            If IsNumeric(wsData.Cells(i, courseworkCols(j)).Value) Then
                courseworkMark = courseworkMark + wsData.Cells(i, courseworkCols(j)).Value * weights(j)
                cwN = cwN + 1
            End If
        Next j
        If cwN > 0 Then
            stats(1) = stats(1) + courseworkMark / (weightSum * cwN / (UBound(courseworkCols) - LBound(courseworkCols) + 1))
            stats(2) = stats(2) + 1
            allStats(1) = allStats(1) + courseworkMark / (weightSum * cwN / (UBound(courseworkCols) - LBound(courseworkCols) + 1))
            allStats(2) = allStats(2) + 1
        End If

        ' Exam
        If IsNumeric(wsData.Cells(i, examCol).Value) Then
            examMark = wsData.Cells(i, examCol).Value
            stats(3) = stats(3) + examMark
            stats(4) = stats(4) + 1
            allStats(3) = allStats(3) + examMark
            allStats(4) = allStats(4) + 1
        End If

        ' Total
        If IsNumeric(wsData.Cells(i, totalCol).Value) Then
            totalMark = wsData.Cells(i, totalCol).Value
            stats(5) = stats(5) + totalMark
            stats(6) = stats(6) + 1
            allStats(5) = allStats(5) + totalMark
            allStats(6) = allStats(6) + 1

            ' Grade distribution
            Select Case totalMark
                Case Is >= 70: gradeCounts(0) = gradeCounts(0) + 1
                Case Is >= 60: gradeCounts(1) = gradeCounts(1) + 1
                Case Is >= 50: gradeCounts(2) = gradeCounts(2) + 1
                Case Is >= 40: gradeCounts(3) = gradeCounts(3) + 1
                Case Else: gradeCounts(4) = gradeCounts(4) + 1
            End Select
        End If

        dictStats(groupName) = stats
    Next i

    ' === Output ===
    Dim rowOut As Long: rowOut = 1
    wsOut.Range("A1:H1").Value = Array("Group", "Total", "CW Avg (%)", "CW n", "Exam Avg (%)", "Exam n", "Total Avg (%)", "Total n")

    rowOut = 2
    Dim key
    For Each key In dictStats.Keys
        stats = dictStats(key)
        wsOut.Cells(rowOut, 1).Value = key
        wsOut.Cells(rowOut, 2).Value = stats(0)
        wsOut.Cells(rowOut, 3).Value = IIf(stats(2) > 0, stats(1) / stats(2) / 100, "")
        wsOut.Cells(rowOut, 4).Value = stats(2)
        wsOut.Cells(rowOut, 5).Value = IIf(stats(4) > 0, stats(3) / stats(4) / 100, "")
        wsOut.Cells(rowOut, 6).Value = stats(4)
        wsOut.Cells(rowOut, 7).Value = IIf(stats(6) > 0, stats(5) / stats(6) / 100, "")
        wsOut.Cells(rowOut, 8).Value = stats(6)
        rowOut = rowOut + 1
    Next key

    ' ALL row
    wsOut.Cells(rowOut, 1).Value = "ALL"
    wsOut.Cells(rowOut, 2).Value = allStats(0)
    wsOut.Cells(rowOut, 3).Value = IIf(allStats(2) > 0, allStats(1) / allStats(2) / 100, "")
    wsOut.Cells(rowOut, 4).Value = allStats(2)
    wsOut.Cells(rowOut, 5).Value = IIf(allStats(4) > 0, allStats(3) / allStats(4) / 100, "")
    wsOut.Cells(rowOut, 6).Value = allStats(4)
    wsOut.Cells(rowOut, 7).Value = IIf(allStats(6) > 0, allStats(5) / allStats(6) / 100, "")
    wsOut.Cells(rowOut, 8).Value = allStats(6)

    wsOut.Range("C2:C" & rowOut).NumberFormat = "0.00%"
    wsOut.Range("E2:E" & rowOut).NumberFormat = "0.00%"
    wsOut.Range("G2:G" & rowOut).NumberFormat = "0.00%"

    rowOut = rowOut + 2
    wsOut.Cells(rowOut, 1).Value = "Grade distribution:"
    wsOut.Range("A" & rowOut + 1 & ":F" & rowOut + 1).Value = Array("1st", "2.1", "2.2", "3rd", "Fail", "Total")

    Dim totalGrades As Long: totalGrades = Application.Sum(gradeCounts)
    If totalGrades > 0 Then
        Dim gradePercents(0 To 4) As Double, rounded(0 To 4) As Double, sumPercents As Double
        For i = 0 To 4
            gradePercents(i) = gradeCounts(i) / totalGrades
            rounded(i) = Round(gradePercents(i), 4)
            sumPercents = sumPercents + rounded(i)
        Next i

        ' Adjust rounding
        Dim delta As Double: delta = 1# - sumPercents
        If Abs(delta) > 0.00001 Then
            Dim maxIdx As Long: maxIdx = 0
            For i = 1 To 4
                If rounded(i) > rounded(maxIdx) Then maxIdx = i
            Next i
            rounded(maxIdx) = rounded(maxIdx) + delta
        End If

        For i = 0 To 4
            wsOut.Cells(rowOut + 2, i + 1).Value = Format(rounded(i), "0.00%")
        Next i
        wsOut.Cells(rowOut + 2, 6).Value = totalGrades
    End If

    wsOut.Columns("A:H").AutoFit
    MsgBox "Summary complete with per-group stats and overall grade distribution.", vbInformation

End Sub


