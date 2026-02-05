Attribute VB_Name = "Module1"
Sub Final_Perfect_Checklist_System_Fixed_v2()
    Dim masterWB As Workbook, sourceWB As Workbook
    Dim masterSheet As Worksheet, resultSheet As Worksheet, sourceSheet As Worksheet
    Dim ruleSheet As Worksheet, teachingSheet As Worksheet, etcSheet As Worksheet
    Dim lastRowS As Long, writeRow As Long, r As Long, sRow As Long, c As Long
    Dim currentID As String, currentName As String, deptName As String
    Dim allDepts As Object, studentData As Object, studentKey As Variant
    Dim studentYear As Integer, startCol As Integer, tStartCol As Integer
    
    Set masterWB = ThisWorkbook
    Set sourceWB = ActiveWorkbook
    
    If sourceWB.Name = masterWB.Name Then
        MsgBox "분석할 '성적표 파일'을 활성화한 상태에서 실행해 주세요!", vbExclamation: Exit Sub
    End If

    On Error Resume Next
    Set ruleSheet = masterWB.Sheets("학과별기준")
    Set teachingSheet = masterWB.Sheets("교직과목기준")
    Set resultSheet = masterWB.Sheets("점검표")
    If resultSheet Is Nothing Then
        Set resultSheet = masterWB.Sheets.Add(After:=masterWB.Sheets(masterWB.Sheets.Count))
        resultSheet.Name = "점검표"
    End If
    On Error GoTo 0
    
    resultSheet.Cells.Clear
    resultSheet.Range("A1:L1").Value = Array("학과", "학번", "성명", "전공충족", "교직충족", "최종판정", "응급처치", "인성검사", "성인지", "전공인정", "교직인정", "특이사항")
    resultSheet.Range("A1:L1").Font.Bold = True
    writeRow = 2

    Set sourceSheet = sourceWB.Sheets(1)
    lastRowS = sourceSheet.Cells(sourceSheet.Rows.Count, 17).End(xlUp).Row
    Set allDepts = CreateObject("Scripting.Dictionary")
    
    For r = 1 To lastRowS
        If InStr(CStr(sourceSheet.Cells(r, 17).Value), "학번") > 0 Then
            deptName = Trim(CStr(sourceSheet.Cells(r, 9).Value))
            If deptName <> "" Then allDepts(deptName) = True
        End If
    Next r

    For Each targetDept In allDepts.Keys
        Dim minMajor As Integer, minTeach As Integer, minEmerg As Integer, minTest As Integer, minGender As Integer
        Dim foundRule As Boolean: foundRule = False
        For r = 2 To ruleSheet.UsedRange.Rows.Count
            If Trim(ruleSheet.Cells(r, 1).Value) = targetDept Then
                minMajor = Val(ruleSheet.Cells(r, 2).Value)
                minTeach = Val(ruleSheet.Cells(r, 5).Value)
                minEmerg = Val(ruleSheet.Cells(r, 6).Value)
                minTest = Val(ruleSheet.Cells(r, 7).Value)
                minGender = Val(ruleSheet.Cells(r, 8).Value)
                foundRule = True: Exit For
            End If
        Next r

        Set masterSheet = Nothing
        On Error Resume Next
        Set masterSheet = masterWB.Sheets(CStr(targetDept))
        On Error GoTo 0
        
        If masterSheet Is Nothing Or Not foundRule Then
            resultSheet.Cells(writeRow, "A").Value = targetDept
            resultSheet.Cells(writeRow, "F").Value = "기준 미설정"
            writeRow = writeRow + 1
        Else
            Set studentData = CreateObject("Scripting.Dictionary")
            For r = 1 To lastRowS
                If InStr(CStr(sourceSheet.Cells(r, 17).Value), "학번") > 0 And InStr(CStr(sourceSheet.Cells(r, 9).Value), targetDept) > 0 Then
                    currentID = Trim(CStr(sourceSheet.Cells(r, 18).Value))
                    If InStr(currentID, ".") > 0 Then currentID = Split(currentID, ".")(0)
                    currentName = Trim(CStr(sourceSheet.Cells(r, 23).Value))
                    
                    If Not studentData.Exists(currentID) Then
                        Set studentData(currentID) = CreateObject("Scripting.Dictionary")
                        studentData(currentID)("Name") = currentName
                        Set studentData(currentID)("RawList") = New Collection
                    End If
                    
                    For sRow = r + 3 To r + 50
                        If InStr(CStr(sourceSheet.Cells(sRow, 17).Value), "학번") > 0 Then Exit For
                        ' 공백만 제거하고 숫자는 나중에 처리하거나 원본으로 비교
                        Dim sL As String: sL = Replace(CStr(sourceSheet.Cells(sRow, 4).Value), " ", "")
                        Dim sR As String: sR = Replace(CStr(sourceSheet.Cells(sRow, 15).Value), " ", "")
                        If sL <> "" Then studentData(currentID)("RawList").Add sL
                        If sR <> "" Then studentData(currentID)("RawList").Add sR
                    Next sRow
                End If
            Next r

            For Each studentKey In studentData.Keys
                studentYear = Val(Left(CStr(studentKey), 4))
                startCol = 3: tStartCol = 3
                For c = 3 To masterSheet.Cells(2, masterSheet.Columns.Count).End(xlToLeft).Column
                    If Val(masterSheet.Cells(2, c).Value) = studentYear Then: startCol = c: Exit For
                Next c
                If Not teachingSheet Is Nothing Then
                    For c = 3 To teachingSheet.Cells(2, teachingSheet.Columns.Count).End(xlToLeft).Column
                        If Val(teachingSheet.Cells(2, c).Value) = studentYear Then: tStartCol = c: Exit For
                    Next c
                End If

                Dim majorMatch As Integer: majorMatch = 0: Dim teachMatch As Integer: teachMatch = 0
                Dim cEmerg As Integer: cEmerg = 0: Dim cTest As Integer: cTest = 0: Dim cGender As Integer: cGender = 0
                Dim majorList As String: majorList = "": Dim teachList As String: teachList = ""
                
                Dim subName As Variant
                For Each subName In studentData(studentKey)("RawList")
                    ' 1. 전공/교직 대조용 (숫자 제거 버전)
                    Dim cleanSub As String: cleanSub = CleanSubNameOnly(CStr(subName))
                    
                    ' 전공 대조
                    For mRow = 3 To masterSheet.UsedRange.Rows.Count
                        If CleanSubNameOnly(CStr(masterSheet.Cells(mRow, startCol).Value)) = cleanSub Then
                            majorMatch = majorMatch + 1
                            majorList = majorList & cleanSub & ", ": Exit For
                        End If
                    Next mRow
                    ' 교직 대조
                    For mRow = 3 To teachingSheet.UsedRange.Rows.Count
                        If CleanSubNameOnly(CStr(teachingSheet.Cells(mRow, tStartCol).Value)) = cleanSub Then
                            teachMatch = teachMatch + 1
                            teachList = teachList & cleanSub & ", ": Exit For
                        End If
                    Next mRow
                    
                    ' 2. 기타 과목 카운트 (포함 문구 체크로 강화)
                    If InStr(subName, "응급처치") > 0 Then cEmerg = cEmerg + 1
                    If InStr(subName, "인성검사") > 0 Then cTest = cTest + 1
                    If InStr(subName, "성인지") > 0 Then cGender = cGender + 1
                Next subName

                resultSheet.Cells(writeRow, "A").Value = targetDept
                resultSheet.Cells(writeRow, "B").Value = "'" & studentKey
                resultSheet.Cells(writeRow, "C").Value = studentData(studentKey)("Name")
                resultSheet.Cells(writeRow, "D").Value = majorMatch
                resultSheet.Cells(writeRow, "E").Value = teachMatch
                ' 결과 표시 방식 변경 (숫자만 표시)
                resultSheet.Cells(writeRow, "G").Value = cEmerg
                resultSheet.Cells(writeRow, "H").Value = cTest
                resultSheet.Cells(writeRow, "I").Value = cGender
                
                Dim isOK As Boolean: isOK = (majorMatch >= minMajor And teachMatch >= minTeach And _
                                            cEmerg >= minEmerg And cTest >= minTest And cGender >= minGender)
                resultSheet.Cells(writeRow, "F").Value = IIf(isOK, "충족", "불충족")
                resultSheet.Cells(writeRow, "J").Value = Left(majorList, IIf(Len(majorList) > 2, Len(majorList) - 2, 0))
                resultSheet.Cells(writeRow, "K").Value = Left(teachList, IIf(Len(teachList) > 2, Len(teachList) - 2, 0))
                
                If Not isOK Then resultSheet.Rows(writeRow).Interior.Color = RGB(255, 245, 245)
                writeRow = writeRow + 1
            Next studentKey
        End If
    Next targetDept
    resultSheet.Columns("A:L").AutoFit
    MsgBox "분석 완료! [점검표] 시트를 확인하세요."
End Sub

' 숫자 제거용 함수 분리
Function CleanSubNameOnly(ByVal txt As String) As String
    Dim res As String: res = Replace(Trim(txt), " ", "")
    If res = "" Then CleanSubNameOnly = "": Exit Function
    Do While Len(res) > 0 And IsNumeric(Right(res, 1))
        res = Left(res, Len(res) - 1)
    Loop
    CleanSubNameOnly = res
End Function

