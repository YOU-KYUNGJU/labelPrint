Attribute VB_Name = "Module9"
Public gAutoPrintAll As Boolean   ' True = 묻지 않고 전체 출력

Sub 자동_양식_출력_20260117()

    Dim wb As Workbook, wbStart As Workbook
    Dim wsMain As Worksheet, wsForm As Worksheet, wsStart As Worksheet
    Dim basePath As String, startFile As String
    Dim row As Long, lastRow As Long, i As Long
    Dim 접수번호 As String, formattedID As String
    Dim txtFile As String, txtLine As String
    Dim B1List As Collection, B1 As Variant
    Dim C1 As String, D1 As String, A1 As String
    Dim cell As Range, original As String
    Dim fso As Object, fileStream As Object
    Dim parts As Variant
    Dim sampleCnt As Long, sampleNo As Long
    Dim ShrinkCodeList As Collection
    Dim shrinkCode As String
    Dim idx As Long
    Dim SampleCntList As Collection



    Set wb = ThisWorkbook
    Set wsMain = wb.Sheets("접수번호")
    Set wsForm = wb.Sheets("양식")
    basePath = ThisWorkbook.Path
    startFile = basePath & "\start.xlsx"
    Dim firstRun As Boolean
    firstRun = True


    row = 1
    Do While wsMain.Cells(row, "B").Value <> ""
        If firstRun Then
            If Not gAutoPrintAll Then
                Dim sel As VbMsgBoxResult
                sel = MsgBox( _
                    "출력 방식을 선택하세요." & vbCrLf & vbCrLf & _
                    "예(Y)  : 전체 자동 출력" & vbCrLf & _
                    "아니오(N) : 건별 확인 출력", _
                    vbYesNo + vbQuestion, "출력 방식 선택")
        
                gAutoPrintAll = (sel = vbYes)
            End If
            firstRun = False
        End If

        wsForm.Range("f1").Value = wsMain.Cells(row, "B").Value
        접수번호 = Replace(wsMain.Cells(row, "B").Value, "@", "")
        formattedID = 접수번호 ' 접수번호 포맷팅: "H2312522028" → "H231-25-22028"
        formattedID = Left(접수번호, 4) _
              & "-" & Mid(접수번호, 5, 2) _
              & "-" & Mid(접수번호, 7)

        ' 텍스트 파일 경로
        txtFile = basePath & "\testNumber\2025\" & 접수번호 & ".txt"
        If Dir(txtFile) = "" Then
            row = row + 1
            GoTo ContinueLoop
        End If

        ' 수축율 리스트 및 A1 값 가져오기
        Set B1List = New Collection
        Set ShrinkCodeList = New Collection
        Set SampleCntList = New Collection

        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fileStream = fso.OpenTextFile(txtFile, 1)
        
        A1 = "" ' A1 초기화

        Do Until fileStream.AtEndOfStream
            txtLine = fileStream.ReadLine
            
            ' 첫 번째 항목을 A1 변수로 추출
            parts = Split(txtLine, vbTab)
            If UBound(parts) >= 0 Then
                If Trim(parts(0)) <> "" And A1 = "" Then
                    A1 = Trim(parts(0))  ' 한 번만 저장 (선택 사항)
                End If
            End If
        
            ' 예: 수축율 관련 B1 값 추출도 병행 가능
            If InStr(txtLine, "수축율") > 0 Then
                parts = Split(txtLine, vbTab)
            
                ' parts(1) = 37/38/39 같은 코드, parts(2) = 시험조건(설명)
                If UBound(parts) >= 2 Then
                    If Trim(parts(2)) <> "" Then
                        B1List.Add Trim(parts(2))          ' 기존처럼 시험조건(설명)
                    End If
                    If Trim(parts(1)) <> "" Then
                        ShrinkCodeList.Add Trim(parts(1))  ' 37/38/39 코드 추가 저장
                    End If
                End If
            
                ' 시료수(요청: 2개라고 가정) → 파일에 값이 있으면 그 값을 사용
                ' 업로드 txt 예시 기준으로 parts(4)에 "2"가 들어있는 형태가 일반적이라 여기서 읽습니다.
                ' 수축율 행마다 sampleCnt를 SampleCntList에 같이 넣기
                Dim oneCnt As Long
                oneCnt = 1
                If UBound(parts) >= 4 Then
                    oneCnt = Val(parts(4))
                End If
                If oneCnt <= 0 Then oneCnt = 1
                
                SampleCntList.Add oneCnt

            End If

        Loop
        fileStream.Close

        ' start.xlsx에서 C1, D1 찾기
        Set wbStart = Workbooks.Open(startFile, ReadOnly:=True)
        Set wsStart = wbStart.Sheets(1)
        lastRow = wsStart.Cells(wsStart.Rows.Count, 1).End(xlUp).row

        C1 = "": D1 = ""
        For i = 1 To lastRow
            If wsStart.Cells(i, 1).Value = formattedID Then
                C1 = wsStart.Cells(i, 12).Value  '신청업체 11, 납품업체 12
                D1 = wsStart.Cells(i, 3).Value
                A1 = wsStart.Cells(i, 5).Value
                Exit For
            End If
        Next i
        wbStart.Close False
        
        ' === (추가) 수동 시료번호 목록 읽기: 접수번호 시트 C열 (예: C1=2, C2=3, ...)
        Dim manualNos() As Long
        manualNos = ReadManualSampleNos(wsMain, row)
        
        Dim hasManual As Boolean
        hasManual = False
        
        On Error Resume Next
        If UBound(manualNos) >= 1 Then
            If manualNos(1) <> -1 Then hasManual = True
        End If
        On Error GoTo 0
        
        Dim manualCnt As Long
        If hasManual Then
            manualCnt = GetArrayCountLong(manualNos)
        Else
            manualCnt = 0
        End If


        ' 각 (수축율 코드/시험조건) × (시료번호) 만큼 반복 출력
        For idx = 1 To B1List.Count
        
            B1 = B1List(idx)
            shrinkCode = ShrinkCodeList(idx)
        
            ' B3 셀용 처리: 괄호 안 문자열에서 첫 번째 '(' ~ ',' 사이 제거 (기존 로직 유지)
            Dim lParen As Long, commaPos As Long
            lParen = InStr(B1, "(")
            commaPos = InStr(B1, ",")
            If lParen > 0 And commaPos > lParen Then
                B1 = Left(B1, lParen) & Mid(B1, commaPos + 1)
            End If
            B1 = "시험조건 " & B1
        
            ' 시료수만큼 반복 (예: 2개면 1,2)
            sampleCnt = SampleCntList(idx)
            If sampleCnt <= 0 Then sampleCnt = 1
            
            Dim k As Long
            Dim realNo As Long
            
            For k = 1 To sampleCnt
            
                ' (추가) 수동 시료번호가 있으면 그 값을 사용, 없으면 기존처럼 1..N
                If manualCnt >= k Then
                    realNo = manualNos(k)
                Else
                    realNo = k
                End If
            
                wsForm.Visible = xlSheetVisible
                wsForm.Range("B2").Value = formattedID & "  #" & realNo
                wsForm.Range("B3").Value = B1
                wsForm.Range("B5").Value = C1
                wsForm.Range("C6").Value = D1
                wsForm.Range("D6").Value = A1
            
                wsForm.Range("F2").Value = "@" & 접수번호 & "@" & realNo & "," & shrinkCode
            
                wsForm.PageSetup.PrintArea = "$A$1:$D$6"
                wsForm.Visible = xlSheetVisible
                wsForm.Activate
                
                wsForm.PageSetup.Orientation = xlLandscape
                
                On Error GoTo PrintFail
                
                If gAutoPrintAll Then
                    wsForm.PrintOut Copies:=1
                Else
                    userChoice = MsgBox("현재 라벨을 출력하시겠습니까?", vbYesNo + vbQuestion, "출력 확인")
                    If userChoice = vbYes Then
                        wsForm.PrintOut Copies:=1
                    End If
                End If
                
                On Error GoTo 0
                DoEvents
                GoTo PrintDone
                
PrintFail:
                    MsgBox "인쇄 실패: " & Err.Number & " / " & Err.Description & vbCrLf & _
                           "ActivePrinter=" & Application.ActivePrinter & vbCrLf & _
                           "PrintArea=" & wsForm.PageSetup.PrintArea, vbExclamation
                    Err.Clear
                    On Error GoTo 0
                
PrintDone:

                DoEvents

            
                wsForm.PageSetup.PrintArea = ""
                wsForm.Visible = xlSheetHidden
            
            Next k

        Next idx

ContinueLoop:
        row = row + 1
    Loop

    'MsgBox "인쇄 완료!"
    wb.Sheets("접수번호").Activate
    
    Sheets("접수번호").Select
    Application.CutCopyMode = False
    Sheets("접수번호").Copy

    Dim basePath1 As String
    Dim savePath As String
    Dim fileName As String
    Dim todayStr As String
    Dim yearStr As String, monthStr As String

    ' 날짜 정보
    todayStr = Format(Date, "yyyymmdd")         ' 예: 20250626
    yearStr = Format(Date, "yyyy")              ' 예: 2025
    monthStr = Format(Date, "mm")               ' 예: 06

    ' 기본 경로
    basePath1 = "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\" & "2025" & "\가공성능평가팀_수축율\" & yearStr & "\" & monthStr & "\"
    NewYearPath = "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\" & "2025" & "\가공성능평가팀_수축율\" & yearStr & "\"
    
    ' 최종 저장 경로: 날짜 폴더 포함
    savePath = basePath1 & todayStr & "\"

    ' 폴더가 없으면 생성
    If Dir(basePath1, vbDirectory) = "" Then
        MkDir basePath1
    End If
    If Dir(NewYearPath, vbDirectory) = "" Then
        MkDir NewYearPath
    End If
    If Dir(savePath, vbDirectory) = "" Then
        MkDir basePath1 & todayStr
    End If

    ' 파일 이름: 현재 날짜_시간초.csv
    fileName = Format(Now, "yyyymmdd_HHMMSS") & ".csv"

    ' 저장
    ActiveWorkbook.SaveAs fileName:=savePath & fileName, FileFormat:=xlCSV

    ' CSV 파일 닫기
    ActiveWorkbook.Close SaveChanges:=False
    ' 저장 경로 및 파일명 기록
    wb.Sheets("접수번호").Range("D11").Value = savePath
    wb.Sheets("접수번호").Range("D12").Value = fileName

    MsgBox "CSV 파일이 '" & savePath & "' 경로에 저장되고 자동으로 닫혔습니다."
    wb.Sheets("접수번호").Activate
    Range("B1:C400").ClearContents
    Range("B1").Select
End Sub
' 접수번호 시트 C열에서 수동 시료번호 읽기
' - 현재 접수번호 행(rowIdx)의 C열부터 아래로 연속된 숫자만 읽음
' - 아무것도 없으면 (-1) 1개짜리 배열 반환  -> "수동값 없음" 표시용
Private Function ReadManualSampleNos(ByVal ws As Worksheet, ByVal rowIdx As Long) As Long()
    Dim arr() As Long
    Dim cnt As Long: cnt = 0

    Dim r As Long
    Dim v As Variant

    r = rowIdx
    Do While True
        v = ws.Cells(r, "C").Value

        If Len(Trim$(CStr(v))) = 0 Then Exit Do      ' 빈칸이면 종료
        If Not IsNumeric(v) Then Exit Do             ' 숫자 아니면 종료

        cnt = cnt + 1
        ReDim Preserve arr(1 To cnt)
        arr(cnt) = CLng(v)

        r = r + 1
    Loop

    If cnt = 0 Then
        ReDim arr(1 To 1)
        arr(1) = -1
    End If

    ReadManualSampleNos = arr
End Function


Private Function GetArrayCountLong(ByRef a() As Long) As Long
    On Error GoTo EmptyArr
    GetArrayCountLong = UBound(a) - LBound(a) + 1
    Exit Function
EmptyArr:
    GetArrayCountLong = 0
End Function

Sub 출력_개별확인()
    gAutoPrintAll = False
    자동_양식_출력_20260117
End Sub


Sub 출력_전체자동()
    gAutoPrintAll = True
    자동_양식_출력_20260117
End Sub
