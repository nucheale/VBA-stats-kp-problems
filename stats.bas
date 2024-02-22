Function removeDublicatesFromTwoDimArray(arr)
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Not dict.Exists(arr(i, 1)) Then dict.Add arr(i, 1), arr(i, 1)
    Next i
    Dim uniqueArr As Variant
    ReDim uniqueArr(1 To dict.Count)
    i = 1
    For Each Key In dict.Keys
        uniqueArr(i) = Key
        i = i + 1
    Next Key
    removeDublicatesFromTwoDimArray = uniqueArr
End Function

Function removeDublicatesFromOneDimArray(arr)
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr) To UBound(arr)
        If Not dict.Exists(arr(i)) Then dict.Add arr(i), arr(i)
    Next i
    Dim uniqueArr As Variant
    ReDim uniqueArr(1 To dict.Count)
    i = 1
    For Each Key In dict.Keys
        uniqueArr(i) = Key
        i = i + 1
    Next Key
    removeDublicatesFromOneDimArray = uniqueArr
End Function

Sub Stats()

    Dim e, element, i, j, fileIndex, listKpRow As Long
    
    Set macrowb = ActiveWorkbook
    
    filesToOpen = Application.GetOpenFilename(FileFilter:="All files (*.*), *.*", MultiSelect:=True, Title:="Выберите файлы")
    If TypeName(filesToOpen) = "Boolean" Then Exit Sub
    
    With Application
        .AskToUpdateLinks = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
    
    With macrowb.Worksheets("Справочник")
        Dim districts, carriers, files As Variant
        districts = .ListObjects("Районы").DataBodyRange.Value
        carriers = .ListObjects("Перевозчики").DataBodyRange.Value
        files = .ListObjects("Файлы").DataBodyRange.Value
    End With
    
    Set listKpWb = Application.Workbooks.Add

    listKpRow = 0
    fileIndex = 1
    For Each file In filesToOpen
        Set currentWb = Application.Workbooks.Open(Filename:=filesToOpen(fileIndex))
        Select Case True
            Case currentWb.Name Like "*Статистика за*"
                Set statsKpWb = currentWb
            Case currentWb.Name Like "*Отчет по срывам*"
                Set failuresKpWb = currentWb
            Case currentWb.Name Like "*Список КП по участкам*"
                lastRow = currentWb.Sheets(1).Cells.SpecialCells(xlLastCell).Row
                lastColumn = currentWb.Sheets(1).Cells.SpecialCells(xlLastCell).Column
                Dim listData As Variant
                listData = currentWb.Worksheets(1).Range(currentWb.Worksheets(1).Cells(4, 1), currentWb.Worksheets(1).Cells(lastRow, lastColumn))
                listKpWb.Sheets(1).Cells(listKpRow + 1, 1).Resize(UBound(listData), UBound(listData, 2)).Value = listData
                listKpRow = listKpWb.Sheets(1).Cells.SpecialCells(xlLastCell).Row
                currentWb.Close SaveChanges:=False
            Case Else
                MsgBox "Неопознанный файл: " & currentWb.Name
                GoTo errorExit
        End Select
        fileIndex = fileIndex + 1
    Next file

    reportDate = CDate(Left(Right(statsKpWb.Name, 15), 10))
    
    With listKpWb.Sheets(1)
        lastRowListKp = .Cells(Rows.Count, 1).End(xlUp).Row
        lastColumnListKp = .Cells.SpecialCells(xlLastCell).Column
        Set findIDCell = .Range(.Cells(1, 1), .Cells(1, lastColumnListKp)).Find(What:="Код КП", LookAt:=xlWhole)
        Set findDistrictCell = .Range(.Cells(1, 1), .Cells(1, lastColumnListKp)).Find(What:="Район", LookAt:=xlWhole)
        Dim listKpIDList, listKpDistrictsList As Variant
        listKpIDList = .Range(.Cells(findIDCell.Row + 1, findIDCell.Column), .Cells(lastRowListKp, findIDCell.Column))
        listKpDistrictsList = .Range(.Cells(findDistrictCell.Row + 1, findDistrictCell.Column), .Cells(lastRowListKp, findDistrictCell.Column))
        ' Debug.Print "listKpIDList: " & UBound(listKpIDList)
        ' Debug.Print "listKpDistrictsList: " & UBound(listKpDistrictsList)
        listKpWb.Close SaveChanges:=False
    End With
    
    With failuresKpWb.Sheets("report")
        lastRowFailures = .Cells(Rows.Count, 3).End(xlUp).Row
        lastColumnFailures = .Cells.SpecialCells(xlLastCell).Column
        Set findIDCell = .Range(.Cells(1, 1), .Cells(4, lastColumnFailures)).Find(What:="Код КП", LookAt:=xlWhole)
        Set findDistrictCell = .Range(.Cells(1, 1), .Cells(4, lastColumnFailures)).Find(What:="Участок", LookAt:=xlWhole)
        Set findCarrierCell = .Range(.Cells(1, 1), .Cells(4, lastColumnFailures)).Find(What:="Перевозчик", LookAt:=xlWhole)
        Set findProblemCell = .Range(.Cells(1, 1), .Cells(4, lastColumnFailures)).Find(What:="Проблема", LookAt:=xlWhole)
        Dim failuresIDList, failuresDistrictsList, failuresProblemsList, failuresCarriersList As Variant
        failuresIDList = .Range(.Cells(findIDCell.Row + 1, findIDCell.Column), .Cells(lastRowFailures, findIDCell.Column))
        failuresDistrictsList = .Range(.Cells(findDistrictCell.Row + 1, findDistrictCell.Column), .Cells(lastRowFailures, findDistrictCell.Column))
        failuresProblemsList = .Range(.Cells(findProblemCell.Row + 1, findProblemCell.Column), .Cells(lastRowFailures, findProblemCell.Column))
        failuresCarriersList = .Range(.Cells(findCarrierCell.Row + 1, findCarrierCell.Column), .Cells(lastRowFailures, findCarrierCell.Column))
        failuresProblemsListWithoutDublicates = removeDublicatesFromTwoDimArray(failuresProblemsList)
    End With
    
   ' On Error Resume Next
    For e = LBound(failuresIDList) + 1 To UBound(failuresIDList) 'заполнение района из реестра кп по коду кп
    failuresIDList(e, 1) = CLng(failuresIDList(e, 1))
        For n = LBound(listKpIDList) + 1 To UBound(listKpIDList)
            If listKpIDList(n, 1) = failuresIDList(e, 1) Then
                failuresDistrictsList(e, 1) = listKpDistrictsList(n, 1)
                Exit For
            Else
                failuresDistrictsList(e, 1) = "КП не найдена"
            End If
        Next n
    Next e
    On Error GoTo 0
        
    With failuresKpWb.Sheets("report")
        .Cells(findDistrictCell.Row + 1, findDistrictCell.Column).Resize(UBound(failuresDistrictsList), UBound(failuresDistrictsList, 2)).Value = failuresDistrictsList 'заполнение района
    End With

    With statsKpWb.Sheets("Вывоз КГ")
        lastRowStatsKp = .Cells(Rows.Count, 2).End(xlUp).Row
        lastColumnStatsKp = .Cells.SpecialCells(xlLastCell).Column
        Set findDistrictCell = .Range(.Cells(1, 1), .Cells(5, lastColumnStatsKp)).Find(What:="Район", LookAt:=xlWhole)
        Set findPlanCell = .Range(.Cells(1, 1), .Cells(5, lastColumnStatsKp)).Find(What:="В план задании", LookAt:=xlWhole)
        Dim statsWbKgDistricts, statsWbKgPlan As Variant
        statsWbKgDistricts = .Range(.Cells(findDistrictCell.Row + 1, findDistrictCell.Column), .Cells(lastRowStatsKp, findDistrictCell.Column))
        statsWbKgPlan = .Range(.Cells(findPlanCell.Row + 1, findPlanCell.Column), .Cells(lastRowStatsKp, findPlanCell.Column))
        Dim resultDistrictsPlan As Variant
        ReDim resultDistrictsPlan(1 To UBound(districts, 1))
        For i = LBound(districts, 1) To UBound(districts, 1)
            sumPlan = 0
            For n = LBound(statsWbKgDistricts, 1) To UBound(statsWbKgDistricts, 1)
                If districts(i, 1) = statsWbKgDistricts(n, 1) Then
                    If statsWbKgPlan(n, 1) = 1 Then sumPlan = sumPlan + 1
                End If
            Next n
            resultDistrictsPlan(i) = sumPlan
        Next i
    End With

    With failuresKpWb.Sheets("report")
        Dim resultDistrictsProblems As Variant
        ReDim resultDistrictsProblems(1 To UBound(districts, 1))
        For i = LBound(districts, 1) To UBound(districts, 1)
            sumProblems = 0
            For n = LBound(failuresDistrictsList, 1) To UBound(failuresDistrictsList, 1)
                If districts(i, 1) = failuresDistrictsList(n, 1) Then sumProblems = sumProblems + 1
            Next n
            resultDistrictsProblems(i) = sumProblems
        Next i

        Dim problems As Variant
        ReDim problems(1 To UBound(districts) * UBound(failuresProblemsListWithoutDublicates), 1 To 3)

        counter = 1
        For d = LBound(districts, 1) To UBound(districts, 1)
            For j = LBound(failuresProblemsListWithoutDublicates) To UBound(failuresProblemsListWithoutDublicates)
                    problems(counter, 1) = districts(d, 1)
                    problems(counter, 2) = failuresProblemsListWithoutDublicates(j)
                    counter = counter + 1
            Next j
        Next d
        
        For i = LBound(problems, 1) To UBound(problems, 1)
            For n = LBound(failuresDistrictsList, 1) To UBound(failuresDistrictsList, 1)
                If failuresDistrictsList(n, 1) = problems(i, 1) Then
                    If failuresProblemsList(n, 1) = problems(i, 2) Then problems(i, 3) = CInt(problems(i, 3)) + 1
                End If
            Next n
            If problems(i, 3) = Empty Then
                problems(i, 1) = Empty
                problems(i, 2) = Empty
            End If
        Next i
        
        For i = LBound(problems, 1) To UBound(problems, 1) - 1 'сортировка по убыванию
            For j = i + 1 To UBound(problems, 1)
                If problems(i, 3) < problems(j, 3) Then
                    For k = 1 To 3
                        tempValue = problems(i, k)
                        problems(i, k) = problems(j, k)
                        problems(j, k) = tempValue
                    Next k
                End If
            Next j
        Next i

        Dim problems2 As Variant
        ReDim problems2(LBound(districts, 1) To UBound(districts, 1))
        For n = LBound(districts, 1) To UBound(districts, 1)
            counter = 1
            otherProblems = 0
            For i = LBound(problems) To UBound(problems)
                If districts(n, 1) = problems(i, 1) Then
                    If counter <= 4 Then
                        problems2(n) = problems2(n) & problems(i, 2) & ": " & problems(i, 3) & vbLf
                        counter = counter + 1
                    Else
                        otherProblems = otherProblems + CInt(problems(i, 3))
                    End If
                End If
            Next i
            problems2(n) = problems2(n) & "Иные" & ": " & otherProblems
        Next n
        
        sumProblemsAll = 0
        For i = LBound(problems) To UBound(problems)
            sumProblemsAll = sumProblemsAll + CInt(problems(i, 3))
        Next i

        Dim resultDistrictsFact As Variant
        Dim effiency As Variant
        ReDim resultDistrictsFact(1 To UBound(districts, 1))
        ReDim effiency(1 To UBound(districts, 1)) As Double
        sumFact = 0
        sumPlan = 0
        For i = LBound(resultDistrictsFact) To UBound(resultDistrictsFact)
            resultDistrictsFact(i) = resultDistrictsPlan(i) - resultDistrictsProblems(i)
            effiency(i) = (resultDistrictsPlan(i) - resultDistrictsProblems(i)) / resultDistrictsPlan(i)
            sumFact = sumFact + resultDistrictsFact(i)
            sumPlan = sumPlan + resultDistrictsPlan(i)
            sumEffiency = sumEffiency + effiency(i)
        Next i
            averageEffiency = sumEffiency / UBound(effiency)
    End With

    
    macrowb.Sheets("Шаблон").Copy After:=macrowb.Sheets(macrowb.Sheets.Count - 1)
    Set newWs = ActiveSheet
    ' Set newWs = macroWb.Sheets.Add(After:=macroWb.Sheets(macroWb.Sheets.Count - 1))
    currTime = Array(Hour(Now), Minute(Now), Second(Now))
    ' For Each e In currTime
    '     If e < 10 Then e = "0" & e
    ' Next e
    newWs.Name = Date & "_" & currTime(0) & "_" & currTime(1) & "_" & currTime(2)
    With newWs
        .Cells(1, 1) = "Отчет за " & reportDate
        For i = LBound(districts) To UBound(districts)
            Cells(i + 2, 1) = districts(i, 1)
            Cells(i + 2, 2) = districts(i, 2)
            Cells(i + 2, 3) = resultDistrictsPlan(i)
            Cells(i + 2, 4) = resultDistrictsFact(i)
            Cells(i + 2, 5) = problems2(i)
            Cells(i + 2, 6) = effiency(i)
        Next i
        lastRowMacroWb = .Cells(Rows.Count, 1).End(xlUp).Row
        lastColumnMacroWb = .Cells.SpecialCells(xlLastCell).Column
        .Cells(lastRowMacroWb + 1, 1) = "Итого"
        .Cells(lastRowMacroWb + 1, 3) = sumPlan
        .Cells(lastRowMacroWb + 1, 4) = sumFact
        .Cells(lastRowMacroWb + 1, 5) = sumProblemsAll
        .Cells(lastRowMacroWb + 1, 6) = averageEffiency
        .Range(.Cells(lastRowMacroWb + 1, 1), .Cells(lastRowMacroWb + 1, 2)).Merge
        .Range(.Cells(lastRowMacroWb + 1, 1), .Cells(lastRowMacroWb + 1, lastColumnMacroWb)).Font.Bold = True
        .Range(.Cells(lastRowMacroWb + 1, 1), .Cells(lastRowMacroWb + 1, lastColumnMacroWb)).VerticalAlignment = xlCenter
        .Range(.Cells(lastRowMacroWb + 1, 1), .Cells(lastRowMacroWb + 1, lastColumnMacroWb)).HorizontalAlignment = xlCenter
        .Range(.Cells(2, 1), .Cells(lastRowMacroWb + 1, lastColumnMacroWb)).Borders.LineStyle = xlContinuous
        
        Set pivotRange = .Range(.Cells(2, 1), .Cells(lastRowMacroWb, lastColumnMacroWb))
        Set pivotDestination = .Cells(lastRowMacroWb + 5, 1)
        Set pivotTableResult = .PivotTableWizard(SourceType:=xlDatabase, SourceData:=pivotRange, TableDestination:=pivotDestination)
        With pivotTableResult
            .ColumnGrand = True
            .HasAutoFormat = True
            .RowGrand = True
            .SaveData = True
            .InGridDropZones = False
            .DisplayFieldCaptions = True
            .TableStyle2 = "PivotStyleDark2"
            .RowAxisLayout xlCompactRow
        End With
        With pivotTableResult.PivotFields("Район")
            .Orientation = xlRowField
            .Position = 1
        End With
        With pivotTableResult.PivotFields("Генеральный перевозчик")
            .Orientation = xlColumnField
            .Position = 1
        End With
        pivotTableResult.AddDataField pivotTableResult.PivotFields("% выполнения"), "Среднее по полю % выполнения", xlAverage
        .Range(.Cells(lastRowMacroWb + 5, 1), .Cells(lastRowMacroWb + 5 + UBound(districts, 1) + 2, lastColumnMacroWb)).NumberFormat = "0%"
        .Range(.Cells(lastRowMacroWb + 6, 1), .Cells(lastRowMacroWb + 5 + UBound(districts, 1) + 2, lastColumnMacroWb)).Borders.LineStyle = xlContinuous
        .Range(.Cells(lastRowMacroWb + 6, 1), .Cells(lastRowMacroWb + 5 + UBound(districts, 1) + 2, lastColumnMacroWb)).HorizontalAlignment = xlCenter
        .Range(.Cells(lastRowMacroWb + 6, 1), .Cells(lastRowMacroWb + 5 + UBound(districts, 1) + 2, lastColumnMacroWb)).VerticalAlignment = xlCenter

        .Range(.Cells(1, 1), .Cells(lastRowMacroWb + 5 + UBound(districts, 1) + 2, lastColumnMacroWb)).Columns.AutoFit
    End With

    statsKpWb.Close SaveChanges:=False
    failuresKpWb.Close SaveChanges:=False
    
errorExit:
    With Application
        .AskToUpdateLinks = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub