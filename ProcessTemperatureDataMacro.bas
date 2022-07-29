Attribute VB_Name = "Modul1"
Public temperatureLowestLevelSize As Long
Public ZfTemperatures As Variant
Public ZfPositions As Variant


Public Sub ReadData()
    Dim inputDataTransient As String, nStep As Integer, nStepLoopCount As Long, inputDataTransientTextline As String, parameterFilesLoopCount As Long
    Dim iiiiDataFile As String, iiii As Integer, iiiiDataFileTextline As String, temperatureFileName As String, temperatureFileNumber As Integer
    Dim electricFileName As String
    
    Dim coreTemperatureData() As Variant, coreTimeData() As String, temporaryCollectionOfCoreFilesData As Collection, yPosCore As Variant
    Dim masterCoreTemperatureData() As Variant, masterCoreTimeData() As Variant
    
    Dim surfaceTemperatureData() As Variant, surfaceTimeData() As Variant, temporaryCollectionOfSurfaceFilesData As Collection, yPosSurface As Variant
    Dim masterSurfaceTemperatureData() As Variant, masterSurfaceTimeData() As Variant
    
    Dim electricDataCollection As Collection
    Dim masterFrequencyData() As Variant, masterVoltageData() As Variant, masterIAData() As Variant, masterCosphiData() As Variant
    Dim masterPgesData() As Variant, masterPwrIndData() As Variant, masterPwrBandgesData() As Variant, masterWirkungsgradData() As Variant
    
    Dim zfResults() As Variant
    
    inputDataTransient = "C:\Users\Gäste_PC_2\Desktop\Tailored_Heating_ML\Subprog_input\input_data_transient.txt"
    iiiiDataFile = "C:\Users\Gäste_PC_2\Desktop\Tailored_Heating_ML\iiii.txt"
    
    nStepLoopCount = 1
    parameterFilesLoopCount = 1
    
    'The following code retrieves the nStep number which is used for the number of parameter files
    Open inputDataTransient For Input As #1
    Do Until nStepLoopCount = 3
        Line Input #1, inputDataTransientTextline
        If nStepLoopCount = 2 Then
            nStep = inputDataTransientTextline
        End If
        nStepLoopCount = nStepLoopCount + 1
    Loop
    Close #1
    
    'The following code retrieves the iiii number which is used for the number of Core/Surface temperature files
    Open iiiiDataFile For Input As #2
    Line Input #2, iiiiDataFileTextline
    iiii = Mid(iiiiDataFileTextline, 15, 2)
    Close #2
    
    ReDim Preserve coreTemperatureData(0)
    ReDim Preserve coreTimeData(0)
    ReDim Preserve masterCoreTemperatureData(0)
    ReDim Preserve masterCoreTimeData(0)
    ReDim Preserve surfaceTemperatureData(0)
    ReDim Preserve surfaceTimeData(0)
    ReDim Preserve masterSurfaceTemperatureData(0)
    ReDim Preserve masterSurfaceTimeData(0)
    
    ReDim Preserve masterFrequencyData(0)
    ReDim Preserve masterVoltageData(0)
    ReDim Preserve masterIAData(0)
    ReDim Preserve masterCosphiData(0)
    ReDim Preserve masterPgesData(0)
    ReDim Preserve masterPwrIndData(0)
    ReDim Preserve masterPwrBandgesData(0)
    ReDim Preserve masterWirkungsgradData(0)
    
    ReDim Preserve zfResults(0)
    
    electricFileName = "C:\Users\Gäste_PC_2\Desktop\Tailored_Heating_ML\PRM_SET1\electric_data.dat"
    
    temperatureFileName = "C:\Users\Gäste_PC_2\Desktop\Tailored_Heating_ML\PRM_SET1\Temp_data\Surface\SurfTemp0.txt"
    temperatureFileNumber = 0
    
    'The following code loops through the nStep count to call the files from each PRM_SET file
    Do Until parameterFilesLoopCount = nStep + 1
        
        'Surface Data ****START****
        Do Until temperatureFileNumber = iiii
            temperatureFileNumber = temperatureFileNumber + 1
            temperatureFileName = StrReverse(Replace(StrReverse(temperatureFileName), StrReverse(temperatureFileNumber - 1), StrReverse(temperatureFileNumber), count:=1))
            Set temporaryCollectionOfSurfaceFilesData = CoreAndSurfaceData(temperatureFileName)
            ReDim Preserve surfaceTemperatureData(0 To temperatureFileNumber - 1)
            ReDim Preserve surfaceTimeData(0 To temperatureFileNumber - 1)
            
            'Get Item 1 into surfaceTempData
            surfaceTemperatureData(temperatureFileNumber - 1) = temporaryCollectionOfSurfaceFilesData.Item(1)
            
            'Get Item 2 into yPosSurface (same for ALL surface files so only happens once)
            If temperatureFileNumber = 1 Then
                yPosSurface = temporaryCollectionOfSurfaceFilesData.Item(2)
            End If
            
            'Get Item 3 into coreTimeData
            surfaceTimeData(temperatureFileNumber - 1) = temporaryCollectionOfSurfaceFilesData.Item(3)
        Loop
        'Surface Data ****END****
        
        'Store Surface data from files permanently in master arrays
        ReDim Preserve masterSurfaceTemperatureData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterSurfaceTimeData(0 To parameterFilesLoopCount - 1)
        masterSurfaceTemperatureData(parameterFilesLoopCount - 1) = surfaceTemperatureData()
        masterSurfaceTimeData(parameterFilesLoopCount - 1) = surfaceTimeData()
        
        'Put Temp_data file back at 0
        temperatureFileName = StrReverse(Replace(StrReverse(temperatureFileName), StrReverse(temperatureFileNumber), StrReverse(0), count:=1))
        temperatureFileNumber = 0
        
        'Change file to Core data instead of Surface
        temperatureFileName = Replace(temperatureFileName, "Surface", "Core", count:=1)
        temperatureFileName = Replace(temperatureFileName, "Surf", "Core", count:=1)
        
        'Core Data ****START****
        Do Until temperatureFileNumber = iiii
            temperatureFileNumber = temperatureFileNumber + 1
            temperatureFileName = StrReverse(Replace(StrReverse(temperatureFileName), StrReverse(temperatureFileNumber - 1), StrReverse(temperatureFileNumber), count:=1))
            Set temporaryCollectionOfCoreFilesData = CoreAndSurfaceData(temperatureFileName)
            ReDim Preserve coreTemperatureData(0 To temperatureFileNumber - 1)
            ReDim Preserve coreTimeData(0 To temperatureFileNumber - 1)
            
            'Get Item 1 into coreTemperatureData
            coreTemperatureData(temperatureFileNumber - 1) = temporaryCollectionOfCoreFilesData.Item(1)
            
            'Get Item 2 into yPosCore (same for ALL core files so only happens once)
            If temperatureFileNumber = 1 Then
                yPosCore = temporaryCollectionOfCoreFilesData.Item(2)
            End If
            
            'Get Item 3 into coreTimeData
            coreTimeData(temperatureFileNumber - 1) = temporaryCollectionOfCoreFilesData.Item(3)
        Loop
        'Core Data ****END****
        
        'Store Core data from files permanently in master arrays
        ReDim Preserve masterCoreTemperatureData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterCoreTimeData(0 To parameterFilesLoopCount - 1)
        masterCoreTemperatureData(parameterFilesLoopCount - 1) = coreTemperatureData()
        masterCoreTimeData(parameterFilesLoopCount - 1) = coreTimeData()
        
        'Electric Data ****START****
        Set electricDataCollection = ElectricData(electricFileName)
        
        ReDim Preserve masterFrequencyData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterVoltageData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterIAData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterCosphiData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterPgesData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterPwrIndData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterPwrBandgesData(0 To parameterFilesLoopCount - 1)
        ReDim Preserve masterWirkungsgradData(0 To parameterFilesLoopCount - 1)
        
        masterFrequencyData(parameterFilesLoopCount - 1) = electricDataCollection.Item(1)
        masterVoltageData(parameterFilesLoopCount - 1) = electricDataCollection.Item(2)
        masterIAData(parameterFilesLoopCount - 1) = electricDataCollection.Item(3)
        masterCosphiData(parameterFilesLoopCount - 1) = electricDataCollection.Item(4)
        masterPgesData(parameterFilesLoopCount - 1) = electricDataCollection.Item(5)
        masterPwrIndData(parameterFilesLoopCount - 1) = electricDataCollection.Item(6)
        masterPwrBandgesData(parameterFilesLoopCount - 1) = electricDataCollection.Item(7)
        masterWirkungsgradData(parameterFilesLoopCount - 1) = electricDataCollection.Item(8)
        'Electric Data ****END****
        
        'Get Target Data
        ReDim Preserve zfResults(0 To parameterFilesLoopCount - 1)
        
        Dim count As Integer
        Dim calculateGraphTemporaryCollection As Collection
        
        count = 0
        
        If temperatureFileNumber = 20 Then
            count = count + 1
            Set calculateGraphTemporaryCollection = CalculateGraph(coreTemperatureData(temperatureFileNumber - 1), surfaceTemperatureData(temperatureFileNumber - 1))
            zfResults(parameterFilesLoopCount - 1) = calculateGraphTemporaryCollection.Item(1)
            If count = 1 Then
                ZfTemperatures = calculateGraphTemporaryCollection.Item(2)
                ZfPositions = calculateGraphTemporaryCollection.Item(3)
            End If
        End If
        
        'Change PRM_SET file number for Temperature Data
        parameterFilesLoopCount = parameterFilesLoopCount + 1
        electricFileName = StrReverse(Replace(StrReverse(electricFileName), StrReverse(parameterFilesLoopCount - 1), StrReverse(parameterFilesLoopCount), count:=1))
        If parameterFilesLoopCount < 3 Then
            temperatureFileName = StrReverse(Replace(StrReverse(temperatureFileName), StrReverse(parameterFilesLoopCount - 1), StrReverse(parameterFilesLoopCount), count:=2))
        Else
            Dim firstThird As String, secondThird As String, thirdThird As String
            firstThird = Left(temperatureFileName, 48)
            If parameterFilesLoopCount < 11 Then
                secondThird = Mid(temperatureFileName, 49, 19)
                thirdThird = Right(temperatureFileName, 19)
            Else
                secondThird = Mid(temperatureFileName, 49, 20)
                thirdThird = Right(temperatureFileName, 19)
            End If
            secondThird = Replace(secondThird, parameterFilesLoopCount - 1, parameterFilesLoopCount, count:=1)
            temperatureFileName = firstThird + secondThird + thirdThird
        End If
        
        'Put Temp_data file back at 0
        temperatureFileName = StrReverse(Replace(StrReverse(temperatureFileName), StrReverse(temperatureFileNumber), StrReverse(0), count:=1))
        temperatureFileNumber = 0
        
        'Change file to Surface data instead of Core
        temperatureFileName = Replace(temperatureFileName, "Core", "Surface", count:=1)
        temperatureFileName = Replace(temperatureFileName, "Core", "Surf", count:=1)
        
    Loop
    
    
    Call WriteData(nStep, iiii, zfResults, yPosSurface, yPosCore, masterSurfaceTemperatureData(), masterSurfaceTimeData(), masterCoreTemperatureData(), masterCoreTimeData(), masterFrequencyData(), masterVoltageData(), masterIAData(), masterCosphiData(), masterPgesData(), masterPwrIndData(), masterPwrBandgesData(), masterWirkungsgradData())
    
End Sub


Public Sub WriteData(nStep As Integer, iiii As Integer, zfResults() As Variant, yPosSurface As Variant, yPosCore As Variant, surfaceTemperatureData() As Variant, surfaceTimeData() As Variant, coreTemperatureData() As Variant, coreTimeData() As Variant, frequencyData() As Variant, voltageData() As Variant, IAData() As Variant, cosphiData() As Variant, pgesData() As Variant, pwrIndData() As Variant, pwrBandgesData() As Variant, wirkungsgradData() As Variant)
    'Set up the sheets
    Dim coreTempNamed As Boolean, surfTempNamed As Boolean, electricDataNamed As Boolean, zfGraphNamed As Boolean, zfDataNamed As Boolean, chosenDataNamed As Boolean, chosenDataGraphNamed As Boolean
    coreTempNamed = False
    surfTempNamed = False
    electricDataNamed = False
    zfGraphNamed = False
    zfDataNamed = False
    chosenDataNamed = False
    chosenDataGraphNamed = False
    For Each Sheet In Worksheets
        If Sheet.Name = "CoreTemp" Then
            coreTempNamed = True
        ElseIf Sheet.Name = "SurfTemp" Then
            surfTempNamed = True
        ElseIf Sheet.Name = "ElectricData" Then
            electricDataNamed = True
        ElseIf Sheet.Name = "ZfGraph" Then
            zfGraphNamed = True
        ElseIf Sheet.Name = "Zf" Then
            zfDataNamed = True
        ElseIf Sheet.Name = "ChosenDataGraph" Then
            chosenDataGraphNamed = True
        ElseIf Sheet.Name = "ChosenData" Then
            chosenDataNamed = True
        End If
    Next Sheet
    
    If Not surfTempNamed Then
        Sheets.Add.Name = "SurfTemp"
    End If
    If Not coreTempNamed Then
        Sheets.Add.Name = "CoreTemp"
    End If
    If Not electricDataNamed Then
        Sheets.Add.Name = "ElectricData"
    End If
    If Not zfGraphNamed Then
        Sheets.Add.Name = "ZfGraph"
    End If
    If Not zfDataNamed Then
        Sheets.Add.Name = "Zf"
    End If
    If Not chosenDataGraphNamed Then
        Sheets.Add.Name = "ChosenDataGraph"
    End If
    If Not chosenDataNamed Then
        Sheets.Add.Name = "ChosenData"
    End If
    
    'Populate CoreTemp
    Sheets("CoreTemp").Select
    Cells.Clear
    Cells(1, 1) = "nStep"
    Cells(2, 1) = nStep
    
    Dim columnNameCounter As Integer, parameterSetFileCounter As Integer, fileWithinParameterSetCounter As Integer, iiiiCounter As Integer
    Dim temperatureLineDataCounter As Integer, rowCounter As Integer
    Dim electricCount As Integer
    
    columnNameCounter = 0
    iiiiCounter = 0
    parameterSetFileCounter = 1
    fileWithinParameterSetCounter = 2
    rowCounter = 7
    temperatureLineDataCounter = 0
    
    temperatureLowestLevelSize = UBound(coreTemperatureData(0)(0)) - LBound(coreTemperatureData(0)(0))
    
    Do Until parameterSetFileCounter = nStep + 1
        
        Cells(4, fileWithinParameterSetCounter) = "PRM_SET" & parameterSetFileCounter
        
        fileWithinParameterSetCounter = fileWithinParameterSetCounter + 20
        
        Do Until iiiiCounter = iiii
            iiiiCounter = iiiiCounter + 1
            columnNameCounter = columnNameCounter + 1
            Cells(5, columnNameCounter + 1) = "CoreTemp" & iiiiCounter
            Cells(6, columnNameCounter + 1) = "Time:" & coreTimeData(parameterSetFileCounter - 1)(iiiiCounter - 1)
            
            Do Until temperatureLineDataCounter = temperatureLowestLevelSize + 1
                Cells(rowCounter, columnNameCounter + 1) = coreTemperatureData(parameterSetFileCounter - 1)(iiiiCounter - 1)(temperatureLineDataCounter)
                temperatureLineDataCounter = temperatureLineDataCounter + 1
                rowCounter = rowCounter + 1
            Loop
            temperatureLineDataCounter = 0
            rowCounter = 7
        Loop
        parameterSetFileCounter = parameterSetFileCounter + 1
        iiiiCounter = 0
        
        fileWithinParameterSetCounter = fileWithinParameterSetCounter + 1
        columnNameCounter = columnNameCounter + 1
    Loop
    
    Do Until temperatureLineDataCounter = temperatureLowestLevelSize + 1
        Cells(rowCounter, 1) = yPosCore(temperatureLineDataCounter)
        temperatureLineDataCounter = temperatureLineDataCounter + 1
        rowCounter = rowCounter + 1
    Loop
    
    'Populate SurfTemp
    Sheets("SurfTemp").Select
    Cells.Clear
    Cells(1, 1) = "nStep"
    Cells(2, 1) = nStep
    
    columnNameCounter = 0
    iiiiCounter = 0
    parameterSetFileCounter = 1
    fileWithinParameterSetCounter = 2
    rowCounter = 7
    temperatureLineDataCounter = 0
    
    Do Until parameterSetFileCounter = nStep + 1
        
        Cells(4, fileWithinParameterSetCounter) = "PRM_SET" & parameterSetFileCounter
        
        fileWithinParameterSetCounter = fileWithinParameterSetCounter + 20
        
        Do Until iiiiCounter = iiii
            iiiiCounter = iiiiCounter + 1
            columnNameCounter = columnNameCounter + 1
            Cells(5, columnNameCounter + 1) = "SurfTemp" & iiiiCounter
            Cells(6, columnNameCounter + 1) = "Time:" & surfaceTimeData(parameterSetFileCounter - 1)(iiiiCounter - 1)
            
            Do Until temperatureLineDataCounter = temperatureLowestLevelSize + 1
                Cells(rowCounter, columnNameCounter + 1) = surfaceTemperatureData(parameterSetFileCounter - 1)(iiiiCounter - 1)(temperatureLineDataCounter)
                temperatureLineDataCounter = temperatureLineDataCounter + 1
                rowCounter = rowCounter + 1
            Loop
            temperatureLineDataCounter = 0
            rowCounter = 7
        Loop
        parameterSetFileCounter = parameterSetFileCounter + 1
        iiiiCounter = 0
        
        fileWithinParameterSetCounter = fileWithinParameterSetCounter + 1
        columnNameCounter = columnNameCounter + 1
    Loop
    
    Do Until temperatureLineDataCounter = temperatureLowestLevelSize + 1
        Cells(rowCounter, 1) = yPosSurface(temperatureLineDataCounter)
        temperatureLineDataCounter = temperatureLineDataCounter + 1
        rowCounter = rowCounter + 1
    Loop
    
    'Populate ElectricData
    Sheets("ElectricData").Select
    Cells.Clear
    Cells(1, 1) = "nStep"
    Cells(2, 1) = nStep
    
    columnNameCounter = 0
    iiiiCounter = 0
    parameterSetFileCounter = 1
    fileWithinParameterSetCounter = 2
    rowCounter = 7
    temperatureLineDataCounter = 0
    
    Do Until parameterSetFileCounter = nStep + 1
        
        Cells(4, fileWithinParameterSetCounter) = "PRM_SET" & parameterSetFileCounter
        
        fileWithinParameterSetCounter = fileWithinParameterSetCounter + 20
        
        electricCount = 1
        Do Until iiiiCounter = iiii
            iiiiCounter = iiiiCounter + 1
            columnNameCounter = columnNameCounter + 1
            Cells(5, columnNameCounter + 1) = "Time:" & surfaceTimeData(parameterSetFileCounter - 1)(iiiiCounter - 1)
            
            Cells(6, columnNameCounter + 1) = frequencyData(parameterSetFileCounter - 1)
            Cells(7, columnNameCounter + 1) = IAData(parameterSetFileCounter - 1)
            
            Cells(8, columnNameCounter + 1) = voltageData(parameterSetFileCounter - 1)(electricCount)
            Cells(9, columnNameCounter + 1) = cosphiData(parameterSetFileCounter - 1)(electricCount)
            Cells(10, columnNameCounter + 1) = pgesData(parameterSetFileCounter - 1)(electricCount)
            Cells(11, columnNameCounter + 1) = pwrIndData(parameterSetFileCounter - 1)(electricCount)
            Cells(12, columnNameCounter + 1) = pwrBandgesData(parameterSetFileCounter - 1)(electricCount)
            Cells(13, columnNameCounter + 1) = wirkungsgradData(parameterSetFileCounter - 1)(electricCount)
            
            electricCount = electricCount + 1
            
        Loop
        parameterSetFileCounter = parameterSetFileCounter + 1
        iiiiCounter = 0
        
        fileWithinParameterSetCounter = fileWithinParameterSetCounter + 1
        columnNameCounter = columnNameCounter + 1
    Loop
    
    Cells(6, 1) = "Frequency"
    Cells(7, 1) = "IA"
    Cells(8, 1) = "Voltage"
    Cells(9, 1) = "Cosphi"
    Cells(10, 1) = "Pges"
    Cells(11, 1) = "PwrInd"
    Cells(12, 1) = "PwrBand_ges"
    Cells(13, 1) = "Wirkungsgrad"
    
    'Create Target Graph
    Sheets("Zf").Select
    Cells.Clear
    
    Dim j As Integer
    j = 0
    
    Do Until j = nStep
        Cells(j + 1, 1) = zfResults(j)
        j = j + 1
    Loop
    
    'Create Target Graph
    Sheets("ZfGraph").Select
    ActiveSheet.ChartObjects.Delete
    
    Dim targetDataEmbeddedChart As ChartObject

    Set targetDataEmbeddedChart = Sheets("ZfGraph").ChartObjects.Add(Left:=200, Width:=600, Top:=10, Height:=400)
    targetDataEmbeddedChart.chart.ChartType = xlLine
    targetDataEmbeddedChart.chart.SetSourceData Source:=Sheets("Zf").Range("A1:A" & nStep)
    
    Dim chooseData As Button, t As Range
    
    Application.ScreenUpdating = False
    ActiveSheet.Buttons.Delete
    
    Set t = ActiveSheet.Range(Cells(12, 1), Cells(13, 3))
    Set chooseData = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    chooseData.OnAction = "GraphChosenData"
    chooseData.Caption = "Choose Data File"
    chooseData.Name = "ChooseDataFile"
    Application.ScreenUpdating = True
End Sub


Public Function GraphChosenData()
    Dim MyValue As String
    MyValue = InputBox("Enter the file number", "PRM_SET File Selection", 1)
    
    Dim nextEmptyRowString As String, nextEmptyRowInteger As Integer, error As Boolean
    
    error = False
    
    Sheets("CoreTemp").Select
    nextEmptyRowString = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Offset(1).Address
    nextEmptyRowInteger = Mid(nextEmptyRowString, 4)
    
    On Error GoTo Canceled:
    ActiveSheet.Range(Cells(8, 1), Cells(nextEmptyRowInteger - 1, 1)).Copy Worksheets("ChosenData").Range("A1:A" & nextEmptyRowInteger - 8)
    ActiveSheet.Range(Cells(8, 21 * MyValue), Cells(nextEmptyRowInteger - 1, 21 * MyValue)).Copy Worksheets("ChosenData").Range("B1:B" & nextEmptyRowInteger - 8)
    
    Sheets("SurfTemp").Select
    On Error GoTo Canceled:
    ActiveSheet.Range(Cells(8, 21 * MyValue), Cells(nextEmptyRowInteger - 1, 21 * MyValue)).Copy Worksheets("ChosenData").Range("C1:C" & nextEmptyRowInteger - 8)
    
    Sheets("ChosenDataGraph").Select
    ActiveSheet.ChartObjects.Delete
    
    Dim chosenDataEmbeddedChart As ChartObject

    Set chosenDataEmbeddedChart = Sheets("ChosenDataGraph").ChartObjects.Add(Left:=200, Width:=600, Top:=10, Height:=400)
    chosenDataEmbeddedChart.chart.ChartType = xlLine
    With chosenDataEmbeddedChart.chart.SeriesCollection.NewSeries
        .Name = "Core Temperature"
        .XValues = Sheets("ChosenData").Range("A1:A" & nextEmptyRowInteger - 8)
        .Values = Sheets("ChosenData").Range("B1:B" & nextEmptyRowInteger - 8)
    End With
    With chosenDataEmbeddedChart.chart.SeriesCollection.NewSeries
        .Name = "Surf Temperature"
        .XValues = Sheets("ChosenData").Range("A1:A" & nextEmptyRowInteger - 8)
        .Values = Sheets("ChosenData").Range("C1:C" & nextEmptyRowInteger - 8)
    End With
    With chosenDataEmbeddedChart.chart.SeriesCollection.NewSeries
        .Name = "Target Temperature"
        .XValues = Sheets("ChosenData").Range("A1:A" & nextEmptyRowInteger - 8)
        .Values = Sheets("ChosenData").Range("D1:D" & nextEmptyRowInteger - 8)
        .MarkerStyle = xlMarkerStyleCircle
    End With
    GoTo endProcess
    
Canceled:
    MsgBox ("Error occured.")
    GoTo endProcess
endProcess:
End Function


Public Function CalculateGraph(coreData As Variant, surfaceData As Variant) As Collection
    Dim Zf As Variant
    Dim Position As Variant
    Dim zfRangeString As String
    Dim positionRangeString As String
    Dim chosenZfData() As Variant
    Dim chosenPositionData(0 To 6) As Variant
    Dim collectionOfData As Collection
    Set collectionOfData = New Collection
    
    temperatureLowestLevelSize = UBound(coreData) - LBound(coreData)
    
    zfRangeString = "B1:B" & temperatureLowestLevelSize
    
    positionRangeString = "A1:A" & temperatureLowestLevelSize
    
    Zf = Sheets("YPos&TempForZf").Range(zfRangeString)
    
    Position = Sheets("YPos&TempForZf").Range(positionRangeString)
    
    ReDim Preserve chosenZfData(0)
    
    Dim positionOne As Integer
    Dim positionTwo As Integer
    Dim positionThree As Integer
    Dim positionFour As Integer
    Dim positionFive As Integer
    Dim positionSix As Integer
    Dim positionSeven As Integer
    Dim equation As Double
    Dim n As Integer
    Dim c As Double
    Dim s As Double
    Dim t As Double
    Dim leftSide As Double
    Dim rightSide As Double
    Dim inside As Double
    
    positionOne = 8
    positionTwo = 19
    positionThree = 27
    positionFour = 39
    positionFive = 51
    positionSix = 60
    positionSeven = 70
    
    chosenPositionData(0) = Position(positionOne + 1, 1)
    chosenPositionData(1) = Position(positionTwo + 1, 1)
    chosenPositionData(2) = Position(positionThree + 1, 1)
    chosenPositionData(3) = Position(positionFour + 1, 1)
    chosenPositionData(4) = Position(positionFive + 1, 1)
    chosenPositionData(5) = Position(positionSix + 1, 1)
    chosenPositionData(6) = Position(positionSeven + 1, 1)
    
    equation = 0
    
    n = 1
    
    leftSide = 0
    rightSide = 0
    inside = 0
    
    Do Until n = 8
        If n = 1 Then
            c = coreData(positionOne)
            s = surfaceData(positionOne)
            t = Zf(positionOne + 1, 1)
            chosenZfData(n - 1) = Zf(positionOne + 1, 1)
        ElseIf n = 2 Then
            c = coreData(positionTwo)
            s = surfaceData(positionTwo)
            t = Zf(positionTwo + 1, 1)
            ReDim Preserve chosenZfData(0 To n - 1)
            chosenZfData(n - 1) = Zf(positionTwo + 1, 1)
        ElseIf n = 3 Then
            c = coreData(positionThree)
            s = surfaceData(positionThree)
            t = Zf(positionThree + 1, 1)
            ReDim Preserve chosenZfData(0 To n - 1)
            chosenZfData(n - 1) = Zf(positionThree + 1, 1)
        ElseIf n = 4 Then
            c = coreData(positionFour)
            s = surfaceData(positionFour)
            t = Zf(positionFour + 1, 1)
            ReDim Preserve chosenZfData(0 To n - 1)
            chosenZfData(n - 1) = Zf(positionFour + 1, 1)
        ElseIf n = 5 Then
            c = coreData(positionFive)
            s = surfaceData(positionFive)
            t = Zf(positionFive + 1, 1)
            ReDim Preserve chosenZfData(0 To n - 1)
            chosenZfData(n - 1) = Zf(positionFive + 1, 1)
        ElseIf n = 6 Then
            c = coreData(positionSix)
            s = surfaceData(positionSix)
            t = Zf(positionSix + 1, 1)
            ReDim Preserve chosenZfData(0 To n - 1)
            chosenZfData(n - 1) = Zf(positionSix + 1, 1)
        ElseIf n = 7 Then
            c = coreData(positionSeven)
            s = surfaceData(positionSeven)
            t = Zf(positionSeven + 1, 1)
            ReDim Preserve chosenZfData(0 To n - 1)
            chosenZfData(n - 1) = Zf(positionSeven + 1, 1)
        End If
        
        leftSide = t - c
        leftSide = leftSide * leftSide
        rightSide = t - s
        rightSide = rightSide * rightSide
        inside = leftSide + rightSide
        equation = equation + Sqr(inside)
        
        n = n + 1
    Loop
    
    equation = equation / 7
    
    
    collectionOfData.Add (equation)
    collectionOfData.Add (chosenZfData())
    collectionOfData.Add (chosenPositionData())
    
    Set CalculateGraph = collectionOfData
End Function



Public Function ElectricData(file As String) As Collection
    Dim textline As String, i As Long
    Dim frequency As String, voltage() As String, IA As String, cosphi() As String, pges() As String, pwrInd() As String, pwrBandges() As String, wirkungsgrad() As String
    Dim collectionOfData As Collection
    Set collectionOfData = New Collection
    
    i = 0
    
    ReDim Preserve voltage(i)
    ReDim Preserve cosphi(i)
    ReDim Preserve pges(i)
    ReDim Preserve pwrInd(i)
    ReDim Preserve pwrBandges(i)
    ReDim Preserve wirkungsgrad(i)
    
    Open file For Input As #1
    
    Do Until EOF(1)
        Line Input #1, textline
        textline = Mid(textline, 9)
        
        ReDim Preserve voltage(0 To i)
        ReDim Preserve cosphi(0 To i)
        ReDim Preserve pges(0 To i)
        ReDim Preserve pwrInd(0 To i)
        ReDim Preserve pwrBandges(0 To i)
        ReDim Preserve wirkungsgrad(0 To i)
        
        voltage(i) = Mid(textline, 19, 14)
        cosphi(i) = Mid(textline, 47, 15)
        
        pges(i) = Mid(textline, 63, 15)
        pwrInd(i) = Mid(textline, 78, 14)
        
        pwrBandges(i) = Mid(textline, 92, 16)
        wirkungsgrad(i) = Right(textline, 12)
        
        If i = 1 Then
            frequency = Left(textline, 18)
            IA = Mid(textline, 32, 14)
        End If
        
        i = i + 1
    Loop
    Close #1
    
    collectionOfData.Add (frequency)
    collectionOfData.Add (voltage())
    collectionOfData.Add (IA)
    collectionOfData.Add (cosphi())
    collectionOfData.Add (pges())
    collectionOfData.Add (pwrInd())
    collectionOfData.Add (pwrBandges())
    collectionOfData.Add (wirkungsgrad())
    
    Set ElectricData = collectionOfData
End Function


Public Function CoreAndSurfaceData(file As String) As Collection
    Dim textline As String, tmp() As String, yPos() As String, time As String, i As Long
    Dim collectionOfData As Collection
    Set collectionOfData = New Collection
    
    i = 0
    
    ReDim Preserve tmp(i)
    ReDim Preserve yPos(i)
    
    Open file For Input As #1
    
    Do Until EOF(1)
        Line Input #1, textline
        
        ReDim Preserve tmp(0 To i)
        ReDim Preserve yPos(0 To i)
        
        tmp(i) = Left(textline, 14)
        yPos(i) = Mid(textline, 14, 10)
        If i = 1 Then
            time = Right(textline, 8)
        End If
        
        i = i + 1
    Loop
    Close #1
    
    collectionOfData.Add (tmp())
    collectionOfData.Add (yPos())
    collectionOfData.Add (time)
    
    Set CoreAndSurfaceData = collectionOfData
End Function
