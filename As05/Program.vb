'------------------------------------------------------------
'-                File Name : Program.vb                    - 
'-                Part of Project: Assign5                  -
'------------------------------------------------------------
'-                Written By: Alex Buckstiegel              -
'-                Written On: 02-18-20                      -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the main application form where the   -
'- user will input a file name and then get the report      - 
'------------------------------------------------------------
'- Program Purpose:                                         -
'- This program creates a report of employee sales using    -
'- generic data containers and LINQ                     
'------------------------------------------------------------
'- Global Variable Dictionary (alphabetically):             -
'- lstEmployee - List of all inputted employees             -
'- strInputFile - String containing the file path of input  -
'------------------------------------------------------------

Imports System
Imports System.Text
Module Program

    Public Class clsEmployee
        Property strFirstName As String
        Property strLastName As String
        Property intOrderID As Integer
        Property intID As Integer
        Property sngGameSales As Single
        Property intGameQuantity As Integer
        Property sngDollSales As Single
        Property intDollQuantity As Integer
        Property sngBuildingSales As Single
        Property intBuildingQuantity As Integer
        Property sngModelSales As Single
        Property intModelQuantity As Integer
        Property sngTotalSales As Single

        '------------------------------------------------------------
        '-                Subprogram Name: New                      -
        '------------------------------------------------------------
        '-                Written By: Alex Buckstiegel              -
        '-                Written On: 02-18-20                      -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This subroutine takes all the variables and assigns them –
        '- to the corresponding property
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- fname - first name
        '- lname - last name
        '- orderID - OrderI ID
        '- id - Employee id
        '- gamesale - Games sales
        '- gameint - Quantity of games sold
        '- dollsale - dolls sales
        '- dollint -Quantity of dolls sold
        '- buildsale - Building sales
        '- buildint - Quantity of Buildings sold
        '- modelsale - Model sales
        '- modelint - Quantity of models sold
        '- totalsale - total sales(sum of gamesale, dollsale, buildsale, and modelsale)
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        Public Sub New(fname As String, lname As String, orderID As Integer, id As Integer, gamesale As Single, gameint As Integer, dollsale As Single, dollint As Integer,
        buildsale As Single, buildint As Integer, modelsale As Single, modelint As Integer, totalsale As Single)
            'Me.Property = input             'where it is in array
            Me.strFirstName = fname         '0
            Me.strLastName = lname          '1
            Me.intOrderID = orderID         '2
            Me.intID = id                   '3
            Me.sngGameSales = gamesale      '4
            Me.intGameQuantity = gameint    '5
            Me.sngDollSales = dollsale      '6
            Me.intDollQuantity = dollint    '7
            Me.sngBuildingSales = buildsale '8
            Me.intBuildingQuantity = buildint '9
            Me.sngModelSales = modelsale     '10
            Me.intModelQuantity = modelint  '11
            Me.sngTotalSales = totalsale    '4+6+8+10
        End Sub
    End Class
    'Public List of Employees
    Public lstEmployees As New List(Of clsEmployee)
    Dim strInputFile As String
    '------------------------------------------------------------
    '-                Subprogram Name: Main                     -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 02-18-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is the main sub from where everythingh runs
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- args - String that does something
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- s - for final read
    '------------------------------------------------------------

    Sub Main(args As String())

        InputFileDirectories()
        InputData()
        'Sorts to intID
        lstEmployees = lstEmployees.OrderBy(Function(x) x.intOrderID).ToList
        Console.Write(PrintReport)
        'Sorts to strLastName
        lstEmployees = lstEmployees.OrderBy(Function(x) x.strLastName).ToList
        Console.Write(PrintReport)
        AssignSalesStats()
        CalcAboveAvg()
        Console.WriteLine("Press 'Enter' to exit the Application")
        Dim s = Console.Read()
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: InputFileDirectories     -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 02-18-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Gets and validates input file that it exists
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (none)
    '------------------------------------------------------------
    Public Sub InputFileDirectories()

        Console.WriteLine("Please enter the path and the name of the file containing the measurements: ")

        'strInputFile = Console.ReadLine()
        strInputFile = "ToyOrder.txt"     'For testing
        If System.IO.File.Exists(strInputFile) Then
        Else
            Console.WriteLine("This File does not exist!")
            InputFileDirectories()
        End If
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: InputData                -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 02-18-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Opens file and puts variables into clsEmployees and lstEmployees
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- strTempArray - temp array of the split line
    '- total - holds value of total sales
    '------------------------------------------------------------
    Sub InputData()
        Dim objStreamReader As System.IO.StreamReader
        objStreamReader = System.IO.File.OpenText(strInputFile)
        While Not (objStreamReader.EndOfStream)
            Dim strTempArray() = Split(objStreamReader.ReadLine(), " ")
            Dim total = CSng(strTempArray(4)) + CSng(strTempArray(6)) + CSng(strTempArray(8)) + CSng(strTempArray(10))
            lstEmployees.Add(New clsEmployee(strTempArray(0), strTempArray(1), strTempArray(2), strTempArray(3), strTempArray(4), strTempArray(5),
            strTempArray(6), strTempArray(7), strTempArray(8), strTempArray(9), strTempArray(10), strTempArray(11), total))
        End While
        objStreamReader.Close()
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: PrintReport              -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 02-18-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Prints the first 2 reports
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- BobTheBuilder - Stringbuilder for sub
    '- strFinal - Final string that is written to the console
    '------------------------------------------------------------
    Function PrintReport()
        Dim BobTheBuilder As New StringBuilder
        Dim strFinal
        BobTheBuilder.AppendLine(vbTab & vbTab & vbTab & "     Waste O' Time Toys")
        BobTheBuilder.AppendLine(vbTab & vbTab & vbTab & "***Sales Report by Order***")
        BobTheBuilder.AppendLine(StrDup(120, "-"))
        BobTheBuilder.AppendLine(" ")
        BobTheBuilder.AppendLine("Name" & StrDup(3, vbTab) & "ID" & StrDup(2, vbTab) & "Games" & StrDup(2, vbTab) & "Dolls" & StrDup(2, vbTab) & "Buildings" & StrDup(1, vbTab) & "Models" & StrDup(2, vbTab) & "Total")
        For Each employee In lstEmployees
            If employee.strFirstName.Length + employee.strLastName.Length > 13 Then     '13 is the magic number that lines them all up. If the Name is longer than that, they will not be lined up
                BobTheBuilder.AppendLine(employee.strLastName & ", " & employee.strFirstName & StrDup(1, vbTab) & employee.intOrderID.ToString("000") & StrDup(2, vbTab) & Format(employee.sngGameSales, "C") & StrDup(2, vbTab) & Format(employee.sngDollSales, "C") & StrDup(2, vbTab) & Format(employee.sngBuildingSales, "C") & StrDup(2, vbTab) & Format(employee.sngModelSales, "C") & StrDup(2, vbTab) & Format(employee.sngTotalSales, "C"))
            Else
                BobTheBuilder.AppendLine(employee.strLastName & ", " & employee.strFirstName & StrDup(2, vbTab) & employee.intOrderID.ToString("000") & StrDup(2, vbTab) & Format(employee.sngGameSales, "C") & StrDup(2, vbTab) & Format(employee.sngDollSales, "C") & StrDup(2, vbTab) & Format(employee.sngBuildingSales, "C") & StrDup(2, vbTab) & Format(employee.sngModelSales, "C") & StrDup(2, vbTab) & Format(employee.sngTotalSales, "C"))
            End If
        Next
        strFinal = BobTheBuilder.ToString()
        Return strFinal
    End Function

    '------------------------------------------------------------
    '-                Subprogram Name: AssignSalesStats         -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 02-18-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Uses LINQ to aggregate all the data and puts everything in
    '- correct arrays
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- allstats() - array of the arrays for use in other subs
    '- avgQ() - Average Quantity
    '- avgS() - Average Sales
    '- hiQ() - Maximum Quantity
    '- hiS() - Maximum Sales
    '- lowQ() - Minimum Quantity
    '- lowS() - Minimum Sales
    '------------------------------------------------------------
    Sub AssignSalesStats()
        Const GAMES = 0
        Const DOLLS = 1
        Const BUILDING = 2
        Const MODEL = 3
        Const OVERALLTOTAL = 4

        'Could be one array, but this is easier to read

        Dim lowQ(3) As Integer
        Dim avgQ(3) As Single
        Dim hiQ(3) As Integer

        Dim lowS(3) As Single
        Dim avgS(3) As Single
        Dim hiS(3) As Single
        Dim totalS(4) As Single

        Dim allstats As New List(Of Array)({lowQ, avgQ, hiQ, lowS, avgS, hiS, totalS})

        'This could definitely be more efficient, I just can't think of a way to loop it correctly and
        'have it work so this will have to do
        lowQ(GAMES) = Aggregate emps In lstEmployees Into Min(emps.intGameQuantity)
        lowQ(DOLLS) = Aggregate emps In lstEmployees Into Min(emps.intDollQuantity)
        lowQ(BUILDING) = Aggregate emps In lstEmployees Into Min(emps.intBuildingQuantity)
        lowQ(MODEL) = Aggregate emps In lstEmployees Into Min(emps.intModelQuantity)

        avgQ(GAMES) = Aggregate emps In lstEmployees Into Average(emps.intGameQuantity)
        avgQ(DOLLS) = Aggregate emps In lstEmployees Into Average(emps.intDollQuantity)
        avgQ(BUILDING) = Aggregate emps In lstEmployees Into Average(emps.intBuildingQuantity)
        avgQ(MODEL) = Aggregate emps In lstEmployees Into Average(emps.intModelQuantity)

        hiQ(GAMES) = Aggregate emps In lstEmployees Into Max(emps.intGameQuantity)
        hiQ(DOLLS) = Aggregate emps In lstEmployees Into Max(emps.intDollQuantity)
        hiQ(BUILDING) = Aggregate emps In lstEmployees Into Max(emps.intBuildingQuantity)
        hiQ(MODEL) = Aggregate emps In lstEmployees Into Max(emps.intModelQuantity)

        lowS(GAMES) = Aggregate emps In lstEmployees Into Min(emps.sngGameSales)
        lowS(DOLLS) = Aggregate emps In lstEmployees Into Min(emps.sngDollSales)
        lowS(BUILDING) = Aggregate emps In lstEmployees Into Min(emps.sngBuildingSales)
        lowS(MODEL) = Aggregate emps In lstEmployees Into Min(emps.sngModelSales)

        avgS(GAMES) = Aggregate emps In lstEmployees Into Average(emps.sngGameSales)
        avgS(DOLLS) = Aggregate emps In lstEmployees Into Average(emps.sngDollSales)
        avgS(BUILDING) = Aggregate emps In lstEmployees Into Average(emps.sngBuildingSales)
        avgS(MODEL) = Aggregate emps In lstEmployees Into Average(emps.sngModelSales)

        hiS(GAMES) = Aggregate emps In lstEmployees Into Max(emps.sngGameSales)
        hiS(DOLLS) = Aggregate emps In lstEmployees Into Max(emps.sngDollSales)
        hiS(BUILDING) = Aggregate emps In lstEmployees Into Max(emps.sngBuildingSales)
        hiS(MODEL) = Aggregate emps In lstEmployees Into Max(emps.sngModelSales)

        totalS(GAMES) = Aggregate emps In lstEmployees Into Sum(emps.sngGameSales)
        totalS(DOLLS) = Aggregate emps In lstEmployees Into Sum(emps.sngDollSales)
        totalS(BUILDING) = Aggregate emps In lstEmployees Into Sum(emps.sngBuildingSales)
        totalS(MODEL) = Aggregate emps In lstEmployees Into Sum(emps.sngModelSales)
        totalS(OVERALLTOTAL) = Aggregate items In totalS Into Sum(items)

        PrintSalesStats2(allstats)

    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: PrintSalesStats2         -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 02-18-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- First one was a lot of copy and pasting so this is my redo
    '- with considerably less copy and pasting
    '- Builds the sales stats reports by inputting the values and
    '- a lot of For Each loops
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- allstats as List(Of Array)) - inputs the stats created in 
    '- AssignSalesStats
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- BobTheBuilder - He's back! and building more Strings!
    '- counter - counter for determining which metric is being shown
    '- quantityList - list of the quanity arrays
    '- salesList - list of the sales arrays
    '- strFinal - Final string to be shown

    '------------------------------------------------------------
    Sub PrintSalesStats2(allStats As List(Of Array))
        Dim BobTheBuilder As New StringBuilder
        Dim quantityList As New List(Of Array)({allStats(0), allStats(1), allStats(2)})
        Dim salesList As New List(Of Array)({allStats(3), allStats(4), allStats(5), allStats(6)})
        Dim counter = 0
        Dim strFinal

        BobTheBuilder.AppendLine(StrDup(120, "-"))
        BobTheBuilder.AppendLine()
        BobTheBuilder.AppendLine(StrDup(3, vbTab) & "Sales Statistics Summary")
        BobTheBuilder.AppendLine(StrDup(120, "-"))
        BobTheBuilder.AppendLine()
        BobTheBuilder.AppendLine(StrDup(3, vbTab) & "Quantity Statistics")
        BobTheBuilder.AppendLine(vbTab & "Games" & StrDup(2, vbTab) & "Dolls" & StrDup(2, vbTab) & "Building" & StrDup(1, vbTab) & "Model")
        For Each array In quantityList
            Select Case counter
                Case 0
                    BobTheBuilder.Append("Low:" & vbTab)
                Case 1
                    BobTheBuilder.Append("Avg:" & vbTab)
                Case 2
                    BobTheBuilder.Append("High:" & vbTab)
            End Select
            counter += 1

            For Each item In array
                BobTheBuilder.Append(Format(item, "Standard") & StrDup(2, vbTab))
            Next
            BobTheBuilder.AppendLine()
        Next
        BobTheBuilder.AppendLine()
        BobTheBuilder.AppendLine(StrDup(3, vbTab) & "Sales Statistics")
        counter = 0
        For Each array In salesList
            Select Case counter
                Case 0
                    BobTheBuilder.Append("Low:" & vbTab)
                Case 1
                    BobTheBuilder.Append("Avg:" & vbTab)
                Case 2
                    BobTheBuilder.Append("High:" & vbTab)
                Case 3
                    BobTheBuilder.AppendLine(StrDup(120, "-"))
                    BobTheBuilder.AppendLine()
                    BobTheBuilder.Append("Total:" & vbTab)
            End Select
            counter += 1

            For Each item In array
                If counter <> 4 Then
                    BobTheBuilder.Append(Format(item, "C") & StrDup(2, vbTab))
                Else
                    BobTheBuilder.Append(Format(item, "C") & StrDup(1, vbTab))
                End If

            Next
            BobTheBuilder.AppendLine()
        Next
        strFinal = BobTheBuilder.ToString
        Console.Write(strFinal)
    End Sub

    '------------------------------------------------------------
    '-                Subprogram Name: CalcAboveAvg             -
    '------------------------------------------------------------
    '-                Written By: Alex Buckstiegel              -
    '-                Written On: 02-18-20                      -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Calculates and prints the employees in alphabetical order
    '- that are performing above average
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (none)
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- avgBuildings - average sales for Buildings
    '- avgDolls - average sales for Dolls
    '- avgGames - Average sales for Games
    '- avgModels - average sales for models
    '- lstAboveAvgEmps - List of employees that are above average
    '------------------------------------------------------------
    Sub CalcAboveAvg()
        'I'm not sure exactly is considered above average in this context, but I will be including any employee who sold an above average amount in ANY SALES category
        'As they are technically above average. If I am wrong it is just a matter of changing what is being compared in the if statement below
        Dim avgGames = Aggregate emps In lstEmployees Into Average(emps.sngGameSales)
        Dim avgDolls = Aggregate emps In lstEmployees Into Average(emps.sngDollSales)
        Dim avgBuildings = Aggregate emps In lstEmployees Into Average(emps.sngBuildingSales)
        Dim avgModels = Aggregate emps In lstEmployees Into Average(emps.sngModelSales)

        Dim lstAboveAvgEmps As New List(Of clsEmployee)

        For Each emp In lstEmployees
            If emp.sngGameSales > avgGames Then
                lstAboveAvgEmps.Add(emp)
            ElseIf emp.sngDollSales > avgDolls Then
                lstAboveAvgEmps.Add(emp)
            ElseIf emp.sngBuildingSales > avgBuildings Then
                lstAboveAvgEmps.Add(emp)
            ElseIf emp.sngModelSales > avgModels Then
                lstAboveAvgEmps.Add(emp)
            End If
        Next

        Console.WriteLine("There are " & lstAboveAvgEmps.Count & " employees who sold above the average sales level")
        Console.WriteLine("The names of the above average selling employees in alphabetical order are:")
        lstAboveAvgEmps = lstAboveAvgEmps.OrderBy(Function(x) x.strLastName).ToList
        For Each emp In lstAboveAvgEmps
            Console.WriteLine(emp.strLastName & ", " & emp.strFirstName)
        Next
    End Sub
End Module
