Imports System.IO
Imports System.Windows.Forms

Module Module1

    Sub Main()

        'General Settings
        Dim SkipFirstLine As Boolean = True

        'Fetch inputfile name
        Dim InputFile As String = Nothing
        Dim ExecutableName As String = Process.GetCurrentProcess.MainModule.FileName
        For Each Arg In Environment.GetCommandLineArgs
            If IO.File.Exists(Arg) Then 'arg is filename
                If Not Arg = ExecutableName Then
                    InputFile = Arg 'Inputfile found and set
                End If
            End If
        Next

        If InputFile = Nothing Then 'No file was given
            Dim FileFinder As New OpenFileDialog() With {
                .DefaultExt = "csv",
                .Filter = "Comma Seperated Values|*.csv|All Files|*.*",
                .FilterIndex = 0,
                .Title = "Please select an input file..."}
            If FileFinder.ShowDialog = DialogResult.OK Then
                InputFile = FileFinder.FileName
            Else 'User abort
                Exit Sub
            End If
        End If

        Dim OutputFile As String = Nothing
        Dim SaveSelector As New SaveFileDialog With {
            .DefaultExt = "csv",
            .FileName = "contractperweek.csv",
            .Filter = "Comma Seperated Values|*.csv",
            .FilterIndex = 0,
            .InitialDirectory = Path.GetDirectoryName(InputFile),
            .Title = "Please select a save file location..."}
        If SaveSelector.ShowDialog = DialogResult.OK Then
            OutputFile = SaveSelector.FileName
        Else 'user abort
            Exit Sub
        End If

        Dim AggregatedData As New Dictionary(Of Integer, Integer) '{WeekNumber,NumSwitchers}

        Try

            Using Input As New StreamReader(InputFile)

                If SkipFirstLine Then Input.ReadLine()

                While Not Input.EndOfStream

                    Dim Line As String = Input.ReadLine
                    Dim LineSplit As String() = Line.Split(";"c)

                    Dim ContractDateText As String = LineSplit(5)
                    ContractDateText = ContractDateText.Split(".")(0)
                    Dim ContractDateDateText As String = ContractDateText.Split(" ")(0)

                    Dim ContractDateDateSplit() As String = ContractDateDateText.Split("-")

                    Dim Year As Integer = CInt(ContractDateDateSplit(0))
                    Dim Month As Integer = CInt(ContractDateDateSplit(1))
                    Dim Day As Integer = CInt(ContractDateDateSplit(2))

                    Dim ContractDate As New Date(Year, Month, Day)
                    Dim Week As Integer = GetWeekNumber(ContractDate)

                    If AggregatedData.ContainsKey(Week) Then
                        AggregatedData(Week) += 1
                    Else
                        AggregatedData.Add(Week, 1)
                    End If

                End While

            End Using

        Catch ex As Exception

            PrintFatalError(ex)

        End Try

        Try

            Using Output As New StreamWriter(OutputFile)

                Output.WriteLine("WeekNumber;NumSwitchers")

                For i As Integer = 1 To 53

                    If AggregatedData.ContainsKey(i) Then
                        Output.WriteLine(i & ";" & AggregatedData(i))
                    Else
                        Output.WriteLine(i & ";0")
                    End If

                Next

            End Using

        Catch ex As Exception

            PrintFatalError(ex)

        End Try

        Try
            Process.Start(OutputFile)
        Catch ex As Exception

        End Try

        Console.WriteLine("All done!!! Press any key to close...")
        Console.ReadKey()

    End Sub

    Sub PrintFatalError(Ex As Exception)
        Console.WriteLine()
        Console.WriteLine("Fatal error:")
        Console.WriteLine(Ex.Message)
        Console.WriteLine()
        Console.WriteLine("Press enter to terminate program...")
        While Not Console.ReadKey.Key = ConsoleKey.Enter
            Console.WriteLine("Lowl wrong key... Better luck next time...")
        End While
        Environment.Exit(1)
    End Sub

    Private Function GetWeekNumber(OurDateTime As DateTime) As Integer
        Return System.Globalization.CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(OurDateTime, Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)
    End Function

End Module
