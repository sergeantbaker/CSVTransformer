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
            .FileName = "aggregated.csv",
            .Filter = "Comma Seperated Values|*.csv",
            .FilterIndex = 0,
            .InitialDirectory = Path.GetDirectoryName(InputFile),
            .Title = "Please select a save file location..."}
        If SaveSelector.ShowDialog = DialogResult.OK Then
            OutputFile = SaveSelector.FileName
        Else 'user abort
            Exit Sub
        End If

        Dim AggregatedData As New Dictionary(Of Integer, Integer()) '{WeekNumber,{NumSwitchers,NumPositiveExpressions}}

        Try

            Using Input As New StreamReader(InputFile)

                If SkipFirstLine Then Input.ReadLine()

                Dim TempWeekNumber As Integer = Nothing
                Dim TempDayNumber As Integer = 0
                Dim TempNumSwitchers As Integer = 0
                Dim TempNumPositiveExpressions As Integer = 0

                While Not Input.EndOfStream

                    Dim Line As String = Input.ReadLine
                    Dim LineSplit As String() = Line.Split(";"c)

                    Dim WeekNumber As Integer = CInt(LineSplit(1))
                    Dim DayNumber As Integer = CInt(LineSplit(2))
                    Dim NumSwitchers As Integer = CInt(LineSplit(4))
                    Dim NumPositiveExpressions As Integer = CInt(LineSplit(5))

                    If TempWeekNumber = Nothing Then TempWeekNumber = WeekNumber

                    If Not WeekNumber = TempWeekNumber Or Input.EndOfStream Then
                        Console.WriteLine("Adding " & TempWeekNumber & " " & TempNumSwitchers & " " & TempNumPositiveExpressions)
                        AggregatedData.Add(TempWeekNumber, {TempNumSwitchers, TempNumPositiveExpressions})
                        TempNumSwitchers = 0
                        TempNumPositiveExpressions = 0
                        TempDayNumber = 0
                    End If

                    TempWeekNumber = WeekNumber
                    TempNumSwitchers += NumSwitchers

                    If Not TempDayNumber = DayNumber Then 'Different day. Reset num expressions
                        TempNumPositiveExpressions += NumPositiveExpressions
                        TempDayNumber = DayNumber
                    End If

                End While

            End Using

        Catch ex As Exception

            PrintFatalError(ex)

        End Try

        Try

            Using Output As New StreamWriter(OutputFile)

                Output.WriteLine("WeekNumber;NumSwitchers;NumPositiveExpressions")

                For Each AggData In AggregatedData

                    Output.WriteLine(AggData.Key & ";" & AggData.Value(0) & ";" & AggData.Value(1))

                Next

            End Using

        Catch ex As Exception

            PrintFatalError(ex)

        End Try

        Try
            Process.Start(OutputFile)
        Catch ex As Exception

        End Try

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

End Module
