Sub ExportWordCountToCSV()
    Application.ScreenUpdating = False
    Dim RngHd As Range, h As Integer, csvData As String
    Dim rngStart As Range, rngEnd As Range, totalWordCount As Long
    Dim doc As Document, fso As Object, outputFile As Object
    Dim desktopPath As String, outputFileName As String

    Set doc = ActiveDocument
    Set fso = CreateObject("Scripting.FileSystemObject")
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    outputFileName = desktopPath & "\WordCountReport.csv"

    ' Loop for Heading 1, Heading 2, and Heading 3
    For h = 1 To 3
        With doc.Range
            With .Find
                .ClearFormatting
                .Text = ""
                .Style = "Heading " & h
                .Format = True
                .Execute
            End With
            Do While .Find.Found
                Set RngHd = .Paragraphs(1).Range
                Set RngHd = RngHd.GoTo(What:=wdGoToBookmark, Name:="\HeadingLevel")
                With RngHd
                    csvData = csvData & "Heading " & h & "," & .ComputeStatistics(wdStatisticWords) - .Paragraphs.First.Range.ComputeStatistics(wdStatisticWords) & ",""" & Trim(.Paragraphs.First.Range.Text) & """" & vbCrLf
                End With
                .Start = RngHd.End
                .Find.Execute
            Loop
        End With
    Next h

    ' Word count from ABSTRACT to REFERENCES
    Set rngStart = doc.Content
    With rngStart.Find
        .ClearFormatting
        .Text = "ABSTRACT"
        .Style = "Heading 1"
        .Execute
    End With

    Set rngEnd = doc.Content
    With rngEnd.Find
        .ClearFormatting
        .Text = "REFERENCES"
        .Style = "Heading 1"
        .Execute
    End With

    If rngStart.Find.Found And rngEnd.Find.Found Then
        totalWordCount = doc.Range(Start:=rngStart.Start, End:=rngEnd.Start).ComputeStatistics(Statistic:=wdStatisticWords)
        csvData = csvData & "Total Word Count from ABSTRACT to REFERENCES," & totalWordCount & vbCrLf
    End If

    ' Create and write to the CSV file
    Set outputFile = fso.CreateTextFile(outputFileName, True)
    outputFile.WriteLine csvData
    outputFile.Close

    MsgBox "Word count report exported to: " & outputFileName
    Application.ScreenUpdating = True
End Sub
