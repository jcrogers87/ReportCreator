Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Module Module1
    'globals
    Dim inputJobSheet() As String, outputFolder As String, delete As Boolean, autoPrint As Boolean, templateFile As String, sortColumn As Integer
    Sub Main()
        ReadConfig()
        createReport()
        writeLog("Finished Cycle")
    End Sub
    Public Function getInput(rawInput As String) As DataTable
        'grab the entire input csv and put it into a dataTable called "data"
        Dim SR As StreamReader = New StreamReader(rawInput)
        Dim line As String = SR.ReadLine()
        Dim strArray As String() = line.Split(",")
        Dim data As System.Data.DataTable = New System.Data.DataTable()
        Dim row As DataRow
        For Each s As String In strArray
            data.Columns.Add(New DataColumn())
        Next
        Do
            If Not line = String.Empty Then
                row = data.NewRow()
                row.ItemArray = line.Split(",")
                data.Rows.Add(row)
            Else
                Exit Do
            End If
            line = SR.ReadLine
        Loop
        SR.Close()

        Dim view As New DataView(data)
        view.Sort = data.Columns(sortColumn - 1).ToString
        data = view.ToTable()

        If delete = True Then
            File.Delete(rawInput)
        End If

        Return data

    End Function
    Public Sub createReport()
        Dim _excel As New Excel.Application
        Dim wBook As Excel.Workbook
        Dim wSheet As Excel.Worksheet

        Try
            wBook = _excel.Workbooks.Open(templateFile)
            If wBook.Worksheets.Count < 2 Then
                writeLog("Template in invalid format")
                wBook.Close(False)
                ReleaseObject(wBook)
                _excel.Quit()
                ReleaseObject(_excel)
                GC.Collect()
                Threading.Thread.Sleep(2000)
                Exit Sub
            End If
            wSheet = wBook.Worksheets(2)

            For Each rawFile As String In inputJobSheet
                If Path.GetExtension(rawFile) = ".csv" Then
                    Dim outputFile As String = outputFolder & Path.GetFileNameWithoutExtension(rawFile) & ".xlsx"
                    If File.Exists(outputFile) Then
                        writeLog("Found existing file called: " & outputFile & " deleting")
                        File.Delete(outputFile)
                    End If
                    ' get the input file data and fill the excel sheet "PasteValuesHere"
                    Dim dt As System.Data.DataTable = getInput(rawFile)
                    Dim dc As System.Data.DataColumn
                    Dim dr As System.Data.DataRow
                    Dim colIndex As Integer = 0
                    Dim rowIndex As Integer = 0
                    For Each dr In dt.Rows
                        colIndex = 0
                        For Each dc In dt.Columns
                            colIndex = colIndex + 1
                            wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                        Next
                        rowIndex = rowIndex + 1
                    Next
                    Dim endRange As Excel.Range = wBook.Worksheets(1).range("A1").end(Excel.XlDirection.xlToRight).offset(rowIndex, 0)
                    wBook.Worksheets(1).PageSetup.PrintArea = wBook.Worksheets(1).Range("A1", endRange).Address
                    wBook.SaveAs(outputFile)
                    writeLog("Saved: " & outputFile)
                    If autoPrint = True Then
                        wBook.Worksheets(1).printout()
                        writeLog("Printed: " & outputFile)
                    End If
                End If
                wSheet.Range("A1:ZZ10000").ClearContents()
            Next

            ReleaseObject(wSheet)
            wBook.Close(False)
            ReleaseObject(wBook)
            _excel.Quit()
            ReleaseObject(_excel)
            GC.Collect()

        Catch ex As Exception
            writeLog(ex.ToString)
            MsgBox("An error has occured with the report generator. Contact Plataine", vbOKOnly, "Error")
        End Try
    End Sub
    Public Sub writeLog(text As String)
        If (Not Directory.Exists("C:\ProgramData\Plataine\")) Then
            Directory.CreateDirectory("C:\ProgramData\Plataine\")
        End If
        Dim sw As New StreamWriter("C:\ProgramData\Plataine\ReportCreatorLog.txt", True)
        sw.WriteLine(DateTime.Now.ToString("MM:dd-HH:mm:ss") & " : " & text)
        sw.Close()
    End Sub
    Private Sub ReleaseObject(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub
    Public Sub ReadConfig()
        If Not File.Exists("C:\ProgramData\Plataine\ReportCreator.config") Then
            BuildConfig()
            MsgBox("A configuration file was created: C:\ProgramData\Plataine\ReportCreator.config", vbOKOnly, "Initial Setup")
            End
        End If
        Try
            delete = False
            autoPrint = False
            sortColumn = 52
            Dim configFile() As String = File.ReadAllLines("C:\ProgramData\Plataine\ReportCreator.config")
            For Each line As String In configFile
                Dim setting() As String = Split(line, "=")
                If UCase(setting(0)) = "INPUTFOLDER" Then
                    inputJobSheet = Directory.GetFiles(setting(1))
                ElseIf UCase(setting(0)) = "OUTPUTFOLDER" Then
                    outputFolder = setting(1).ToString
                    If Not Right(outputFolder, 1) = "\" Then outputFolder = outputFolder & "\"
                ElseIf UCase(setting(0)) = "DELETE" Then
                    If UCase(setting(1).ToString) = "TRUE" Then delete = True Else delete = False
                ElseIf UCase(setting(0)) = "PRINT" Then
                    If UCase(setting(1).ToString) = "TRUE" Then autoPrint = True Else autoPrint = False
                ElseIf UCase(setting(0)) = "TEMPLATE" Then
                    templateFile = setting(1).ToString
                    If Not File.Exists(templateFile) Or Path.GetExtension(templateFile) <> ".xlsx" Then
                        templateFile = Nothing
                    End If
                ElseIf UCase(setting(0)) = "SORTCOLUMN" Then
                    sortColumn = CInt(UCase(setting(1).ToString))
                End If
            Next
            If IsNothing(inputJobSheet) Or IsNothing(outputFolder) Or IsNothing(templateFile) Then
                Call MsgBox("Your config file is invalid. Must be of the form:" _
                           & Chr(13) & "inputfile=pathtojobs\" _
                           & Chr(13) & "outputFolder=path" _
                           & Chr(13) & "template=<valid xlsx file>" _
                           & Chr(13) & "Config location must be: C:\ProgramData\Plataine\ReportCreator.config")
                End
            End If
        Catch ex As Exception
            Call MsgBox("Your config file is missing, or missing required column mappings." _
                               & Chr(13) & "Config location must be: C:\ProgramData\Plataine\ReportCreator.config")
            End
        End Try
        CleanLog()
        writeLog("Started Program")
    End Sub
    Public Sub CleanLog()
        Dim logFile As String = "C:\ProgramData\Plataine\ReportCreatorLog.txt"
        If File.Exists(logFile) Then
            Dim fileInfo As FileInfo = My.Computer.FileSystem.GetFileInfo(logFile)
            If fileInfo.Length > 20000000 Then
                writeLog("Log file has reached maximum length: " & fileInfo.Length & " bytes.")
                Dim newName As String = Path.GetDirectoryName(logFile) & "\" & Path.GetFileNameWithoutExtension(logFile) & "_old.txt"
                If File.Exists(newName) Then My.Computer.FileSystem.DeleteFile(newName)
                My.Computer.FileSystem.RenameFile(logFile, Path.GetFileName(newName))
                writeLog("Log file maximum length reached, new file created.")
            End If
        End If
    End Sub
    Public Sub BuildConfig()
        If (Not Directory.Exists("C:\ProgramData\Plataine\")) Then
            Directory.CreateDirectory("C:\ProgramData\Plataine\")
        End If
        Dim sw As New StreamWriter("C:\ProgramData\Plataine\ReportCreator.config", False)
        sw.WriteLine("##Auto Config##")
        sw.WriteLine("InputFolder=C:\InputFiles")
        sw.WriteLine("OutputFolder=C:\OutputFiles")
        sw.WriteLine("template=C:\programdata\plataine\reports\template.xlsx")
        sw.WriteLine("Print=false")
        sw.WriteLine("delete=false")
        sw.WriteLine("SortColumn=1")
        sw.Close()
        writeLog("Built config")
    End Sub
End Module
