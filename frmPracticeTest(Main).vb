'**************************************************************************
' Name: frmPracticeTest(Main).vb
' Programmer: Curtis N Frank
' Date: 4/2/2016
' Assignment: Advanced VB.NET ITSE 2349 Individual Project
' Purpose: The Practice Test application allows the user to edit and
'          maintain a database of test questions and answers. The UI
'          can create a quick digital quiz for your desktop, or output
'          a test as a text file. 
'**************************************************************************

Imports System.IO
Imports System.IO.File
Imports System.Threading
Imports System.Data
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmPracticeTest

    ' Excel connection strings
    Private Excel03ConString As String =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};" +
            "Extended Properties='Excel 8.0;HDR={1}'"
    Private Excel07ConString As String =
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};" +
            "Extended Properties='Excel 8.0;HDR={1}'"

    ' class level variable
    Private headers As String

    ' instantiate class-level lists of strings
    Public myQuestions As New List(Of String)
    Public myAnswers As New List(Of String)

    ' class-level variables
    Public Count As Integer = 0
    Public testName As String

    Private Sub frmPracticeTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' form startup procedures

        ' instantiate splash screen thread
        Dim t1 As New Thread(New ThreadStart(AddressOf SplashStart))

        ' start thread
        t1.Start()

        ' pause 5.3 seconds
        Thread.Sleep(5300)

        ' terminate thread
        t1.Abort()

        ' start Vivaldi's guitar concerto in D major
        'PlayBackgroundSoundResource()

        'TODO: This line of code loads data into the 'ProblemsDataSet.tblProblems' table. You can move, or remove it, as needed.
        Me.TblProblemsTableAdapter.Fill(Me.ProblemsDataSet.tblProblems)

        ' call helper method to
        ' populate listbox
        RefreshListBox()

        ' default selected value
        lstTestNames.SelectedIndex = 0

    End Sub

    Private Sub SplashStart()
        ' open the splash screen form

        Application.Run(PracticeTestSplash)

    End Sub

    Public Sub RefreshListBox()

        ' refresh the database table adapter
        Me.TblProblemsTableAdapter.Fill(Me.ProblemsDataSet.tblProblems)

        ' boolean flag and list of strings
        Dim originalName As Boolean = True
        Dim testNames As New List(Of String)

        ' LINQ query to get all test names
        Dim names = From name In ProblemsDataSet.tblProblems
                    Order By name.TestName
                    Select name.TestName Distinct

        ' loop
        For Each n In names

            ' copy test name to list
            lstTestNames.Items.Add(n)

        Next n

        '' loop...
        'For Each n In names

        '    ' copy test name to list
        '    testNames.Add(n)

        'Next n

        '' loop to check names in list...
        'For Each t In testNames

        '    ' reset boolean flag...
        '    originalName = True

        '    ' loop to compare names to listbox items...
        '    For Each item In lstTestNames.Items

        '        ' if a match is found...
        '        If t = item Then

        '            ' toggle boolean flag
        '            originalName = False

        '        End If

        '    Next item

        '    ' if no matches were found...
        '    If originalName = True Then

        '        ' add unique name to listbox
        '        lstTestNames.Items.Add(t)

        '    End If

        'Next t

        ' default selected value
        lstTestNames.SelectedIndex = 0

    End Sub

    Private Sub btnRunTest_Click(sender As Object, e As EventArgs) Handles btnRunTest.Click,
        RunTestToolStripMenuItem.Click
        ' query test questions and answers, then store in a List of Problem Objects
        ' call the test form to display questions and answers

        ' clear the lists
        myQuestions.Clear()
        myAnswers.Clear()

        ' reset the counter
        Count = 0

        ' load test name
        testName = lstTestNames.SelectedItem.ToString()

        ' LINQ query to filter the test questions and answers
        Dim records = From record In ProblemsDataSet.tblProblems
                      Where record.TestName = lstTestNames.SelectedItem
                      Select record.Question, record.Answer

        ' loop...
        For Each r In records

            ' copy records to lists
            myQuestions.Add(r.Question)
            myAnswers.Add(r.Answer)

            ' increment counter
            Count += 1

        Next r

        ' open the test form
        frmTest.Show()

    End Sub

    Private Sub OpenHelpMenu(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        ' open the Help form

        frmHelp.Show()

    End Sub

    Private Sub ExportTextFile()
        ' create a text file with title, questions, and answers and open file

        ' list of Problem objects
        Dim myTest = New List(Of Problem)

        ' local string variable to hold selected test name
        Dim test As String = lstTestNames.SelectedItem.ToString()

        ' LINQ query to retrieve test questions
        Dim questions = From q In ProblemsDataSet.tblProblems
                        Where q.TestName = test
                        Select q.Question

        ' LINQ query to retrieve test answers
        Dim answers = From a In ProblemsDataSet.tblProblems
                      Where a.TestName = test
                      Select a.Answer

        ' copy query results into Problem list
        For x As Integer = 0 To (questions.Count() - 1)

            myTest.Add(New Problem(questions(x), answers(x)))

        Next x

        ' Save dialog box
        Dim myStream As Stream
        Dim saveFileDialog1 As New SaveFileDialog()
        ' counter variable
        Dim lineCount As Integer = 1
        Dim outFile As StreamWriter

        saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        saveFileDialog1.FilterIndex = 2
        saveFileDialog1.RestoreDirectory = True

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            myStream = saveFileDialog1.OpenFile()

            If (myStream IsNot Nothing) Then
                ' instantiate file output object


                

                ' open file for output
                outFile = CreateText(test + ".txt")

                ' test title
                outFile.WriteLine(test)
                outFile.WriteLine()

                ' questions header
                outFile.WriteLine("QUESTIONS:")

                ' loop...
                For Each p In myTest

                    ' write number and record to file
                    outFile.WriteLine(lineCount.ToString() + ". " + p.Question)

                    ' increment counter
                    lineCount += 1

                Next p

                ' reset counter
                lineCount = 1

                ' answers header
                outFile.WriteLine()
                outFile.WriteLine("ANSWERS:")

                ' loop...
                For Each p In myTest

                    ' write number and record to file
                    outFile.WriteLine(lineCount.ToString() + ". " + p.Answer)

                    ' increment counter
                    lineCount += 1

                Next p

                ' close file
                outFile.Close()
                myStream.Close()
            End If
        End If




        '' instantiate file output object
        'Dim outFile As StreamWriter

        ' counter variable
        lineCount = 1

        ' open file for output
        outFile = CreateText(test + ".txt")

        ' test title
        outFile.WriteLine(test)
        outFile.WriteLine()

        ' questions header
        outFile.WriteLine("QUESTIONS:")

        ' loop...
        For Each p In myTest

            ' write number and record to file
            outFile.WriteLine(lineCount.ToString() + ". " + p.Question)

            ' increment counter
            lineCount += 1

        Next p

        ' reset counter
        lineCount = 1

        ' answers header
        outFile.WriteLine()
        outFile.WriteLine("ANSWERS:")

        ' loop...
        For Each p In myTest

            ' write number and record to file
            outFile.WriteLine(lineCount.ToString() + ". " + p.Answer)

            ' increment counter
            lineCount += 1

        Next p

        ' close file
        outFile.Close()

        ' open text file in Notepad
        Process.Start("notepad.exe", test + ".txt")

    End Sub

    Private Sub btnRunTest_MouseHover(sender As Object, e As EventArgs) Handles btnRunTest.MouseHover
        ' when mouse pointer over button, change control colors

        btnRunTest.BackColor = Color.Orange
        btnRunTest.ForeColor = Color.White

    End Sub

    Private Sub btnRunTest_MouseLeave(sender As Object, e As EventArgs) Handles btnRunTest.MouseLeave
        ' when mouse pointer leaves button, return control colors to default settings

        btnRunTest.BackColor = Color.LightGray
        btnRunTest.ForeColor = Color.Black

    End Sub

    Private Sub btnExit_MouseHover(sender As Object, e As EventArgs) Handles btnExit.MouseHover
        ' when mouse pointer over button, change control colors

        btnExit.BackColor = Color.Orange
        btnExit.ForeColor = Color.White

    End Sub

    Private Sub btnExit_MouseLeave(sender As Object, e As EventArgs) Handles btnExit.MouseLeave
        ' when mouse pointer leaves button, return control colors to default settings

        btnExit.BackColor = Color.LightGray
        btnExit.ForeColor = Color.Black

    End Sub

    Private Sub ExportTextFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportTextFileToolStripMenuItem.Click
        ' call the ExportTextFile() subprocedure

        ExportTextFile()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ' exits the program

        Me.Close()

    End Sub

    Private Sub MouseLightToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles MouseLightToolStripMenuItem.Click
        ' open the MouseLight C# program

        Process.Start("MouseLight.exe")

    End Sub

    Private Sub PlayBackgroundSoundResource()
        ' play Vivaldi's guitar concerto in D major in a loop

        My.Computer.Audio.Play("Vivaldi.wav",
                               AudioPlayMode.BackgroundLoop)

    End Sub

    Private Sub VivaldiToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' open the music control form

        frmVivaldi.Show()

    End Sub

    Private Sub mnuImport_Click(sender As Object, e As EventArgs) Handles mnuImport.Click

        ' get user input
        'fileName = InputBox("Filename:",
        '                    "Import Text File")

        ' user warning message
        MessageBox.Show("WARNING - Do not try to import a text file" +
                        vbNewLine + "that was generated by this application!" +
                        vbNewLine + "To import a text file you MUST use a specific" +
                        vbNewLine + "format. See HELP menu for details.",
                        "Import Text File",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)

        ' open dialog box
        OpenFileDialog1.ShowDialog()

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

        ' local variables
        Dim fileName, testName, id As String
        Dim line As String = "x"
        Dim question As String = ""
        Dim answer As String = ""
        Dim number As Integer = 1

        ' list collection
        Dim input = New List(Of String)

        ' store filename
        fileName = OpenFileDialog1.FileName

        Try
            ' instantiate StreamReader object
            Dim inFile As StreamReader = New StreamReader(fileName)
            ' loop to traverse file...
            While Not IsNothing(line)

                ' read and store line
                line = inFile.ReadLine()

                ' verify line exists...
                If Not IsNothing(line) Then

                    ' add line to list collection
                    input.Add(line)

                End If

            End While

            ' close file
            inFile.Close()

            ' assign test name and id (first chars of id)
            testName = input(0)
            id = input(1)

            ' loop to traverse list
            For i As Integer = 2 To (input.Count - 1)

                ' if even index...
                If i Mod 2 = 0 Then

                    ' store question
                    question = input(i)

                End If

                ' if odd index...
                If i Mod 2 = 1 Then

                    ' store answer
                    answer = input(i)

                    ' write record to database
                    ProblemsDataSet.tblProblems.AddtblProblemsRow(id + number.ToString(), testName,
                                                                  number, question,
                                                                  answer)

                    ' increment question number
                    number += 1

                End If

            Next i

            ' update table adapter and backend database
            Me.Validate()
            Me.TblProblemsBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(Me.ProblemsDataSet)

            ' user confirmation message
            MessageBox.Show(fileName + " was imported successfully.",
                            "Import Text File",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

            ' add new test to listbox
            lstTestNames.Items.Clear()
            RefreshListBox()

        Catch ex As Exception

            ' error message
            MessageBox.Show("ERROR - Cannot import the file: " + fileName + "." +
                            vbNewLine + "Check the Database Manager and" +
                            vbNewLine + "verify file has not already been imported." +
                            vbNewLine + "Verify the Unique ID has not already been used.",
                            "Import Text File",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub OpenFileDialog2_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        ' import the Excel file

        ' local varibles
        Dim filePath As String = OpenFileDialog2.FileName
        Dim extension As String = Path.GetExtension(filePath)
        Dim conStr As String = ""
        Dim sheetName As String = ""
        Dim drList As New List(Of DataRow)()
        Dim dlgResult As DialogResult

        ' get user input
        dlgResult = MessageBox.Show("Does the Excel file have headers?",
                                    "Import Excel File",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question)

        ' assign header flag for connection string parameter
        headers =
        If(dlgResult = Windows.Forms.DialogResult.Yes, "YES", "NO")

        ' clear connection string
        conStr = ""

        ' determine excel version and appropriate
        ' connection string
        Select Case extension

            Case ".xls"
                'Excel 97-03
                conStr = String.Format(Excel03ConString, filePath, headers)
                Exit Select

            Case ".xlsx"
                'Excel 07
                conStr = String.Format(Excel07ConString, filePath, headers)
                Exit Select

        End Select

        ' get the name of the first sheet
        Using con As New OleDbConnection(conStr)
            Using cmd As New OleDbCommand()

                ' open connection
                cmd.Connection = con
                con.Open()

                ' import Excel data
                Dim dtExcelSchema As DataTable =
                    con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                                            Nothing)

                ' store the sheet name
                sheetName = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()

                ' close connection
                con.Close()

            End Using
        End Using

        ' read data from the first sheet
        Using con As New OleDbConnection(conStr)
            Using cmd As New OleDbCommand()
                Using oda As New OleDbDataAdapter()

                    ' instantiate new DataTable
                    Dim dt As New DataTable()

                    ' SQL query
                    cmd.CommandText =
                        (Convert.ToString("SELECT * FROM [") &
                         sheetName) + "]"

                    ' open connection
                    cmd.Connection = con
                    con.Open()

                    ' fill new data table with data
                    oda.SelectCommand = cmd
                    oda.Fill(dt)

                    ' close connection
                    con.Close()

                    ' loop to traverse data table...
                    For Each row As DataRow In dt.Rows

                        ' store rows in list collection
                        drList.Add(CType(row, DataRow))

                    Next row

                End Using
            End Using
        End Using

        ' error handling
        Try
            ' loop to traverse list...
            For Each item In drList

                ' add list item fields to dataset row
                ProblemsDataSet.tblProblems.AddtblProblemsRow(
                    item(0).ToString(),
                    item(1).ToString(),
                    Convert.ToInt32(item(2)),
                    item(3).ToString(),
                    item(4).ToString())

            Next item

            ' update table adapter and backend database
            Me.Validate()
            Me.TblProblemsBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(Me.ProblemsDataSet)

            ' user confirmation message
            MessageBox.Show(filePath + " was imported successfully.",
                            "Import Excel File",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

            ' add new test to listbox
            lstTestNames.Items.Clear()
            RefreshListBox()

        Catch ex As Exception

            ' error message
            MessageBox.Show("ERROR - Cannot import the file: " + filePath + "." +
                            vbNewLine + "Check the Database Manager and" +
                            vbNewLine + "verify file has not already been imported." +
                            vbNewLine + "Verify the Unique ID has not already been used." +
                            vbNewLine + "Check the fields and data types to match required" +
                            vbNewLine + "schema (string, string, int, string, string)",
                            "Import Text File",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub ExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelToolStripMenuItem.Click

        ' open dialog box
        OpenFileDialog2.ShowDialog()

    End Sub

    Private Sub mnuDbManager_Click(sender As Object, e As EventArgs) Handles mnuDbManager.Click
        ' open the Database Manager form

        frmDatabaseManager.Show()

    End Sub

    Private Sub AddTestMenuItem_Click(sender As Object, e As EventArgs) Handles AddTestMenuItem.Click
        ' open the Add Test form

        frmAddTest.Show()

    End Sub

    Private Sub mnuDeleteTest_Click(sender As Object, e As EventArgs) Handles mnuDeleteTest.Click
        ' open the Delete Test form

        frmDeleteTest.Show()

    End Sub

End Class
