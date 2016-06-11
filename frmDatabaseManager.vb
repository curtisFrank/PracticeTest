'**************************************************************************
' Name: frmDatabaseManager.vb
' Programmer: Curtis N Frank
' Date: 4/2/2016
' Assignment: Advanced VB.NET ITSE 2349 Individual Project
' Purpose: The Database Manager allows the user to add, edit, or delete
'          records of test questions and answers. The Data Grid
'          interface is bound to an Access database, and uses LINQ to
'          query records for easy database management.
'**************************************************************************

Public Class frmDatabaseManager

    ' class level boolean variable
    Private _failed As Boolean = False

    ' class level boolean variable
    Public TriedToDelete As Boolean = False

    Private Sub TblProblemsBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs) Handles TblProblemsBindingNavigatorSaveItem.Click
        ' save datatable

        Try
            ' reset private boolean flag
            _failed = False

            ' save and update with error handling
            Try
                Me.Validate()
                Me.TblProblemsBindingSource.EndEdit()
                Me.TableAdapterManager.UpdateAll(Me.ProblemsDataSet)

                'frmPracticeTest.Validate()
                'frmPracticeTest.TblProblemsBindingSource.EndEdit()
                'frmPracticeTest.TableAdapterManager.UpdateAll(frmPracticeTest.ProblemsDataSet)

                ' if no errors, display confirmation message
                If _failed = False Then
                    ShowMessage("Changes saved.")
                End If

            Catch ex As Exception
                ' display an error message

                If TriedToDelete = False Then

                    ShowMessage("DATA ERROR: Please follow the database schema carefully." +
                                vbNewLine + "For guidelines, go to HELP in the Menu Toolbar.")

                End If

            End Try

        Catch ex As NoNullAllowedException

            ShowMessage("Null values are not allowed.")

        End Try

    End Sub

    Private Sub frmDatabaseManager_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' confirm user wants to exit the program

        If TriedToDelete = False Then
            ' local variable
            Dim dlgButton As DialogResult

            ' get user input via message box
            dlgButton = MessageBox.Show("Are you sure you want to exit Database Manager?" + vbNewLine +
                                        "Any unsaved changes could be lost.",
                                        "Pracice Test",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Exclamation)

            ' if user selected no...
            If dlgButton = Windows.Forms.DialogResult.No Then

                ' cancel form closing event
                e.Cancel = True
            Else
                ' open Main form
                frmPracticeTest.Show()

                'TODO: This line of code loads data into the 'ProblemsDataSet.tblProblems' table
                frmPracticeTest.TblProblemsTableAdapter.Fill(Me.ProblemsDataSet.tblProblems)

                ' refresh listbox
                frmPracticeTest.lstTestNames.Items.Clear()
                frmPracticeTest.RefreshListBox()

            End If

        Else
            ' open Main form
            frmPracticeTest.Show()

            'TODO: This line of code loads data into the 'ProblemsDataSet.tblProblems' table
            frmPracticeTest.TblProblemsTableAdapter.Fill(Me.ProblemsDataSet.tblProblems)

            ' refresh listbox
            frmPracticeTest.lstTestNames.Items.Clear()
            frmPracticeTest.RefreshListBox()

        End If

    End Sub

    Private Sub frmDatabaseManager_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' form start up procedures

            ' hide the main form
            frmPracticeTest.Hide()

            ' populate the Data Grid control with database records
            TblProblemsTableAdapter.Fill(ProblemsDataSet.tblProblems)

            ' LINQ query to select all records
            Dim allRecords = From record In ProblemsDataSet.tblProblems
                             Order By record.TestName, record.Number
                             Select record

            ' display all records
            TblProblemsBindingSource.DataSource = allRecords.AsDataView()

    End Sub

    Public Sub TblProblemsDataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles TblProblemsDataGridView.DataError
        ' handle the Data Grid control's Data Error event

        ' toggle private boolean flag
        _failed = True

        ' display an error message box
        ShowMessage("DATA ERROR: Please follow the database schema carefully." +
                    vbNewLine + "ID must be unique. Number must be a number" +
                    vbNewLine + "Row: " + (e.RowIndex + 1).ToString() +
                        vbNewLine + "For guidelines, go to HELP in the Menu Toolbar," +
                        vbNewLine + "or press Alt + H.")

    End Sub

    Private Sub ShowMessage(ByVal input As String)
        ' accepts a string argument, then displays it in a message box

        MessageBox.Show(input,
                        "Practice Test",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

    End Sub

    Private Sub ToolStripHelpButton_Click(sender As Object, e As EventArgs) Handles ToolStripHelpButton.Click
        ' open the Help page form

        frmHelp.Show()

    End Sub

    Private Sub ToolStripButtonGo_Click(sender As Object, e As EventArgs) Handles ToolStripButtonGo.Click
        ' run test name query, then update data grid

        ' store textbox input in string variable
        Dim testName As String = txtTestName.Text

        ' LINQ query to select all records
        Dim allRecords = From record In ProblemsDataSet.tblProblems
                         Order By record.TestName, record.Number
                         Select record

        ' LINQ query to select records with matching test name
        Dim query = From record In ProblemsDataSet.tblProblems
                    Where record.TestName.ToUpper() Like
                    txtTestName.Text.ToUpper & "*"
                    Order By record.TestName, record.Number
                    Select record

        ' if textbox was empty...
        If testName = "" Then

            ' display all records
            TblProblemsBindingSource.DataSource = allRecords.AsDataView()

            ' textbox had a test name...
        Else

            ' display test name query
            TblProblemsBindingSource.DataSource = query.AsDataView()

        End If

    End Sub

End Class