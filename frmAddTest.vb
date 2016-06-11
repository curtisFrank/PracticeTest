'**************************************************************************
' Name: frmAddTest.vb
' Programmer: Curtis N Frank
' Date: 4/4/2016
' Purpose: UI for adding new records to database. Checks for unique
'          primary key field and any fields left blank to handle
'          input errors, in the form and in the database table.
'**************************************************************************

Public Class frmAddTest

    ' instantiate class-level collections of strings
    Private _newQuestions As New List(Of String) From {""}
    Private _newAnswers As New List(Of String) From {""}
    Private _newIDs As New List(Of String) From {""}

    ' class-level counters
    Private _count As Integer = 0
    Private _index = 0

    Private Sub frmAddTest_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' reopen the Main form

        ' open Main form
        frmPracticeTest.Show()

        'TODO: This line of code loads data into the 'ProblemsDataSet.tblProblems' table
        frmPracticeTest.TblProblemsTableAdapter.Fill(Me.ProblemsDataSet.tblProblems)

        ' refresh listbox
        frmPracticeTest.lstTestNames.Items.Clear()
        frmPracticeTest.RefreshListBox()

    End Sub

    Private Sub frmAddTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'ProblemsDataSet.tblProblems' table. You can move, or remove it, as needed.

        Me.TblProblemsTableAdapter.Fill(Me.ProblemsDataSet.tblProblems)

        ' start with question number 1
        lblQuestionNumber.Text = 1.ToString()

        ' hide the main form
        frmPracticeTest.Hide()

        ' send the focus to the textbox
        ' and select all text
        txtTestName.Focus()
        txtTestName.SelectAll()

    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        ' save textbox input and navigate forward through collections

        ' boolean flag
        Dim isUnique As Boolean = True

        ' check for empty fields...
        If (txtAnswer.Text <> "") AndAlso
           (txtQuestion.Text <> "") AndAlso
           (txtTestName.Text <> "") AndAlso
           (txtID.Text <> "") Then

            ' if at the end of the lists...
            If (_index + 1 = _newQuestions.Count()) Then

                ' loop...
                For intI As Integer = 0 To (_newIDs.Count() - 2)

                    ' search for matching IDs
                    If (_newIDs(intI) = txtID.Text) Then

                        ' toggle boolean flag
                        isUnique = False

                    End If

                Next intI

                ' if no matching IDs...
                If isUnique = True Then

                    ' add input to collections
                    '_newIDs(_index) = txtID.Text()
                    '_newQuestions(_index) = txtQuestion.Text
                    '_newAnswers(_index) = txtAnswer.Text

                    ' add empty record
                    _newQuestions.Add("")
                    _newAnswers.Add("")
                    _newIDs.Add("")

                Else

                    ' input error user message
                    ShowMessage("DATA ERROR: ID field must be unique")

                    ' send focus to textbox
                    txtID.Focus()
                    txtID.SelectAll()

                End If

            End If

            ' if the ID field is unique...
            If isUnique = True Then

                ' add input to collections
                _newIDs(_index) = txtID.Text()
                _newQuestions(_index) = txtQuestion.Text
                _newAnswers(_index) = txtAnswer.Text

                ' increment counter
                _index += 1

                ' update textboxes
                txtQuestion.Text = _newQuestions(_index).ToString()
                txtAnswer.Text = _newAnswers(_index).ToString()
                txtID.Text = _newIDs(_index).ToString()

                ' update label control
                lblQuestionNumber.Text = ((_index + 1).ToString())

            End If

        Else

            ' input error message
            ShowMessage("Cannot leave any fields blank")

        End If

        ' send focus to question textbox
        txtID.Focus()

    End Sub

    Private Sub ShowMessage(ByVal input As String)
        ' accepts a string argument, then displays it in a message box

        MessageBox.Show(input,
                        "Practice Test",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        ' exits the program without saving

        Me.Close()

    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        ' will automatically save last record if at end of lists, then navigates
        ' to previous record in collection

        ' check to see if at beginning of lists...
        If (_index = 0) Then

            ' display user message
            ShowMessage("Cannot navigate to previous record.")

        Else

            ' check to see if at end of collection...
            If (_index + 1) = _newQuestions.Count() Then

                ' save the last record that was input
                _newAnswers(_index) = txtAnswer.Text
                _newQuestions(_index) = txtQuestion.Text
                _newIDs(_index) = txtID.Text

            End If

            ' decrement counter
            _index -= 1

            ' update textboxes
            txtAnswer.Text = _newAnswers(_index).ToString()
            txtQuestion.Text = _newQuestions(_index).ToString()
            txtID.Text = _newIDs(_index).ToString()

            ' update label control
            lblQuestionNumber.Text = (_index + 1).ToString()

        End If

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' save records to database and exit program

        ' boolean flags
        Dim isUnique2Form As Boolean = True
        Dim isUnique2Table As Boolean = True

        ' check for empty fields...
        If (txtAnswer.Text <> "") AndAlso
           (txtQuestion.Text <> "") AndAlso
           (txtTestName.Text <> "") AndAlso
           (txtID.Text <> "") Then

            ' add input to collections
            _newIDs(_index) = txtID.Text()
            _newQuestions(_index) = txtQuestion.Text
            _newAnswers(_index) = txtAnswer.Text

            ' if at the end of the lists...
            'If (_index + 1 = _newQuestions.Count()) Then

            Dim tempID As String = _newIDs(0)

            ' loop...
            For intIndex = 1 To (_newIDs.Count() - 2)

                ' search for matching IDs
                If (_newIDs(intIndex) = tempID) Then

                    ' toggle boolean flag
                    isUnique2Form = False

                End If

                tempID = _newIDs(intIndex - 1)

            Next intIndex

            ' LINQ query to collect all records from database
            Dim records = From record In ProblemsDataSet.tblProblems
                          Select record

            ' loop...
            For Each r In records

                ' loop...
                For Each id In _newIDs

                    ' test each ID in the database
                    ' against each new ID to find
                    ' any matches
                    If r.ID = id Then

                        ' toggle boolean flag
                        isUnique2Table = False

                    End If

                Next id

            Next r

            ' if no matching IDs...
            If isUnique2Form = True AndAlso
                isUnique2Table = True Then

                ' add input to collections
                _newIDs(_index) = txtID.Text()
                _newQuestions(_index) = txtQuestion.Text
                _newAnswers(_index) = txtAnswer.Text

                ' add empty record
                _newQuestions.Add("")
                _newAnswers.Add("")
                _newIDs.Add("")

                Try
                    ' loop...
                    For intI As Integer = 0 To _newIDs.Count() - 2

                        ' write record to database
                        ProblemsDataSet.tblProblems.AddtblProblemsRow(_newIDs(intI), txtTestName.Text,
                                                                      (intI + 1), _newQuestions(intI),
                                                                      _newAnswers(intI))
                    Next intI

                    ' update table adapter and backend database
                    Me.Validate()
                    Me.TblProblemsBindingSource.EndEdit()
                    Me.TableAdapterManager.UpdateAll(Me.ProblemsDataSet)

                    ' confirmation message
                    ShowMessage("New test " + txtTestName.Text + " saved.")

                    ' exit form
                    Me.Close()

                Catch ex As Exception

                    ' input error user message
                    ShowMessage("ID field must be unique." +
                                vbNewLine + "Check the Database Manager for a duplicate ID.")

                    ' send the focus to textbox
                    txtID.Focus()
                    txtID.SelectAll()

                End Try

            Else
                ' if the error local to the form...
                If isUnique2Form = False Then

                    ' input error user message
                    ShowMessage("ID field must be unique." +
                                vbNewLine + "Check the IDs in the Add Test form.")

                    ' send the focus to textbox
                    txtID.Focus()
                    txtID.SelectAll()

                    ' if the error was in the table...
                ElseIf isUnique2Table = False Then

                    ' input error user message
                    ShowMessage("ID field must be unique." +
                                vbNewLine + "Check the Database Manager for a duplicate ID.")

                    ' send the focus to textbox
                    txtID.Focus()
                    txtID.SelectAll()

                End If

            End If

            'End If

        Else
            ' input error message
            ShowMessage("Cannot leave any fields empty")

        End If

    End Sub

End Class