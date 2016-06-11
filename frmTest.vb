'**************************************************************************
' Name: frmTest.vb
' Programmer: Curtis N Frank
' Date: 4/2/2016
' Assignment: Advanced VB.NET ITSE 2349 Individual Project
' Purpose: The test form is a desktop tool for quick memorization and
'          enhanced studying and test-taking performance. The UI uses
'          a LINQ query to access the database and display questions
'          and answers for the test name that was selected in the
'          main form's list box.
'**************************************************************************

Public Class frmTest

    ' class level counters
    Private _index As Integer = 0
    Private _clickCounter = 0

    Private Sub btnClick_Click(sender As Object, e As EventArgs) Handles btnClick.Click
        ' display questions and answers sequentially in label controls
        ' when list is finished, call the class method testFinished()

        ' if reached end of list...
        If _index = (frmPracticeTest.Count) Then

            ' call class method
            testFinished()

        Else
            ' if the click counter is set to an odd number...
            If _clickCounter Mod 2 = 1 Then

                ' display the appropriate question with no answer
                lblQuestion.Text = frmPracticeTest.myQuestions(_index).ToString()
                lblAnswer.Text = ""

                ' if the click counter is set to an even number...
            ElseIf _clickCounter Mod 2 = 0 Then

                ' display the appropriate answer and question
                lblQuestion.Text = frmPracticeTest.myQuestions(_index).ToString()
                lblAnswer.Text = frmPracticeTest.myAnswers(_index).ToString()

                ' increment the index counter
                _index += 1

            End If

            ' increment the click counter
            _clickCounter += 1

        End If

    End Sub

    Private Sub frmTest_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' show the main form

        frmPracticeTest.Show()

    End Sub

    Private Sub frmTest_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' menu bar title
        Me.Text = frmPracticeTest.testName

        ' close the Database Manager (NO CHEATING!! HAHA!)
        frmDatabaseManager.Close()

        ' hide the main form
        frmPracticeTest.Hide()

        ' load the first question to the form
        lblQuestion.Text = frmPracticeTest.myQuestions(0).ToString()
        lblAnswer.Text = ""

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ' exits the program

        Me.Close()

    End Sub

    Private Sub testFinished()
        ' display a user message box

        MessageBox.Show("Test finished",
                        "Practice Test",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information)

        ' close the form
        Me.Close()

    End Sub

End Class