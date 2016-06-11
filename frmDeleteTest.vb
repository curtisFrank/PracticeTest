'**************************************************************************
' Name: frmDeleteTest.vb
' Programmer: Curtis N Frank
' Date: 4/4/2016
' Purpose: The Delete Test UI will delete all records associated with
'          the selected test name.
'**************************************************************************

Public Class frmDeleteTest

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click

        ' local variableS
        Dim dlgButton As DialogResult
        Dim myTestName As String = lstTestNames.SelectedItem

        ' play an alert sound (the question icon doesn't play a sound)
        My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Asterisk)

        ' confirm user wants to delete records
        dlgButton = MessageBox.Show("Are you sure you want to delete test " +
                                    myTestName.ToString() + vbNewLine +
                                    "and all related records?",
                                    "Practice Test",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question)

        ' if user chose yes...
        If dlgButton = Windows.Forms.DialogResult.Yes Then

            ' LINQ query to find all records with selected test name
            Dim records = From record In ProblemsDataSet.tblProblems
                          Where record.TestName = myTestName
                          Select record

            ' loop while records still exist...
            Do While records.Count() > 0

                ' delete first record
                records(0).Delete()

                ' update table adapter and backend database
                Me.Validate()
                Me.TblProblemsBindingSource.EndEdit()
                Me.TableAdapterManager.UpdateAll(Me.ProblemsDataSet)

            Loop

            ' confirmation message
            MessageBox.Show(myTestName + " was successfully deleted from the database.",
                            "Practice Test",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

            ' exit Delete Test UI
            Me.Close()

        End If

    End Sub

    Private Sub frmDeleteTest_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' reopen the Main form

        ' open Main form
        frmPracticeTest.Show()

        'TODO: This line of code loads data into the 'ProblemsDataSet.tblProblems' table
        frmPracticeTest.TblProblemsTableAdapter.Fill(Me.ProblemsDataSet.tblProblems)

        ' refresh listbox
        frmPracticeTest.lstTestNames.Items.Clear()
        frmPracticeTest.RefreshListBox()

    End Sub

    Private Sub frmDeleteTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' populate the listbox with the test names in the database

        ' hide the Main form
        frmPracticeTest.Hide()

        'TODO: This line of code loads data into the 'ProblemsDataSet.tblProblems' table. You can move, or remove it, as needed.
        Me.TblProblemsTableAdapter.Fill(Me.ProblemsDataSet.tblProblems)

        ' boolean flag and list of strings
        Dim originalName As Boolean = True
        Dim testNames = New List(Of String)

        ' LINQ query to get all test names
        Dim names = From name In ProblemsDataSet.tblProblems
                    Order By name.TestName
                    Select name.TestName

        ' loop...
        For Each n In names

            ' copy test name to list
            testNames.Add(n)

        Next n

        ' loop to check names in list...
        For Each t In testNames

            ' reset boolean flag...
            originalName = True

            ' loop to compare names to listbox items...
            For Each item In lstTestNames.Items

                ' if a match is found...
                If t = item Then

                    ' toggle boolean flag
                    originalName = False

                End If

            Next item

            ' if no matches were found...
            If originalName = True Then

                ' add unique name to listbox
                lstTestNames.Items.Add(t)

            End If

        Next t

        ' default selected value
        lstTestNames.SelectedIndex = 0

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        ' exits the Delete Test UI

        Me.Close()

    End Sub

End Class