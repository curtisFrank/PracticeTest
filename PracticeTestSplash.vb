'**************************************************************************
' Name: PracticeTestSplash.vb
' Programmer: Curtis N Frank
' Date: 4/2/2016
' Assignment: Advanced VB.NET ITSE 2349 Individual Project
' Purpose: The splash screen for the Practice Test application.
'**************************************************************************

Public Class PracticeTestSplash

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' manage splash screen progress bar

        ' increment progress bar
        pbrProgress.Increment(1)

        ' when progress bar finished...
        If pbrProgress.Value = 100 Then

            ' stop timer
            Timer1.Stop()

        End If

    End Sub

End Class