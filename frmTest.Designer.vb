<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTest
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.lblQuestion = New System.Windows.Forms.Label()
        Me.lblAnswer = New System.Windows.Forms.Label()
        Me.btnClick = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'lblQuestion
        '
        Me.lblQuestion.Font = New System.Drawing.Font("Segoe UI", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblQuestion.Location = New System.Drawing.Point(13, 13)
        Me.lblQuestion.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblQuestion.Name = "lblQuestion"
        Me.lblQuestion.Size = New System.Drawing.Size(852, 182)
        Me.lblQuestion.TabIndex = 0
        Me.lblQuestion.Text = "Label1"
        '
        'lblAnswer
        '
        Me.lblAnswer.Font = New System.Drawing.Font("Segoe UI", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAnswer.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lblAnswer.Location = New System.Drawing.Point(13, 212)
        Me.lblAnswer.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAnswer.Name = "lblAnswer"
        Me.lblAnswer.Size = New System.Drawing.Size(852, 72)
        Me.lblAnswer.TabIndex = 1
        Me.lblAnswer.Text = "Label2"
        '
        'btnClick
        '
        Me.btnClick.Location = New System.Drawing.Point(561, 307)
        Me.btnClick.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnClick.Name = "btnClick"
        Me.btnClick.Size = New System.Drawing.Size(124, 37)
        Me.btnClick.TabIndex = 2
        Me.btnClick.Text = "&Click"
        Me.ToolTip1.SetToolTip(Me.btnClick, "Click to display answers and" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "move to the next question.")
        Me.btnClick.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(725, 307)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(124, 37)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "E&xit"
        Me.ToolTip1.SetToolTip(Me.btnExit, "Click to exit the program.")
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'frmTest
        '
        Me.AcceptButton = Me.btnClick
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(878, 367)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnClick)
        Me.Controls.Add(Me.lblAnswer)
        Me.Controls.Add(Me.lblQuestion)
        Me.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frmTest"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Test"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblQuestion As System.Windows.Forms.Label
    Friend WithEvents lblAnswer As System.Windows.Forms.Label
    Friend WithEvents btnClick As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
