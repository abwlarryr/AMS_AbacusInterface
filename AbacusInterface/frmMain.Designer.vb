<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.components = New System.ComponentModel.Container
        Me.tbcMain = New System.Windows.Forms.TabControl
        Me.tabMain = New System.Windows.Forms.TabPage
        Me.btnStop = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.txtText = New System.Windows.Forms.TextBox
        Me.tmrMain = New System.Windows.Forms.Timer(Me.components)
        Me.dstXMLData = New System.Data.DataSet
        Me.txtEnvironment = New System.Windows.Forms.TextBox
        Me.lblEnvironment = New System.Windows.Forms.Label
        Me.tbcMain.SuspendLayout()
        Me.tabMain.SuspendLayout()
        CType(Me.dstXMLData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'tbcMain
        '
        Me.tbcMain.Controls.Add(Me.tabMain)
        Me.tbcMain.Location = New System.Drawing.Point(12, 5)
        Me.tbcMain.Name = "tbcMain"
        Me.tbcMain.SelectedIndex = 0
        Me.tbcMain.Size = New System.Drawing.Size(721, 477)
        Me.tbcMain.TabIndex = 0
        '
        'tabMain
        '
        Me.tabMain.Controls.Add(Me.lblEnvironment)
        Me.tabMain.Controls.Add(Me.txtEnvironment)
        Me.tabMain.Controls.Add(Me.btnStop)
        Me.tabMain.Controls.Add(Me.btnStart)
        Me.tabMain.Controls.Add(Me.txtText)
        Me.tabMain.Location = New System.Drawing.Point(4, 22)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.Padding = New System.Windows.Forms.Padding(3)
        Me.tabMain.Size = New System.Drawing.Size(713, 451)
        Me.tabMain.TabIndex = 0
        Me.tabMain.Text = "Main"
        Me.tabMain.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(615, 423)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(75, 23)
        Me.btnStop.TabIndex = 2
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(199, 422)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 2
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'txtText
        '
        Me.txtText.Location = New System.Drawing.Point(3, 3)
        Me.txtText.Multiline = True
        Me.txtText.Name = "txtText"
        Me.txtText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtText.Size = New System.Drawing.Size(704, 410)
        Me.txtText.TabIndex = 0
        '
        'tmrMain
        '
        '
        'dstXMLData
        '
        Me.dstXMLData.DataSetName = "Dataset1"
        '
        'txtEnvironment
        '
        Me.txtEnvironment.Location = New System.Drawing.Point(95, 422)
        Me.txtEnvironment.Name = "txtEnvironment"
        Me.txtEnvironment.Size = New System.Drawing.Size(25, 20)
        Me.txtEnvironment.TabIndex = 1
        '
        'lblEnvironment
        '
        Me.lblEnvironment.AutoSize = True
        Me.lblEnvironment.Location = New System.Drawing.Point(7, 425)
        Me.lblEnvironment.Name = "lblEnvironment"
        Me.lblEnvironment.Size = New System.Drawing.Size(82, 13)
        Me.lblEnvironment.TabIndex = 4
        Me.lblEnvironment.Text = "Prod, Test, Dev"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(745, 507)
        Me.Controls.Add(Me.tbcMain)
        Me.Name = "frmMain"
        Me.Text = "AbacusInterface"
        Me.tbcMain.ResumeLayout(False)
        Me.tabMain.ResumeLayout(False)
        Me.tabMain.PerformLayout()
        CType(Me.dstXMLData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tbcMain As System.Windows.Forms.TabControl
    Friend WithEvents tabMain As System.Windows.Forms.TabPage
    Friend WithEvents txtText As System.Windows.Forms.TextBox
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents tmrMain As System.Windows.Forms.Timer
    Friend WithEvents dstXMLData As System.Data.DataSet
    Friend WithEvents lblEnvironment As System.Windows.Forms.Label
    Friend WithEvents txtEnvironment As System.Windows.Forms.TextBox

End Class
