<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SpeechControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SpeechControl))
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.speedTrackBar = New System.Windows.Forms.TrackBar
        Me.volumeTrackBar = New System.Windows.Forms.TrackBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.useHighlight = New System.Windows.Forms.CheckBox
        Me.errorLabel = New System.Windows.Forms.Label
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.playimg = New System.Windows.Forms.PictureBox
        Me.pauseimg = New System.Windows.Forms.PictureBox
        Me.stopimg = New System.Windows.Forms.PictureBox
        Me.continuousBox = New System.Windows.Forms.CheckBox
        CType(Me.speedTrackBar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.volumeTrackBar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.playimg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pauseimg, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.stopimg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.AccessibleName = "Speech Voice"
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(21, 32)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'speedTrackBar
        '
        Me.speedTrackBar.AccessibleName = "Reading Speed"
        Me.speedTrackBar.Location = New System.Drawing.Point(21, 238)
        Me.speedTrackBar.Minimum = -10
        Me.speedTrackBar.Name = "speedTrackBar"
        Me.speedTrackBar.Size = New System.Drawing.Size(121, 45)
        Me.speedTrackBar.TabIndex = 3
        Me.speedTrackBar.Value = -1
        '
        'volumeTrackBar
        '
        Me.volumeTrackBar.AccessibleName = "Volume"
        Me.volumeTrackBar.Location = New System.Drawing.Point(21, 302)
        Me.volumeTrackBar.Maximum = 100
        Me.volumeTrackBar.Name = "volumeTrackBar"
        Me.volumeTrackBar.Size = New System.Drawing.Size(121, 45)
        Me.volumeTrackBar.TabIndex = 4
        Me.volumeTrackBar.Value = 50
        '
        'Label1
        '
        Me.Label1.AccessibleName = "Select Voice"
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Speech Voice:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Location = New System.Drawing.Point(18, 222)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Reading Speed:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(18, 286)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Volume:"
        '
        'useHighlight
        '
        Me.useHighlight.AccessibleName = "Enable Highlighting"
        Me.useHighlight.AutoSize = True
        Me.useHighlight.Location = New System.Drawing.Point(21, 353)
        Me.useHighlight.Name = "useHighlight"
        Me.useHighlight.Size = New System.Drawing.Size(117, 17)
        Me.useHighlight.TabIndex = 9
        Me.useHighlight.Text = "Enable Highlighting"
        Me.useHighlight.UseVisualStyleBackColor = False
        '
        'errorLabel
        '
        Me.errorLabel.AutoSize = True
        Me.errorLabel.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.errorLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.errorLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.errorLabel.ForeColor = System.Drawing.Color.Black
        Me.errorLabel.Location = New System.Drawing.Point(21, 412)
        Me.errorLabel.Name = "errorLabel"
        Me.errorLabel.Size = New System.Drawing.Size(124, 41)
        Me.errorLabel.TabIndex = 10
        Me.errorLabel.Text = "Error: No voices" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "installed! Add" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "voices before using!"
        Me.errorLabel.Visible = False
        '
        'ComboBox2
        '
        Me.ComboBox2.AccessibleName = "Speech Amount"
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(21, 82)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox2.TabIndex = 12
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(86, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Speech Amount:"
        '
        'playimg
        '
        Me.playimg.AccessibleName = "play"
        Me.playimg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.playimg.Image = CType(resources.GetObject("playimg.Image"), System.Drawing.Image)
        Me.playimg.Location = New System.Drawing.Point(35, 136)
        Me.playimg.Name = "playimg"
        Me.playimg.Size = New System.Drawing.Size(45, 45)
        Me.playimg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.playimg.TabIndex = 16
        Me.playimg.TabStop = False
        '
        'pauseimg
        '
        Me.pauseimg.AccessibleName = "pause"
        Me.pauseimg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pauseimg.Image = CType(resources.GetObject("pauseimg.Image"), System.Drawing.Image)
        Me.pauseimg.Location = New System.Drawing.Point(35, 136)
        Me.pauseimg.Name = "pauseimg"
        Me.pauseimg.Size = New System.Drawing.Size(45, 45)
        Me.pauseimg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pauseimg.TabIndex = 17
        Me.pauseimg.TabStop = False
        '
        'stopimg
        '
        Me.stopimg.AccessibleName = "stop"
        Me.stopimg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.stopimg.Image = CType(resources.GetObject("stopimg.Image"), System.Drawing.Image)
        Me.stopimg.Location = New System.Drawing.Point(79, 136)
        Me.stopimg.Name = "stopimg"
        Me.stopimg.Size = New System.Drawing.Size(45, 45)
        Me.stopimg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.stopimg.TabIndex = 18
        Me.stopimg.TabStop = False
        '
        'continuousBox
        '
        Me.continuousBox.AccessibleName = "Continuous Reading"
        Me.continuousBox.AutoSize = True
        Me.continuousBox.Location = New System.Drawing.Point(21, 376)
        Me.continuousBox.Name = "continuousBox"
        Me.continuousBox.Size = New System.Drawing.Size(122, 17)
        Me.continuousBox.TabIndex = 19
        Me.continuousBox.Text = "Continuous Reading"
        Me.continuousBox.UseVisualStyleBackColor = True
        '
        'SpeechControl
        '
        Me.AccessibleDescription = "Need to install voices"
        Me.AccessibleName = "Error"
        Me.AccessibleRole = System.Windows.Forms.AccessibleRole.Alert
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.continuousBox)
        Me.Controls.Add(Me.stopimg)
        Me.Controls.Add(Me.pauseimg)
        Me.Controls.Add(Me.playimg)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.errorLabel)
        Me.Controls.Add(Me.useHighlight)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.volumeTrackBar)
        Me.Controls.Add(Me.speedTrackBar)
        Me.Controls.Add(Me.ComboBox1)
        Me.Name = "SpeechControl"
        Me.Size = New System.Drawing.Size(173, 549)
        CType(Me.speedTrackBar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.volumeTrackBar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.playimg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pauseimg, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.stopimg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents speedTrackBar As System.Windows.Forms.TrackBar
    Friend WithEvents volumeTrackBar As System.Windows.Forms.TrackBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents useHighlight As System.Windows.Forms.CheckBox
    Friend WithEvents errorLabel As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents playimg As System.Windows.Forms.PictureBox
    Friend WithEvents pauseimg As System.Windows.Forms.PictureBox
    Friend WithEvents stopimg As System.Windows.Forms.PictureBox
    Friend WithEvents continuousBox As System.Windows.Forms.CheckBox

End Class
