<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSample
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmSample))
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.TSSLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TSProgressBar = New System.Windows.Forms.ToolStripProgressBar()
        Me.TSSBtnStop = New System.Windows.Forms.ToolStripSplitButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSSLabel, Me.TSProgressBar, Me.TSSBtnStop})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 539)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(784, 23)
        Me.StatusStrip.TabIndex = 0
        Me.StatusStrip.Text = "StatusStrip1"
        '
        'TSSLabel
        '
        Me.TSSLabel.Name = "TSSLabel"
        Me.TSSLabel.Size = New System.Drawing.Size(635, 18)
        Me.TSSLabel.Spring = True
        Me.TSSLabel.Text = "ToolStripStatusLabel"
        Me.TSSLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TSProgressBar
        '
        Me.TSProgressBar.Name = "TSProgressBar"
        Me.TSProgressBar.Size = New System.Drawing.Size(100, 17)
        '
        'TSSBtnStop
        '
        Me.TSSBtnStop.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TSSBtnStop.Image = CType(resources.GetObject("TSSBtnStop.Image"), System.Drawing.Image)
        Me.TSSBtnStop.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSSBtnStop.Name = "TSSBtnStop"
        Me.TSSBtnStop.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never
        Me.TSSBtnStop.Size = New System.Drawing.Size(32, 21)
        Me.TSSBtnStop.Text = "ToolStripSplitButton1"
        Me.TSSBtnStop.ToolTipText = "中止"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(697, 513)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.StatusStrip)
        Me.Name = "FrmMain"
        Me.Text = "Form1"
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents TSSLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TSProgressBar As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TSSBtnStop As System.Windows.Forms.ToolStripSplitButton

End Class
