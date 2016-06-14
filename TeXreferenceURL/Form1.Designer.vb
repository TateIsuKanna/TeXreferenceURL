<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
	Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		Try
			If disposing AndAlso components IsNot Nothing Then
				components.Dispose()
			End If
		Finally
			MyBase.Dispose(disposing)
		End Try
	End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
	Private Sub InitializeComponent()
		Me.TextBox2 = New System.Windows.Forms.TextBox()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
		Me.ListBox1 = New System.Windows.Forms.ListBox()
		CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SplitContainer1.Panel1.SuspendLayout()
		Me.SplitContainer1.Panel2.SuspendLayout()
		Me.SplitContainer1.SuspendLayout()
		Me.SuspendLayout()
		'
		'TextBox2
		'
		Me.TextBox2.Dock = System.Windows.Forms.DockStyle.Fill
		Me.TextBox2.Location = New System.Drawing.Point(0, 0)
		Me.TextBox2.Multiline = True
		Me.TextBox2.Name = "TextBox2"
		Me.TextBox2.ReadOnly = True
		Me.TextBox2.ScrollBars = System.Windows.Forms.ScrollBars.Both
		Me.TextBox2.Size = New System.Drawing.Size(744, 162)
		Me.TextBox2.TabIndex = 0
		Me.TextBox2.WordWrap = False
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(12, 9)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(747, 12)
		Me.Label1.TabIndex = 0
		Me.Label1.Text = "URLか書名を列挙して下さい．Ctrl+V:貼り付け Ctrl+A:手入力追加 ↑↓:選択 Del:削除 Ctrl+Shift+Enter:生成 ※書籍情報取得に" &
	"は割りと掛かっちゃうかも"
		'
		'SplitContainer1
		'
		Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.SplitContainer1.Location = New System.Drawing.Point(14, 33)
		Me.SplitContainer1.Name = "SplitContainer1"
		Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
		'
		'SplitContainer1.Panel1
		'
		Me.SplitContainer1.Panel1.Controls.Add(Me.ListBox1)
		'
		'SplitContainer1.Panel2
		'
		Me.SplitContainer1.Panel2.Controls.Add(Me.TextBox2)
		Me.SplitContainer1.Size = New System.Drawing.Size(744, 331)
		Me.SplitContainer1.SplitterDistance = 165
		Me.SplitContainer1.TabIndex = 8
		'
		'ListBox1
		'
		Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
		Me.ListBox1.FormattingEnabled = True
		Me.ListBox1.ItemHeight = 12
		Me.ListBox1.Location = New System.Drawing.Point(0, 0)
		Me.ListBox1.Name = "ListBox1"
		Me.ListBox1.Size = New System.Drawing.Size(744, 165)
		Me.ListBox1.TabIndex = 9
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(770, 376)
		Me.Controls.Add(Me.SplitContainer1)
		Me.Controls.Add(Me.Label1)
		Me.Name = "Form1"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "TeXreferenceURL"
		Me.SplitContainer1.Panel1.ResumeLayout(False)
		Me.SplitContainer1.Panel2.ResumeLayout(False)
		Me.SplitContainer1.Panel2.PerformLayout()
		CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.SplitContainer1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents TextBox2 As TextBox
	Friend WithEvents Label1 As Label
	Friend WithEvents SplitContainer1 As SplitContainer
	Friend WithEvents ListBox1 As ListBox
End Class
