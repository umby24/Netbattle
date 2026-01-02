<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class MSListing
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents KickTimer As System.Windows.Forms.Timer
	Public WithEvents cmdMassMsg As System.Windows.Forms.Button
	Public WithEvents ChannelScanner As System.Windows.Forms.Timer
	Public WithEvents QueueTimer As System.Windows.Forms.Timer
	Public WithEvents _ServerSocket_0 As AxMSWinsockLib.AxWinsock
	Public WithEvents _ClientSocket_0 As AxMSWinsockLib.AxWinsock
	Public WithEvents cmdDisconnect As System.Windows.Forms.Button
	Public WithEvents _ListDisplay_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListDisplay_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListDisplay_ColumnHeader_3 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListDisplay_ColumnHeader_4 As System.Windows.Forms.ColumnHeader
	Public WithEvents _ListDisplay_ColumnHeader_5 As System.Windows.Forms.ColumnHeader
	Public WithEvents ListDisplay As System.Windows.Forms.ListView
	Public WithEvents _StatusBar1_Panel1 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents _StatusBar1_Panel2 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents _StatusBar1_Panel3 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents StatusBar1 As System.Windows.Forms.StatusStrip
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents ClientSocket As AxWinsockArray
	Public WithEvents ServerSocket As AxWinsockArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MSListing))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.KickTimer = New System.Windows.Forms.Timer(components)
		Me.cmdMassMsg = New System.Windows.Forms.Button
		Me.ChannelScanner = New System.Windows.Forms.Timer(components)
		Me.QueueTimer = New System.Windows.Forms.Timer(components)
		Me._ServerSocket_0 = New AxMSWinsockLib.AxWinsock
		Me._ClientSocket_0 = New AxMSWinsockLib.AxWinsock
		Me.cmdDisconnect = New System.Windows.Forms.Button
		Me.ListDisplay = New System.Windows.Forms.ListView
		Me._ListDisplay_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
		Me._ListDisplay_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
		Me._ListDisplay_ColumnHeader_3 = New System.Windows.Forms.ColumnHeader
		Me._ListDisplay_ColumnHeader_4 = New System.Windows.Forms.ColumnHeader
		Me._ListDisplay_ColumnHeader_5 = New System.Windows.Forms.ColumnHeader
		Me.StatusBar1 = New System.Windows.Forms.StatusStrip
		Me._StatusBar1_Panel1 = New System.Windows.Forms.ToolStripStatusLabel
		Me._StatusBar1_Panel2 = New System.Windows.Forms.ToolStripStatusLabel
		Me._StatusBar1_Panel3 = New System.Windows.Forms.ToolStripStatusLabel
		Me.Label1 = New System.Windows.Forms.Label
		Me.ClientSocket = New AxWinsockArray(components)
		Me.ServerSocket = New AxWinsockArray(components)
		Me.ListDisplay.SuspendLayout()
		Me.StatusBar1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me._ServerSocket_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me._ClientSocket_0, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.ClientSocket, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.ServerSocket, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Text = "NetBattle Server List"
		Me.ClientSize = New System.Drawing.Size(498, 289)
		Me.Location = New System.Drawing.Point(12, 37)
		Me.Icon = CType(resources.GetObject("MSListing.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "MSListing"
		Me.KickTimer.Interval = 1000
		Me.KickTimer.Enabled = True
		Me.cmdMassMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdMassMsg.Text = "Send Mass Message"
		Me.cmdMassMsg.Size = New System.Drawing.Size(113, 25)
		Me.cmdMassMsg.Location = New System.Drawing.Point(376, 232)
		Me.cmdMassMsg.TabIndex = 4
		Me.cmdMassMsg.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdMassMsg.BackColor = System.Drawing.SystemColors.Control
		Me.cmdMassMsg.CausesValidation = True
		Me.cmdMassMsg.Enabled = True
		Me.cmdMassMsg.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdMassMsg.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdMassMsg.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdMassMsg.TabStop = True
		Me.cmdMassMsg.Name = "cmdMassMsg"
		Me.ChannelScanner.Interval = 60000
		Me.ChannelScanner.Enabled = True
		Me.QueueTimer.Interval = 5
		Me.QueueTimer.Enabled = True
		_ServerSocket_0.OcxState = CType(resources.GetObject("_ServerSocket_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._ServerSocket_0.Location = New System.Drawing.Point(440, 0)
		Me._ServerSocket_0.Name = "_ServerSocket_0"
		_ClientSocket_0.OcxState = CType(resources.GetObject("_ClientSocket_0.OcxState"), System.Windows.Forms.AxHost.State)
		Me._ClientSocket_0.Location = New System.Drawing.Point(408, 0)
		Me._ClientSocket_0.Name = "_ClientSocket_0"
		Me.cmdDisconnect.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdDisconnect.Text = "Disconnect Server"
		Me.cmdDisconnect.Size = New System.Drawing.Size(113, 25)
		Me.cmdDisconnect.Location = New System.Drawing.Point(376, 200)
		Me.cmdDisconnect.TabIndex = 2
		Me.cmdDisconnect.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDisconnect.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDisconnect.CausesValidation = True
		Me.cmdDisconnect.Enabled = True
		Me.cmdDisconnect.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDisconnect.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDisconnect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDisconnect.TabStop = True
		Me.cmdDisconnect.Name = "cmdDisconnect"
		Me.ListDisplay.Size = New System.Drawing.Size(481, 177)
		Me.ListDisplay.Location = New System.Drawing.Point(8, 8)
		Me.ListDisplay.TabIndex = 0
		Me.ListDisplay.View = System.Windows.Forms.View.Details
		Me.ListDisplay.Alignment = System.Windows.Forms.ListViewAlignment.Left
		Me.ListDisplay.LabelEdit = False
		Me.ListDisplay.LabelWrap = True
		Me.ListDisplay.HideSelection = False
		Me.ListDisplay.FullRowSelect = True
		Me.ListDisplay.ForeColor = System.Drawing.SystemColors.WindowText
		Me.ListDisplay.BackColor = System.Drawing.SystemColors.Window
		Me.ListDisplay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ListDisplay.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.ListDisplay.Name = "ListDisplay"
		Me._ListDisplay_ColumnHeader_1.Text = "#"
		Me._ListDisplay_ColumnHeader_1.Width = 48
		Me._ListDisplay_ColumnHeader_2.Text = "Name"
		Me._ListDisplay_ColumnHeader_2.Width = 212
		Me._ListDisplay_ColumnHeader_3.Text = "Address / IP"
		Me._ListDisplay_ColumnHeader_3.Width = 236
		Me._ListDisplay_ColumnHeader_4.Text = "Main Admin"
		Me._ListDisplay_ColumnHeader_4.Width = 200
		Me._ListDisplay_ColumnHeader_5.Text = "Users/Max"
		Me._ListDisplay_ColumnHeader_5.Width = 117
		Me.StatusBar1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.StatusBar1.Size = New System.Drawing.Size(498, 18)
		Me.StatusBar1.Location = New System.Drawing.Point(0, 271)
		Me.StatusBar1.TabIndex = 1
		Me.StatusBar1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StatusBar1.Name = "StatusBar1"
		Me._StatusBar1_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
		Me._StatusBar1_Panel1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
		Me._StatusBar1_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._StatusBar1_Panel1.Size = New System.Drawing.Size(303, 18)
		Me._StatusBar1_Panel1.Spring = True
		Me._StatusBar1_Panel1.AutoSize = True
		Me._StatusBar1_Panel1.BorderSides = CType(System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom, System.Windows.Forms.ToolStripStatusLabelBorderSides)
		Me._StatusBar1_Panel1.Margin = New System.Windows.Forms.Padding(0)
		Me._StatusBar1_Panel1.AutoSize = False
		Me._StatusBar1_Panel2.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
		Me._StatusBar1_Panel2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
		Me._StatusBar1_Panel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._StatusBar1_Panel2.AutoSize = True
		Me._StatusBar1_Panel2.BorderSides = CType(System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom, System.Windows.Forms.ToolStripStatusLabelBorderSides)
		Me._StatusBar1_Panel2.Margin = New System.Windows.Forms.Padding(0)
		Me._StatusBar1_Panel2.Size = New System.Drawing.Size(96, 18)
		Me._StatusBar1_Panel2.AutoSize = False
		Me._StatusBar1_Panel3.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
		Me._StatusBar1_Panel3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
		Me._StatusBar1_Panel3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._StatusBar1_Panel3.AutoSize = True
		Me._StatusBar1_Panel3.BorderSides = CType(System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom, System.Windows.Forms.ToolStripStatusLabelBorderSides)
		Me._StatusBar1_Panel3.Margin = New System.Windows.Forms.Padding(0)
		Me._StatusBar1_Panel3.Size = New System.Drawing.Size(96, 18)
		Me._StatusBar1_Panel3.AutoSize = False
		Me.Label1.BackColor = System.Drawing.Color.White
		Me.Label1.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Label1.Size = New System.Drawing.Size(347, 64)
		Me.Label1.Location = New System.Drawing.Point(16, 200)
		Me.Label1.TabIndex = 3
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(cmdMassMsg)
		Me.Controls.Add(_ServerSocket_0)
		Me.Controls.Add(_ClientSocket_0)
		Me.Controls.Add(cmdDisconnect)
		Me.Controls.Add(ListDisplay)
		Me.Controls.Add(StatusBar1)
		Me.Controls.Add(Label1)
		Me.ListDisplay.Columns.Add(_ListDisplay_ColumnHeader_1)
		Me.ListDisplay.Columns.Add(_ListDisplay_ColumnHeader_2)
		Me.ListDisplay.Columns.Add(_ListDisplay_ColumnHeader_3)
		Me.ListDisplay.Columns.Add(_ListDisplay_ColumnHeader_4)
		Me.ListDisplay.Columns.Add(_ListDisplay_ColumnHeader_5)
		Me.StatusBar1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._StatusBar1_Panel1})
		Me.StatusBar1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._StatusBar1_Panel2})
		Me.StatusBar1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._StatusBar1_Panel3})
		Me.ClientSocket.SetIndex(_ClientSocket_0, CType(0, Short))
		Me.ServerSocket.SetIndex(_ServerSocket_0, CType(0, Short))
		CType(Me.ServerSocket, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.ClientSocket, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._ClientSocket_0, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me._ServerSocket_0, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ListDisplay.ResumeLayout(False)
		Me.StatusBar1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class