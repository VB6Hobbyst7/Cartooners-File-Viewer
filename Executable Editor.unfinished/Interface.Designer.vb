<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InterfaceWindow
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
      Me.MenuBar = New System.Windows.Forms.MenuStrip()
      Me.FileMainMenu = New System.Windows.Forms.ToolStripMenuItem()
      Me.LoadExecutableMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileMainMenuSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.LoadRegionsMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileMainMenuSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.QuitMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.SplitterBox = New System.Windows.Forms.SplitContainer()
        Me.DataBox = New System.Windows.Forms.TextBox()
        Me.DisassemblyBox = New System.Windows.Forms.TextBox()
        Me.StatusBar = New System.Windows.Forms.StatusStrip()
        Me.StatusPositionLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusRegionLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuBar.SuspendLayout()
        CType(Me.SplitterBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitterBox.Panel1.SuspendLayout()
        Me.SplitterBox.Panel2.SuspendLayout()
        Me.SplitterBox.SuspendLayout()
        Me.StatusBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuBar
        '
        Me.MenuBar.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileMainMenu})
        Me.MenuBar.Location = New System.Drawing.Point(0, 0)
        Me.MenuBar.Name = "MenuBar"
        Me.MenuBar.Padding = New System.Windows.Forms.Padding(4, 2, 0, 2)
        Me.MenuBar.Size = New System.Drawing.Size(912, 24)
        Me.MenuBar.TabIndex = 0
        '
        'FileMainMenu
        '
        Me.FileMainMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LoadExecutableMenu, Me.FileMainMenuSeparator1, Me.LoadRegionsMenu, Me.FileMainMenuSeparator2, Me.QuitMenu})
        Me.FileMainMenu.Name = "FileMainMenu"
        Me.FileMainMenu.Size = New System.Drawing.Size(37, 20)
        Me.FileMainMenu.Text = "&File"
        '
        'LoadExecutableMenu
        '
        Me.LoadExecutableMenu.Name = "LoadExecutableMenu"
        Me.LoadExecutableMenu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.LoadExecutableMenu.Size = New System.Drawing.Size(200, 22)
        Me.LoadExecutableMenu.Text = "Load &Executable"
        '
        'FileMainMenuSeparator1
        '
        Me.FileMainMenuSeparator1.Name = "FileMainMenuSeparator1"
        Me.FileMainMenuSeparator1.Size = New System.Drawing.Size(197, 6)
        '
        'LoadRegionsMenu
        '
        Me.LoadRegionsMenu.Name = "LoadRegionsMenu"
        Me.LoadRegionsMenu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.LoadRegionsMenu.Size = New System.Drawing.Size(200, 22)
        Me.LoadRegionsMenu.Text = "Load &Regions"
        '
        'FileMainMenuSeparator2
        '
        Me.FileMainMenuSeparator2.Name = "FileMainMenuSeparator2"
        Me.FileMainMenuSeparator2.Size = New System.Drawing.Size(197, 6)
        '
        'QuitMenu
        '
        Me.QuitMenu.Name = "QuitMenu"
        Me.QuitMenu.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Q), System.Windows.Forms.Keys)
        Me.QuitMenu.Size = New System.Drawing.Size(200, 22)
        Me.QuitMenu.Text = "&Quit"
        '
        'SplitterBox
        '
        Me.SplitterBox.Dock = System.Windows.Forms.DockStyle.Top
        Me.SplitterBox.Location = New System.Drawing.Point(0, 24)
        Me.SplitterBox.Name = "SplitterBox"
        Me.SplitterBox.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitterBox.Panel1
        '
        Me.SplitterBox.Panel1.Controls.Add(Me.DataBox)
        '
        'SplitterBox.Panel2
        '
        Me.SplitterBox.Panel2.Controls.Add(Me.DisassemblyBox)
        Me.SplitterBox.Size = New System.Drawing.Size(912, 426)
        Me.SplitterBox.SplitterDistance = 265
        Me.SplitterBox.TabIndex = 1
        '
        'DataBox
        '
        Me.DataBox.AllowDrop = True
        Me.DataBox.BackColor = System.Drawing.SystemColors.Window
        Me.DataBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataBox.Font = New System.Drawing.Font("Consolas", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataBox.Location = New System.Drawing.Point(0, 0)
        Me.DataBox.Multiline = True
        Me.DataBox.Name = "DataBox"
        Me.DataBox.ReadOnly = True
        Me.DataBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataBox.Size = New System.Drawing.Size(912, 265)
        Me.DataBox.TabIndex = 0
        '
        'DisassemblyBox
        '
        Me.DisassemblyBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DisassemblyBox.Font = New System.Drawing.Font("Consolas", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DisassemblyBox.Location = New System.Drawing.Point(0, 0)
        Me.DisassemblyBox.Multiline = True
        Me.DisassemblyBox.Name = "DisassemblyBox"
        Me.DisassemblyBox.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.DisassemblyBox.Size = New System.Drawing.Size(912, 157)
        Me.DisassemblyBox.TabIndex = 3
        '
        'StatusBar
        '
        Me.StatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusPositionLabel, Me.StatusRegionLabel})
        Me.StatusBar.Location = New System.Drawing.Point(0, 450)
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.Size = New System.Drawing.Size(912, 22)
        Me.StatusBar.TabIndex = 2
        '
        'StatusPositionLabel
        '
        Me.StatusPositionLabel.Name = "StatusPositionLabel"
        Me.StatusPositionLabel.Size = New System.Drawing.Size(73, 17)
        Me.StatusPositionLabel.Text = "Position: 0/0"
        '
        'StatusRegionLabel
        '
        Me.StatusRegionLabel.Name = "StatusRegionLabel"
        Me.StatusRegionLabel.Size = New System.Drawing.Size(47, 17)
        Me.StatusRegionLabel.Text = "Region:"
        '
        'InterfaceWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(912, 472)
        Me.Controls.Add(Me.StatusBar)
        Me.Controls.Add(Me.SplitterBox)
        Me.Controls.Add(Me.MenuBar)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuBar
        Me.Name = "InterfaceWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.MenuBar.ResumeLayout(False)
        Me.MenuBar.PerformLayout()
        Me.SplitterBox.Panel1.ResumeLayout(False)
        Me.SplitterBox.Panel1.PerformLayout()
        Me.SplitterBox.Panel2.ResumeLayout(False)
        Me.SplitterBox.Panel2.PerformLayout()
        CType(Me.SplitterBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitterBox.ResumeLayout(False)
        Me.StatusBar.ResumeLayout(False)
        Me.StatusBar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuBar As System.Windows.Forms.MenuStrip
   Friend WithEvents FileMainMenu As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents LoadExecutableMenu As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents SplitterBox As System.Windows.Forms.SplitContainer
   Friend WithEvents DataBox As System.Windows.Forms.TextBox
   Friend WithEvents DisassemblyBox As System.Windows.Forms.TextBox
    Friend WithEvents FileMainMenuSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents QuitMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LoadRegionsMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FileMainMenuSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents StatusBar As System.Windows.Forms.StatusStrip
    Friend WithEvents StatusPositionLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusRegionLabel As System.Windows.Forms.ToolStripStatusLabel
End Class
