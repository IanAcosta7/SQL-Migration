<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMigration
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblServer1 = New System.Windows.Forms.Label()
        Me.lblDB1 = New System.Windows.Forms.Label()
        Me.lblUser1 = New System.Windows.Forms.Label()
        Me.lblPass1 = New System.Windows.Forms.Label()
        Me.lblServer2 = New System.Windows.Forms.Label()
        Me.lblDB2 = New System.Windows.Forms.Label()
        Me.lblUser2 = New System.Windows.Forms.Label()
        Me.lblPass2 = New System.Windows.Forms.Label()
        Me.txtServer1 = New System.Windows.Forms.TextBox()
        Me.txtDB1 = New System.Windows.Forms.TextBox()
        Me.txtUser1 = New System.Windows.Forms.TextBox()
        Me.txtPass1 = New System.Windows.Forms.TextBox()
        Me.txtServer2 = New System.Windows.Forms.TextBox()
        Me.txtDB2 = New System.Windows.Forms.TextBox()
        Me.txtUser2 = New System.Windows.Forms.TextBox()
        Me.txtPass2 = New System.Windows.Forms.TextBox()
        Me.gbOrigin = New System.Windows.Forms.GroupBox()
        Me.gbDestination = New System.Windows.Forms.GroupBox()
        Me.pbMigration = New System.Windows.Forms.ProgressBar()
        Me.bgwMigrate = New System.ComponentModel.BackgroundWorker()
        Me.btnAnalyze = New System.Windows.Forms.Button()
        Me.bgwAnalyze = New System.ComponentModel.BackgroundWorker()
        Me.lblAnalyze = New System.Windows.Forms.Label()
        Me.btnMigrate = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblAmountAnalyzed = New System.Windows.Forms.Label()
        Me.clbAnalyzedTables = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblAmountInserted = New System.Windows.Forms.Label()
        Me.lbInsertedTables = New System.Windows.Forms.ListBox()
        Me.cbReseedAndDelete = New System.Windows.Forms.CheckBox()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.gbOrigin.SuspendLayout()
        Me.gbDestination.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblServer1
        '
        Me.lblServer1.AutoSize = True
        Me.lblServer1.Location = New System.Drawing.Point(6, 27)
        Me.lblServer1.Name = "lblServer1"
        Me.lblServer1.Size = New System.Drawing.Size(46, 13)
        Me.lblServer1.TabIndex = 0
        Me.lblServer1.Text = "Servidor"
        '
        'lblDB1
        '
        Me.lblDB1.AutoSize = True
        Me.lblDB1.Location = New System.Drawing.Point(198, 27)
        Me.lblDB1.Name = "lblDB1"
        Me.lblDB1.Size = New System.Drawing.Size(31, 13)
        Me.lblDB1.TabIndex = 1
        Me.lblDB1.Text = "Base"
        '
        'lblUser1
        '
        Me.lblUser1.AutoSize = True
        Me.lblUser1.Location = New System.Drawing.Point(375, 27)
        Me.lblUser1.Name = "lblUser1"
        Me.lblUser1.Size = New System.Drawing.Size(43, 13)
        Me.lblUser1.TabIndex = 2
        Me.lblUser1.Text = "Usuario"
        '
        'lblPass1
        '
        Me.lblPass1.AutoSize = True
        Me.lblPass1.Location = New System.Drawing.Point(564, 27)
        Me.lblPass1.Name = "lblPass1"
        Me.lblPass1.Size = New System.Drawing.Size(61, 13)
        Me.lblPass1.TabIndex = 3
        Me.lblPass1.Text = "Contraseña"
        '
        'lblServer2
        '
        Me.lblServer2.AutoSize = True
        Me.lblServer2.Location = New System.Drawing.Point(6, 26)
        Me.lblServer2.Name = "lblServer2"
        Me.lblServer2.Size = New System.Drawing.Size(46, 13)
        Me.lblServer2.TabIndex = 4
        Me.lblServer2.Text = "Servidor"
        '
        'lblDB2
        '
        Me.lblDB2.AutoSize = True
        Me.lblDB2.Location = New System.Drawing.Point(198, 26)
        Me.lblDB2.Name = "lblDB2"
        Me.lblDB2.Size = New System.Drawing.Size(31, 13)
        Me.lblDB2.TabIndex = 5
        Me.lblDB2.Text = "Base"
        '
        'lblUser2
        '
        Me.lblUser2.AutoSize = True
        Me.lblUser2.Location = New System.Drawing.Point(375, 26)
        Me.lblUser2.Name = "lblUser2"
        Me.lblUser2.Size = New System.Drawing.Size(43, 13)
        Me.lblUser2.TabIndex = 6
        Me.lblUser2.Text = "Usuario"
        '
        'lblPass2
        '
        Me.lblPass2.AutoSize = True
        Me.lblPass2.Location = New System.Drawing.Point(564, 26)
        Me.lblPass2.Name = "lblPass2"
        Me.lblPass2.Size = New System.Drawing.Size(61, 13)
        Me.lblPass2.TabIndex = 7
        Me.lblPass2.Text = "Contraseña"
        '
        'txtServer1
        '
        Me.txtServer1.Location = New System.Drawing.Point(58, 23)
        Me.txtServer1.Name = "txtServer1"
        Me.txtServer1.Size = New System.Drawing.Size(134, 20)
        Me.txtServer1.TabIndex = 8
        Me.txtServer1.Text = "localhost"
        '
        'txtDB1
        '
        Me.txtDB1.Location = New System.Drawing.Point(235, 23)
        Me.txtDB1.Name = "txtDB1"
        Me.txtDB1.Size = New System.Drawing.Size(134, 20)
        Me.txtDB1.TabIndex = 9
        '
        'txtUser1
        '
        Me.txtUser1.Location = New System.Drawing.Point(424, 23)
        Me.txtUser1.Name = "txtUser1"
        Me.txtUser1.Size = New System.Drawing.Size(134, 20)
        Me.txtUser1.TabIndex = 10
        '
        'txtPass1
        '
        Me.txtPass1.Location = New System.Drawing.Point(631, 23)
        Me.txtPass1.Name = "txtPass1"
        Me.txtPass1.Size = New System.Drawing.Size(134, 20)
        Me.txtPass1.TabIndex = 11
        Me.txtPass1.UseSystemPasswordChar = True
        '
        'txtServer2
        '
        Me.txtServer2.Location = New System.Drawing.Point(58, 22)
        Me.txtServer2.Name = "txtServer2"
        Me.txtServer2.Size = New System.Drawing.Size(134, 20)
        Me.txtServer2.TabIndex = 12
        Me.txtServer2.Text = "localhost"
        '
        'txtDB2
        '
        Me.txtDB2.Location = New System.Drawing.Point(235, 22)
        Me.txtDB2.Name = "txtDB2"
        Me.txtDB2.Size = New System.Drawing.Size(134, 20)
        Me.txtDB2.TabIndex = 13
        '
        'txtUser2
        '
        Me.txtUser2.Location = New System.Drawing.Point(424, 22)
        Me.txtUser2.Name = "txtUser2"
        Me.txtUser2.Size = New System.Drawing.Size(134, 20)
        Me.txtUser2.TabIndex = 14
        '
        'txtPass2
        '
        Me.txtPass2.Location = New System.Drawing.Point(631, 22)
        Me.txtPass2.Name = "txtPass2"
        Me.txtPass2.Size = New System.Drawing.Size(134, 20)
        Me.txtPass2.TabIndex = 15
        Me.txtPass2.UseSystemPasswordChar = True
        '
        'gbOrigin
        '
        Me.gbOrigin.Controls.Add(Me.lblServer1)
        Me.gbOrigin.Controls.Add(Me.lblDB1)
        Me.gbOrigin.Controls.Add(Me.lblUser1)
        Me.gbOrigin.Controls.Add(Me.lblPass1)
        Me.gbOrigin.Controls.Add(Me.txtServer1)
        Me.gbOrigin.Controls.Add(Me.txtDB1)
        Me.gbOrigin.Controls.Add(Me.txtPass1)
        Me.gbOrigin.Controls.Add(Me.txtUser1)
        Me.gbOrigin.Location = New System.Drawing.Point(12, 12)
        Me.gbOrigin.Name = "gbOrigin"
        Me.gbOrigin.Size = New System.Drawing.Size(776, 60)
        Me.gbOrigin.TabIndex = 16
        Me.gbOrigin.TabStop = False
        Me.gbOrigin.Text = "Origen"
        '
        'gbDestination
        '
        Me.gbDestination.Controls.Add(Me.lblServer2)
        Me.gbDestination.Controls.Add(Me.lblDB2)
        Me.gbDestination.Controls.Add(Me.txtPass2)
        Me.gbDestination.Controls.Add(Me.lblUser2)
        Me.gbDestination.Controls.Add(Me.txtUser2)
        Me.gbDestination.Controls.Add(Me.lblPass2)
        Me.gbDestination.Controls.Add(Me.txtDB2)
        Me.gbDestination.Controls.Add(Me.txtServer2)
        Me.gbDestination.Location = New System.Drawing.Point(12, 78)
        Me.gbDestination.Name = "gbDestination"
        Me.gbDestination.Size = New System.Drawing.Size(776, 60)
        Me.gbDestination.TabIndex = 17
        Me.gbDestination.TabStop = False
        Me.gbDestination.Text = "Destino"
        '
        'pbMigration
        '
        Me.pbMigration.Location = New System.Drawing.Point(12, 440)
        Me.pbMigration.Name = "pbMigration"
        Me.pbMigration.Size = New System.Drawing.Size(776, 23)
        Me.pbMigration.TabIndex = 21
        '
        'bgwMigrate
        '
        Me.bgwMigrate.WorkerReportsProgress = True
        '
        'btnAnalyze
        '
        Me.btnAnalyze.Location = New System.Drawing.Point(12, 144)
        Me.btnAnalyze.Name = "btnAnalyze"
        Me.btnAnalyze.Size = New System.Drawing.Size(75, 23)
        Me.btnAnalyze.TabIndex = 24
        Me.btnAnalyze.Text = "Analizar"
        Me.btnAnalyze.UseVisualStyleBackColor = True
        '
        'bgwAnalyze
        '
        '
        'lblAnalyze
        '
        Me.lblAnalyze.AutoSize = True
        Me.lblAnalyze.BackColor = System.Drawing.Color.Transparent
        Me.lblAnalyze.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblAnalyze.Location = New System.Drawing.Point(6, 239)
        Me.lblAnalyze.Name = "lblAnalyze"
        Me.lblAnalyze.Size = New System.Drawing.Size(0, 13)
        Me.lblAnalyze.TabIndex = 27
        '
        'btnMigrate
        '
        Me.btnMigrate.Enabled = False
        Me.btnMigrate.Location = New System.Drawing.Point(509, 143)
        Me.btnMigrate.Name = "btnMigrate"
        Me.btnMigrate.Size = New System.Drawing.Size(75, 23)
        Me.btnMigrate.TabIndex = 28
        Me.btnMigrate.Text = "Migrar"
        Me.btnMigrate.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblAmountAnalyzed)
        Me.GroupBox1.Controls.Add(Me.clbAnalyzedTables)
        Me.GroupBox1.Controls.Add(Me.lblAnalyze)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 174)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(385, 260)
        Me.GroupBox1.TabIndex = 29
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tablas Analizadas"
        '
        'lblAmountAnalyzed
        '
        Me.lblAmountAnalyzed.Location = New System.Drawing.Point(235, 235)
        Me.lblAmountAnalyzed.Name = "lblAmountAnalyzed"
        Me.lblAmountAnalyzed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAmountAnalyzed.Size = New System.Drawing.Size(144, 21)
        Me.lblAmountAnalyzed.TabIndex = 28
        Me.lblAmountAnalyzed.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'clbAnalyzedTables
        '
        Me.clbAnalyzedTables.FormattingEnabled = True
        Me.clbAnalyzedTables.Location = New System.Drawing.Point(6, 19)
        Me.clbAnalyzedTables.Name = "clbAnalyzedTables"
        Me.clbAnalyzedTables.Size = New System.Drawing.Size(373, 214)
        Me.clbAnalyzedTables.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblAmountInserted)
        Me.GroupBox2.Controls.Add(Me.lbInsertedTables)
        Me.GroupBox2.Location = New System.Drawing.Point(403, 174)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(385, 260)
        Me.GroupBox2.TabIndex = 30
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Tablas Migradas"
        '
        'lblAmountInserted
        '
        Me.lblAmountInserted.Location = New System.Drawing.Point(235, 235)
        Me.lblAmountInserted.Name = "lblAmountInserted"
        Me.lblAmountInserted.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAmountInserted.Size = New System.Drawing.Size(144, 21)
        Me.lblAmountInserted.TabIndex = 29
        Me.lblAmountInserted.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbInsertedTables
        '
        Me.lbInsertedTables.FormattingEnabled = True
        Me.lbInsertedTables.Location = New System.Drawing.Point(6, 19)
        Me.lbInsertedTables.Name = "lbInsertedTables"
        Me.lbInsertedTables.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lbInsertedTables.Size = New System.Drawing.Size(373, 212)
        Me.lbInsertedTables.Sorted = True
        Me.lbInsertedTables.TabIndex = 26
        '
        'cbReseedAndDelete
        '
        Me.cbReseedAndDelete.AutoSize = True
        Me.cbReseedAndDelete.Enabled = False
        Me.cbReseedAndDelete.Location = New System.Drawing.Point(590, 148)
        Me.cbReseedAndDelete.Name = "cbReseedAndDelete"
        Me.cbReseedAndDelete.Size = New System.Drawing.Size(156, 17)
        Me.cbReseedAndDelete.TabIndex = 31
        Me.cbReseedAndDelete.Text = "Reseed y borrado de tablas"
        Me.cbReseedAndDelete.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(409, 144)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(94, 23)
        Me.btnExport.TabIndex = 32
        Me.btnExport.Text = "Exportar a Excel"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'frmMigration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 474)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.cbReseedAndDelete)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnMigrate)
        Me.Controls.Add(Me.btnAnalyze)
        Me.Controls.Add(Me.gbDestination)
        Me.Controls.Add(Me.gbOrigin)
        Me.Controls.Add(Me.pbMigration)
        Me.MaximumSize = New System.Drawing.Size(816, 513)
        Me.MinimumSize = New System.Drawing.Size(816, 513)
        Me.Name = "frmMigration"
        Me.Text = "BizOne -Migration"
        Me.gbOrigin.ResumeLayout(False)
        Me.gbOrigin.PerformLayout()
        Me.gbDestination.ResumeLayout(False)
        Me.gbDestination.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblServer1 As Label
    Friend WithEvents lblDB1 As Label
    Friend WithEvents lblUser1 As Label
    Friend WithEvents lblPass1 As Label
    Friend WithEvents lblServer2 As Label
    Friend WithEvents lblDB2 As Label
    Friend WithEvents lblUser2 As Label
    Friend WithEvents lblPass2 As Label
    Friend WithEvents txtServer1 As TextBox
    Friend WithEvents txtDB1 As TextBox
    Friend WithEvents txtUser1 As TextBox
    Friend WithEvents txtPass1 As TextBox
    Friend WithEvents txtServer2 As TextBox
    Friend WithEvents txtDB2 As TextBox
    Friend WithEvents txtUser2 As TextBox
    Friend WithEvents txtPass2 As TextBox
    Friend WithEvents gbOrigin As GroupBox
    Friend WithEvents gbDestination As GroupBox
    Friend WithEvents pbMigration As ProgressBar
    Friend WithEvents bgwMigrate As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnAnalyze As Button
    Friend WithEvents bgwAnalyze As System.ComponentModel.BackgroundWorker
    Friend WithEvents lblAnalyze As Label
    Friend WithEvents btnMigrate As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents cbReseedAndDelete As CheckBox
    Friend WithEvents btnExport As Button
    Friend WithEvents lblAmountAnalyzed As Label
    Friend WithEvents clbAnalyzedTables As CheckedListBox
    Friend WithEvents lbInsertedTables As ListBox
    Friend WithEvents lblAmountInserted As Label
End Class
