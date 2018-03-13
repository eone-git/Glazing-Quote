<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGlazingNote
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
        Me.uscGlassQuote = New Infragistics.Win.UltraWinSpellChecker.UltraSpellChecker(Me.components)
        Me.brnAdd = New System.Windows.Forms.Button()
        Me.UltraFormattedTextWordWriter1 = New Infragistics.Win.UltraWinFormattedText.WordWriter.UltraFormattedTextWordWriter(Me.components)
        Me.btnSpellCheck = New System.Windows.Forms.Button()
        Me.UltraFormattedTextWordWriter2 = New Infragistics.Win.UltraWinFormattedText.WordWriter.UltraFormattedTextWordWriter(Me.components)
        Me.frmGlazingNote_Fill_Panel = New Infragistics.Win.Misc.UltraPanel()
        Me.utxtNoteText = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        CType(Me.uscGlassQuote, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.frmGlazingNote_Fill_Panel.ClientArea.SuspendLayout()
        Me.frmGlazingNote_Fill_Panel.SuspendLayout()
        CType(Me.utxtNoteText, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'uscGlassQuote
        '
        Me.uscGlassQuote.ContainingControl = Me
        Me.uscGlassQuote.Mode = Infragistics.Win.UltraWinSpellChecker.SpellCheckingMode.AsYouType
        Me.uscGlassQuote.ShowDialogsModal = False
        Me.uscGlassQuote.SpellOptions.AllowCapitalizedWords = True
        Me.uscGlassQuote.SpellOptions.AllowMixedCase = True
        Me.uscGlassQuote.SpellOptions.AllowWordsWithDigits = True
        Me.uscGlassQuote.SpellOptions.CheckCompoundWords = True
        Me.uscGlassQuote.SpellOptions.CheckHyphenatedText = False
        Me.uscGlassQuote.SpellOptions.SeparateHyphenWords = True
        Me.uscGlassQuote.UnderlineSpellingErrorColor = System.Drawing.Color.DarkRed
        Me.uscGlassQuote.UnderlineSpellingErrorStyle = Infragistics.Win.UltraWinSpellChecker.UnderlineErrorsStyle.SingleLine
        '
        'brnAdd
        '
        Me.brnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.brnAdd.BackColor = System.Drawing.Color.DodgerBlue
        Me.brnAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.brnAdd.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.brnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.brnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.brnAdd.ForeColor = System.Drawing.Color.White
        Me.brnAdd.Location = New System.Drawing.Point(906, 405)
        Me.brnAdd.Name = "brnAdd"
        Me.brnAdd.Size = New System.Drawing.Size(75, 24)
        Me.brnAdd.TabIndex = 2
        Me.brnAdd.Text = "OK"
        Me.brnAdd.UseVisualStyleBackColor = False
        '
        'btnSpellCheck
        '
        Me.btnSpellCheck.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSpellCheck.BackColor = System.Drawing.Color.DodgerBlue
        Me.btnSpellCheck.FlatAppearance.BorderColor = System.Drawing.Color.DodgerBlue
        Me.btnSpellCheck.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSpellCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSpellCheck.ForeColor = System.Drawing.Color.White
        Me.btnSpellCheck.Location = New System.Drawing.Point(825, 405)
        Me.btnSpellCheck.Name = "btnSpellCheck"
        Me.btnSpellCheck.Size = New System.Drawing.Size(75, 24)
        Me.btnSpellCheck.TabIndex = 2
        Me.btnSpellCheck.Text = "Spell Check"
        Me.btnSpellCheck.UseVisualStyleBackColor = False
        '
        'frmGlazingNote_Fill_Panel
        '
        '
        'frmGlazingNote_Fill_Panel.ClientArea
        '
        Me.frmGlazingNote_Fill_Panel.ClientArea.Controls.Add(Me.utxtNoteText)
        Me.frmGlazingNote_Fill_Panel.ClientArea.Controls.Add(Me.btnSpellCheck)
        Me.frmGlazingNote_Fill_Panel.ClientArea.Controls.Add(Me.brnAdd)
        Me.frmGlazingNote_Fill_Panel.Cursor = System.Windows.Forms.Cursors.Default
        Me.frmGlazingNote_Fill_Panel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.frmGlazingNote_Fill_Panel.Location = New System.Drawing.Point(0, 0)
        Me.frmGlazingNote_Fill_Panel.Name = "frmGlazingNote_Fill_Panel"
        Me.frmGlazingNote_Fill_Panel.Size = New System.Drawing.Size(984, 432)
        Me.frmGlazingNote_Fill_Panel.TabIndex = 0
        '
        'utxtNoteText
        '
        Me.utxtNoteText.AcceptsTab = True
        Me.utxtNoteText.Dock = System.Windows.Forms.DockStyle.Top
        Me.utxtNoteText.Location = New System.Drawing.Point(0, 0)
        Me.utxtNoteText.Multiline = True
        Me.utxtNoteText.Name = "utxtNoteText"
        Me.utxtNoteText.Size = New System.Drawing.Size(984, 402)
        Me.utxtNoteText.SpellChecker = Me.uscGlassQuote
        Me.utxtNoteText.TabIndex = 3
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'frmGlazingNote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(984, 432)
        Me.Controls.Add(Me.frmGlazingNote_Fill_Panel)
        Me.Name = "frmGlazingNote"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        CType(Me.uscGlassQuote, System.ComponentModel.ISupportInitialize).EndInit()
        Me.frmGlazingNote_Fill_Panel.ClientArea.ResumeLayout(False)
        Me.frmGlazingNote_Fill_Panel.ClientArea.PerformLayout()
        Me.frmGlazingNote_Fill_Panel.ResumeLayout(False)
        CType(Me.utxtNoteText, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents brnAdd As System.Windows.Forms.Button
    Friend WithEvents UltraFormattedTextWordWriter1 As Infragistics.Win.UltraWinFormattedText.WordWriter.UltraFormattedTextWordWriter
    Friend WithEvents uscGlassQuote As Infragistics.Win.UltraWinSpellChecker.UltraSpellChecker
    Friend WithEvents btnSpellCheck As System.Windows.Forms.Button
    Friend WithEvents frmGlazingNote_Fill_Panel As Infragistics.Win.Misc.UltraPanel
    Friend WithEvents UltraFormattedTextWordWriter2 As Infragistics.Win.UltraWinFormattedText.WordWriter.UltraFormattedTextWordWriter
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents utxtNoteText As Infragistics.Win.UltraWinEditors.UltraTextEditor
End Class
