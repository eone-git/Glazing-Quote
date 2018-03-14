Public Class frmGlazingNote

    Public selectedDocDes As String
    Public isJobDescriptionActive As Boolean = False

    Dim frmGlazingQuote As frmGlazingQuote

    Public Sub New(ByRef frmGlazingQuote As frmGlazingQuote)
        ' This call is required by the designer.
        InitializeComponent()
        Me.frmGlazingQuote = frmGlazingQuote
        'selectedRow = frmGlazingQuote.UG2.ActiveRow
        'If selectedRow.Cells("ItemType").Value > 0 Then
        '    isLoading = True
        'End If
    End Sub
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub utxtNoteText_KeyDown(sender As Object, e As KeyEventArgs) Handles utxtNoteText.KeyDown
        GlazingNote(e)
    End Sub

    Public Sub GlazingNote(e As KeyEventArgs)
        Try
            If e.KeyCode = 119 Then
                Dim newGlazingDocDescription As New frmGlazingDocDescription
                newGlazingDocDescription.DocDesTypeName = "Footer Note"
                newGlazingDocDescription.ShowDialog()

                If utxtNoteText.Text <> "" Then
                    utxtNoteText.Value = utxtNoteText.Text + vbCrLf + frmGlazingQuote.selectedDocDes
                ElseIf utxtNoteText.Text = "" Then
                    utxtNoteText.Value = frmGlazingQuote.selectedDocDes
                End If
            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        Finally
            selectedDocDes = ""

        End Try
    End Sub

    Private Sub brnAdd_Click(sender As Object, e As EventArgs) Handles brnAdd.Click
        Dim utxtNoteTextValue As String

        If IsNothing(utxtNoteText.Text) = False Then
            utxtNoteTextValue = utxtNoteText.Text

        Else
            utxtNoteTextValue = ""

        End If

        If isJobDescriptionActive = False Then
            frmGlazingQuote.utxtNoteText.Value = utxtNoteTextValue

        Else
            frmGlazingQuote.jobDescription = utxtNoteTextValue

        End If

    End Sub

    Private Sub frmGlazingNote_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If isJobDescriptionActive Then
                'Me.ShowIcon = True
                'Me.ShowInTaskbar = True
                Me.Text = "Job Description"

            Else
                'Me.ShowIcon = False
                'Me.ShowInTaskbar = False
                Me.Text = "Footer Note"

            End If
        Catch ex As Exception
            modGlazingQuoteExtension.GQShowMessage(ex.Message, Me.Text, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub RepopulateAutoCorrectionList()
        ''For eahc entry in the auto corrections hashtable
        'Dim entry As DictionaryEntry
        'For Each entry In Me.autoCorrections

        '    'Add the auto correction entry to the list view
        '    Dim autoCorrection As UltraListViewItem = New UltraListViewItem(entry.Key, New Object() {entry.Value})
        '    Me.lvAutoCorrections.Items.Add(autoCorrection)

        'Next
    End Sub


    Private Sub btnSpellCheck_Click(sender As Object, e As EventArgs) Handles btnSpellCheck.Click
        Me.uscGlassQuote.ShowSpellCheckDialog(utxtNoteText)
        'uscGlassQuote.ShowDialogsMo()
    End Sub

    Private Sub uscGlassQuote_SpellCheckDialogOpening(sender As Object, e As Infragistics.Win.UltraWinSpellChecker.SpellCheckDialogOpeningEventArgs) Handles uscGlassQuote.SpellCheckDialogOpening

    End Sub
End Class