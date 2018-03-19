Public Class frmGlazingMessageBox
    Dim messagFor As String = ""
    Public returnText As String = ""
    Dim defaultModulelName = "SPIL GLass"
    Dim newclsGlazingQuoteExtensionObj As clsGlazingQuoteExtension

    Public Sub New(ByRef clsGlazingQuoteExtensionObj As clsGlazingQuoteExtension, Optional ByRef messageBoxTitle As String = "", Optional ByRef messageLable As String = "")
        ' This call is required by the designer.
        InitializeComponent()
        Me.newclsGlazingQuoteExtensionObj = clsGlazingQuoteExtensionObj

        If Me.Text = "" Then
            Me.Text = defaultModulelName
        Else
            Me.Text = messageBoxTitle
        End If

        If Me.lblReason.Text = "" Then
            Me.lblReason.Text = "Reason"
        Else
            Me.lblReason.Text = messageLable
        End If

    End Sub

    Public Property SkuText() As String
        Get
            Return txtInputText.Text
        End Get
        Set(ByVal value As String)
            txtInputText.Text = value
        End Set
    End Property

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        GetReturnText()
    End Sub

    Public Sub GetReturnText()
        Try
            If IsNothing(txtInputText.Text) = False Then
                newclsGlazingQuoteExtensionObj.reuturnedMessage = txtInputText.Text

            End If

        Catch ex As Exception
            GQShowMessage(ex.Message, messagFor, MsgBoxStyle.Critical)

        End Try
    End Sub

    Public Sub SetMessageBoxData(Optional ByRef messageBoxTitle As String = "", Optional ByRef messageLable As String = "")


    End Sub

End Class