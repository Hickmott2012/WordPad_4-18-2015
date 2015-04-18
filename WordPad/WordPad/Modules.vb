Imports System.Speech.Synthesis

Module SpeakModule

    Dim speaker As New SpeechSynthesizer
    Dim checkedBoolen As Boolean

    Sub Main()
        AddHandler speaker.SpeakCompleted, AddressOf SpeachComplete
        If checkedBoolen = False Then
            checkedBoolen = True
            speaker.SpeakAsync(WordPadForm.RichTextBox1.SelectedText)
        Else
            speaker.SpeakAsyncCancelAll()
        End If
        'Console.ReadLine()
    End Sub

    Private Sub SpeachComplete(ByVal sender As Object, ByVal e As SpeakCompletedEventArgs)
        checkedBoolen = False
    End Sub
End Module

Module TextValidation
    Dim DeterminedFont, DeterminedSize As String
    Sub Main()
        Try
            ' Determine the font style
            DeterminedFont = Convert.ToString(WordPadForm.RichTextBox1.SelectionFont.Name)
            ' Put the appropriate font into fontstylecombobox
            WordPadForm.FontStyleStripComboBox.Text = DeterminedFont
            ' Determine the size
            DeterminedSize = Convert.ToSingle(WordPadForm.RichTextBox1.SelectionFont.Size)
            ' Put the appropraite font into fontsizecombobox
            WordPadForm.FontSizeToolStripComboBox.SelectedItem = DeterminedSize
        Catch ex As Exception
            WordPadForm.FontStyleStripComboBox.Text = ""
        End Try
    End Sub
End Module