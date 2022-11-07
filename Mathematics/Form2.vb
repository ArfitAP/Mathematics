
' Redaka: 49

Public Class Form2

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim greska As Boolean

        If IsNumeric(TextBox1.Text) Then
            Kut.stupnjevi = Int(TextBox1.Text)
        Else
            greska = True
        End If
        If IsNumeric(TextBox2.Text) Then
            Kut.minute = Int(TextBox2.Text)
        Else
            greska = True
        End If
        If IsNumeric(TextBox3.Text) Then
            Kut.sekunde = TextBox3.Text
        Else
            greska = True
        End If


        If greska = True Or Kut.Provjeri() Then
            TextBox1.Text = ""
            TextBox2.Text = "0"
            TextBox3.Text = "0"
            Kut.stupnjevi = 0
            Kut.minute = 0
            Kut.sekunde = 0
            Label5.Text = "Unijeli ste pogresne podatke, pokusajte ponovo !"
        Else
            Me.Close()
            Label5.Text = ""
            TextBox1.Text = ""
            TextBox2.Text = "0"
            TextBox3.Text = "0"
        End If

    End Sub

    Private Sub Form2_Show(sender As Object, e As EventArgs) Handles MyBase.Shown
        TextBox1.Focus()
        TextBox1.SelectionStart = TextBox1.TextLength
        TextBox1.ScrollToCaret()
    End Sub
End Class