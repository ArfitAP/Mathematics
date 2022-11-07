
' Redaka: 194

Imports System.Math
Public Class Form3
    Dim koef As Double
    Dim tocke(900) As Point
    Dim tocke1(900) As Point
    Dim x As Double
    Dim k As Boolean = False
    Dim a, b, c, x1, x2, y1, y2 As Double
    Dim T As Point

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label5.Text = "Tocke presjeka:"
        Label5.Visible = False
        koef = 250 / Kut.mj

        For i = 0 To 900
            x = (Kut.mj * (i - 450)) / 250
            If Abs(koef * (Kut.ax * x * x + Kut.bx * x + Kut.cx)) < 425 Then
                tocke(i).Y = 425 - Round(koef * (Kut.ax * x * x + Kut.bx * x + Kut.cx))
                k = True
            Else
                If ax < 0 Then
                    tocke(i).Y = 850
                ElseIf ax > 0 Then
                    tocke(i).Y = -1
                ElseIf bx > 0 And k = False Then
                    tocke(i).Y = 850
                ElseIf bx > 0 And k = True Then
                    tocke(i).Y = -1
                ElseIf bx < 0 And k = False Then
                    tocke(i).Y = -1
                ElseIf bx < 0 And k = True Then
                    tocke(i).Y = 850
                End If
            End If
            tocke(i).X = i

        Next i
        k = False
        If Abs(Kut.ax1) > 0 Or Abs(Kut.bx1) > 0 Or Abs(Kut.cx1) > 0 Then
            a = Kut.ax - Kut.ax1
            b = Kut.bx - Kut.bx1
            c = Kut.cx - Kut.cx1

            Dim koef1 As Double
            Dim koef2 As Double
            Dim kutp As Double
            koef1 = 0
            koef2 = 0

            If Abs(a) > 0 Then
                If b * b - 4 * a * c >= 0 Then
                    x1 = (-b + Sqrt(b * b - 4 * a * c)) / (2 * a)
                    x2 = (-b - Sqrt(b * b - 4 * a * c)) / (2 * a)
                    y1 = Kut.ax * x1 * x1 + Kut.bx * x1 + Kut.cx
                    y2 = Kut.ax * x2 * x2 + Kut.bx * x2 + Kut.cx
                    Label5.Text = Label5.Text & vbCrLf & "T1 (" & FormatNumber(x1, 2) & " , " & FormatNumber(y1, 2) & ")"
                    Label5.Text = Label5.Text & vbCrLf & "T2 (" & FormatNumber(x2, 2) & " , " & FormatNumber(y2, 2) & ")" & vbCrLf
                    koef1 = 2 * Kut.ax * x1 + Kut.bx
                    koef2 = 2 * Kut.ax1 * x1 + Kut.bx1
                    kutp = Atan(Abs((koef2 - koef1) / (1 + koef1 * koef2)))
                    Label5.Text = Label5.Text & "Kut presjeka (x=" & FormatNumber(x1, 2) & "): " & Form1.Kut_Ispis(kutp, False) & vbCrLf

                    koef1 = 2 * Kut.ax * x2 + Kut.bx
                    koef2 = 2 * Kut.ax1 * x2 + Kut.bx1
                    kutp = Atan(Abs((koef2 - koef1) / (1 + koef1 * koef2)))
                    Label5.Text = Label5.Text & "Kut presjeka (x=" & FormatNumber(x2, 2) & "): " & Form1.Kut_Ispis(kutp, False) & vbCrLf
                    If x1 <> x2 Then
                        Dim povrsina As Double
                        povrsina = Abs((x2 * x2 * x2 * ((Kut.ax - Kut.ax1) / 3) + x2 * x2 * ((Kut.bx - Kut.bx1) / 2) + x2 * (Kut.cx - Kut.cx1)) - (x1 * x1 * x1 * ((Kut.ax - Kut.ax1) / 3) + x1 * x1 * ((Kut.bx - Kut.bx1) / 2) + x1 * (Kut.cx - Kut.cx1)))
                        Label5.Text = Label5.Text & "Povrsina izmedu krivulja: " & FormatNumber(povrsina, 2)
                    End If
                Else
                    Label5.Text = Label5.Text & vbCrLf & "-"
                End If
            ElseIf Abs(b) > 0 Then
                x1 = -c / b
                y1 = Kut.bx * x1 + Kut.cx
                Label5.Text = Label5.Text & vbCrLf & "T1 (" & FormatNumber(x1, 2) & " , " & FormatNumber(y1, 2) & ")" & vbCrLf
                koef1 = 2 * Kut.ax * x1 + Kut.bx
                koef2 = 2 * Kut.ax1 * x1 + Kut.bx1
                kutp = Atan(Abs((koef2 - koef1) / (1 + koef1 * koef2)))
                Label5.Text = Label5.Text & "Kut presjeka: " & Form1.Kut_Ispis(kutp, False)
            Else
                Label5.Text = Label5.Text & vbCrLf & "-"
            End If



            Label5.Visible = True

            For i = 0 To 900
                x = (Kut.mj * (i - 450)) / 250
                If Abs(koef * (Kut.ax1 * x * x + Kut.bx1 * x + Kut.cx1)) < 425 Then
                    tocke1(i).Y = 425 - Round(koef * (Kut.ax1 * x * x + Kut.bx1 * x + Kut.cx1))
                    k = True
                Else
                    If ax1 < 0 Then
                        tocke1(i).Y = 850
                    ElseIf ax1 > 0 Then
                        tocke1(i).Y = -1
                    ElseIf bx1 > 0 And k = False Then
                        tocke1(i).Y = 850
                    ElseIf bx1 > 0 And k = True Then
                        tocke1(i).Y = -1
                    ElseIf bx1 < 0 And k = False Then
                        tocke1(i).Y = -1
                    ElseIf bx1 < 0 And k = True Then
                        tocke1(i).Y = 850
                    End If
                End If
                tocke1(i).X = i

            Next i
        End If


        Label1.Text = Kut.tekst
        Label3.Text = "-" & Kut.mj
        Label4.Text = Kut.mj

    End Sub

    Private Sub Form3_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        T.X = Control.MousePosition.X.ToString() - Me.Bounds.Location.X - 8
        T.Y = Control.MousePosition.Y.ToString() - Me.Bounds.Location.Y - 31
        If T.Y > 5 And T.Y < 800 Then poruka()
    End Sub

    Private Sub Form3_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        Dim Osi As New Pen(Color.Black, 2)
        Dim krivulja As New Pen(Color.Black, 1)

        e.Graphics.DrawLine(Osi, 0, 425, 900, 425)
        e.Graphics.DrawLine(Osi, 450, 0, 450, 850)

        For i = 0 To 900 Step 50
            e.Graphics.DrawLine(Osi, i, 420, i, 430)
        Next
        For i = 25 To 850 Step 50
            e.Graphics.DrawLine(Osi, 445, i, 455, i)
        Next

        e.Graphics.DrawCurve(krivulja, tocke)
        If Abs(ax1) > 0 Or Abs(bx1) > 0 Or Abs(cx1) > 0 Then
            e.Graphics.DrawCurve(krivulja, tocke1)
        End If

    End Sub

    Sub poruka()
        Dim d As Double
        Dim x, y As Integer
        Dim lab As Point
        For i = 0 To 900
            x = T.X - tocke(i).X
            y = T.Y - tocke(i).Y
            d = Sqrt(x * x + y * y)
            If d < 7 Then
                lab.X = tocke(i).X + Me.Bounds.Location.X + 8
                lab.Y = tocke(i).Y + Me.Bounds.Location.Y + 31
                Cursor.Position = lab
                lab.X = tocke(i).X + 8
                lab.Y = tocke(i).Y
                Label6.Location = lab
                Label6.Text = "X = " & ((tocke(i).X - 450) / 250) * Kut.mj & vbCrLf & "Y = " & ((425 - tocke(i).Y) / 250) * Kut.mj
                GoTo kraj
            End If
        Next
        If Kut.ax1 > 0 Or Kut.bx1 > 0 Or Kut.cx1 > 0 Then
            For i = 0 To 900
                x = T.X - tocke1(i).X
                y = T.Y - tocke1(i).Y
                d = Sqrt(x * x + y * y)
                If d < 7 Then
                    lab.X = tocke1(i).X + Me.Bounds.Location.X + 8
                    lab.Y = tocke1(i).Y + Me.Bounds.Location.Y + 31
                    Cursor.Position = lab
                    lab.X = tocke1(i).X + 8
                    lab.Y = tocke1(i).Y
                    Label6.Location = lab
                    Label6.Text = "X = " & ((tocke1(i).X - 450) / 250) * Kut.mj & vbCrLf & "Y = " & ((425 - tocke1(i).Y) / 250) * Kut.mj
                    GoTo kraj
                End If
            Next
        End If
        Label6.Text = ""
kraj:
    End Sub
End Class