
' Math Help by Antonio
' v1.0 15.5.2016.
' Redaka: 6434

Imports System.Math
Public Class Form1

    Dim pStupanj As Boolean = False
    Dim Ispravan_unos As Boolean = True
    Dim d_u As Boolean = False
    Dim Sustav As Boolean = False
    Dim uneseno As Integer = 0
    Dim promjena As Boolean = False
    Dim susjed As Boolean = False
    Dim Pt As Boolean

    Dim rispis As String
    Dim gispis As String
    Dim pispis As String

    Public activeTextBox As TextBox
    Dim strHelpPath As String = System.IO.Path.Combine(Application.StartupPath, "Mathematics.chm")

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        HelpProvider1.HelpNamespace = strHelpPath
    End Sub

    Private Sub Form1_Show(sender As Object, e As EventArgs) Handles MyBase.Shown
        TextBox1.Focus()
        TextBox1.SelectionStart = TextBox1.TextLength
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub Tab_change(sender As Object, e As EventArgs) Handles TabControl2.SelectedIndexChanged, TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            If TabControl2.SelectedIndex = 0 Then
                TextBox1.Focus()
                TextBox1.SelectionStart = TextBox1.TextLength
                TextBox1.ScrollToCaret()
            ElseIf TabControl2.SelectedIndex = 1 Then
                TextBox4.Focus()
                TextBox4.SelectionStart = TextBox4.TextLength
                TextBox4.ScrollToCaret()
            ElseIf TabControl2.SelectedIndex = 2 Then
                TextBox5.Focus()
                TextBox5.SelectionStart = TextBox5.TextLength
                TextBox5.ScrollToCaret()
            End If
        ElseIf TabControl1.SelectedIndex = 3 Then
            TextBox7.Focus()
            TextBox7.SelectionStart = TextBox7.TextLength
            TextBox7.ScrollToCaret()
        End If
    End Sub

    Private Sub HelpF1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpF1ToolStripMenuItem.Click
        Help.ShowHelp(Me, HelpProvider1.HelpNamespace, HelpNavigator.TableOfContents)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim D, a, b, c, pom As Double

        If Sustav = False Or Ispravan_unos = False Or Label10.Text = "Rezultat:" Or uneseno = 2 Then
            Label10.Text = Nothing
            uneseno = 0
            d_u = False
            Kut.ax1 = 0
            Kut.bx1 = 0
            Kut.cx1 = 0
        End If

        If (IsNumeric(TextBox1.Text) = False And TextBox1.Text <> "") Or (IsNumeric(TextBox2.Text) = False And TextBox2.Text <> "") Or (IsNumeric(TextBox3.Text) = False And TextBox3.Text <> "") Or (TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "") Then
            Label10.Text = Label10.Text & "Pogresan unos !" & vbCrLf & "Pokusajte ponovo"
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            Button2.Visible = False
            Ispravan_unos = False
        Else
            Ispravan_unos = True
            If TextBox1.Text <> "" Then
                a = TextBox1.Text
            Else
                a = 0
            End If
            If TextBox2.Text <> "" Then
                b = TextBox2.Text
            Else
                b = 0
            End If
            If TextBox3.Text <> "" Then
                c = TextBox3.Text
            Else
                c = 0
            End If

            If a = 0 And Abs(b) > 0 Then
                Dim x As String = Nothing
                If b > 0 Then x = "rastući."
                If b < 0 Then x = "padajući."

                Label10.Text = Label10.Text & Graf_Ispis(a, b, c) & vbCrLf & "Pravac je " & x & vbCrLf & "Nul-točka: ( " & FormatNumber(-c / b, 2) & " , 0 )" & vbCrLf & "Odsjecak na Y osi: " & c

                If Abs(-c / b) < Abs(c) Then
                    pom = Abs(c)
                Else
                    pom = Abs(-c / b)
                End If

                Button2.Visible = True

            ElseIf a = 0 And b = 0 Then
                pom = Abs(c)
                Label10.Text = Label10.Text & Graf_Ispis(a, b, c) & vbCrLf & "Pravac je konstantan" & vbCrLf & "Odsjecak na Y osi: " & c
                Button2.Visible = True
            ElseIf Abs(a) > 0 Then
                D = b * b - 4 * a * c
                If D < 0 Then
                    Label10.Text = Label10.Text & Graf_Ispis(a, b, c) & vbCrLf & "Ova funkcija nema realnih rjesenja !" & vbCrLf & "Tjeme parabole: ( " & FormatNumber(-b / (2 * a), 2) & " , " & FormatNumber(-D / (4 * a), 2) & " )"

                    If Abs(-b / (2 * a)) < Abs(-D / (4 * a)) Then
                        pom = Abs(-D / (4 * a))
                    Else
                        pom = Abs(-b / (2 * a))
                    End If

                    Button2.Visible = True

                ElseIf D = 0 Then

                    Label10.Text = Label10.Text & Graf_Ispis(a, b, c) & vbCrLf & "Nul-točka: ( " & FormatNumber(-b / (2 * a), 2) & " , 0 )" & vbCrLf & "Tjeme parabole: ( " & FormatNumber(-b / (2 * a), 2) & " , " & FormatNumber(-D / (4 * a), 2) & " )"

                    D = -4 * a * a / 2

                    If Abs(-b / (2 * a)) < Abs(-D / (4 * a)) Then
                        pom = Abs(-D / (4 * a))
                    Else
                        pom = Abs(-b / (2 * a))
                    End If

                    Button2.Visible = True

                ElseIf D > 0 Then

                    Label10.Text = Label10.Text & Graf_Ispis(a, b, c) & vbCrLf & "Nul-točke: ( " & FormatNumber((-b + Sqrt(D)) / (2 * a), 2) & " , 0 ), ( " & FormatNumber((-b - Sqrt(D)) / (2 * a), 2) & " , 0)" & vbCrLf & "Tjeme parabole: ( " & FormatNumber(-b / (2 * a), 2) & " , " & FormatNumber(-D / (4 * a), 2) & " )"

                    If Abs((-b + Sqrt(D)) / (2 * a)) < Abs((-b - Sqrt(D)) / (2 * a)) Then
                        pom = Abs((-b - Sqrt(D)) / (2 * a))
                    Else
                        pom = Abs((-b + Sqrt(D)) / (2 * a))
                    End If
                    If pom < Abs(-D / (4 * a)) Then pom = Abs(-D / (4 * a))

                    Button2.Visible = True

                End If

            End If

            Label10.Text = Label10.Text & vbCrLf & vbCrLf

            If d_u = True Then
                Kut.tekst = Kut.tekst & " i " & Graf_Ispis(a, b, c)
                pom = Round(pom)
                If pom = 0 Then pom = 1
                If Kut.mj < pom Then Kut.mj = pom
                Kut.ax1 = a
                Kut.bx1 = b
                Kut.cx1 = c
            Else
                Kut.tekst = Graf_Ispis(a, b, c)
                pom = Round(pom)
                If pom = 0 Then pom = 1
                Kut.mj = pom
                Kut.ax = a
                Kut.bx = b
                Kut.cx = c
            End If
            uneseno = uneseno + 1
            If d_u = False Then
                d_u = True
            Else
                d_u = False
            End If

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox1.Focus()
            TextBox1.SelectionStart = TextBox1.TextLength
            TextBox1.ScrollToCaret()

        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form3.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        Kut.ax = 0
        Kut.bx = 0
        Kut.cx = 0
        Kut.ax1 = 0
        Kut.bx1 = 0
        Kut.cx1 = 0
        Label10.Text = "Rezultat:"
        Button2.Visible = False
        TextBox1.Focus()
        TextBox1.SelectionStart = TextBox1.TextLength
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim p_l As New List(Of Double)
        Dim n_l As New List(Of Double)
        Dim p_d As New List(Of Double)
        Dim n_d As New List(Of Double)
        Dim kn_d As New List(Of Double)
        Dim kn_l As New List(Of Double)
        Dim predznak As Integer = 0
        Dim pb As Integer = -1
        Dim zb As Integer = -1
        Dim broj As Double
        Dim strana As Integer = 0
        Dim izracunaj As Boolean = False
        Dim jed As String
        Dim pom As Boolean = False
        Dim kon As Integer = 0
        Dim rj1, rj2 As Double

        ListBox1.Items.Clear()
        Label12.Text = "Rjesenja:"

        jed = TextBox4.Text

        Dim kz As Integer = 0
        Dim kz1 As Integer = 0
        Dim kz2 As Integer = 0
        Dim bj As Integer = 0
        If jed.Length = 0 Then
            ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
            GoTo greska
        End If
        For i = 0 To jed.Length - 1

            If IsNumeric(jed(i)) = False And jed(i) <> "(" And jed(i) <> ")" And jed(i) <> "x" And jed(i) <> "K" And jed(i) <> "+" And jed(i) <> "-" And jed(i) <> "/" And jed(i) <> "*" And jed(i) <> "=" And jed(i) <> "," Then
                ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                GoTo greska
            End If
                If i = 0 Then
                    If IsNumeric(jed(i)) = False And jed(i) <> "+" And jed(i) <> "-" Then
                        ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
                If i > 0 Then
                If (IsNumeric(jed(i - 1)) = False And IsNumeric(jed(i)) = False) And ((jed(i - 1) = "(" And jed(i) = "-") Or (jed(i - 1) = "(" And jed(i) = "+") Or (jed(i - 1) = "x" And jed(i) = ")") Or (jed(i - 1) = "x" And jed(i) = "+") Or (jed(i - 1) = "x" And jed(i) = "-") Or (jed(i - 1) = ")" And jed(i) = ")") Or (jed(i - 1) = ")" And jed(i) = "=") Or (jed(i - 1) = ")" And jed(i) = "+") Or (jed(i - 1) = ")" And jed(i) = "-") Or (jed(i - 1) = "x" And jed(i) = "(") Or (jed(i - 1) = "=" And jed(i) = "-") Or (jed(i - 1) = "=" And jed(i) = "+") Or (jed(i - 1) = "x" And jed(i) = "=") Or (jed(i - 1) = "K" And jed(i) = "=") Or (jed(i - 1) = "K" And jed(i) = "+") Or (jed(i - 1) = "K" And jed(i) = "(") Or (jed(i - 1) = "K" And jed(i) = "-") Or (jed(i - 1) = "K" And jed(i) = ")")) = False Then
                    ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
                End If
                If jed(i) = "(" Then
                    kz = kz + 1
                kz1 = kz1 + 1
                If i < jed.Length - 1 Then
                    If IsNumeric(jed(i + 1)) Then
                        ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
                End If
                If jed(i) = ")" Then
                    kz = kz - 1
                kz2 = kz2 + 1
                If i < jed.Length - 1 Then
                    If IsNumeric(jed(i + 1)) Then
                        ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
                End If
                If jed(i) = ")" And kz < 0 Then
                    ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
                If jed(i) = "=" Then
                bj = bj + 1

                If bj > 1 Or kz <> 0 Then
                    ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed(i) = "=" And i < jed.Length - 2 Then
                If IsNumeric(jed(i + 1)) = True Then
                    ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If i = jed.Length - 1 Then
                If (kz1 <> kz2) Or bj = 0 Then
                    ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed(i) = "x" Or jed(i) = "K" Then
                If i < jed.Length - 1 Then
                    If IsNumeric(jed(i + 1)) Then
                        ListBox1.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If

        Next


        ListBox1.Items.Add(jed)

        jed = dijeli(jed)
        jed = mnozi(jed)

        kz = 0
        kz1 = 0
        kz2 = 0
        bj = 0
        For i = 0 To jed.Length - 1

            If IsNumeric(jed(i)) = False And jed(i) <> "(" And jed(i) <> ")" And jed(i) <> "x" And jed(i) <> "K" And jed(i) <> "+" And jed(i) <> "-" And jed(i) <> "/" And jed(i) <> "*" And jed(i) <> "=" And jed(i) <> "," Then
                ListBox1.Items.Clear()
                ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                GoTo greska
            End If
            If i = 0 Then
                If IsNumeric(jed(i)) = False And jed(i) <> "+" And jed(i) <> "-" Then
                    ListBox1.Items.Clear()
                    ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If i > 0 Then
                If (IsNumeric(jed(i - 1)) = False And IsNumeric(jed(i)) = False) And ((jed(i - 1) = "(" And jed(i) = "-") Or (jed(i - 1) = "(" And jed(i) = "+") Or (jed(i - 1) = "x" And jed(i) = ")") Or (jed(i - 1) = "x" And jed(i) = "+") Or (jed(i - 1) = "x" And jed(i) = "-") Or (jed(i - 1) = ")" And jed(i) = ")") Or (jed(i - 1) = ")" And jed(i) = "=") Or (jed(i - 1) = ")" And jed(i) = "+") Or (jed(i - 1) = ")" And jed(i) = "-") Or (jed(i - 1) = "x" And jed(i) = "(") Or (jed(i - 1) = "=" And jed(i) = "-") Or (jed(i - 1) = "=" And jed(i) = "+") Or (jed(i - 1) = "x" And jed(i) = "=") Or (jed(i - 1) = "K" And jed(i) = "=") Or (jed(i - 1) = "K" And jed(i) = "+") Or (jed(i - 1) = "K" And jed(i) = "(") Or (jed(i - 1) = "K" And jed(i) = "-") Or (jed(i - 1) = "K" And jed(i) = ")")) = False Then
                    ListBox1.Items.Clear()
                    ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If jed(i) = "(" Then
                kz = kz + 1
                kz1 = kz1 + 1
                If i < jed.Length - 1 Then
                    If IsNumeric(jed(i + 1)) Then
                        ListBox1.Items.Clear()
                        ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed(i) = ")" Then
                kz = kz - 1
                kz2 = kz2 + 1
                If i < jed.Length - 1 Then
                    If IsNumeric(jed(i + 1)) Then
                        ListBox1.Items.Clear()
                        ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed(i) = ")" And kz < 0 Then
                ListBox1.Items.Clear()
                ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                GoTo greska
            End If
            If jed(i) = "=" Then
                bj = bj + 1

                If bj > 1 Or kz <> 0 Then
                    ListBox1.Items.Clear()
                    ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed(i) = "=" And i < jed.Length - 2 Then
                If IsNumeric(jed(i + 1)) = True Then
                    ListBox1.Items.Clear()
                    ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If i = jed.Length - 1 Then
                If (kz1 <> kz2) Or bj = 0 Then
                    ListBox1.Items.Clear()
                    ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed(i) = "x" Or jed(i) = "K" Then
                If i < jed.Length - 1 Then
                    If IsNumeric(jed(i + 1)) Then
                        ListBox1.Items.Clear()
                        ListBox1.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If

        Next

        Do
            For i = 0 To jed.Length - 1
                If jed(i) = "(" Then
                    pb = i
                    pom = True
                End If

                If pb > -1 Then
                    For j = pb To 0 Step -1
                        If izracunaj = False And (jed(j) = "+" Or jed(j) = "-" Or jed(j) = "=" Or j = 0) Then
                            If jed(j) = "=" Then
                                pb = j + 1
                            Else : pb = j
                            End If
                            izracunaj = True
                        End If
                    Next
                    izracunaj = False
                    For j = pb To jed.Length - 1
                        If izracunaj = False And jed(j) = ")" Then
                            zb = j
                            izracunaj = True
                        End If
                    Next
                End If
                izracunaj = False
            Next

            If pb > -1 Then
                jed = zamjeni(jed, zagrada(jed, pb, zb), pb, zb)
                If pStupanj = True Then
                    ListBox1.Items.Add("Pogreska!")
                    ListBox1.Items.Add("Unosite jednadzbe najvise drugog reda")
                    GoTo greska
                End If
                ListBox1.Items.Add(jed)
            End If

            pb = -1
            zb = -1
            pom = False
            For i = 0 To jed.Length - 1
                If jed(i) = "(" Then
                    pom = True
                End If
            Next

        Loop While pom = True

        izracunaj = False

        For i = 0 To jed.Length - 1
            If i = jed.Length - 1 Then
                izracunaj = True
            End If
            If i > 0 Then
                If jed(i) = "=" And (jed(i - 1) = "x" Or jed(i - 1) = "K") Then
                    kon = 0
                    strana = 1
                    GoTo lab
                End If
            End If

            If IsNumeric(jed(i)) Or jed(i) = "," Or ((jed(i) = "+" Or jed(i) = "-") And zb = -1) Then
                If pb = -1 Then pb = i
                zb = i
                If izracunaj = False Then
                    GoTo lab
                End If
            End If
            If izracunaj = True Or (IsNumeric(jed(i)) Or jed(i) = "," Or ((jed(i) = "+" Or jed(i) = "-") And zb = -1)) = False Then
                If kon = 0 Then
                    Double.TryParse(izdvoji_niz(jed, pb, zb), broj)
                    kon = 1
                ElseIf pb > 0 Then
                    If jed(pb - 1) = "x" Or jed(pb - 1) = "K" Then
                        Double.TryParse(izdvoji_niz(jed, pb, zb), broj)
                    Else
                        Double.TryParse(izdvoji_niz(jed, pb - 1, zb), broj)
                    End If

                End If

                If jed(i) = "x" Then
                    If strana = 0 Then n_l.Add(broj)
                    If strana = 1 Then n_d.Add(broj)
                ElseIf jed(i) = "K" Then
                    If strana = 0 Then kn_l.Add(broj)
                    If strana = 1 Then kn_d.Add(broj)
                ElseIf jed(i) = "+" Or jed(i) = "-" Or jed(i) = "=" Or izracunaj = True Then
                    If strana = 0 Then p_l.Add(broj)
                    If strana = 1 Then p_d.Add(broj)
                End If
                If jed(i) = "=" Then
                    strana = 1
                    kon = 0
                End If
                pb = -1
                zb = -1

            End If
lab:
        Next

        Dim pozl() As Double = p_l.ToArray()
        Dim pozd() As Double = p_d.ToArray()
        Dim nepl() As Double = n_l.ToArray()
        Dim nepd() As Double = n_d.ToArray()
        Dim knepl() As Double = kn_l.ToArray()
        Dim knepd() As Double = kn_d.ToArray()

        If knepl.Length >= 1 Or knepd.Length >= 1 Then
            Dim txt As String = Nothing
            Dim a, b, c As Double
            a = 0
            b = 0
            c = 0

            For Each i As Double In knepl
                txt = txt & pz(i) & Abs(i) & "K"
                a = a + i
            Next
            For Each i As Double In knepd
                txt = txt & pz(-i) & Abs(i) & "K"
                a = a - i
            Next
            For Each i As Double In nepl
                txt = txt & pz(i) & Abs(i) & "x"
                b = b + i
            Next
            For Each i As Double In nepd
                txt = txt & pz(-i) & Abs(i) & "x"
                b = b - i
            Next
            For Each i As Double In pozl
                txt = txt & pz(i) & Abs(i)
                c = c + i
            Next
            For Each i As Double In pozd
                txt = txt & pz(-i) & Abs(i)
                c = c - i
            Next
            txt = txt & "=0"
            ListBox1.Items.Add(txt)
            txt = pz(a) & Abs(a) & "K" & pz(b) & Abs(b) & "x" & pz(c) & Abs(c) & "=0"
            ListBox1.Items.Add(txt)
            If Abs(a) > 0 Then
                txt = "x = ( -" & b & "+/- sqrt(" & b & "*" & b & "-4*" & a & "*" & c & ") ) / 2*" & a
                ListBox1.Items.Add(txt)
                txt = "x = ( -" & b & "+/- sqrt(" & b * b - 4 * a * c & ") ) / " & 2 * a
                ListBox1.Items.Add(txt)
                If (b * b - 4 * a * c) >= 0 Then
                    ListBox1.Items.Add("x1 = " & FormatNumber((-b + Sqrt(b * b - 4 * a * c)) / (2 * a), 5) & ", x2 = " & FormatNumber((-b - Sqrt(b * b - 4 * a * c)) / (2 * a), 5))
                    rj1 = (-b + Sqrt(b * b - 4 * a * c)) / (2 * a)
                    rj2 = (-b - Sqrt(b * b - 4 * a * c)) / (2 * a)
                    Label12.Text = Label12.Text & vbCrLf & "x1 = " & FormatNumber(rj1, 4) & "," & vbCrLf & "x2 = " & FormatNumber(rj2, 4)
                Else
                    ListBox1.Items.Add("x1,x2 = prazan skup")
                    rj1 = Nothing
                    rj2 = Nothing
                End If
            Else
                txt = pz(b) & Abs(b) & "x=" & pz(-c) & Abs(c)
                ListBox1.Items.Add(txt)
                If b = 0 Then
                    If c = 0 Then
                        txt = "x = R"
                        Label12.Text = Label12.Text & vbCrLf & "x = R"
                    Else
                        txt = "x = Prazan skup"
                    End If
                    rj1 = Nothing
                    rj2 = Nothing
                Else
                    txt = "x = " & -c & "/" & b
                    ListBox1.Items.Add(txt)
                    txt = "x = " & FormatNumber(-c / b, 5)
                    rj1 = -c / b
                    rj2 = -c / b
                    Label12.Text = Label12.Text & vbCrLf & "x = " & FormatNumber(rj1, 4)
                End If
                ListBox1.Items.Add(txt)
            End If
        Else
            Dim txt As String = Nothing
            Dim b, c As Double

            b = 0
            c = 0

            For Each i As Double In nepl
                txt = txt & pz(i) & Abs(i) & "x"
                b = b + i
            Next
            For Each i As Double In nepd
                txt = txt & pz(-i) & Abs(i) & "x"
                b = b - i
            Next
            txt = txt & "="
            For Each i As Double In pozl
                txt = txt & pz(-i) & Abs(i)
                c = c - i
            Next
            For Each i As Double In pozd
                txt = txt & pz(i) & Abs(i)
                c = c + i
            Next
            If pozl.Length = 0 Or pozd.Length = 0 Then txt = txt & 0
            ListBox1.Items.Add(txt)

            txt = pz(b) & Abs(b) & "x=" & pz(c) & Abs(c)
            ListBox1.Items.Add(txt)

            If b = 0 Then
                If c = 0 Then
                    txt = "x = R"
                    Label12.Text = Label12.Text & vbCrLf & "x = R"
                Else
                    txt = "x = Prazan skup"
                End If
                rj1 = Nothing
                rj2 = Nothing
            Else
                txt = "x = " & c & "/" & b
                ListBox1.Items.Add(txt)
                txt = "x = " & FormatNumber(c / b, 5)
                rj1 = c / b
                rj2 = c / b
                Label12.Text = Label12.Text & vbCrLf & "x = " & FormatNumber(rj1, 4)
            End If
            ListBox1.Items.Add(txt)

        End If

greska:
        pStupanj = False
    End Sub


    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Dim p_l As New List(Of Double)
        Dim n_l As New List(Of Double)
        Dim p_d As New List(Of Double)
        Dim n_d As New List(Of Double)
        Dim kn_d As New List(Of Double)
        Dim kn_l As New List(Of Double)
        Dim y_d As New List(Of Double)
        Dim y_l As New List(Of Double)

        Dim predznak As Integer = 0
        Dim pb As Integer = -1
        Dim zb As Integer = -1
        Dim broj As Double
        Dim strana As Integer = 0
        Dim izracunaj As Boolean = False
        Dim jed1, jed2 As String
        Dim pom As Boolean = False
        Dim kon As Integer = 0
        Dim rj1, rj2, ry1, ry2 As Double

        ListBox2.Items.Clear()
        Label21.Text = "Rjesenja:"
        susjed = True
        promjena = False

        jed1 = TextBox5.Text
        jed2 = TextBox6.Text

        Dim kz As Integer = 0
        Dim kz1 As Integer = 0
        Dim kz2 As Integer = 0
        Dim bj As Integer = 0

        If jed1.Length = 0 Then
            ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
            GoTo greska
        End If
        For i = 0 To jed1.Length - 1

            If IsNumeric(jed1(i)) = False And jed1(i) <> "(" And jed1(i) <> ")" And jed1(i) <> "x" And jed1(i) <> "K" And jed1(i) <> "+" And jed1(i) <> "-" And jed1(i) <> "/" And jed1(i) <> "*" And jed1(i) <> "=" And jed1(i) <> "y" And jed1(i) <> "," Then
                ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                GoTo greska
            End If
            If i = 0 Then
                If IsNumeric(jed1(i)) = False And jed1(i) <> "+" And jed1(i) <> "-" Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If i > 0 Then
                If (IsNumeric(jed1(i - 1)) = False And IsNumeric(jed1(i)) = False) And ((jed1(i - 1) = "(" And jed1(i) = "-") Or (jed1(i - 1) = "(" And jed1(i) = "+") Or (jed1(i - 1) = "x" And jed1(i) = ")") Or (jed1(i - 1) = "x" And jed1(i) = "+") Or (jed1(i - 1) = "x" And jed1(i) = "-") Or (jed1(i - 1) = ")" And jed1(i) = ")") Or (jed1(i - 1) = ")" And jed1(i) = "=") Or (jed1(i - 1) = ")" And jed1(i) = "+") Or (jed1(i - 1) = ")" And jed1(i) = "-") Or (jed1(i - 1) = "x" And jed1(i) = "(") Or (jed1(i - 1) = "=" And jed1(i) = "-") Or (jed1(i - 1) = "=" And jed1(i) = "+") Or (jed1(i - 1) = "x" And jed1(i) = "=") Or (jed1(i - 1) = "K" And jed1(i) = "=") Or (jed1(i - 1) = "K" And jed1(i) = "+") Or (jed1(i - 1) = "K" And jed1(i) = "(") Or (jed1(i - 1) = "K" And jed1(i) = "-") Or (jed1(i - 1) = "K" And jed1(i) = ")") Or (jed1(i - 1) = "y" And jed1(i) = "=") Or (jed1(i - 1) = "y" And jed1(i) = "+") Or (jed1(i - 1) = "y" And jed1(i) = "(") Or (jed1(i - 1) = "y" And jed1(i) = "-") Or (jed1(i - 1) = "y" And jed1(i) = ")")) = False Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If jed1(i) = "(" Then
                kz = kz + 1
                kz1 = kz1 + 1
                If i < jed1.Length - 1 Then
                    If IsNumeric(jed1(i + 1)) Then
                        ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed1(i) = ")" Then
                kz = kz - 1
                kz2 = kz2 + 1
                If i < jed1.Length - 1 Then
                    If IsNumeric(jed1(i + 1)) Then
                        ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed1(i) = ")" And kz < 0 Then
                ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                GoTo greska
            End If
            If jed1(i) = "=" Then
                bj = bj + 1

                If bj > 1 Or kz <> 0 Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If i = jed1.Length - 1 Then
                If (kz1 <> kz2) Or bj = 0 Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed1(i) = "=" And i < jed1.Length - 2 Then
                If IsNumeric(jed1(i + 1)) = True Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed1(i) = "x" Or jed1(i) = "K" Or jed1(i) = "y" Then
                If i < jed1.Length - 1 Then
                    If IsNumeric(jed1(i + 1)) Then
                        ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
        Next

        'provjera 2 jednadzbe
        kz = 0
        kz1 = 0
        kz2 = 0
        bj = 0

        If jed2.Length = 0 Then
            ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
            GoTo greska
        End If
        For i = 0 To jed2.Length - 1

            If IsNumeric(jed2(i)) = False And jed2(i) <> "(" And jed2(i) <> ")" And jed2(i) <> "x" And jed2(i) <> "K" And jed2(i) <> "+" And jed2(i) <> "-" And jed2(i) <> "/" And jed2(i) <> "*" And jed2(i) <> "=" And jed2(i) <> "y" And jed2(i) <> "," Then
                ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                GoTo greska
            End If
            If i = 0 Then
                If IsNumeric(jed2(i)) = False And jed2(i) <> "+" And jed2(i) <> "-" Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If i > 0 Then
                If (IsNumeric(jed2(i - 1)) = False And IsNumeric(jed2(i)) = False) And ((jed2(i - 1) = "(" And jed2(i) = "-") Or (jed2(i - 1) = "(" And jed2(i) = "+") Or (jed2(i - 1) = "x" And jed2(i) = ")") Or (jed2(i - 1) = "x" And jed2(i) = "+") Or (jed2(i - 1) = "x" And jed2(i) = "-") Or (jed2(i - 1) = ")" And jed2(i) = ")") Or (jed2(i - 1) = ")" And jed2(i) = "=") Or (jed2(i - 1) = ")" And jed2(i) = "+") Or (jed2(i - 1) = ")" And jed2(i) = "-") Or (jed2(i - 1) = "x" And jed2(i) = "(") Or (jed2(i - 1) = "=" And jed2(i) = "-") Or (jed2(i - 1) = "=" And jed2(i) = "+") Or (jed2(i - 1) = "x" And jed2(i) = "=") Or (jed2(i - 1) = "K" And jed2(i) = "=") Or (jed2(i - 1) = "K" And jed2(i) = "+") Or (jed2(i - 1) = "K" And jed2(i) = "(") Or (jed2(i - 1) = "K" And jed2(i) = "-") Or (jed2(i - 1) = "K" And jed2(i) = ")") Or (jed2(i - 1) = "y" And jed2(i) = "=") Or (jed2(i - 1) = "y" And jed2(i) = "+") Or (jed2(i - 1) = "y" And jed2(i) = "(") Or (jed2(i - 1) = "y" And jed2(i) = "-") Or (jed2(i - 1) = "y" And jed2(i) = ")")) = False Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If jed2(i) = "(" Then
                kz = kz + 1
                kz1 = kz1 + 1
                If i < jed2.Length - 1 Then
                    If IsNumeric(jed2(i + 1)) Then
                        ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed2(i) = ")" Then
                kz = kz - 1
                kz2 = kz2 + 1
                If i < jed2.Length - 1 Then
                    If IsNumeric(jed2(i + 1)) Then
                        ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed2(i) = ")" And kz < 0 Then
                ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                GoTo greska
            End If
            If jed2(i) = "=" Then
                bj = bj + 1

                If bj > 1 Or kz <> 0 Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If i = jed2.Length - 1 Then
                If (kz1 <> kz2) Or bj = 0 Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed2(i) = "=" And i < jed2.Length - 2 Then
                If IsNumeric(jed2(i + 1)) = True Then
                    ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed2(i) = "x" Or jed2(i) = "K" Or jed2(i) = "y" Then
                If i < jed2.Length - 1 Then
                    If IsNumeric(jed2(i + 1)) Then
                        ListBox2.Items.Add("Pogresan unos, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
        Next


        ListBox2.Items.Add(jed1)
        ListBox2.Items.Add(jed2)
        ListBox2.Items.Add("--------------------")

        jed1 = dijeli(jed1)
        jed1 = mnozi(jed1)
        If promjena = False Then
            ListBox2.Items.Add(jed1)
        End If
        promjena = False
        ListBox2.Items.Add(" ")
        jed2 = dijeli(jed2)
        jed2 = mnozi(jed2)
        If promjena = False Then
            ListBox2.Items.Add(jed2)
        End If
        ListBox2.Items.Add("--------------------")
        promjena = False

        kz = 0
        kz1 = 0
        kz2 = 0
        bj = 0

        For i = 0 To jed1.Length - 1

            If IsNumeric(jed1(i)) = False And jed1(i) <> "(" And jed1(i) <> ")" And jed1(i) <> "x" And jed1(i) <> "K" And jed1(i) <> "+" And jed1(i) <> "-" And jed1(i) <> "/" And jed1(i) <> "*" And jed1(i) <> "=" And jed1(i) <> "y" And jed1(i) <> "," Then
                ListBox2.Items.Clear()
                ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                GoTo greska
            End If
            If i = 0 Then
                If IsNumeric(jed1(i)) = False And jed1(i) <> "+" And jed1(i) <> "-" Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If i > 0 Then
                If (IsNumeric(jed1(i - 1)) = False And IsNumeric(jed1(i)) = False) And ((jed1(i - 1) = "(" And jed1(i) = "-") Or (jed1(i - 1) = "(" And jed1(i) = "+") Or (jed1(i - 1) = "x" And jed1(i) = ")") Or (jed1(i - 1) = "x" And jed1(i) = "+") Or (jed1(i - 1) = "x" And jed1(i) = "-") Or (jed1(i - 1) = ")" And jed1(i) = ")") Or (jed1(i - 1) = ")" And jed1(i) = "=") Or (jed1(i - 1) = ")" And jed1(i) = "+") Or (jed1(i - 1) = ")" And jed1(i) = "-") Or (jed1(i - 1) = "x" And jed1(i) = "(") Or (jed1(i - 1) = "=" And jed1(i) = "-") Or (jed1(i - 1) = "=" And jed1(i) = "+") Or (jed1(i - 1) = "x" And jed1(i) = "=") Or (jed1(i - 1) = "K" And jed1(i) = "=") Or (jed1(i - 1) = "K" And jed1(i) = "+") Or (jed1(i - 1) = "K" And jed1(i) = "(") Or (jed1(i - 1) = "K" And jed1(i) = "-") Or (jed1(i - 1) = "K" And jed1(i) = ")") Or (jed1(i - 1) = "y" And jed1(i) = "=") Or (jed1(i - 1) = "y" And jed1(i) = "+") Or (jed1(i - 1) = "y" And jed1(i) = "(") Or (jed1(i - 1) = "y" And jed1(i) = "-") Or (jed1(i - 1) = "y" And jed1(i) = ")")) = False Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If jed1(i) = "(" Then
                kz = kz + 1
                kz1 = kz1 + 1
                If i < jed1.Length - 1 Then
                    If IsNumeric(jed1(i + 1)) Then
                        ListBox2.Items.Clear()
                        ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed1(i) = ")" Then
                kz = kz - 1
                kz2 = kz2 + 1
                If i < jed1.Length - 1 Then
                    If IsNumeric(jed1(i + 1)) Then
                        ListBox2.Items.Clear()
                        ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed1(i) = ")" And kz < 0 Then
                ListBox2.Items.Clear()
                ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                GoTo greska
            End If
            If jed1(i) = "=" Then
                bj = bj + 1

                If bj > 1 Or kz <> 0 Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If i = jed1.Length - 1 Then
                If (kz1 <> kz2) Or bj = 0 Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed1(i) = "=" And i < jed1.Length - 2 Then
                If IsNumeric(jed1(i + 1)) = True Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed1(i) = "x" Or jed1(i) = "K" Or jed1(i) = "y" Then
                If i < jed1.Length - 1 Then
                    If IsNumeric(jed1(i + 1)) Then
                        ListBox2.Items.Clear()
                        ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
        Next

        'provjera 2 jednadzbe
        kz = 0
        kz1 = 0
        kz2 = 0
        bj = 0

        For i = 0 To jed2.Length - 1

            If IsNumeric(jed2(i)) = False And jed2(i) <> "(" And jed2(i) <> ")" And jed2(i) <> "x" And jed2(i) <> "K" And jed2(i) <> "+" And jed2(i) <> "-" And jed2(i) <> "/" And jed2(i) <> "*" And jed2(i) <> "=" And jed2(i) <> "y" And jed2(i) <> "," Then
                ListBox2.Items.Clear()
                ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                GoTo greska
            End If
            If i = 0 Then
                If IsNumeric(jed2(i)) = False And jed2(i) <> "+" And jed2(i) <> "-" Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If i > 0 Then
                If (IsNumeric(jed2(i - 1)) = False And IsNumeric(jed2(i)) = False) And ((jed2(i - 1) = "(" And jed2(i) = "-") Or (jed2(i - 1) = "(" And jed2(i) = "+") Or (jed2(i - 1) = "x" And jed2(i) = ")") Or (jed2(i - 1) = "x" And jed2(i) = "+") Or (jed2(i - 1) = "x" And jed2(i) = "-") Or (jed2(i - 1) = ")" And jed2(i) = ")") Or (jed2(i - 1) = ")" And jed2(i) = "=") Or (jed2(i - 1) = ")" And jed2(i) = "+") Or (jed2(i - 1) = ")" And jed2(i) = "-") Or (jed2(i - 1) = "x" And jed2(i) = "(") Or (jed2(i - 1) = "=" And jed2(i) = "-") Or (jed2(i - 1) = "=" And jed2(i) = "+") Or (jed2(i - 1) = "x" And jed2(i) = "=") Or (jed2(i - 1) = "K" And jed2(i) = "=") Or (jed2(i - 1) = "K" And jed2(i) = "+") Or (jed2(i - 1) = "K" And jed2(i) = "(") Or (jed2(i - 1) = "K" And jed2(i) = "-") Or (jed2(i - 1) = "K" And jed2(i) = ")") Or (jed2(i - 1) = "y" And jed2(i) = "=") Or (jed2(i - 1) = "y" And jed2(i) = "+") Or (jed2(i - 1) = "y" And jed2(i) = "(") Or (jed2(i - 1) = "y" And jed2(i) = "-") Or (jed2(i - 1) = "y" And jed2(i) = ")")) = False Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If
            If jed2(i) = "(" Then
                kz = kz + 1
                kz1 = kz1 + 1
                If i < jed2.Length - 1 Then
                    If IsNumeric(jed2(i + 1)) Then
                        ListBox2.Items.Clear()
                        ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed2(i) = ")" Then
                kz = kz - 1
                kz2 = kz2 + 1
                If i < jed2.Length - 1 Then
                    If IsNumeric(jed2(i + 1)) Then
                        ListBox2.Items.Clear()
                        ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
            If jed2(i) = ")" And kz < 0 Then
                ListBox2.Items.Clear()
                ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                GoTo greska
            End If
            If jed2(i) = "=" Then
                bj = bj + 1

                If bj > 1 Or kz <> 0 Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If i = jed2.Length - 1 Then
                If (kz1 <> kz2) Or bj = 0 Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed2(i) = "=" And i < jed2.Length - 2 Then
                If IsNumeric(jed2(i + 1)) = True Then
                    ListBox2.Items.Clear()
                    ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                    GoTo greska
                End If
            End If

            If jed2(i) = "x" Or jed2(i) = "K" Or jed2(i) = "y" Then
                If i < jed2.Length - 1 Then
                    If IsNumeric(jed2(i + 1)) Then
                        ListBox2.Items.Clear()
                        ListBox2.Items.Add("Nazalost, dogodila se pogreska, pokusajte ponovo !")
                        GoTo greska
                    End If
                End If
            End If
        Next

        Do
            For i = 0 To jed1.Length - 1
                If jed1(i) = "(" Then
                    pb = i
                    pom = True
                End If

                If pb > -1 Then
                    For j = pb To 0 Step -1
                        If izracunaj = False And (jed1(j) = "+" Or jed1(j) = "-" Or jed1(j) = "=" Or j = 0) Then
                            If jed1(j) = "=" Then
                                pb = j + 1
                            Else : pb = j
                            End If
                            izracunaj = True
                        End If
                    Next
                    izracunaj = False
                    For j = pb To jed1.Length - 1
                        If izracunaj = False And jed1(j) = ")" Then
                            zb = j
                            izracunaj = True
                        End If
                    Next
                End If
                izracunaj = False
            Next

            If pb > -1 Then
                jed1 = zamjeni(jed1, zagrada(jed1, pb, zb), pb, zb)
                If pStupanj = True Then
                    ListBox2.Items.Add("Pogreska!")
                    ListBox2.Items.Add("Unosite jednadzbe najvise drugog reda")
                    GoTo greska
                End If
                If promjena = True Then
                    ListBox2.Items.Add(jed1)
                End If
            End If

            pb = -1
            zb = -1
            pom = False
            For i = 0 To jed1.Length - 1
                If jed1(i) = "(" Then
                    pom = True
                End If
            Next

        Loop While pom = True

        izracunaj = False
        Dim pp As Boolean
        pp = promjena
        '2 jednadzba

        If promjena = True Then
            ListBox2.Items.Add(" ")
            promjena = False
        End If

        pStupanj = False
        Dim kont As Integer = 0
        Do
            For i = 0 To jed2.Length - 1
                If jed2(i) = "(" Then
                    pb = i
                    pom = True
                End If

                If pb > -1 Then
                    For j = pb To 0 Step -1
                        If izracunaj = False And (jed2(j) = "+" Or jed2(j) = "-" Or jed2(j) = "=" Or j = 0) Then
                            If jed2(j) = "=" Then
                                pb = j + 1
                            Else : pb = j
                            End If
                            izracunaj = True
                        End If
                    Next
                    izracunaj = False
                    For j = pb To jed2.Length - 1
                        If izracunaj = False And jed2(j) = ")" Then
                            zb = j
                            izracunaj = True
                        End If
                    Next
                End If
                izracunaj = False
            Next

            If pb > -1 Then
                jed2 = zamjeni(jed2, zagrada(jed2, pb, zb), pb, zb)
                If pStupanj = True Then
                    ListBox2.Items.Add("Pogreska!")
                    ListBox2.Items.Add("Unosite jednadzbe najvise drugog reda")
                    GoTo greska
                End If
            End If

            If promjena = True Then
                If pp = False And kont = 0 Then
                    ListBox2.Items.Add(jed1)
                    ListBox2.Items.Add(" ")
                    ListBox2.Items.Add(jed2)
                    kont = 1
                Else
                    ListBox2.Items.Add(jed2)
                End If
            ElseIf promjena = False And pp = True Then
                ListBox2.Items.Add(jed2)
            End If

            pb = -1
            zb = -1
            pom = False
            For i = 0 To jed2.Length - 1
                If jed2(i) = "(" Then
                    pom = True
                End If
            Next
        Loop While pom = True

        If promjena = True Or pp = True Then
            ListBox2.Items.Add("--------------------")
        End If

        promjena = False
        izracunaj = False


        For i = 0 To jed1.Length - 1
            If i = jed1.Length - 1 Then
                izracunaj = True
            End If
            If i > 0 Then
                If jed1(i) = "=" And (jed1(i - 1) = "x" Or jed1(i - 1) = "K" Or jed1(i - 1) = "y") Then
                    kon = 0
                    strana = 1
                    GoTo lab
                End If
            End If

            If IsNumeric(jed1(i)) Or jed1(i) = "," Or ((jed1(i) = "+" Or jed1(i) = "-") And zb = -1) Then
                If pb = -1 Then pb = i
                zb = i
                If izracunaj = False Then
                    GoTo lab
                End If
            End If
            If izracunaj = True Or (IsNumeric(jed1(i)) Or jed1(i) = "," Or ((jed1(i) = "+" Or jed1(i) = "-") And zb = -1)) = False Then
                If kon = 0 Then
                    Double.TryParse(izdvoji_niz(jed1, pb, zb), broj)
                    kon = 1
                ElseIf pb > 0 Then
                    If jed1(pb - 1) = "x" Or jed1(pb - 1) = "K" Or jed1(pb - 1) = "y" Then
                        Double.TryParse(izdvoji_niz(jed1, pb, zb), broj)
                    Else
                        Double.TryParse(izdvoji_niz(jed1, pb - 1, zb), broj)
                    End If

                End If

                If jed1(i) = "x" Then
                    If strana = 0 Then n_l.Add(broj)
                    If strana = 1 Then n_d.Add(broj)
                ElseIf jed1(i) = "K" Then
                    If strana = 0 Then kn_l.Add(broj)
                    If strana = 1 Then kn_d.Add(broj)
                ElseIf jed1(i) = "y" Then
                    If strana = 0 Then y_l.Add(broj)
                    If strana = 1 Then y_d.Add(broj)
                ElseIf jed1(i) = "+" Or jed1(i) = "-" Or jed1(i) = "=" Or izracunaj = True Then
                    If strana = 0 Then p_l.Add(broj)
                    If strana = 1 Then p_d.Add(broj)
                End If
                If jed1(i) = "=" Then
                    strana = 1
                    kon = 0
                End If
                pb = -1
                zb = -1

            End If
lab:

        Next

        Dim pozl() As Double = p_l.ToArray()
        Dim pozd() As Double = p_d.ToArray()
        Dim nepl() As Double = n_l.ToArray()
        Dim nepd() As Double = n_d.ToArray()
        Dim knepl() As Double = kn_l.ToArray()
        Dim knepd() As Double = kn_d.ToArray()
        Dim yl() As Double = y_l.ToArray()
        Dim yd() As Double = y_d.ToArray()

        Dim txt As String = Nothing
        Dim a, b, c, d As Double
        a = 0
        b = 0
        c = 0
        d = 0

        Dim ka, kb, kc As Double

        For Each i As Double In yl
            txt = txt & pz(i) & Abs(i) & "y"
            d = d + i
        Next
        For Each i As Double In yd
            txt = txt & pz(-i) & Abs(i) & "y"
            d = d - i
        Next
        txt = txt & "="
        For Each i As Double In knepl
            txt = txt & pz(-i) & Abs(i) & "K"
            a = a - i
        Next
        For Each i As Double In knepd
            txt = txt & pz(i) & Abs(i) & "K"
            a = a + i
        Next
        For Each i As Double In nepl
            txt = txt & pz(-i) & Abs(i) & "x"
            b = b - i
        Next
        For Each i As Double In nepd
            txt = txt & pz(i) & Abs(i) & "x"
            b = b + i
        Next
        For Each i As Double In pozl
            txt = txt & pz(-i) & Abs(i)
            c = c - i
        Next
        For Each i As Double In pozd
            txt = txt & pz(i) & Abs(i)
            c = c + i
        Next
        ListBox2.Items.Add(txt)
        txt = pz(d) & Abs(d) & "y=" & jedispis(a, b, c)
        ListBox2.Items.Add(txt & " /" & d)
        If d = 0 Then
            ListBox2.Items.Clear()
            ListBox2.Items.Add("Koeficijent nepoznanice y ne moze biti 0")
            GoTo greska
        End If
        a = FormatNumber(a / d, 3)
        b = FormatNumber(b / d, 3)
        c = FormatNumber(c / d, 3)
        ka = a
        kb = b
        kc = c
        d = 1
        txt = "y=" & jedispis(a, b, c)
        ListBox2.Items.Add(txt)

        ListBox2.Items.Add(" ")

        p_l.Clear()
        n_l.Clear()
        p_d.Clear()
        n_d.Clear()
        kn_l.Clear()
        kn_d.Clear()
        y_l.Clear()
        y_d.Clear()

        strana = 0
        kon = 0
        izracunaj = False

        For i = 0 To jed2.Length - 1
            If i = jed2.Length - 1 Then
                izracunaj = True
            End If
            If i > 0 Then
                If jed2(i) = "=" And (jed2(i - 1) = "x" Or jed2(i - 1) = "K" Or jed2(i - 1) = "y") Then
                    kon = 0
                    strana = 1
                    GoTo lab1
                End If
            End If

            If IsNumeric(jed2(i)) Or jed2(i) = "," Or ((jed2(i) = "+" Or jed2(i) = "-") And zb = -1) Then
                If pb = -1 Then pb = i
                zb = i
                If izracunaj = False Then
                    GoTo lab1
                End If
            End If
            If izracunaj = True Or (IsNumeric(jed2(i)) Or jed2(i) = "," Or ((jed2(i) = "+" Or jed2(i) = "-") And zb = -1)) = False Then
                If kon = 0 Then
                    Double.TryParse(izdvoji_niz(jed2, pb, zb), broj)
                    kon = 1
                ElseIf pb > 0 Then
                    If jed2(pb - 1) = "x" Or jed2(pb - 1) = "K" Or jed2(pb - 1) = "y" Then
                        Double.TryParse(izdvoji_niz(jed2, pb, zb), broj)
                    Else
                        Double.TryParse(izdvoji_niz(jed2, pb - 1, zb), broj)
                    End If

                End If

                If jed2(i) = "x" Then
                    If strana = 0 Then n_l.Add(broj)
                    If strana = 1 Then n_d.Add(broj)
                ElseIf jed2(i) = "K" Then
                    If strana = 0 Then kn_l.Add(broj)
                    If strana = 1 Then kn_d.Add(broj)
                ElseIf jed2(i) = "y" Then
                    If strana = 0 Then y_l.Add(broj)
                    If strana = 1 Then y_d.Add(broj)
                ElseIf jed2(i) = "+" Or jed2(i) = "-" Or jed2(i) = "=" Or izracunaj = True Then
                    If strana = 0 Then p_l.Add(broj)
                    If strana = 1 Then p_d.Add(broj)
                End If
                If jed2(i) = "=" Then
                    strana = 1
                    kon = 0
                End If
                pb = -1
                zb = -1

            End If
lab1:

        Next

        pozl = p_l.ToArray()
        pozd = p_d.ToArray()
        nepl = n_l.ToArray()
        nepd = n_d.ToArray()
        knepl = kn_l.ToArray()
        knepd = kn_d.ToArray()
        yl = y_l.ToArray()
        yd = y_d.ToArray()

        txt = ""
        a = 0
        b = 0
        c = 0
        d = 0

        For Each i As Double In yl
            txt = txt & pz(i) & Abs(i) & "y"
            d = d + i
        Next
        For Each i As Double In yd
            txt = txt & pz(-i) & Abs(i) & "y"
            d = d - i
        Next
        txt = txt & "="
        For Each i As Double In knepl
            txt = txt & pz(-i) & Abs(i) & "K"
            a = a - i
        Next
        For Each i As Double In knepd
            txt = txt & pz(i) & Abs(i) & "K"
            a = a + i
        Next
        For Each i As Double In nepl
            txt = txt & pz(-i) & Abs(i) & "x"
            b = b - i
        Next
        For Each i As Double In nepd
            txt = txt & pz(i) & Abs(i) & "x"
            b = b + i
        Next
        For Each i As Double In pozl
            txt = txt & pz(-i) & Abs(i)
            c = c - i
        Next
        For Each i As Double In pozd
            txt = txt & pz(i) & Abs(i)
            c = c + i
        Next
        ListBox2.Items.Add(txt)
        txt = pz(d) & Abs(d) & "y=" & jedispis(a, b, c)
        ListBox2.Items.Add(txt & " /" & d)
        If d = 0 Then
            ListBox2.Items.Clear()
            ListBox2.Items.Add("Koeficijentt nepoznanice y ne moze biti 0")
            GoTo greska
        End If
        a = FormatNumber(a / d, 3)
        b = FormatNumber(b / d, 3)
        c = FormatNumber(c / d, 3)
        d = 1
        txt = "y=" & jedispis(a, b, c)
        ListBox2.Items.Add(txt)
        ListBox2.Items.Add("--------------------")

        txt = jedispis(ka, kb, kc) & "=" & jedispis(a, b, c)
        ListBox2.Items.Add(txt)
        ka = ka - a
        kb = kb - b
        kc = kc - c
        txt = jedispis(ka, kb, kc) & "=0"
        ListBox2.Items.Add(txt)

        Dim nemarj As Boolean = False

        If Abs(ka) > 0 Then
            txt = "x = ( -" & kb & "+/- sqrt(" & kb & "*" & kb & "-4*" & ka & "*" & kc & ") ) / 2*" & ka
            ListBox2.Items.Add(txt)
            txt = "x = ( -" & kb & "+/- sqrt(" & kb * kb - 4 * ka * kc & ") ) / " & 2 * ka
            ListBox2.Items.Add(txt)

            If (kb * kb - 4 * ka * kc) >= 0 Then
                ListBox2.Items.Add("x1 = " & FormatNumber((-kb + Sqrt(kb * kb - 4 * ka * kc)) / (2 * ka), 5))
                ListBox2.Items.Add("x2 = " & FormatNumber((-kb - Sqrt(kb * kb - 4 * ka * kc)) / (2 * ka), 5))
                rj1 = (-kb + Sqrt(kb * kb - 4 * ka * kc)) / (2 * ka)
                rj2 = (-kb - Sqrt(kb * kb - 4 * ka * kc)) / (2 * ka)
            Else
                ListBox2.Items.Add("x1,x2 = prazan skup")
                rj1 = Nothing
                rj2 = Nothing
                nemarj = True
            End If
        Else
            txt = pz(kb) & Abs(kb) & "x=" & pz(-kc) & Abs(kc)
            ListBox2.Items.Add(txt)
            If kb = 0 Then
                If kc = 0 Then
                    txt = "x = R"
                Else
                    txt = "x = Prazan skup"
                End If
                rj1 = Nothing
                rj2 = Nothing
                nemarj = True
            Else
                txt = "x = " & -kc & "/" & kb
                ListBox2.Items.Add(txt)
                txt = "x = " & FormatNumber(-kc / kb, 5)
                rj1 = -kc / kb
                rj2 = -kc / kb

            End If
            ListBox2.Items.Add(txt)
        End If

        ListBox2.Items.Add(" ")

        txt = "y=" & jedispis(a, b, c)
        ListBox2.Items.Add(txt)
        If rj1 <> rj2 And nemarj = False Then
            txt = "y1=" & pz(a) & Abs(a) & "*" & pz(rj1) & FormatNumber(Abs(rj1), 4) & "*" & pz(rj1) & FormatNumber(Abs(rj1), 4) & pz(b) & Abs(b) & "*" & pz(rj1) & FormatNumber(Abs(rj1), 4) & pz(c) & Abs(c)
            ListBox2.Items.Add(txt)
            txt = "y2=" & pz(a) & Abs(a) & "*" & pz(rj2) & FormatNumber(Abs(rj2), 4) & "*" & pz(rj2) & FormatNumber(Abs(rj2), 4) & pz(b) & Abs(b) & "*" & pz(rj2) & FormatNumber(Abs(rj2), 4) & pz(c) & Abs(c)
            ListBox2.Items.Add(txt)
            ry1 = FormatNumber(a * rj1 * rj1 + b * rj1 + c, 5)
            ry2 = FormatNumber(a * rj2 * rj2 + b * rj2 + c, 5)
            txt = "y1=" & ry1
            ListBox2.Items.Add(txt)
            txt = "y2=" & ry2
            ListBox2.Items.Add(txt)
            Label21.Text = Label21.Text & vbCrLf & "( " & FormatNumber(rj1, 4) & " , " & FormatNumber(ry1, 4) & " )" & vbCrLf & "( " & FormatNumber(rj2, 4) & " , " & FormatNumber(ry2, 4) & " )"
        ElseIf rj1 = rj2 And nemarj = False Then
            txt = "y=" & pz(a) & Abs(a) & "*" & pz(rj1) & FormatNumber(Abs(rj1), 4) & "*" & pz(rj1) & FormatNumber(Abs(rj1), 4) & pz(b) & Abs(b) & "*" & pz(rj1) & FormatNumber(Abs(rj1), 4) & pz(c) & Abs(c)
            ListBox2.Items.Add(txt)
            ry1 = FormatNumber(a * rj1 * rj1 + b * rj1 + c, 5)
            txt = "y=" & ry1
            ListBox2.Items.Add(txt)
            Label21.Text = Label21.Text & vbCrLf & "( " & FormatNumber(rj1, 4) & " , " & FormatNumber(ry1, 4) & " )"
        Else
            txt = "y= prazan skup"
            ListBox2.Items.Add(txt)
            Label21.Text = Label21.Text & vbCrLf & "(-,-)"
        End If

greska:
        pStupanj = False
        susjed = False
        promjena = False
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim a, b, c, Va, Vb, Vc, Ro, ru, ta, tb, tc, alp, bet, gam, Pov, o, s As Double
        Dim greska As Boolean
        Dim ispis As String
        a = 0
        b = 0
        c = 0
        Va = 0
        Vb = 0
        Vc = 0
        Ro = 0
        ru = 0
        ta = 0
        tb = 0
        tc = 0
        alp = -1
        bet = -1
        gam = -1
        Pov = 0
        o = 0
        greska = False
        rispis = ""
        gispis = ""
        pispis = ""

        Pt = False

        If Trokut_Unos(a, b, c, Va, Vb, Vc, 0, 0, Ro, ru, ta, tb, tc, alp, bet, gam, greska) = 0 Then GoTo prekid1
        Trokut(a, b, c, Va, Vb, Vc, -1, -1, Ro, ru, ta, tb, tc, alp, bet, gam, Pov, o, s, greska)

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav trokut ne postoji" & vbCrLf
            If gispis.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gispis
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Trokut")
        Else : MsgBox(rispis, MsgBoxStyle.Information, "Trokut")
        End If
prekid1:
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim a, b, c, Vc, ta, tb, tc, alp, bet, gam, Pov, o, p, q As Double
        Dim greska As Boolean
        Dim ispis As String
        a = 0
        b = 0
        c = 0
        p = 0
        q = 0
        Vc = 0
        ta = 0
        tb = 0
        tc = 0
        alp = -1
        bet = -1
        p = 0
        o = 0
        gam = System.Math.PI / 2
        greska = False
        rispis = ""
        gispis = ""
        pispis = ""

        Pt = True

        If Trokut_Unos(a, b, c, b, a, Vc, p, q, 0, 0, ta, tb, tc, alp, bet, gam, greska) = 0 Then GoTo prekid2
        Trokut(a, b, c, b, a, Vc, p, q, 0, 0, ta, tb, tc, alp, bet, gam, Pov, o, 0, greska)

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav pravokutan trokut ne postoji" & vbCrLf
            If gispis.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gispis
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Pravokutan trokut")
        Else : MsgBox(rispis, MsgBoxStyle.Information, "Pravokutan trokut")
        End If
prekid2:
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim n, bet, Kn, gam, Dn, alp As Double
        Dim greska As Boolean
        Dim unos, ispis As String
        n = 0
        Kn = 0
        Dn = -1
        alp = -1
        bet = -1
        gam = -1
        greska = False
        ispis = ""
        rispis = ""
        gispis = ""
        pispis = ""

        ' Unos i Racun

        If RadioButton1.Checked = True Then
            Do
                unos = InputBox("Upisi broj stranica / kuteva pravilnog mnogokuta:", "Unos podataka")
                If IsNumeric(unos) = True Then n = unos
                If unos = "" Then GoTo prekid
            Loop While n <= 2 Or Int(n) - n <> 0
            rispis = rispis & "Broj stranica n = " & n & vbCrLf & vbCrLf
            pispis = rispis
            bet = 360 / n
            rispis = rispis & "Kut β = " & Kut_Ispis(bet, True) & ", β = 360° / n" & vbCrLf
            Kn = (n - 2) * 180
            rispis = rispis & "Zbroj unutarnjih kuteva Kn = " & Kut_Ispis(Kn, True) & ", Kn = (n - 2) * 180°" & vbCrLf
            gam = (n - 2) * 180 / n
            rispis = rispis & "Kut ɣ = " & Kut_Ispis(gam, True) & ", ɣ = (n - 2) * 180° / n" & vbCrLf
            Dn = (n - 3) * n / 2
            rispis = rispis & "Broj dijagonala Dn = " & Int(Dn) & ", Dn = (n - 3) * n / 2" & vbCrLf
            alp = 360 / n
            rispis = rispis & "Kut α = " & Kut_Ispis(alp, True) & ", α = 360° / n" & vbCrLf

        ElseIf RadioButton2.Checked = True Then
            Do
                Form2.Label1.Text = "Unesi velicinu vanjskog kuta β pravilnog" & vbCrLf & "mnogokuta u stupnjevima:"
                Kut.Reset()
                Form2.ShowDialog()
                bet = Kut.Izracunaj()
                If bet = 0 Then GoTo prekid
            Loop While bet <= 0
            rispis = rispis & "Kut β = " & Kut_Ispis(bet, True) & vbCrLf & vbCrLf
            pispis = rispis
            If bet > 120 Then
                greska = True
                gispis = gispis & "β > 120°" & vbCrLf
            End If

            n = 360 / bet
            rispis = rispis & "Broj stranica n = " & Int(n) & ", n = 360° / β" & vbCrLf
            Kn = (n - 2) * 180
            rispis = rispis & "Zbroj unutarnjih kuteva Kn = " & Kut_Ispis(Kn, True) & ", Kn = (n - 2) * 180°" & vbCrLf
            gam = (n - 2) * 180 / n
            rispis = rispis & "Kut ɣ = " & Kut_Ispis(gam, True) & ", ɣ = (n - 2) * 180° / n" & vbCrLf
            Dn = (n - 3) * n / 2
            rispis = rispis & "Broj dijagonala Dn = " & Int(Dn) & ", Dn = (n - 3) * n / 2" & vbCrLf
            alp = 360 / n
            rispis = rispis & "Kut α = " & Kut_Ispis(alp, True) & ", α = 360° / n" & vbCrLf

        ElseIf RadioButton3.Checked = True Then
            Do
                Kut.Kontrola = False
                Form2.Label1.Text = "Unesi Zbroj svih unutarnjih kutova Kn" & vbCrLf & "u stupnjevima:"
                Kut.Reset()
                Form2.ShowDialog()
                Kn = Kut.Izracunaj()
                Kut.Kontrola = True
                If Kn = 0 Then GoTo prekid
            Loop While Kn <= 0
            rispis = rispis & "Zbroj unutarnjih kuteva Kn = " & Kut_Ispis(Kn, True) & vbCrLf & vbCrLf
            pispis = rispis
            n = (Kn + 360) / 180
            rispis = rispis & "Broj stranica n = " & Int(n) & ", n = (Kn + 360°) / 180°" & vbCrLf
            bet = 360 / n
            rispis = rispis & "Kut β = " & Kut_Ispis(bet, True) & ", β = 360° / n" & vbCrLf
            gam = (n - 2) * 180 / n
            rispis = rispis & "Kut ɣ = " & Kut_Ispis(gam, True) & ", ɣ = (n - 2) * 180° / n" & vbCrLf
            Dn = (n - 3) * n / 2
            rispis = rispis & "Broj dijagonala Dn = " & Int(Dn) & ", Dn = (n - 3) * n / 2" & vbCrLf
            alp = 360 / n
            rispis = rispis & "Kut α = " & Kut_Ispis(alp, True) & ", α = 360° / n" & vbCrLf

        ElseIf RadioButton4.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu unutarnjeg kuta ɣ pravilnog" & vbCrLf & "mnogokuta u stupnjevima."
                Kut.Reset()
                Form2.ShowDialog()
                gam = Kut.Izracunaj()
                If gam = 0 Then GoTo prekid
            Loop While gam <= 0
            rispis = rispis & "Kut ɣ = " & Kut_Ispis(gam, True) & vbCrLf & vbCrLf
            pispis = rispis
            n = (360) / (180 - gam)
            rispis = rispis & "Broj stranica n = " & Int(n) & ", n = 360° / (180° - ɣ)" & vbCrLf
            bet = 360 / n
            rispis = rispis & "Kut β = " & Kut_Ispis(bet, True) & ", β = 360° / n" & vbCrLf
            Kn = (n - 2) * 180
            rispis = rispis & "Zbroj unutarnjih kuteva Kn = " & Kut_Ispis(Kn, True) & ", Kn = (n - 2) * 180°" & vbCrLf
            Dn = (n - 3) * n / 2
            rispis = rispis & "Broj dijagonala Dn = " & Int(Dn) & ", Dn = (n - 3) * n / 2" & vbCrLf
            alp = 360 / n
            rispis = rispis & "Kut α = " & Kut_Ispis(alp, True) & ", α = 360° / n" & vbCrLf

        ElseIf RadioButton5.Checked = True Then
            Do
                unos = InputBox("Upisi ukupan broj dijagonala pravilnog mnogokuta:", "Unos podataka")
                If IsNumeric(unos) = True Then Dn = unos
                If unos = "" Then GoTo prekid
            Loop While Dn < 0 Or Int(Dn) - Dn <> 0
            rispis = rispis & "Broj dijagonala Dn = " & Dn & vbCrLf & vbCrLf
            pispis = rispis
            n = (3 + Sqrt(9 + 4 * 2 * Dn)) / 2
            rispis = rispis & "Broj stranica n = " & Int(n) & ", n = (3 + sqrt(9 + 8*Dn)) / 2" & vbCrLf
            bet = 360 / n
            rispis = rispis & "Kut β = " & Kut_Ispis(bet, True) & ", β = 360° / n" & vbCrLf
            Kn = (n - 2) * 180
            rispis = rispis & "Zbroj unutarnjih kuteva Kn = " & Kut_Ispis(Kn, True) & ", Kn = (n - 2) * 180°" & vbCrLf
            gam = (n - 2) * 180 / n
            rispis = rispis & "Kut ɣ = " & Kut_Ispis(gam, True) & ", ɣ = (n - 2) * 180° / n" & vbCrLf
            alp = 360 / n
            rispis = rispis & "Kut α = " & Kut_Ispis(alp, True) & ", α = 360° / n" & vbCrLf

        ElseIf RadioButton6.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu sredisnjeg kuta α pravilnog" & vbCrLf & "mnogokuta u stupnjevima."
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            rispis = rispis & "Kut α = " & Kut_Ispis(alp, True) & vbCrLf & vbCrLf
            pispis = rispis
            If alp > 120 Then
                greska = True
                gispis = gispis & "α > 120°" & vbCrLf
            End If

            n = 360 / alp
            rispis = rispis & "Broj stranica n = " & Int(n) & ", n = 360° / α" & vbCrLf
            Kn = (n - 2) * 180
            rispis = rispis & "Zbroj unutarnjih kuteva Kn = " & Kut_Ispis(Kn, True) & ", Kn = (n - 2) * 180°" & vbCrLf
            gam = (n - 2) * 180 / n
            rispis = rispis & "Kut ɣ = " & Kut_Ispis(gam, True) & ", ɣ = (n - 2) * 180° / n" & vbCrLf
            Dn = (n - 3) * n / 2
            rispis = rispis & "Broj dijagonala Dn = " & Int(Dn) & ", Dn = (n - 3) * n / 2" & vbCrLf
            bet = 360 / n
            rispis = rispis & "Kut β = " & Kut_Ispis(bet, True) & ", β = 360° / n" & vbCrLf

        End If

        ' Provjera

        If n < 0 Or bet < 0 Or Kn < 0 Or gam < 0 Or alp < 0 Or Dn < 0 Then
            greska = True
            gispis = gispis & "Velicine moraju biti pozitivne!" & vbCrLf
        End If
        If Int(n) - n < 0 Then
            greska = True
            gispis = gispis & "Broj stranica/vrhova mora biti prirodan broj" & vbCrLf
        End If
        If Int(Dn) - Dn < 0 Then
            greska = True
            gispis = gispis & "Broj dijagonala mora biti prirodan broj" & vbCrLf
        End If
        If bet < 360 / n - 0.0001 Or bet > 360 / n + 0.0001 Then
            greska = True
            gispis = gispis & "β != 360° / n" & vbCrLf
        End If
        If alp < 360 / n - 0.0001 Or alp > 360 / n + 0.0001 Then
            gispis = gispis & "α != 360° / n" & vbCrLf
            greska = True
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav pravilan mnogokut ne postoji" & vbCrLf
            If gispis.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gispis
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Mnogokut")
        Else
            MsgBox(rispis, MsgBoxStyle.Information, "Mnogokut")
        End If
prekid:
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim r, o, P As Double
        Dim unos As String
        r = 0
        o = 0
        P = 0
        rispis = ""

        ' Unos i Racun

        If RadioButton7.Checked = True Then
            Do
                unos = InputBox("Upisi duljinu polumjera r:", "Unos podataka")
                If IsNumeric(unos) = True Then r = unos
                If unos = "" Then GoTo prekid
            Loop While r <= 0
            rispis = rispis & "Polumjer r = " & FormatNumber(r, 4) & vbCrLf & vbCrLf

            o = 2 * r * System.Math.PI
            rispis = rispis & "o = " & FormatNumber(o, 4) & ", o = 2*r*π" & vbCrLf
            P = r * r * System.Math.PI
            rispis = rispis & "P = " & FormatNumber(P, 4) & ", P = r*r*π" & vbCrLf

        ElseIf RadioButton8.Checked = True Then
            Do
                unos = InputBox("Upisi opseg kruga o:", "Unos podataka")
                If IsNumeric(unos) = True Then o = unos
                If unos = "" Then GoTo prekid
            Loop While o <= 0
            rispis = rispis & "Opseg o = " & FormatNumber(o, 4) & vbCrLf & vbCrLf

            r = o / (2 * System.Math.PI)
            rispis = rispis & "r = " & FormatNumber(r, 4) & ", r = o / (2*π)" & vbCrLf
            P = r * r * System.Math.PI
            rispis = rispis & "P = " & FormatNumber(P, 4) & ", P = r*r*π" & vbCrLf

        ElseIf RadioButton9.Checked = True Then
            Do
                unos = InputBox("Upisi povrsinu kruga P:", "Unos podataka")
                If IsNumeric(unos) = True Then P = unos
                If unos = "" Then GoTo prekid
            Loop While P <= 0
            rispis = rispis & "Povrsina P = " & FormatNumber(P, 4) & vbCrLf & vbCrLf

            r = Sqrt(P / System.Math.PI)
            rispis = rispis & "r = " & FormatNumber(r, 4) & ", r = sqrt(P / π)" & vbCrLf
            o = 2 * r * System.Math.PI
            rispis = rispis & "o = " & FormatNumber(o, 4) & ", o = 2*r*π" & vbCrLf

        End If

        ' Ispis

        MsgBox(rispis, MsgBoxStyle.Information, "Krug")
prekid:
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dim alp, bet, a, b, d, P, o As Double
        Dim greska As Boolean
        Dim unos, ispis As String
        a = 0
        b = 0
        d = 0
        P = 0
        o = 0
        alp = -1
        bet = -1
        greska = False
        rispis = ""
        gispis = ""
        pispis = ""


        ' Unos

        If CheckBox27.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice a u pravokutniku:", "Unos podataka")
                If IsNumeric(unos) = True Then a = unos
                If unos = "" Then GoTo prekid
            Loop While a <= 0
            rispis = rispis & "Stranica a = " & FormatNumber(a, 4) & vbCrLf
        End If
        If CheckBox28.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice b u pravokutniku:", "Unos podataka")
                If IsNumeric(unos) = True Then b = unos
                If unos = "" Then GoTo prekid
            Loop While b <= 0
            rispis = rispis & "Stranica b = " & FormatNumber(b, 4) & vbCrLf
        End If
        If CheckBox31.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale d u pravokutniku:", "Unos podataka")
                If IsNumeric(unos) = True Then d = unos
                If unos = "" Then GoTo prekid
            Loop While d <= 0
            rispis = rispis & "Dijagonala d = " & FormatNumber(d, 4) & vbCrLf
        End If
        If CheckBox32.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta izmedu dijagonala α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            rispis = rispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
            If alp >= System.Math.PI Then
                gispis = gispis & "α > 180°" & vbCrLf
                greska = True
            End If
        End If
        If CheckBox33.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta izmedu dijagonala β:"
                Kut.Reset()
                Form2.ShowDialog()
                bet = Kut.Izracunaj()
                bet = (bet * System.Math.PI) / 180.0
                If bet = 0 Then GoTo prekid
            Loop While bet <= 0
            rispis = rispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
            If bet >= System.Math.PI Then
                gispis = gispis & "β > 180°" & vbCrLf
                greska = True
            End If
        End If

        rispis = rispis & vbCrLf
        pispis = rispis
        ' Racun

        For i = 1 To 3
            If a = 0 Then
                If b > 0 And d > 0 Then
                    a = Sqrt(d * d - b * b)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(d*d - b*b)" & vbCrLf
                ElseIf d > 0 And alp > 0 Then
                    a = Sqrt((d * d) / 2 - ((d * d) / 2) * Cos(alp))
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = Sqrt((d*d) / 2 - ((d*d) / 2)*Cos(α))" & vbCrLf
                End If
            End If
            If b = 0 Then
                If a > 0 And d > 0 Then
                    b = Sqrt(d * d - a * a)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(d*d - a*a)" & vbCrLf
                ElseIf d > 0 And bet > 0 Then
                    b = Sqrt((d * d) / 2 - ((d * d) / 2) * Cos(bet))
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = Sqrt((d*d) / 2 - ((d*d) / 2)*Cos(β))" & vbCrLf
                End If
            End If
            If d = 0 Then
                If a > 0 And b > 0 Then
                    d = Sqrt(b * b + a * a)
                    rispis = rispis & "d = " & FormatNumber(d, 4) & ", d = sqrt(a*a + b*b)" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If bet > 0 Then
                    alp = System.Math.PI - bet
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = 180° - β" & vbCrLf
                ElseIf a > 0 And d > 0 Then
                    If 1 - ((2 * a) / (d * d)) >= 1 Then
                        greska = True
                        gispis = gispis & "1 - ((2*a) / (d*d)) >= 1" & vbCrLf
                    End If
                    alp = Acos(1 - ((2 * a) / (d * d)))
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos(1 - ((2*a) / (d*d)))" & vbCrLf
                End If
            End If
            If bet = -1 Then
                If alp > 0 Then
                    bet = System.Math.PI - alp
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = 180° - α" & vbCrLf
                ElseIf b > 0 And d > 0 Then
                    If 1 - ((2 * b) / (d * d)) >= 1 Then
                        greska = True
                        gispis = gispis & "1 - ((2*b) / (d*d)) >= 1" & vbCrLf
                    End If
                    bet = Acos(1 - ((2 * b) / (d * d)))
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos(1 - ((2*b) / (d*d)))" & vbCrLf
                End If
            End If
        Next

        rispis = rispis & vbCrLf

        If P = 0 Then
            If a > 0 And b > 0 Then
                P = a * b
                rispis = rispis & "P = " & FormatNumber(P, 4) & ", P = a*b" & vbCrLf
            End If
        End If
        If o = 0 Then
            If a > 0 And b > 0 Then
                o = 2 * (a + b)
                rispis = rispis & "o = " & FormatNumber(o, 4) & ", o = 2*(a + b)" & vbCrLf
            End If
        End If

        If P = 0 Or o = 0 Then
            rispis = rispis & "Niste unijeli dovoljno podataka za izracunavanje opsega ili povrsine!" & vbCrLf
        End If

        ' Provjera

        If d > 0 And a > 0 And b > 0 Then
            If d * d < a * a + b * b - 0.00001 Or d * d > a * a + b * b + 0.00001 Then
                gispis = gispis & "a*a + b*b != d*d" & vbCrLf
                greska = True
            End If
        End If
        If d > 0 And a > 0 Then
            If d < a Then
                gispis = gispis & "d < a" & vbCrLf
                greska = True
            End If
        End If
        If d > 0 And b > 0 Then
            If d < b Then
                gispis = gispis & "d < b" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And bet > 0 Then
            If alp + bet < System.Math.PI - 0.00001 Or alp + bet > System.Math.PI + 0.00001 Then
                gispis = gispis & "α + β != 180°" & vbCrLf
                greska = True
            End If
        End If
        If alp = 0 Or bet = 0 Then
            gispis = gispis & "Kutevi α i β ne smiju biti 0°" & vbCrLf
            greska = True
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav pravokutnik ne postoji" & vbCrLf
            If gispis.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gispis
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Pravokutnik")
        Else
            MsgBox(rispis, MsgBoxStyle.Information, "Pravokutnik")
        End If
prekid:
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim alp, bet, a, b, de, df, p1, p2, P, o As Double
        Dim greska As Boolean
        Dim unos, ispis, gi As String
        a = 0
        b = 0
        de = 0
        df = 0
        P = 0
        o = 0
        p1 = -1
        p2 = -1
        alp = -1
        bet = -1
        greska = False
        rispis = ""
        ispis = ""
        pispis = ""
        gispis = ""
        gi = ""

        ' Unos

        If CheckBox40.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice a u paralelogramu:", "Unos podataka")
                If IsNumeric(unos) = True Then a = unos
                If unos = "" Then GoTo prekid
            Loop While a <= 0
            ispis = ispis & "Stranica a = " & FormatNumber(a, 4) & vbCrLf
        End If
        If CheckBox39.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice b u paralelogramu:", "Unos podataka")
                If IsNumeric(unos) = True Then b = unos
                If unos = "" Then GoTo prekid
            Loop While b <= 0
            ispis = ispis & "Stranica b = " & FormatNumber(b, 4) & vbCrLf
        End If
        If CheckBox35.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            If alp >= System.Math.PI Then
                gi = gi & "α > 180°" & vbCrLf
                greska = True
            End If
            ispis = ispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
        End If
        If CheckBox34.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta β:"
                Kut.Reset()
                Form2.ShowDialog()
                bet = Kut.Izracunaj()
                bet = (bet * System.Math.PI) / 180.0
                If bet = 0 Then GoTo prekid
            Loop While bet <= 0
            If bet > System.Math.PI Then
                gi = gi & "β > 180°" & vbCrLf
                greska = True
            End If
            ispis = ispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
        End If
        If CheckBox36.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale e u paralelogramu:", "Unos podataka")
                If IsNumeric(unos) = True Then de = unos
                If unos = "" Then GoTo prekid
            Loop While de <= 0
            ispis = ispis & "Dijagonala e = " & FormatNumber(de, 4) & vbCrLf
        End If
        If CheckBox41.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale f u paralelogramu:", "Unos podataka")
                If IsNumeric(unos) = True Then df = unos
                If unos = "" Then GoTo prekid
            Loop While df <= 0
            ispis = ispis & "Dijagonala f = " & FormatNumber(df, 4) & vbCrLf
        End If
        If CheckBox42.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta izmedu dijagonala φ1:"
                Kut.Reset()
                Form2.ShowDialog()
                p1 = Kut.Izracunaj()
                p1 = (p1 * System.Math.PI) / 180.0
                If p1 = 0 Then GoTo prekid
            Loop While p1 <= 0
            If p1 >= System.Math.PI Then
                gi = gi & "φ1 > 180°" & vbCrLf
                greska = True
            End If
            ispis = ispis & "Kut φ1 = " & Kut_Ispis(p1, False) & vbCrLf
        End If
        If CheckBox43.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta izmedu dijagonala φ2:"
                Kut.Reset()
                Form2.ShowDialog()
                p2 = Kut.Izracunaj()
                p2 = (p2 * System.Math.PI) / 180.0
                If p2 = 0 Then GoTo prekid
            Loop While p2 <= 0
            If p2 >= System.Math.PI Then
                gi = gi & "φ2 > 180°" & vbCrLf
                greska = True
            End If
            ispis = ispis & "Kut φ2 = " & Kut_Ispis(p2, False) & vbCrLf
        End If
        ispis = ispis & vbCrLf
        pispis = ispis
        ' Racun

        For i = 1 To 3
            If a = 0 Then
                If de > 0 And df > 0 And p1 > 0 Then
                    a = Sqrt((de * de) / 4 + (df * df) / 4 - ((de * df) / 2) * Cos(p1))
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = sqrt((e*e)/4 + (f*f)/4 - ((e*f)/2)*cos(φ1))" & vbCrLf
                ElseIf b > 0 And de > 0 And bet > 0 Then
                    Trokut(a, b, de, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, bet, 0, 0, 0, greska)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = (2*b*cos(β) + 2*sqrt(e*e - b*b*sin(β)*sin(β))) / 2" & vbCrLf
                ElseIf b > 0 And df > 0 And alp > 0 Then
                    Trokut(a, b, df, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, alp, 0, 0, 0, greska)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = (2*b*cos(α) + 2*sqrt(f*f - b*b*sin(α)*sin(α))) / 2" & vbCrLf
                End If
            End If
            If b = 0 Then
                If de > 0 And df > 0 And p2 > 0 Then
                    b = Sqrt((de * de) / 4 + (df * df) / 4 - ((de * df) / 2) * Cos(p2))
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt((e*e)/4 + (f*f)/4 - ((e*f)/2)*cos(φ2))" & vbCrLf
                ElseIf a > 0 And de > 0 And bet > 0 Then
                    Trokut(a, b, de, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, bet, 0, 0, 0, greska)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = (2*a*cos(β) + 2*sqrt(e*e - a*a*sin(β)*sin(β))) / 2" & vbCrLf
                ElseIf a > 0 And df > 0 And alp > 0 Then
                    Trokut(a, b, df, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, alp, 0, 0, 0, greska)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = (2*a*cos(α) + 2*sqrt(f*f - a*a*sin(α)*sin(α))) / 2" & vbCrLf
                End If
            End If
            If de = 0 Then
                If a > 0 And b > 0 And bet > 0 Then
                    de = Sqrt(a * a + b * b - 2 * a * b * Cos(bet))
                    ispis = ispis & "e = " & FormatNumber(de, 4) & ", e = sqrt(a*a + b*b - 2*a*b*cos(β))" & vbCrLf
                ElseIf a > 0 And df > 0 And p1 > 0 Then
                    Trokut(a, df / 2, de, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, p1, -1, -1, 0, 0, 0, greska)
                    de = de * 2
                    ispis = ispis & "e = " & FormatNumber(de, 4) & ", e = (2*f*cos(φ1) + 2*sqrt(4*a*a - f*f*sin(φ1)*sin(φ1))) / 2" & vbCrLf
                ElseIf b > 0 And df > 0 And p2 > 0 Then
                    Trokut(b, df / 2, de, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, p2, -1, -1, 0, 0, 0, greska)
                    de = de * 2
                    ispis = ispis & "e = " & FormatNumber(de, 4) & ", e = (2*f*cos(φ2) + 2*sqrt(4*b*b - f*f*sin(φ2)*sin(φ2))) / 2" & vbCrLf
                End If
            End If
            If df = 0 Then
                If a > 0 And b > 0 And alp > 0 Then
                    df = Sqrt(a * a + b * b - 2 * a * b * Cos(alp))
                    ispis = ispis & "f = " & FormatNumber(df, 4) & ", f = sqrt(a*a + b*b - 2*a*b*cos(α))" & vbCrLf
                ElseIf a > 0 And de > 0 And p1 > 0 Then
                    Trokut(a, de / 2, df, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, p1, -1, -1, 0, 0, 0, greska)
                    df = df * 2
                    ispis = ispis & "f = " & FormatNumber(df, 4) & ", f = (2*e*cos(φ1) + 2*sqrt(4*a*a - e*e*sin(φ1)*sin(φ1))) / 2" & vbCrLf
                ElseIf b > 0 And de > 0 And p2 > 0 Then
                    Trokut(b, de / 2, df, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, p2, -1, -1, 0, 0, 0, greska)
                    df = df * 2
                    ispis = ispis & "f = " & FormatNumber(df, 4) & ", f = (2*e*cos(φ2) + 2*sqrt(4*b*b - e*e*sin(φ2)*sin(φ2))) / 2" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If bet > 0 Then
                    alp = System.Math.PI - bet
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = 180° - β" & vbCrLf
                ElseIf a > 0 And b > 0 And df > 0 Then
                    If ((a * a + b * b - df * df) / (2 * a * b)) > 1 Then
                        greska = True
                        gi = gi & "((a*a + b*b - f*f) / (2*a*b)) > 1" & vbCrLf
                    End If
                    alp = Acos((a * a + b * b - df * df) / (2 * a * b))
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos((a*a + b*b - f*f) / (2*a*b))" & vbCrLf
                End If
            End If
            If bet = -1 Then
                If alp > 0 Then
                    bet = System.Math.PI - alp
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = 180° - α" & vbCrLf
                ElseIf a > 0 And b > 0 And de > 0 Then
                    If ((a * a + b * b - de * de) / (2 * a * b)) > 1 Then
                        greska = True
                        gi = gi & "((a*a + b*b - e*e) / (2*a*b)) > 1" & vbCrLf
                    End If
                    bet = Acos((a * a + b * b - de * de) / (2 * a * b))
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos((a*a + b*b - e*e) / (2*a*b))" & vbCrLf
                End If
            End If
            If p1 = -1 Then
                If p2 > 0 Then
                    p1 = System.Math.PI - p2
                    If greska = False Then ispis = ispis & "φ1 = " & Kut_Ispis(p1, False) & ", φ1 = 180° - φ2" & vbCrLf
                ElseIf a > 0 And de > 0 And df > 0 Then
                    If ((de * de + df * df - 4 * a * a) / (2 * de * df)) > 1 Then
                        greska = True
                        gi = gi & "((e*e + f*f - 4*a*a) / (2*e*f)) > 1" & vbCrLf
                    End If
                    p1 = Acos((de * de + df * df - 4 * a * a) / (2 * de * df))
                    If greska = False Then ispis = ispis & "φ1 = " & Kut_Ispis(p1, False) & ", φ1 = arccos((e*e + f*f - 4*a*a) / (2*e*f))" & vbCrLf
                End If
            End If
            If p2 = -1 Then
                If p1 > 0 Then
                    p2 = System.Math.PI - p1
                    If greska = False Then ispis = ispis & "φ2 = " & Kut_Ispis(p2, False) & ", φ2 = 180° - φ1" & vbCrLf
                ElseIf b > 0 And de > 0 And df > 0 Then
                    If ((de * de + df * df - 4 * b * b) / (2 * de * df)) > 1 Then
                        greska = True
                        gi = gi & "((e*e + f*f - 4*b*b) / (2*e*f)) > 1" & vbCrLf
                    End If
                    p2 = Acos((de * de + df * df - 4 * b * b) / (2 * de * df))
                    If greska = False Then ispis = ispis & "φ2 = " & Kut_Ispis(p2, False) & ", φ2 = arccos((e*e + f*f - 4*b*b) / (2*e*f))" & vbCrLf
                End If
            End If
        Next
        ispis = ispis & vbCrLf
        If P = 0 Then
            If a > 0 And b > 0 And alp > 0 Then
                P = a * b * Sin(alp)
                ispis = ispis & "P = " & FormatNumber(P, 4) & ", P = a*b*sin(α)" & vbCrLf
            ElseIf de > 0 And df > 0 And p1 > 0 Then
                P = 0.5 * de * df * Sin(p1)
                ispis = ispis & "P = " & FormatNumber(P, 4) & ", P = 0,5*e*f*sin(φ1)" & vbCrLf
            End If
        End If
        If o = 0 Then
            If a > 0 And b > 0 Then
                o = 2 * (a + b)
                ispis = ispis & "o = " & FormatNumber(o, 4) & ", o = 2*(a + b)" & vbCrLf
            End If
        End If

        If P = 0 Or o = 0 Then
            ispis = ispis & "Niste unijeli dovoljno podataka za izracunavanje opsega ili povrsine!" & vbCrLf
        End If


        ' Provjera

        If de > 0 And a > 0 And b > 0 And bet > 0 Then
            If de * de < a * a + b * b - 2 * a * b * Cos(bet) - 0.00001 Or de * de > a * a + b * b - 2 * a * b * Cos(bet) + 0.00001 Then
                gi = gi & "a*a + b*b - 2*a*b*cos(β) != e*e" & vbCrLf
                greska = True
            End If
        End If
        If df > 0 And a > 0 And b > 0 And alp > 0 Then
            If df * df < a * a + b * b - 2 * a * b * Cos(alp) - 0.00001 Or df * df > a * a + b * b - 2 * a * b * Cos(alp) + 0.00001 Then
                gi = gi & "a*a + b*b - 2*a*b*cos(α) != f*f" & vbCrLf
                greska = True
            End If
        End If

        If de > 0 And a > 0 And alp > 0 And alp < System.Math.PI / 2 Then
            If de < a Then
                gi = gi & "e < a, α < 90°" & vbCrLf
                greska = True
            End If
        End If
        If de > 0 And b > 0 And alp > 0 And alp < System.Math.PI / 2 Then
            If de < b Then
                gi = gi & "e < b, α < 90°" & vbCrLf
                greska = True
            End If
        End If
        If df > 0 And a > 0 And alp > System.Math.PI / 2 Then
            If df < a Then
                gi = gi & "f < a, α > 90°" & vbCrLf
                greska = True
            End If
        End If
        If df > 0 And b > 0 And alp > System.Math.PI / 2 Then
            If df < b Then
                gi = gi & "f < b, α > 90°" & vbCrLf
                greska = True
            End If
        End If

        If alp > 0 And bet > 0 Then
            If alp + bet < System.Math.PI - 0.00001 Or alp + bet > System.Math.PI + 0.00001 Then
                gi = gi & "α + β != 180°" & vbCrLf
                greska = True
            End If
        End If
        If alp = 0 Or bet = 0 Then
            gi = gi & "Kutevi α i β ne smiju biti 0°" & vbCrLf
            greska = True
        End If
        If p1 > 0 And p2 > 0 Then
            If p1 + p2 < System.Math.PI - 0.00001 Or p1 + p2 > System.Math.PI + 0.00001 Then
                gi = gi & "φ1 + φ2 != 180°" & vbCrLf
                greska = True
            End If
        End If
        If p1 = 0 Or p2 = 0 Then
            gi = gi & "Kutevi φ1 i φ2 ne smiju biti 0°" & vbCrLf
            greska = True
        End If
        If de > 0 And df > 0 And a > 0 And b > 0 Then
            If de * de + df * df < 2 * (a * a + b * b) - 0.00001 Or de * de + df * df > 2 * (a * a + b * b) + 0.00001 Then
                gi = gi & "e*e + f*f < 2*(a*a + b*b)" & vbCrLf
                greska = True
            End If
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav paralelogram ne postoji" & vbCrLf
            If gi.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gi
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Paralelogram")
        Else
            MsgBox(ispis, MsgBoxStyle.Information, "Paralelogram")
        End If
        rispis = ""
prekid:
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Dim alp, bet, gam, del, a, b, c, d, de, df, v, P, o As Double
        Dim greska As Boolean
        Dim unos, ispis, gi As String
        a = 0
        b = 0
        c = 0
        d = 0
        de = 0
        df = 0
        P = 0
        o = 0
        v = 0
        alp = -1
        bet = -1
        gam = -1
        del = -1
        greska = False
        ispis = ""
        rispis = ""
        pispis = ""
        gispis = ""
        gi = ""

        ' Unos

        If CheckBox53.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice a u trapezu:", "Unos podataka")
                If IsNumeric(unos) = True Then a = unos
                If unos = "" Then GoTo prekid
            Loop While a <= 0
            ispis = ispis & "Stranica a = " & FormatNumber(a, 4) & vbCrLf
        End If
        If CheckBox52.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice b u trapezu:", "Unos podataka")
                If IsNumeric(unos) = True Then b = unos
                If unos = "" Then GoTo prekid
            Loop While b <= 0
            ispis = ispis & "Stranica b = " & FormatNumber(b, 4) & vbCrLf
        End If
        If CheckBox55.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice c u trapezu:", "Unos podataka")
                If IsNumeric(unos) = True Then c = unos
                If unos = "" Then GoTo prekid
            Loop While c <= 0
            ispis = ispis & "Stranica c = " & FormatNumber(c, 4) & vbCrLf
        End If
        If CheckBox54.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice d u trapezu:", "Unos podataka")
                If IsNumeric(unos) = True Then d = unos
                If unos = "" Then GoTo prekid
            Loop While d <= 0
            ispis = ispis & "Stranica d = " & FormatNumber(d, 4) & vbCrLf
        End If
        If CheckBox48.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            If alp >= System.Math.PI Then
                gi = gi & "α > 180°" & vbCrLf
                greska = True
            End If
            ispis = ispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
        End If
        If CheckBox47.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta β:"
                Kut.Reset()
                Form2.ShowDialog()
                bet = Kut.Izracunaj()
                bet = (bet * System.Math.PI) / 180.0
                If bet = 0 Then GoTo prekid
            Loop While bet <= 0
            If bet > System.Math.PI Then
                gi = gi & "β > 180°" & vbCrLf
                greska = True
            End If
            ispis = ispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
        End If
        If CheckBox57.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta ɣ:"
                Kut.Reset()
                Form2.ShowDialog()
                gam = Kut.Izracunaj()
                gam = (gam * System.Math.PI) / 180.0
                If gam = 0 Then GoTo prekid
            Loop While gam <= 0
            If gam >= System.Math.PI Then
                gi = gi & "ɣ > 180°" & vbCrLf
                greska = True
            End If
            ispis = ispis & "Kut ɣ = " & Kut_Ispis(gam, False) & vbCrLf
        End If
        If CheckBox56.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta δ:"
                Kut.Reset()
                Form2.ShowDialog()
                del = Kut.Izracunaj()
                del = (del * System.Math.PI) / 180.0
                If del = 0 Then GoTo prekid
            Loop While del <= 0
            If del > System.Math.PI Then
                gi = gi & "δ > 180°" & vbCrLf
                greska = True
            End If
            ispis = ispis & "Kut δ = " & Kut_Ispis(del, False) & vbCrLf
        End If
        If CheckBox49.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale e u trapezu:", "Unos podataka")
                If IsNumeric(unos) = True Then de = unos
                If unos = "" Then GoTo prekid
            Loop While de <= 0
            ispis = ispis & "Dijagonala e = " & FormatNumber(de, 4) & vbCrLf
        End If
        If CheckBox46.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale f u trapezu:", "Unos podataka")
                If IsNumeric(unos) = True Then df = unos
                If unos = "" Then GoTo prekid
            Loop While df <= 0
            ispis = ispis & "Dijagonala f = " & FormatNumber(df, 4) & vbCrLf
        End If
        If CheckBox29.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h u trapezu:", "Unos podataka")
                If IsNumeric(unos) = True Then v = unos
                If unos = "" Then GoTo prekid
            Loop While v <= 0
            ispis = ispis & "Visina h = " & FormatNumber(v, 4) & vbCrLf
        End If
        ispis = ispis & vbCrLf
        pispis = ispis
        ' Racun

        For i = 1 To 3
            If a = 0 Then
                If de > 0 And b > 0 And bet > 0 Then
                    Trokut(a, de, b, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, bet, -1, 0, 0, 0, False)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = b*cos(β) + sqrt(e*e - b*b*sin(β))" & vbCrLf
                ElseIf df > 0 And d > 0 And alp > 0 Then
                    Trokut(a, df, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, False)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = d*cos(α) + sqrt(f*f - d*d*sin(α))" & vbCrLf
                ElseIf b > 0 And d > 0 And bet > 0 And c > 0 Then
                    Trokut(a, b, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, bet, 0, 0, 0, False)
                    a = a + c
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = b*cos(β) + sqrt(d*d - b*b*sin(β)) + c" & vbCrLf
                ElseIf b > 0 And d > 0 And alp > 0 And c > 0 Then
                    Trokut(a, b, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, False)
                    a = a + c
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = d*cos(α) + sqrt(b*b - d*d*sin(α)) + c" & vbCrLf
                ElseIf b > 0 And bet > 0 And alp > 0 And c > 0 Then
                    Trokut(a, b, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, False)
                    a = a + c
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = (b*sin(180° - α - β) / sin(α)) + c" & vbCrLf
                ElseIf d > 0 And bet > 0 And alp > 0 And c > 0 Then
                    Trokut(a, d, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, bet, alp, 0, 0, 0, False)
                    a = a + c
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = (d*sin(180° - α - β) / sin(β)) + c" & vbCrLf
                End If
            End If
            If c = 0 Then
                If de > 0 And d > 0 And del > 0 Then
                    Trokut(c, de, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, del, -1, 0, 0, 0, False)
                    ispis = ispis & "c = " & FormatNumber(c, 4) & ", c = d*cos(δ) + sqrt(e*e - d*d*sin(δ))" & vbCrLf
                ElseIf df > 0 And b > 0 And gam > 0 Then
                    Trokut(c, df, b, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, gam, -1, 0, 0, 0, False)
                    ispis = ispis & "c = " & FormatNumber(c, 4) & ", c = b*cos(γ) + sqrt(f*f - b*b*sin(γ))" & vbCrLf
                ElseIf b > 0 And d > 0 And bet > 0 And a > 0 Then
                    Trokut(c, b, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, bet, 0, 0, 0, False)
                    c = a - c
                    ispis = ispis & "c = " & FormatNumber(c, 4) & ", c = a - (b*cos(β) + sqrt(d*d - b*b*sin(β)))" & vbCrLf
                ElseIf b > 0 And d > 0 And alp > 0 And a > 0 Then
                    Trokut(c, b, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, False)
                    c = a - c
                    ispis = ispis & "c = " & FormatNumber(c, 4) & ", c = a - (d*cos(α) + sqrt(b*b - d*d*sin(α)))" & vbCrLf
                ElseIf b > 0 And bet > 0 And alp > 0 And a > 0 Then
                    Trokut(c, b, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, False)
                    c = a - c
                    ispis = ispis & "c = " & FormatNumber(c, 4) & ", c = a - (b*sin(180° - α - β) / sin(α))" & vbCrLf
                ElseIf d > 0 And bet > 0 And alp > 0 And a > 0 Then
                    Trokut(c, d, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, bet, alp, 0, 0, 0, False)
                    c = a - c
                    ispis = ispis & "c = " & FormatNumber(c, 4) & ", c = a - (d*sin(180° - α - β) / sin(β))" & vbCrLf
                End If
            End If
            If b = 0 Then
                If de > 0 And a > 0 And bet > 0 Then
                    Trokut(b, de, a, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, bet, -1, 0, 0, 0, False)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = a*cos(β) + sqrt(e*e - a*a*sin(β))" & vbCrLf
                ElseIf df > 0 And c > 0 And gam > 0 Then
                    Trokut(b, df, c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, gam, -1, 0, 0, 0, False)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = c*cos(γ) + sqrt(f*f - c*c*sin(γ))" & vbCrLf
                ElseIf v > 0 And bet > 0 Then
                    Trokut(b, v, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, System.Math.PI / 2, bet, -1, 0, 0, 0, False)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = v / sin(β)" & vbCrLf
                ElseIf v > 0 And gam > 0 Then
                    Trokut(b, v, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, System.Math.PI / 2, -1, Abs(gam - System.Math.PI / 2), 0, 0, 0, False)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = v / cos(γ - 90°)" & vbCrLf
                ElseIf alp > 0 And bet > 0 And d > 0 Then
                    Trokut(b, d, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, alp, bet, -1, 0, 0, 0, False)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = d*sin(α) / sin(β)" & vbCrLf
                ElseIf alp > 0 And bet > 0 And a > 0 And c > 0 Then
                    Trokut(b, a - c, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, alp, -1, bet, 0, 0, 0, False)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = (a - c)*sin(α) / sin(180° - α - β)" & vbCrLf
                ElseIf alp > 0 And a > 0 And c > 0 And d > 0 Then
                    Trokut(b, d, a - c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, alp, -1, -1, 0, 0, 0, False)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt((a-c)*(a-c) + d*d - 2*d*(a-c)*cos(α))" & vbCrLf
                ElseIf bet > 0 And a > 0 And c > 0 And d > 0 Then
                    Trokut(b, d, a - c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, bet, -1, 0, 0, 0, False)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = (a-c)*cos(β) + sqrt(d*d - (a-c)*(a-c)*sin(β))" & vbCrLf
                End If
            End If
            If d = 0 Then
                If de > 0 And c > 0 And del > 0 Then
                    Trokut(d, de, c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, del, -1, 0, 0, 0, False)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = c*cos(δ) + sqrt(e*e - c*c*sin(δ))" & vbCrLf
                ElseIf df > 0 And a > 0 And alp > 0 Then
                    Trokut(d, df, a, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, False)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = a*cos(α) + sqrt(f*f - a*a*sin(α))" & vbCrLf
                ElseIf v > 0 And alp > 0 Then
                    Trokut(d, v, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, System.Math.PI / 2, alp, -1, 0, 0, 0, False)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = v / sin(α)" & vbCrLf
                ElseIf v > 0 And del > 0 Then
                    Trokut(d, v, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, System.Math.PI / 2, -1, Abs(del - System.Math.PI / 2), 0, 0, 0, False)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = v / cos(δ - 90°)" & vbCrLf
                ElseIf alp > 0 And bet > 0 And b > 0 Then
                    Trokut(d, b, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, alp, -1, 0, 0, 0, False)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = b*sin(β) / sin(α)" & vbCrLf
                ElseIf alp > 0 And bet > 0 And a > 0 And c > 0 Then
                    Trokut(d, a - c, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, -1, alp, 0, 0, 0, False)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = (a-c)*sin(β) / sin(180° - α - β)" & vbCrLf
                ElseIf alp > 0 And a > 0 And c > 0 And b > 0 Then
                    Trokut(d, b, a - c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, greska)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = (a-c)*cos(α) + sqrt(b*b - (a-c)*(a-c)*sin(α))" & vbCrLf
                ElseIf bet > 0 And a > 0 And c > 0 And b > 0 Then
                    Trokut(d, b, a - c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, -1, -1, 0, 0, 0, False)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = sqrt((a-c)*(a-c) + b*b - 2*b*(a-c)*cos(β))" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If del > 0 Then
                    alp = System.Math.PI - del
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = 180° - δ" & vbCrLf
                ElseIf a > 0 And d > 0 And df > 0 Then
                    Trokut(df, a, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, alp, -1, -1, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos((a*a + d*d - f*f) / (2*a*d))" & vbCrLf
                ElseIf d > 0 And v > 0 Then
                    If v / d > 1 Then greska = True
                    alp = Asin(v / d)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(h / d)" & vbCrLf
                ElseIf bet > 0 And b > 0 And d > 0 Then
                    Trokut(d, b, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, alp, -1, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(b*sin(β) / d)" & vbCrLf
                ElseIf bet > 0 And b > 0 And a > 0 And c > 0 Then
                    Trokut(a - c, b, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(b*sin(β) / d), d = sqrt((a-c)*(a-c) + b*b - 2*b*(a-c)*cos(β))" & vbCrLf
                ElseIf bet > 0 And c > 0 And a > 0 And d > 0 Then
                    Trokut(d, a - c, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, -1, alp, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(b*sin(β) / d), b = (a-c)*cos(β) + sqrt(d*d - (a-c)*(a-c)*sin(β))" & vbCrLf
                ElseIf a > 0 And c > 0 And b > 0 And d > 0 Then
                    Trokut(d, b, a - c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos(((a-c)*(a-c) + d*d - b*b) / (2*(a-c)*d))" & vbCrLf
                End If
            End If
            If bet = -1 Then
                If gam > 0 Then
                    bet = System.Math.PI - gam
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = 180° - γ" & vbCrLf
                ElseIf a > 0 And b > 0 And de > 0 Then
                    Trokut(de, a, b, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, -1, -1, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos((a*a + b*b - e*e) / (2*a*b))" & vbCrLf
                ElseIf b > 0 And v > 0 Then
                    If v / b > 1 Then greska = True
                    bet = Asin(v / b)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(v / b)" & vbCrLf
                ElseIf alp > 0 And b > 0 And d > 0 Then
                    Trokut(d, b, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, alp, -1, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β =  arcsin(d*sin(α) / b)" & vbCrLf
                ElseIf alp > 0 And b > 0 And a > 0 And c > 0 Then
                    Trokut(a - c, b, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(d*sin(α) / b), d = (a-c)*cos(α) + sqrt(b*b - (a-c)*(a-c)*sin(α))" & vbCrLf
                ElseIf alp > 0 And c > 0 And a > 0 And d > 0 Then
                    Trokut(d, a - c, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, -1, alp, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(d*sin(α) / b), b = sqrt((a-c)*(a-c) + d*d - 2*d*(a-c)*cos(α))" & vbCrLf
                ElseIf a > 0 And c > 0 And b > 0 And d > 0 Then
                    Trokut(d, b, a - c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, -1, -1, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β =  arccos(((a-c)*(a-c) + b*b - d*d) / (2*(a-c)*b))" & vbCrLf
                End If
            End If
            If gam = -1 Then
                If c > 0 And b > 0 And df > 0 Then
                    Trokut(df, c, b, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, gam, -1, -1, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "γ = " & Kut_Ispis(gam, False) & ", γ = arccos((c*c + b*b - f*f) / (2*c*b))" & vbCrLf
                End If
            End If
            If del = -1 Then
                If c > 0 And d > 0 And de > 0 Then
                    Trokut(de, c, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, del, -1, -1, 0, 0, 0, False)
                    If greska = False Then ispis = ispis & "δ = " & Kut_Ispis(del, False) & ", δ = arccos((c*c + d*d - e*e) / (2*c*d))" & vbCrLf
                End If
            End If
            If de = 0 Then
                If d > 0 And c > 0 And del > 0 Then
                    Trokut(de, d, c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, del, -1, -1, 0, 0, 0, False)
                    ispis = ispis & "e = " & FormatNumber(de, 4) & ", e = sqrt(c*c + d*d - 2*c*d*cos(δ))" & vbCrLf
                ElseIf a > 0 And b > 0 And bet > 0 Then
                    Trokut(de, a, b, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, bet, -1, -1, 0, 0, 0, False)
                    ispis = ispis & "e = " & FormatNumber(de, 4) & ", e = sqrt(a*a + b*b - 2*a*b*cos(β))" & vbCrLf
                End If
            End If
            If df = 0 Then
                If b > 0 And c > 0 And gam > 0 Then
                    Trokut(df, b, c, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, gam, -1, -1, 0, 0, 0, False)
                    ispis = ispis & "f = " & FormatNumber(df, 4) & ", f = sqrt(c*c + b*b - 2*b*c*cos(γ))" & vbCrLf
                ElseIf a > 0 And d > 0 And alp > 0 Then
                    Trokut(df, a, d, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, alp, -1, -1, 0, 0, 0, False)
                    ispis = ispis & "f = " & FormatNumber(df, 4) & ", f = sqrt(a*a + d*d - 2*a*d*cos(α))" & vbCrLf
                End If
            End If
            If v = 0 Then
                If alp > 0 And d > 0 Then
                    v = d * Sin(alp)
                    ispis = ispis & "h = " & FormatNumber(v, 4) & ", h = d*sin(α)" & vbCrLf
                ElseIf bet > 0 And b > 0 Then
                    v = b * Sin(bet)
                    ispis = ispis & "h = " & FormatNumber(v, 4) & ", h = b*sin(β)" & vbCrLf
                End If
            End If
        Next
        ispis = ispis & vbCrLf
        ' Provjera

        If a < 0 Then
            gi = gi & "a < 0" & vbCrLf
            greska = True
        End If
        If d < 0 Then
            gi = gi & "d < 0" & vbCrLf
            greska = True
        End If
        If c < 0 Then
            gi = gi & "c < 0" & vbCrLf
            greska = True
        End If
        If b < 0 Then
            gi = gi & "b < 0" & vbCrLf
            greska = True
        End If
        If d > 0 And v > 0 And alp > 0 Then
            If Sin(alp) < v / d - 0.00001 Or Sin(alp) > v / d + 0.00001 Then
                gi = gi & "sin(α) != v / d" & vbCrLf
                greska = True
            End If
        End If
        If b > 0 And v > 0 And bet > 0 Then
            If Sin(bet) < v / b - 0.00001 Or Sin(bet) > v / b + 0.00001 Then
                gi = gi & "sin(β) != v / b" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And del > 0 Then
            If alp + del < System.Math.PI - 0.00001 Or alp + del > System.Math.PI + 0.00001 Then
                gi = gi & "α + δ != 180°" & vbCrLf
                greska = True
            End If
        End If
        If bet > 0 And gam > 0 Then
            If bet + gam < System.Math.PI - 0.00001 Or bet + gam > System.Math.PI + 0.00001 Then
                gi = gi & "β + γ != 180°" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And del > 0 And bet > 0 And gam > 0 Then
            If alp + del + bet + gam < 2 * System.Math.PI - 0.00001 Or alp + del + bet + gam > 2 * System.Math.PI + 0.00001 Then
                gi = gi & "α + β + γ + δ != 360°" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And c > 0 Then
            If c > a Then
                MsgBox("Stranica a ne smije biti kraca od stranice c !", MsgBoxStyle.Exclamation)
                GoTo pog
            End If
        End If
        If alp >= System.Math.PI / 2 Then
            MsgBox("Kut α mora biti manji od 90° !", MsgBoxStyle.Exclamation)
            GoTo pog
        End If
        If Kut_Ispis(alp, False) = "NaN° NaN' NaN''" Or Kut_Ispis(bet, False) = "NaN° NaN' NaN''" Or Kut_Ispis(gam, False) = "NaN° NaN' NaN''" Or Kut_Ispis(del, False) = "NaN° NaN' NaN''" Then
            greska = True
            gi = gi & "Sa zadanim stranicama se ne moze sastaviti trapez" & vbCrLf
        End If

        ' Povrsina i opseg

        If a > 0 And c > 0 And v > 0 Then
            P = ((a + c) * v) / 2
            ispis = ispis & "P = " & FormatNumber(P, 4) & ", P = (v*(a + c)) / 2" & vbCrLf
        End If
        If a > 0 And b > 0 And c > 0 And d > 0 Then
            o = a + b + c + d
            ispis = ispis & "o = " & FormatNumber(o, 4) & ", o = a + b + c + d" & vbCrLf
        End If

        If P = 0 Or o = 0 Then
            ispis = ispis & "Niste unijeli dovoljno podataka za izracunavanje opsega ili povrsine!" & vbCrLf
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav trapez ne postoji" & vbCrLf
            If gi.Length > 0 Then
                ispis = ispis & "Razlog" & vbCrLf & gi
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Trapez")
        Else
            MsgBox(ispis, MsgBoxStyle.Information, "Trapez")
        End If
pog:
prekid:
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Dim alp, bet, gam, fi, a, b, c, d1, d2, d3, D, O, V As Double
        Dim greska As Boolean
        Dim unos, ispis As String
        a = 0
        b = 0
        c = 0
        d1 = 0
        d2 = 0
        d3 = 0
        V = 0
        O = 0
        D = 0
        alp = -1
        bet = -1
        gam = -1
        fi = -1
        greska = False
        ispis = ""
        rispis = ""
        gispis = ""
        pispis = ""

        ' Unos

        If CheckBox30.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice a u kvadru:", "Unos podataka")
                If IsNumeric(unos) = True Then a = unos
                If unos = "" Then GoTo prekid
            Loop While a <= 0
            rispis = rispis & "Stranica a = " & FormatNumber(a, 4) & vbCrLf
        End If
        If CheckBox44.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice b u kvadru:", "Unos podataka")
                If IsNumeric(unos) = True Then b = unos
                If unos = "" Then GoTo prekid
            Loop While b <= 0
            rispis = rispis & "Stranica b = " & FormatNumber(b, 4) & vbCrLf
        End If
        If CheckBox37.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice c u kvadru:", "Unos podataka")
                If IsNumeric(unos) = True Then c = unos
                If unos = "" Then GoTo prekid
            Loop While c <= 0
            rispis = rispis & "Stranica c = " & FormatNumber(c, 4) & vbCrLf
        End If
        If CheckBox45.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale d1 u kvadru:", "Unos podataka")
                If IsNumeric(unos) = True Then d1 = unos
                If unos = "" Then GoTo prekid
            Loop While d1 <= 0
            rispis = rispis & "Dijagonala d1 = " & FormatNumber(d1, 4) & vbCrLf
        End If
        If CheckBox50.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale d2 u kvadru:", "Unos podataka")
                If IsNumeric(unos) = True Then d2 = unos
                If unos = "" Then GoTo prekid
            Loop While d2 <= 0
            rispis = rispis & "Dijagonala d2 = " & FormatNumber(d2, 4) & vbCrLf
        End If
        If CheckBox51.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale d3 u kvadru:", "Unos podataka")
                If IsNumeric(unos) = True Then d3 = unos
                If unos = "" Then GoTo prekid
            Loop While d3 <= 0
            rispis = rispis & "Dijagonala d3 = " & FormatNumber(d3, 4) & vbCrLf
        End If
        If CheckBox61.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu prostorne dijagonale D u kvadru:", "Unos podataka")
                If IsNumeric(unos) = True Then D = unos
                If unos = "" Then GoTo prekid
            Loop While D <= 0
            rispis = rispis & "Dijagonala D = " & FormatNumber(D, 4) & vbCrLf
        End If
        If CheckBox38.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            If alp >= System.Math.PI / 2 Then
                gispis = gispis & "α >= 90°" & vbCrLf
                greska = True
            End If
            rispis = rispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
        End If
        If CheckBox58.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta β:"
                Kut.Reset()
                Form2.ShowDialog()
                bet = Kut.Izracunaj()
                bet = (bet * System.Math.PI) / 180.0
                If bet = 0 Then GoTo prekid
            Loop While bet <= 0
            If bet >= System.Math.PI / 2 Then
                gispis = gispis & "β >= 90°" & vbCrLf
                greska = True
            End If
            rispis = rispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
        End If
        If CheckBox59.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta ɣ:"
                Kut.Reset()
                Form2.ShowDialog()
                gam = Kut.Izracunaj()
                gam = (gam * System.Math.PI) / 180.0
                If gam = 0 Then GoTo prekid
            Loop While gam <= 0
            If gam >= System.Math.PI / 2 Then
                gispis = gispis & "ɣ >= 90°" & vbCrLf
                greska = True
            End If
            rispis = rispis & "Kut ɣ = " & Kut_Ispis(gam, False) & vbCrLf
        End If
        If CheckBox60.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta φ:"
                Kut.Reset()
                Form2.ShowDialog()
                fi = Kut.Izracunaj()
                fi = (fi * System.Math.PI) / 180.0
                If fi = 0 Then GoTo prekid
            Loop While fi <= 0
            If fi >= System.Math.PI / 2 Then
                gispis = gispis & "φ >= 90°" & vbCrLf
                greska = True
            End If
            rispis = rispis & "Kut φ = " & Kut_Ispis(fi, False) & vbCrLf
        End If
        rispis = rispis & vbCrLf
        pispis = rispis
        ' Racun

        For i = 0 To 3
            If a = 0 Then
                If c > 0 And d1 > 0 Then
                    If d1 * d1 - c * c <= 0 Then
                        gispis = gispis & "d1 <= c" & vbCrLf
                        greska = True
                    End If
                    a = Sqrt(d1 * d1 - c * c)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(d1*d1 - c*c)" & vbCrLf
                ElseIf b > 0 And d2 > 0 Then
                    If d2 * d2 - b * b <= 0 Then
                        gispis = gispis & "d2 <= b" & vbCrLf
                        greska = True
                    End If
                    a = Sqrt(d2 * d2 - b * b)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(d2*d2 - b*b)" & vbCrLf
                ElseIf D > 0 And b > 0 And c > 0 Then
                    If D * D - b * b - c * c <= 0 Then
                        gispis = gispis & "D - b - c <= 0" & vbCrLf
                        greska = True
                    End If
                    a = Sqrt(D * D - b * b - c * c)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(D*D - b*b - c*c)" & vbCrLf
                ElseIf c > 0 And gam > 0 Then
                    a = c / Tan(gam)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = c / tg(ɣ)" & vbCrLf
                ElseIf b > 0 And alp > 0 Then
                    a = b / Tan(alp)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = b / tg(α)" & vbCrLf
                End If
            End If
            If b = 0 Then
                If a > 0 And d2 > 0 Then
                    If d2 * d2 - a * a <= 0 Then
                        gispis = gispis & "d2 <= a" & vbCrLf
                        greska = True
                    End If
                    b = Sqrt(d2 * d2 - a * a)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(d2*d2 - a*a)" & vbCrLf
                ElseIf c > 0 And d3 > 0 Then
                    If d3 * d3 - c * c <= 0 Then
                        gispis = gispis & "d3 <= c" & vbCrLf
                        greska = True
                    End If
                    b = Sqrt(d3 * d3 - c * c)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(d3*d3 - c*c)" & vbCrLf
                ElseIf D > 0 And a > 0 And c > 0 Then
                    If D * D - a * a - c * c <= 0 Then
                        gispis = gispis & "D - a - c <= 0" & vbCrLf
                        greska = True
                    End If
                    b = Sqrt(D * D - a * a - c * c)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(D*D - a*a - c*c)" & vbCrLf
                ElseIf a > 0 And alp > 0 Then
                    b = a * Tan(alp)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = a*tg(α)" & vbCrLf
                ElseIf c > 0 And bet > 0 Then
                    b = c / Tan(bet)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = c / tg(β)" & vbCrLf
                End If
            End If
            If c = 0 Then
                If a > 0 And d1 > 0 Then
                    If d1 * d1 - a * a <= 0 Then
                        gispis = gispis & "d1 <= a" & vbCrLf
                        greska = True
                    End If
                    c = Sqrt(d1 * d1 - a * a)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = sqrt(d1*d1 - a*a)" & vbCrLf
                ElseIf b > 0 And d3 > 0 Then
                    If d3 * d3 - b * b <= 0 Then
                        gispis = gispis & "d3 <= b" & vbCrLf
                        greska = True
                    End If
                    c = Sqrt(d3 * d3 - b * b)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = sqrt(d3*d3 - b*b)" & vbCrLf
                ElseIf D > 0 And a > 0 And b > 0 Then
                    If D * D - a * a - b * b <= 0 Then
                        gispis = gispis & "D - a - b <= 0" & vbCrLf
                        greska = True
                    End If
                    c = Sqrt(D * D - a * a - b * b)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = sqrt(D*D - a*a - b*b)" & vbCrLf
                ElseIf a > 0 And gam > 0 Then
                    c = a * Tan(gam)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = a*tg(ɣ)" & vbCrLf
                ElseIf b > 0 And bet > 0 Then
                    c = b * Tan(bet)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = b*tg(β)" & vbCrLf
                End If
            End If
            If d1 = 0 Then
                If a > 0 And c > 0 Then
                    d1 = Sqrt(c * c + a * a)
                    rispis = rispis & "d1 = " & FormatNumber(d1, 4) & ", d1 = sqrt(c*c + a*a)" & vbCrLf
                ElseIf c > 0 And gam > 0 Then
                    d1 = c / Sin(gam)
                    rispis = rispis & "d1 = " & FormatNumber(d1, 4) & ", d1 = c / sin(γ)" & vbCrLf
                ElseIf a > 0 And gam > 0 Then
                    d1 = a / Cos(gam)
                    rispis = rispis & "d1 = " & FormatNumber(d1, 4) & ", d1 = a / cos(γ)" & vbCrLf
                End If
            End If
            If d2 = 0 Then
                If a > 0 And b > 0 Then
                    d2 = Sqrt(b * b + a * a)
                    rispis = rispis & "d2 = " & FormatNumber(d2, 4) & ", d2 = sqrt(b*b + a*a)" & vbCrLf
                ElseIf b > 0 And alp > 0 Then
                    d2 = b / Sin(alp)
                    rispis = rispis & "d2 = " & FormatNumber(d2, 4) & ", d2 = b / sin(α)" & vbCrLf
                ElseIf a > 0 And alp > 0 Then
                    d2 = a / Cos(alp)
                    rispis = rispis & "d2 = " & FormatNumber(d2, 4) & ", d2 = a / cos(α)" & vbCrLf
                End If
            End If
            If d3 = 0 Then
                If b > 0 And c > 0 Then
                    d3 = Sqrt(c * c + b * b)
                    rispis = rispis & "d3 = " & FormatNumber(d3, 4) & ", d3 = sqrt(c*c + b*b)" & vbCrLf
                ElseIf c > 0 And bet > 0 Then
                    d3 = c / Sin(bet)
                    rispis = rispis & "d3 = " & FormatNumber(d3, 4) & ", d3 = c / sin(β)" & vbCrLf
                ElseIf b > 0 And bet > 0 Then
                    d3 = b / Cos(bet)
                    rispis = rispis & "d3 = " & FormatNumber(d3, 4) & ", d3 = b / cos(β)" & vbCrLf
                End If
            End If
            If D = 0 Then
                If b > 0 And c > 0 And a > 0 Then
                    D = Sqrt(a * a + c * c + b * b)
                    rispis = rispis & "D = " & FormatNumber(D, 4) & ", D = sqrt(a*a + b*b + c*c)" & vbCrLf
                ElseIf d2 > 0 And fi > 0 Then
                    D = d2 / Cos(fi)
                    rispis = rispis & "D = " & FormatNumber(D, 4) & ", D = d2 / cos(φ)" & vbCrLf
                ElseIf c > 0 And fi > 0 Then
                    D = c / Sin(fi)
                    rispis = rispis & "D = " & FormatNumber(D, 4) & ", D = c / sin(φ)" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If b > 0 And a > 0 Then
                    alp = Atan(b / a)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arctg(b / a)" & vbCrLf
                ElseIf d2 > 0 And a > 0 Then
                    If a / d2 >= 1 Then
                        gispis = gispis & "a / d2 >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Acos(a / d2)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos(a / d2)" & vbCrLf
                ElseIf d2 > 0 And b > 0 Then
                    If b / d2 >= 1 Then
                        gispis = gispis & "b / d2 >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Asin(b / d2)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(b / d2)" & vbCrLf
                End If
            End If
            If bet = -1 Then
                If c > 0 And b > 0 Then
                    bet = Atan(c / b)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arctg(c / b)" & vbCrLf
                ElseIf d3 > 0 And c > 0 Then
                    If c / d3 >= 1 Then
                        greska = True
                        gispis = gispis & "c / d3 >= 1" & vbCrLf
                    End If
                    bet = Asin(c / d3)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(c / d3)" & vbCrLf
                ElseIf d3 > 0 And b > 0 Then
                    If b / d3 >= 1 Then
                        greska = True
                        gispis = gispis & "b / d3 >= 1" & vbCrLf
                    End If
                    bet = Acos(b / d3)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos(b / d3)" & vbCrLf
                End If
            End If
            If gam = -1 Then
                If a > 0 And c > 0 Then
                    gam = Atan(c / a)
                    If greska = False Then rispis = rispis & "γ = " & Kut_Ispis(gam, False) & ", γ = arctg(c / a)" & vbCrLf
                ElseIf d1 > 0 And a > 0 Then
                    If a / d1 >= 1 Then
                        greska = True
                        gispis = gispis & "a / d1 >= 1" & vbCrLf
                    End If
                    gam = Acos(a / d1)
                    If greska = False Then rispis = rispis & "γ = " & Kut_Ispis(gam, False) & ", γ = arccos(a / d1)" & vbCrLf
                ElseIf d1 > 0 And c > 0 Then
                    If c / d1 >= 1 Then
                        gispis = gispis & "c / d1 >= 1" & vbCrLf
                        greska = True
                    End If
                    gam = Asin(c / d1)
                    If greska = False Then rispis = rispis & "γ = " & Kut_Ispis(gam, False) & ", γ = arcsin(c / d1)" & vbCrLf
                End If
            End If
            If fi = -1 Then
                If d2 > 0 And c > 0 Then
                    fi = Atan(c / d2)
                    If greska = False Then rispis = rispis & "φ = " & Kut_Ispis(fi, False) & ", φ = arctg(c / d2)" & vbCrLf
                ElseIf D > 0 And c > 0 Then
                    If c / D >= 1 Then
                        gispis = gispis & "c / D >= 1" & vbCrLf
                        greska = True
                    End If
                    fi = Asin(c / D)
                    If greska = False Then rispis = rispis & "φ = " & Kut_Ispis(fi, False) & ", φ = arcsin(c / D)" & vbCrLf
                ElseIf D > 0 And d2 > 0 Then
                    If d2 / D >= 1 Then
                        gispis = gispis & "d2 / D >= 1" & vbCrLf
                        greska = True
                    End If
                    fi = Acos(d2 / D)
                    If greska = False Then rispis = rispis & "φ = " & Kut_Ispis(fi, False) & ", φ = arccos(d2 / D)" & vbCrLf
                End If
            End If
        Next
        rispis = rispis & vbCrLf

        If a > 0 And b > 0 And c > 0 Then
            V = a * b * c
            rispis = rispis & "V = " & FormatNumber(V, 4) & ", V = a*b*c" & vbCrLf
            O = 2 * (a * b + a * c + b * c)
            rispis = rispis & "O = " & FormatNumber(O, 4) & ", O = 2*(a*b + a*c + b*c)" & vbCrLf
        End If

        If V = 0 Or O = 0 Then
            rispis = rispis & "Niste unijeli dovoljno podataka za izracunavanje volumena ili oplosja!" & vbCrLf
        End If

        ' Provjera

        If a > 0 And c > 0 And d1 > 0 Then
            If a * a + c * c < d1 * d1 - 0.00001 Or a * a + c * c > d1 * d1 + 0.00001 Then
                gispis = gispis & "a*a + c*c != d1*d1" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And b > 0 And d2 > 0 Then
            If a * a + b * b < d2 * d2 - 0.00001 Or a * a + b * b > d2 * d2 + 0.00001 Then
                gispis = gispis & "a*a + b*b != d2*d2" & vbCrLf
                greska = True
            End If
        End If
        If c > 0 And b > 0 And d3 > 0 Then
            If c * c + b * b < d3 * d3 - 0.00001 Or c * c + b * b > d3 * d3 + 0.00001 Then
                gispis = gispis & "c*c + b*b != d3*d3" & vbCrLf
                greska = True
            End If
        End If
        If c > 0 And d2 > 0 And D > 0 Then
            If d2 * d2 + c * c < D * D - 0.00001 Or d2 * d2 + c * c > D * D + 0.00001 Then
                gispis = gispis & "d2*d2 + c*c != D*D" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And b > 0 And c > 0 And D > 0 Then
            If a * a + b * b + c * c < D * D - 0.00001 Or a * a + b * b + c * c > D * D + 0.00001 Then
                gispis = gispis & "a*a + b*b + c*c != D*D" & vbCrLf
                greska = True
            End If
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav kvadar ne postoji" & vbCrLf
            If gispis.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gispis
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Kvadar")
        Else
            MsgBox(rispis, MsgBoxStyle.Information, "Kvadar")
        End If
prekid:
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Dim alp, bet, a, b, h, h1, O, V, pl, Baz As Double
        Dim greska As Boolean
        Dim unos, ispis, gi As String
        a = 0
        b = 0
        V = 0
        O = 0
        pl = 0
        Baz = 0
        h = 0
        h1 = 0
        alp = -1
        bet = -1
        greska = False
        ispis = ""
        rispis = ""
        pispis = ""
        gispis = ""
        gi = ""

        ' Unos

        If CheckBox72.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice a u trostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then a = unos
                If unos = "" Then GoTo prekid
            Loop While a <= 0
            ispis = ispis & "Stranica a = " & FormatNumber(a, 4) & vbCrLf
        End If
        If CheckBox69.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice b u trostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then b = unos
                If unos = "" Then GoTo prekid
            Loop While b <= 0
            ispis = ispis & "Stranica b = " & FormatNumber(b, 4) & vbCrLf
        End If
        If CheckBox68.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h u trostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then h = unos
                If unos = "" Then GoTo prekid
            Loop While h <= 0
            ispis = ispis & "Visina h = " & FormatNumber(h, 4) & vbCrLf
        End If
        If CheckBox62.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h1 u trostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then h1 = unos
                If unos = "" Then GoTo prekid
            Loop While h1 <= 0
            ispis = ispis & "Visina h1 = " & FormatNumber(h1, 4) & vbCrLf
        End If
        If CheckBox70.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta uz stranicu α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            If alp >= System.Math.PI / 2 Then
                greska = True
                gi = gi & "α >= 90°" & vbCrLf
            End If
            ispis = ispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
        End If
        If CheckBox65.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta uz brid β:"
                Kut.Reset()
                Form2.ShowDialog()
                bet = Kut.Izracunaj()
                bet = (bet * System.Math.PI) / 180.0
                If bet = 0 Then GoTo prekid
            Loop While bet <= 0
            If bet >= System.Math.PI / 2 Then
                greska = True
                gi = gi & "β >= 90°" & vbCrLf
            End If
            ispis = ispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
        End If
        ispis = ispis & vbCrLf
        pispis = ispis
        'Racun

        For i = 0 To 3
            If a = 0 Then
                If h > 0 And h1 > 0 Then
                    a = Sqrt(12 * h1 * h1 - 12 * h * h)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(12*h1*h1 - 12*h*h)" & vbCrLf
                ElseIf h > 0 And b > 0 Then
                    a = Sqrt(3 * b * b - 3 * h * h)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(3*b*b - 3*h*h)" & vbCrLf
                ElseIf b > 0 And h1 > 0 Then
                    a = 2 * Sqrt(b * b - h1 * h1)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = 2*sqrt(b*b - h1*h1)" & vbCrLf
                End If
            End If
            If b = 0 Then
                If h > 0 And a > 0 Then
                    b = Sqrt(h * h + (a * a) / 3)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(h*h + (a*a) / 3)" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    b = Sqrt((a / 2) * (a / 2) + h1 * h1)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt((a/2)*(a/2) + h1*h1)" & vbCrLf
                ElseIf a > 0 And bet > 0 And alp > 0 Then
                    Trokut(Sqrt(3) * a / 2, b, 0, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, greska)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = ((a*sqrt(3)) / 2)*sin(α) / sin(180° - α - β)" & vbCrLf
                ElseIf a > 0 And bet > 0 And h1 > 0 Then
                    Trokut(Sqrt(3) * a / 2, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, bet, 0, 0, 0, greska)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = ((a*sqrt(3)) / 2)*cos(β) + sqrt(h1*h1 - ((a*sqrt(3)) / 2)*((a*sqrt(3)) / 2)*sin(β)*sin(β))" & vbCrLf
                ElseIf h1 > 0 And bet > 0 And alp > 0 Then
                    Trokut(0, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, greska)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = h1*sin(α) / sin(β)" & vbCrLf
                ElseIf a > 0 And h1 > 0 And alp > 0 Then
                    Trokut(Sqrt(3) * a / 2, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, greska)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(((a*sqrt(3)) / 2)*((a*sqrt(3)) / 2) + h1*h1 - 2*((a*sqrt(3)) / 2)*h1*cos(α))" & vbCrLf
                End If
            End If
            If h = 0 Then
                If alp > 0 And h1 > 0 Then
                    h = h1 * Sin(alp)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = h1*sin(α)" & vbCrLf
                ElseIf a > 0 And alp > 0 Then
                    h = (Sqrt(3) * a / 2) * Tan(alp)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = (sqrt(3)*a / 2)*tg(α)" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    h = Sqrt(h1 * h1 - (Sqrt(3) * a / 6) * (Sqrt(3) * a / 6))
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = sqrt(h1*h1 - (Sqrt(3) * a / 6)*(Sqrt(3) * a / 6))" & vbCrLf
                ElseIf b > 0 And a > 0 Then
                    h = Sqrt(b * b - (Sqrt(3) * a / 3) * (Sqrt(3) * a / 3))
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = sqrt(b*b - (Sqrt(3) * a / 3)*(Sqrt(3) * a / 3))" & vbCrLf
                ElseIf b > 0 And bet > 0 Then
                    h = b * Sin(bet)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = b*sin(β)" & vbCrLf
                ElseIf bet > 0 And a > 0 Then
                    h = Tan(bet) * (Sqrt(3) * a / 3)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = tg(β)*sqrt(sqrt(3)*a / 3)" & vbCrLf
                End If
            End If
            If h1 = 0 Then
                If h > 0 And a > 0 Then
                    h1 = Sqrt(h * h + (Sqrt(3) * a / 6) * (Sqrt(3) * a / 6))
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = sqrt(h*h + (Sqrt(3) * a / 6)*(Sqrt(3) * a / 6))" & vbCrLf
                ElseIf b > 0 And a > 0 Then
                    h1 = Sqrt(b * b - (a / 2) * (a / 2))
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = sqrt(b*b - (a/2)*(a/2))" & vbCrLf
                ElseIf a > 0 And alp > 0 Then
                    h1 = (Sqrt(3) * a / 6) / Cos(alp)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = (sqrt(3)*a / 6) / cos(α)" & vbCrLf
                ElseIf h > 0 And alp > 0 Then
                    h1 = h / Sin(alp)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = h / sin(α)" & vbCrLf
                ElseIf b > 0 And a > 0 And bet > 0 Then
                    Trokut(Sqrt(3) * a / 2, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, bet, 0, 0, 0, greska)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = sqrt(b*b + (sqrt(3)*a / 2)*(sqrt(3)*a / 2) - 2*b*(sqrt(3)*a / 2)*cos(β))" & vbCrLf
                ElseIf b > 0 And a > 0 And alp > 0 Then
                    Trokut(Sqrt(3) * a / 2, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, greska)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = ((a*sqrt(3)) / 2)*cos(α) + sqrt(b*b - ((a*sqrt(3)) / 2)*((a*sqrt(3)) / 2)*sin(α)*sin(α))" & vbCrLf
                ElseIf b > 0 And alp > 0 And bet > 0 Then
                    Trokut(0, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, greska)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = (b*sin(β)) / sin(α)" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If h > 0 And h1 > 0 Then
                    If h / h1 >= 1 Then
                        gi = gi & "h / h1 >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Asin(h / h1)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(h / h1)" & vbCrLf
                ElseIf a > 0 And h > 0 Then
                    alp = Atan(h / (Sqrt(3) * a / 6))
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arctg(h / (sqrt(3)*a / 6))" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    If (Sqrt(3) * a / 6) / h1 >= 1 Then
                        gi = gi & "(Sqrt(3)*a / 6) / h1 >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Acos((Sqrt(3) * a / 6) / h1)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos((sqrt(3)*a / 6) / h1)" & vbCrLf
                ElseIf b > 0 And a > 0 And h1 > 0 Then
                    Trokut(a, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, -1, 0, 0, 0, greska)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos((((a*sqrt(3)) / 2)*((a*sqrt(3)) / 2) + h1*h1 - b*b) / (2*h1*((a*sqrt(3)) / 2)))" & vbCrLf
                ElseIf b > 0 And h1 > 0 And bet > 0 Then
                    Trokut(0, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, greska)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = b*sin(β) / h1" & vbCrLf
                End If
            End If
            If bet = -1 Then
                If h > 0 And b > 0 Then
                    If h / b >= 1 Then
                        gi = gi & "h / b >= 1" & vbCrLf
                        greska = True
                    End If
                    bet = Asin(h / b)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(h / b)" & vbCrLf
                ElseIf a > 0 And h > 0 Then
                    bet = Atan(h / (Sqrt(3) * a / 3))
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arctg(h / (Sqrt(3)*a / 3))" & vbCrLf
                ElseIf a > 0 And b > 0 Then
                    If (Sqrt(3) * a / 3) / b >= 1 Then
                        gi = gi & "(Sqrt(3)*a / 3) / b >= 1" & vbCrLf
                        greska = True
                    End If
                    bet = Acos((Sqrt(3) * a / 3) / b)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos((Sqrt(3)*a / 3) / b)" & vbCrLf
                ElseIf b > 0 And a > 0 And h1 > 0 Then
                    Trokut(a, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, -1, bet, 0, 0, 0, greska)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos((((a*sqrt(3)) / 2)*((a*sqrt(3)) / 2) + b*b - h1*h1) / (2*b*((a*sqrt(3)) / 2)))" & vbCrLf
                ElseIf b > 0 And h1 > 0 And alp > 0 Then
                    Trokut(0, b, h1, 0, 0, 0, -1, -1, 0, 0, 0, 0, 0, -1, alp, bet, 0, 0, 0, greska)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = (h1*sin(α)) / b" & vbCrLf
                End If
            End If
        Next
        ispis = ispis & vbCrLf

        If a > 0 And h > 0 Then
            V = (a * a * Sqrt(3) * h) / 12
            ispis = ispis & "V = " & FormatNumber(V, 4) & ", V = (a*a*sqrt(3)*h) / 12" & vbCrLf
        End If
        If a > 0 And h1 > 0 Then
            O = (a * a * Sqrt(3)) / 4 + (3 * a * h1) / 2
            ispis = ispis & "O = " & FormatNumber(O, 4) & ", O = (a*a*sqrt(3)) / 4 + (3*a*h1) / 2" & vbCrLf
            pl = (3 * a * h1) / 2
            ispis = ispis & "Plast = " & FormatNumber(pl, 4) & ", P = (3*a*h1) / 2" & vbCrLf
        End If
        If a > 0 Then
            Baz = (a * a * Sqrt(3)) / 4
            ispis = ispis & "Baza = " & FormatNumber(Baz, 4) & ", B = (a*a*sqrt(3)) / 4" & vbCrLf
        End If

        If V = 0 Or O = 0 Then
            ispis = ispis & "Niste unijeli dovoljno podataka za izracunavanje volumena ili oplosja!" & vbCrLf
        End If

        ' Provjera

        If a > 0 And b > 0 And h > 0 Then
            If (Sqrt(3) * a / 3) * (Sqrt(3) * a / 3) + h * h < b * b - 0.00001 Or (Sqrt(3) * a / 3) * (Sqrt(3) * a / 3) + h * h > b * b + 0.00001 Then
                gi = gi & "(Sqrt(3)*a / 3)*(Sqrt(3)*a / 3) + h*h != b*b" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And h1 > 0 And h > 0 Then
            If (Sqrt(3) * a / 6) * (Sqrt(3) * a / 6) + h * h < h1 * h1 - 0.00001 Or (Sqrt(3) * a / 6) * (Sqrt(3) * a / 6) + h * h > h1 * h1 + 0.00001 Then
                gi = gi & "(Sqrt(3)*a / 6)*(Sqrt(3)*a / 6) + h*h != h1*h1" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And h1 > 0 And b > 0 Then
            If (a / 2) * (a / 2) + h1 * h1 < b * b - 0.00001 Or (a / 2) * (a / 2) + h1 * h1 > b * b + 0.00001 Then
                gi = gi & "(a / 2)*(a / 2) + h1*h1 != b*b" & vbCrLf
                greska = True
            End If
        End If
        If bet > 0 And h > 0 And b > 0 Then
            If h / b < Sin(bet) - 0.00001 Or h / b > Sin(bet) + 0.00001 Then
                gi = gi & "h / b != Sin(β)" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And h > 0 And h1 > 0 Then
            If h / h1 < Sin(alp) - 0.00001 Or h / h1 > Sin(alp) + 0.00001 Then
                gi = gi & "h / h1 != Sin(α)" & vbCrLf
                greska = True
            End If
        End If
        If h > 0 And h1 > 0 Then
            If h1 <= h Then
                gi = gi & "h1 <= h" & vbCrLf
                greska = True
            End If
        End If
        If h > 0 And b > 0 Then
            If b <= h Then
                gi = gi & "b <= h" & vbCrLf
                greska = True
            End If
        End If
        If b > 0 And h1 > 0 Then
            If b <= h1 Then
                gi = gi & "b <= h1" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And bet > 0 Then
            If bet >= alp Then
                gi = gi & "β >= α" & vbCrLf
                greska = True
            End If
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takva trostrana piramida ne postoji" & vbCrLf
            If gi.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gi
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Trostrana piramida")
        Else
            MsgBox(ispis, MsgBoxStyle.Information, "Trostrana piramida")
        End If
prekid:
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Dim alp, bet, a, b, d, h, h1, O, V, pl, Baz As Double
        Dim greska As Boolean
        Dim unos, ispis, gi As String
        a = 0
        b = 0
        V = 0
        d = 0
        O = 0
        pl = 0
        Baz = 0
        h = 0
        h1 = 0
        alp = -1
        bet = -1
        greska = False
        ispis = ""
        rispis = ""
        pispis = ""
        gispis = ""
        gi = ""

        ' Unos

        If CheckBox73.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice a u cetverostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then a = unos
                If unos = "" Then GoTo prekid
            Loop While a <= 0
            ispis = ispis & "Stranica a = " & FormatNumber(a, 4) & vbCrLf
        End If
        If CheckBox67.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice b u cetverostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then b = unos
                If unos = "" Then GoTo prekid
            Loop While b <= 0
            ispis = ispis & "Stranica b = " & FormatNumber(b, 4) & vbCrLf
        End If
        If CheckBox66.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h u cetverostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then h = unos
                If unos = "" Then GoTo prekid
            Loop While h <= 0
            ispis = ispis & "Visina h = " & FormatNumber(h, 4) & vbCrLf
        End If
        If CheckBox63.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h1 u cetverostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then h1 = unos
                If unos = "" Then GoTo prekid
            Loop While h1 <= 0
            ispis = ispis & "Visina h1 = " & FormatNumber(h1, 4) & vbCrLf
        End If
        If CheckBox80.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale d u cetverostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then d = unos
                If unos = "" Then GoTo prekid
            Loop While d <= 0
            ispis = ispis & "Dijagonala d = " & FormatNumber(d, 4) & vbCrLf
        End If
        If CheckBox71.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta uz stranicu α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            If alp >= System.Math.PI / 2 Then
                greska = True
                gi = gi & "α >= 90°" & vbCrLf
            End If
            ispis = ispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
        End If
        If CheckBox64.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta uz brid β:"
                Kut.Reset()
                Form2.ShowDialog()
                bet = Kut.Izracunaj()
                bet = (bet * System.Math.PI) / 180.0
                If bet = 0 Then GoTo prekid
            Loop While bet <= 0
            If bet >= System.Math.PI / 2 Then
                greska = True
                gi = gi & "β >= 90°" & vbCrLf
            End If
            ispis = ispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
        End If
        ispis = ispis & vbCrLf
        pispis = ispis
        'Racun

        For i = 0 To 3
            If a = 0 Then
                If h > 0 And h1 > 0 Then
                    a = 2 * Sqrt(h1 * h1 - h * h)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = 2*sqrt(h1*h1 - h*h)" & vbCrLf
                ElseIf d > 0 Then
                    a = d / Sqrt(2)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = d / sqrt(2)" & vbCrLf
                ElseIf h > 0 And b > 0 Then
                    a = Sqrt(2 * b * b - 2 * h * h)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(2*b*b - 2*h*h)" & vbCrLf
                ElseIf b > 0 And h1 > 0 Then
                    a = 2 * Sqrt(b * b - h1 * h1)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = 2*sqrt(b*b - h1*h1)" & vbCrLf
                End If
            End If
            If b = 0 Then
                If h > 0 And a > 0 Then
                    b = Sqrt(h * h + (a * a) / 2)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(h*h + (a*a) / 2)" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    b = Sqrt((a / 2) * (a / 2) + h1 * h1)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt((a / 2)*(a / 2) + h1*h1)" & vbCrLf
                ElseIf a > 0 And bet > 0 Then
                    b = ((Sqrt(2) / 2) * a) / Cos(bet)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = ((Sqrt(2) / 2)*a) / cos(β)" & vbCrLf
                ElseIf h > 0 And bet > 0 Then
                    b = h / Sin(bet)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = h / sin(β)" & vbCrLf
                ElseIf a > 0 And h > 0 Then
                    b = Sqrt(((Sqrt(2) / 2) * a) * ((Sqrt(2) / 2) * a) + h * h)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = Sqrt(((Sqrt(2) / 2)*a)*((Sqrt(2) / 2)*a) + h*h)" & vbCrLf
                End If
            End If
            If h = 0 Then
                If alp > 0 And h1 > 0 Then
                    h = h1 * Sin(alp)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = h1*sin(α)" & vbCrLf
                ElseIf a > 0 And alp > 0 Then
                    h = (a / 2) * Tan(alp)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = (a / 2)*tg(α)" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    h = Sqrt(h1 * h1 - (a / 2) * (a / 2))
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = sqrt(h1*h1 - (a / 2)*(a / 2))" & vbCrLf
                ElseIf b > 0 And bet > 0 Then
                    h = b * Sin(bet)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = b*sin(β)" & vbCrLf
                ElseIf bet > 0 And a > 0 Then
                    h = Tan(bet) * ((Sqrt(2) / 2) * a)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = tg(β)*((Sqrt(2) / 2)*a)" & vbCrLf
                ElseIf b > 0 And a > 0 Then
                    h = Sqrt(b * b - ((Sqrt(2) / 2) * a) * ((Sqrt(2) / 2) * a))
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = Sqrt(b*b - ((Sqrt(2) / 2)*a)*((Sqrt(2) / 2)*a))" & vbCrLf
                End If
            End If
            If h1 = 0 Then
                If h > 0 And a > 0 Then
                    h1 = Sqrt(h * h + (a / 2) * (a / 2))
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = Sqrt(h*h + (a / 2)*(a / 2))" & vbCrLf
                ElseIf b > 0 And a > 0 Then
                    h1 = Sqrt(b * b - (a / 2) * (a / 2))
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = Sqrt(b*b - (a / 2)*(a / 2))" & vbCrLf
                ElseIf a > 0 And alp > 0 Then
                    h1 = (a / 2) / Cos(alp)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = (a / 2) / cos(α)" & vbCrLf
                ElseIf h > 0 And alp > 0 Then
                    h1 = h / Sin(alp)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = h / sin(α)" & vbCrLf
                End If
            End If
            If d = 0 Then
                If a > 0 Then
                    d = a * Sqrt(2)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = a*sqrt(2)" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If h > 0 And h1 > 0 Then
                    If h / h1 >= 1 Then
                        gi = gi & "h / h1 >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Asin(h / h1)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(h / h1)" & vbCrLf
                ElseIf a > 0 And h > 0 Then
                    alp = Atan(h / (a / 2))
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arctg(h / (a / 2))" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    If (a / 2) / h1 >= 1 Then
                        gi = gi & "(a / 2) / h1 >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Acos((a / 2) / h1)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos((a / 2) / h1)" & vbCrLf
                End If
            End If
            If bet = -1 Then
                If h > 0 And b > 0 Then
                    If h / b >= 1 Then
                        gi = gi & "h / b >= 1" & vbCrLf
                        greska = True
                    End If
                    bet = Asin(h / b)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(h / b)" & vbCrLf
                ElseIf a > 0 And h > 0 Then
                    bet = Atan(h / ((Sqrt(2) / 2) * a))
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arctg(h / ((Sqrt(2) / 2)*a))" & vbCrLf
                ElseIf a > 0 And b > 0 Then
                    If ((Sqrt(2) / 2) * a) / b >= 1 Then
                        gi = gi & "((Sqrt(2) / 2)*a) / b >= 1" & vbCrLf
                        greska = True
                    End If
                    bet = Acos(((Sqrt(2) / 2) * a) / b)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos(((Sqrt(2) / 2)*a) / b)" & vbCrLf
                End If
            End If
        Next
        ispis = ispis & vbCrLf

        If a > 0 And h > 0 Then
            V = (a * a * h) / 3
            ispis = ispis & "V = " & FormatNumber(V, 4) & ", V = (a*a*h) / 3" & vbCrLf
        End If
        If a > 0 And h1 > 0 Then
            O = a * a + 2 * a * h1
            ispis = ispis & "O = " & FormatNumber(O, 4) & ", O = a*a + 2*a*h1" & vbCrLf
            pl = 2 * a * h1
            ispis = ispis & "Plast = " & FormatNumber(pl, 4) & ", P = 2*a*h1" & vbCrLf
        End If
        If a > 0 Then
            Baz = a * a
            ispis = ispis & "Baza = " & FormatNumber(Baz, 4) & ", B = a*a" & vbCrLf
        End If

        If V = 0 Or O = 0 Then
            ispis = ispis & "Niste unijeli dovoljno podataka za izracunavanje volumena ili oplosja!" & vbCrLf
        End If

        ' Provjera

        If a > 0 And b > 0 And h > 0 Then
            If ((Sqrt(2) / 2) * a) * ((Sqrt(2) / 2) * a) + h * h < b * b - 0.00001 Or ((Sqrt(2) / 2) * a) * ((Sqrt(2) / 2) * a) + h * h > b * b + 0.00001 Then
                gi = gi & "((Sqrt(2) / 2)*a)*((Sqrt(2) / 2)*a) + h*h != b*b" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And h1 > 0 And h > 0 Then
            If (a / 2) * (a / 2) + h * h < h1 * h1 - 0.00001 Or (a / 2) * (a / 2) + h * h > h1 * h1 + 0.00001 Then
                gi = gi & "(a / 2)*(a / 2) + h*h != h1*h1" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And h1 > 0 And b > 0 Then
            If (a / 2) * (a / 2) + h1 * h1 < b * b - 0.00001 Or (a / 2) * (a / 2) + h1 * h1 > b * b + 0.00001 Then
                gi = gi & "(a / 2)*(a / 2) + h1*h1 != b*b" & vbCrLf
                greska = True
            End If
        End If
        If bet > 0 And h > 0 And b > 0 Then
            If h / b < Sin(bet) - 0.00001 Or h / b > Sin(bet) + 0.00001 Then
                gi = gi & "h / b != sin(β)" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And h > 0 And h1 > 0 Then
            If h / h1 < Sin(alp) - 0.00001 Or h / h1 > Sin(alp) + 0.00001 Then
                gi = gi & "h / h1 != sin(α)" & vbCrLf
                greska = True
            End If
        End If
        If h > 0 And h1 > 0 Then
            If h1 <= h Then
                gi = gi & "h1 < h" & vbCrLf
                greska = True
            End If
        End If
        If h > 0 And b > 0 Then
            If b <= h Then
                gi = gi & "b < h" & vbCrLf
                greska = True
            End If
        End If
        If b > 0 And h1 > 0 Then
            If b <= h1 Then
                gi = gi & "b < h" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And bet > 0 Then
            If bet >= alp Then
                gi = gi & "β > α" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And d > 0 Then
            If Sqrt(2) * a < d - 0.00001 Or Sqrt(2) * a > d + 0.00001 Then
                gi = gi & "Sqrt(2)*a != d" & vbCrLf
                greska = True
            End If
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takva cetverostrana piramida ne postoji" & vbCrLf
            If gi.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gi
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Cetverostrana piramida")
        Else
            MsgBox(ispis, MsgBoxStyle.Information, "Cetverostrana piramida")
        End If
prekid:
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        Dim alp, bet, a, b, h, h1, D, d1, O, V, pl, Baz As Double
        Dim greska As Boolean
        Dim unos, ispis, gi As String
        a = 0
        b = 0
        V = 0
        D = 0
        d1 = 0
        O = 0
        pl = 0
        Baz = 0
        h = 0
        h1 = 0
        alp = -1
        bet = -1
        greska = False
        ispis = ""
        rispis = ""
        gispis = ""
        pispis = ""
        gi = ""

        ' Unos

        If CheckBox79.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice a u sesterostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then a = unos
                If unos = "" Then GoTo prekid
            Loop While a <= 0
            ispis = ispis & "Stranica a = " & FormatNumber(a, 4) & vbCrLf
        End If
        If CheckBox77.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu stranice b u sesterostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then b = unos
                If unos = "" Then GoTo prekid
            Loop While b <= 0
            ispis = ispis & "Stranica b = " & FormatNumber(b, 4) & vbCrLf
        End If
        If CheckBox76.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h u sesterostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then h = unos
                If unos = "" Then GoTo prekid
            Loop While h <= 0
            ispis = ispis & "Visina h = " & FormatNumber(h, 4) & vbCrLf
        End If
        If CheckBox74.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h1 u sesterostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then h1 = unos
                If unos = "" Then GoTo prekid
            Loop While h1 <= 0
            ispis = ispis & "Visina h1 = " & FormatNumber(h1, 4) & vbCrLf
        End If
        If CheckBox81.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale D u sesterostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then D = unos
                If unos = "" Then GoTo prekid
            Loop While D <= 0
            ispis = ispis & "Dijagonala D = " & FormatNumber(D, 4) & vbCrLf
        End If
        If CheckBox82.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale d u sesterostranoj piramidi:", "Unos podataka")
                If IsNumeric(unos) = True Then d1 = unos
                If unos = "" Then GoTo prekid
            Loop While d1 <= 0
            ispis = ispis & "Dijagonala d1 = " & FormatNumber(d1, 4) & vbCrLf
        End If
        If CheckBox78.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta uz stranicu α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            If alp >= System.Math.PI / 2 Then
                greska = True
                gi = gi & "α >= 90°" & vbCrLf
            End If
            ispis = ispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
        End If
        If CheckBox75.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta uz brid β:"
                Kut.Reset()
                Form2.ShowDialog()
                bet = Kut.Izracunaj()
                bet = (bet * System.Math.PI) / 180.0
                If bet = 0 Then GoTo prekid
            Loop While bet <= 0
            If bet >= System.Math.PI / 2 Then
                greska = True
                gi = gi & "β >= 90°" & vbCrLf
            End If
            ispis = ispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
        End If
        ispis = ispis & vbCrLf
        pispis = ispis

        'Racun

        For i = 0 To 3
            If a = 0 Then
                If h > 0 And h1 > 0 Then
                    a = Sqrt((4 / 3) * (h1 * h1 - h * h))
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = sqrt((4/3)*(h1*h1 - h*h))" & vbCrLf
                ElseIf D > 0 Then
                    a = D / 2
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = D / 2" & vbCrLf
                ElseIf d1 > 0 Then
                    a = d1 / Sqrt(3)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = d1 / sqrt(3)" & vbCrLf
                ElseIf h > 0 And b > 0 Then
                    a = Sqrt(b * b - h * h)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(b*b - h*h)" & vbCrLf
                ElseIf b > 0 And h1 > 0 Then
                    a = 2 * Sqrt(b * b - h1 * h1)
                    ispis = ispis & "a = " & FormatNumber(a, 4) & ", a = 2*sqrt(b*b - h1*h1)" & vbCrLf
                End If
            End If
            If b = 0 Then
                If h > 0 And a > 0 Then
                    b = Sqrt(h * h + a * a)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(h*h + a*a)" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    b = Sqrt((a / 2) * (a / 2) + h1 * h1)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt((a / 2)*(a / 2) + h1*h1)" & vbCrLf
                ElseIf a > 0 And bet > 0 Then
                    b = a / Cos(bet)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = a / cos(β)" & vbCrLf
                ElseIf h > 0 And bet > 0 Then
                    b = h / Sin(bet)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = h / sin(β)" & vbCrLf
                ElseIf a > 0 And h > 0 Then
                    b = Sqrt(a * a + h * h)
                    ispis = ispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(a*a + h*h)" & vbCrLf
                End If
            End If
            If h = 0 Then
                If alp > 0 And h1 > 0 Then
                    h = h1 * Sin(alp)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = h1*sin(α)" & vbCrLf
                ElseIf a > 0 And alp > 0 Then
                    h = ((a * Sqrt(3)) / 2) * Tan(alp)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = ((a*Sqrt(3)) / 2)*Tg(α)" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    h = Sqrt(h1 * h1 - ((a * Sqrt(3)) / 2) * ((a * Sqrt(3)) / 2))
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = Sqrt(h1*h1 - ((a*Sqrt(3)) / 2)*((a*Sqrt(3)) / 2))" & vbCrLf
                ElseIf b > 0 And bet > 0 Then
                    h = b * Sin(bet)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = b*sin(β)" & vbCrLf
                ElseIf bet > 0 And a > 0 Then
                    h = Tan(bet) * a
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = a*tg(β)" & vbCrLf
                ElseIf b > 0 And a > 0 Then
                    h = Sqrt(b * b - a * a)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = sqrt(b*b - a*a)" & vbCrLf
                End If
            End If
            If h1 = 0 Then
                If h > 0 And a > 0 Then
                    h1 = Sqrt(h * h + ((a * Sqrt(3)) / 2) * ((a * Sqrt(3)) / 2))
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = Sqrt(h*h + ((a*Sqrt(3)) / 2)*((a*Sqrt(3)) / 2))" & vbCrLf
                ElseIf b > 0 And a > 0 Then
                    h1 = Sqrt(b * b - (a / 2) * (a / 2))
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = Sqrt(b*b - (a / 2)*(a / 2))" & vbCrLf
                ElseIf a > 0 And alp > 0 Then
                    h1 = ((a * Sqrt(3)) / 2) / Cos(alp)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = ((a*Sqrt(3)) / 2) / Cos(α)" & vbCrLf
                ElseIf h > 0 And alp > 0 Then
                    h1 = h / Sin(alp)
                    ispis = ispis & "h1 = " & FormatNumber(h1, 4) & ", h1 = h / sin(α)" & vbCrLf
                End If
            End If
            If D = 0 Then
                If a > 0 Then
                    D = a * 2
                    ispis = ispis & "D = " & FormatNumber(D, 4) & ", D = 2*a" & vbCrLf
                End If
            End If
            If d1 = 0 Then
                If a > 0 Then
                    d1 = a * Sqrt(3)
                    ispis = ispis & "d = " & FormatNumber(d1, 4) & ", d = sqrt(3)*a" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If h > 0 And h1 > 0 Then
                    If h / h1 >= 1 Then
                        gi = gi & "h / h1 >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Asin(h / h1)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(h / h1)" & vbCrLf
                ElseIf a > 0 And h > 0 Then
                    alp = Atan(h / ((a * Sqrt(3)) / 2))
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arctg(h / ((a*Sqrt(3)) / 2))" & vbCrLf
                ElseIf a > 0 And h1 > 0 Then
                    If ((a * Sqrt(3)) / 2) / h1 >= 1 Then
                        gi = gi & "((a*Sqrt(3)) / 2) / h1 >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Acos(((a * Sqrt(3)) / 2) / h1)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos(((a*Sqrt(3)) / 2) / h1)" & vbCrLf
                End If
            End If
            If bet = -1 Then
                If h > 0 And b > 0 Then
                    If h / b >= 1 Then
                        gi = gi & "h / b >= 1" & vbCrLf
                        greska = True
                    End If
                    bet = Asin(h / b)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(h / b)" & vbCrLf
                ElseIf a > 0 And h > 0 Then
                    bet = Atan(h / a)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arctg(h / a)" & vbCrLf
                ElseIf a > 0 And b > 0 Then
                    If a / b >= 1 Then
                        gi = gi & "b / b >= 1" & vbCrLf
                        greska = True
                    End If
                    bet = Acos(a / b)
                    If greska = False Then ispis = ispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos(a / b)" & vbCrLf
                End If
            End If
        Next
        ispis = ispis & vbCrLf

        If a > 0 And h > 0 Then
            V = (a * a * h * Sqrt(3)) / 2
            ispis = ispis & "V = " & FormatNumber(V, 4) & ", V = (a*a*h*Sqrt(3)) / 2" & vbCrLf
        End If
        If a > 0 And h1 > 0 Then
            O = ((3 * a * a * Sqrt(3)) / 2) + 3 * a * h1
            ispis = ispis & "O = " & FormatNumber(O, 4) & ", O = ((3*a*a*Sqrt(3)) / 2) + 3*a*h1" & vbCrLf
            pl = 3 * a * h1
            ispis = ispis & "Plast = " & FormatNumber(pl, 4) & ", P = 3*a*h1" & vbCrLf
        End If
        If a > 0 Then
            Baz = (3 * a * a * Sqrt(3)) / 2
            ispis = ispis & "Baza = " & FormatNumber(Baz, 4) & ", B = ((3*a*a*Sqrt(3)) / 2)" & vbCrLf
        End If

        If V = 0 Or O = 0 Then
            ispis = ispis & "Niste unijeli dovoljno podataka za izracunavanje volumena ili oplosja!" & vbCrLf
        End If

        ' Provjera

        If a > 0 And b > 0 And h > 0 Then
            If a * a + h * h < b * b - 0.00001 Or a * a + h * h > b * b + 0.00001 Then
                gi = gi & "a*a + h*h != b*b" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And h1 > 0 And h > 0 Then
            If ((a * Sqrt(3)) / 2) * ((a * Sqrt(3)) / 2) + h * h < h1 * h1 - 0.00001 Or ((a * Sqrt(3)) / 2) * ((a * Sqrt(3)) / 2) + h * h > h1 * h1 + 0.00001 Then
                gi = gi & "((a*Sqrt(3)) / 2)*((a*Sqrt(3)) / 2) + h*h != h1*h1" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And h1 > 0 And b > 0 Then
            If (a / 2) * (a / 2) + h1 * h1 < b * b - 0.00001 Or (a / 2) * (a / 2) + h1 * h1 > b * b + 0.00001 Then
                gi = gi & "(a / 2)*(a / 2) + h1*h1 != b*b" & vbCrLf
                greska = True
            End If
        End If
        If bet > 0 And h > 0 And b > 0 Then
            If h / b < Sin(bet) - 0.00001 Or h / b > Sin(bet) + 0.00001 Then
                gi = gi & "h / b != sin(β)" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And h > 0 And h1 > 0 Then
            If h / h1 < Sin(alp) - 0.00001 Or h / h1 > Sin(alp) + 0.00001 Then
                gi = gi & "h / h1 != sin(α)" & vbCrLf
                greska = True
            End If
        End If
        If h > 0 And h1 > 0 Then
            If h1 <= h Then
                gi = gi & "h1 < h" & vbCrLf
                greska = True
            End If
        End If
        If h > 0 And b > 0 Then
            If b <= h Then
                gi = gi & "b < h" & vbCrLf
                greska = True
            End If
        End If
        If b > 0 And h1 > 0 Then
            If b <= h1 Then
                gi = gi & "b < h1" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And bet > 0 Then
            If bet >= alp Then
                gi = gi & "β >= α" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And D > 0 Then
            If 2 * a < D - 0.00001 Or 2 * a > D + 0.00001 Then
                gi = gi & "2*a != D" & vbCrLf
                greska = True
            End If
        End If
        If a > 0 And d1 > 0 Then
            If Sqrt(3) * a < d1 - 0.00001 Or Sqrt(3) * a > d1 + 0.00001 Then
                gi = gi & "sqrt(3)*a != d" & vbCrLf
                greska = True
            End If
        End If
        If D > 0 And d1 > 0 Then
            If D / d1 < 2 / Sqrt(3) - 0.00001 Or D / d1 > 2 / Sqrt(3) + 0.00001 Then
                gi = gi & "D / d != 2 / sqrt(3)" & vbCrLf
                greska = True
            End If
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takva sesterostrana piramida ne postoji" & vbCrLf
            If gi.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gi
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Sesterostrana piramida")
        Else
            MsgBox(ispis, MsgBoxStyle.Information, "Sesterostrana piramida")
        End If
prekid:
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        Dim alp, r, h, d, V, O, pl, Baz As Double
        Dim greska As Boolean
        Dim unos, ispis As String
        r = 0
        V = 0
        d = 0
        O = 0
        pl = 0
        Baz = 0
        h = 0
        alp = -1
        greska = False
        ispis = ""
        rispis = ""
        gispis = ""
        pispis = ""

        ' Unos

        If CheckBox83.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu polumjera (radijusa) r u valjku:", "Unos podataka")
                If IsNumeric(unos) = True Then r = unos
                If unos = "" Then GoTo prekid
            Loop While r <= 0
            ispis = ispis & "Polumjer r = " & FormatNumber(r, 4) & vbCrLf
        End If
        If CheckBox84.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h u valjku:", "Unos podataka")
                If IsNumeric(unos) = True Then h = unos
                If unos = "" Then GoTo prekid
            Loop While h <= 0
            ispis = ispis & "Visina h = " & FormatNumber(h, 4) & vbCrLf
        End If
        If CheckBox85.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu dijagonale d u valjku:", "Unos podataka")
                If IsNumeric(unos) = True Then d = unos
                If unos = "" Then GoTo prekid
            Loop While d <= 0
            ispis = ispis & "Dijagonala d = " & FormatNumber(d, 4) & vbCrLf
        End If
        If CheckBox86.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            If alp >= System.Math.PI / 2 Then
                greska = True
                gispis = gispis & "α >= 90°" & vbCrLf
            End If
            ispis = ispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
        End If
        ispis = ispis & vbCrLf
        pispis = ispis

        ' Racun

        For i = 0 To 2
            If r = 0 Then
                If d > 0 And alp > 0 Then
                    r = (d * Cos(alp)) / 2
                    ispis = ispis & "r = " & FormatNumber(r, 4) & ", r = (d*cos(α)) / 2" & vbCrLf
                ElseIf h > 0 And d > 0 Then
                    r = Sqrt(d * d - h * h) / 2
                    ispis = ispis & "r = " & FormatNumber(r, 4) & ", r = sqrt(d*d - h*h) / 2" & vbCrLf
                End If
            End If
            If h = 0 Then
                If d > 0 And alp > 0 Then
                    h = d * Sin(alp)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = (d*sin(α))" & vbCrLf
                ElseIf d > 0 And r > 0 Then
                    h = Sqrt(d * d - 4 * r * r)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = sqrt(d*d - 4*r*r)" & vbCrLf
                End If
            End If
            If d = 0 Then
                If r > 0 And h > 0 Then
                    d = Sqrt(4 * r * r + h * h)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = sqrt(4*r*r + h*h)" & vbCrLf
                ElseIf r > 0 And alp > 0 Then
                    d = (2 * r) / Cos(alp)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = (2*r) / cos(α)" & vbCrLf
                ElseIf h > 0 And alp > 0 Then
                    d = h / Sin(alp)
                    ispis = ispis & "d = " & FormatNumber(d, 4) & ", d = h / sin(α)" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If r > 0 And d > 0 Then
                    If (2 * r) / d >= 1 Then
                        gispis = gispis & "(2*r) / d >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Acos((2 * r) / d)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos((2*r) / d)" & vbCrLf
                ElseIf h > 0 And d > 0 Then
                    If h / d >= 1 Then
                        gispis = gispis & "h / d >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Asin(h / d)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(h / d)" & vbCrLf
                End If
            End If
        Next
        ispis = ispis & vbCrLf

        If r > 0 And h > 0 Then
            V = r * r * h * System.Math.PI
            ispis = ispis & "V = " & FormatNumber(V, 4) & ", V = r*r*h*π" & vbCrLf
            O = 2 * r * System.Math.PI * (r + h)
            ispis = ispis & "O = " & FormatNumber(O, 4) & ", O = 2*r*π*(r + h)" & vbCrLf
            pl = 2 * r * h * System.Math.PI
            ispis = ispis & "Plast = " & FormatNumber(pl, 4) & ", P = 2*r*h*π" & vbCrLf
        End If
        If r > 0 Then
            Baz = r * r * System.Math.PI
            ispis = ispis & "Baza = " & FormatNumber(Baz, 4) & ", B = r*r*π" & vbCrLf
        End If

        If V = 0 Or O = 0 Then
            ispis = ispis & "Niste unijeli dovoljno podataka za izracunavanje volumena ili oplosja!" & vbCrLf
        End If

        ' Provjera

        If d > 0 And r > 0 And h > 0 Then
            If 4 * r * r + h * h < d * d - 0.00001 Or 4 * r * r + h * h > d * d + 0.00001 Then
                gispis = gispis & "4*r*r + h*h != d*d" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And r > 0 And d > 0 Then
            If (2 * r) / d < Cos(alp) - 0.00001 Or (2 * r) / d > Cos(alp) + 0.00001 Then
                gispis = gispis & "(2*r) / d != cos(α)" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And r > 0 And h > 0 Then
            If h / (2 * r) < Tan(alp) - 0.00001 Or h / (2 * r) > Tan(alp) + 0.00001 Then
                gispis = gispis & "h / (2*r) != tg(α)" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And d > 0 And h > 0 Then
            If h / d < Sin(alp) - 0.00001 Or h / d > Sin(alp) + 0.00001 Then
                gispis = gispis & "h / d != sin(α)" & vbCrLf
                greska = True
            End If
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav valjak ne postoji" & vbCrLf
            If gispis.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gispis
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Valjak")
        Else
            MsgBox(ispis, MsgBoxStyle.Information, "Valjak")
        End If
prekid:
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Dim alp, r, h, d, V, O, pl, Baz As Double
        Dim greska As Boolean
        Dim unos, ispis As String
        r = 0
        V = 0
        d = 0
        O = 0
        pl = 0
        Baz = 0
        h = 0
        alp = -1
        greska = False
        ispis = ""
        rispis = ""
        gispis = ""
        pispis = ""

        ' Unos

        If CheckBox90.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu polumjera (radijusa) r u stoscu:", "Unos podataka")
                If IsNumeric(unos) = True Then r = unos
                If unos = "" Then GoTo prekid
            Loop While r <= 0
            ispis = ispis & "Polumjer r = " & FormatNumber(r, 4) & vbCrLf
        End If
        If CheckBox89.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu visine h u stoscu:", "Unos podataka")
                If IsNumeric(unos) = True Then h = unos
                If unos = "" Then GoTo prekid
            Loop While h <= 0
            ispis = ispis & "Visina h = " & FormatNumber(h, 4) & vbCrLf
        End If
        If CheckBox88.Checked = True Then
            Do
                unos = InputBox("Unesi duljinu izvodnice s u stoscu:", "Unos podataka")
                If IsNumeric(unos) = True Then d = unos
                If unos = "" Then GoTo prekid
            Loop While d <= 0
            ispis = ispis & "Izvodnica s = " & FormatNumber(d, 4) & vbCrLf
        End If
        If CheckBox87.Checked = True Then
            Do
                Form2.Label1.Text = "Upisi velicinu kuta α:"
                Kut.Reset()
                Form2.ShowDialog()
                alp = Kut.Izracunaj()
                alp = (alp * System.Math.PI) / 180.0
                If alp = 0 Then GoTo prekid
            Loop While alp <= 0
            If alp >= System.Math.PI / 2 Then
                greska = True
                gispis = gispis & "α >= 90°" & vbCrLf
            End If
            ispis = ispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
        End If
        ispis = ispis & vbCrLf
        pispis = ispis

        ' Racun

        For i = 0 To 2
            If r = 0 Then
                If d > 0 And alp > 0 Then
                    r = d * Cos(alp)
                    ispis = ispis & "r = " & FormatNumber(r, 4) & ", r = d*cos(α)" & vbCrLf
                ElseIf h > 0 And d > 0 Then
                    r = Sqrt(d * d - h * h)
                    ispis = ispis & "r = " & FormatNumber(r, 4) & ", r = sqrt(d*d - h*h)" & vbCrLf
                End If
            End If
            If h = 0 Then
                If d > 0 And alp > 0 Then
                    h = d * Sin(alp)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = d*sin(α)" & vbCrLf
                ElseIf d > 0 And r > 0 Then
                    h = Sqrt(d * d - r * r)
                    ispis = ispis & "h = " & FormatNumber(h, 4) & ", h = sqrt(d*d - r*r)" & vbCrLf
                End If
            End If
            If d = 0 Then
                If r > 0 And h > 0 Then
                    d = Sqrt(r * r + h * h)
                    ispis = ispis & "s = " & FormatNumber(d, 4) & ", s = sqrt(r*r + h*h)" & vbCrLf
                ElseIf r > 0 And alp > 0 Then
                    d = r / Cos(alp)
                    ispis = ispis & "s = " & FormatNumber(d, 4) & ", s = r / cos(α)" & vbCrLf
                ElseIf h > 0 And alp > 0 Then
                    d = h / Sin(alp)
                    ispis = ispis & "s = " & FormatNumber(d, 4) & ", s = h / sin(α)" & vbCrLf
                End If
            End If
            If alp = -1 Then
                If r > 0 And d > 0 Then
                    If r / d >= 1 Then
                        gispis = gispis & "r / d >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Acos(r / d)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos(r / d)" & vbCrLf
                ElseIf h > 0 And d > 0 Then
                    If h / d >= 1 Then
                        gispis = gispis & "h / d >= 1" & vbCrLf
                        greska = True
                    End If
                    alp = Asin(h / d)
                    If greska = False Then ispis = ispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(h / d)" & vbCrLf
                End If
            End If
        Next
        ispis = ispis & vbCrLf

        If r > 0 And h > 0 Then
            V = (r * r * h * System.Math.PI) / 3
            ispis = ispis & "V = " & FormatNumber(V, 4) & ", V = (r*r*h*π) / 3" & vbCrLf
        End If
        If r > 0 And d > 0 Then
            O = r * System.Math.PI * (r + d)
            ispis = ispis & "O = " & FormatNumber(O, 4) & ", O = r*π*(r + d)" & vbCrLf
            pl = r * d * System.Math.PI
            ispis = ispis & "Plast = " & FormatNumber(pl, 4) & ", P = r*d*π" & vbCrLf
        End If
        If r > 0 Then
            Baz = r * r * System.Math.PI
            ispis = ispis & "Baza = " & FormatNumber(Baz, 4) & ", B = r*r*π" & vbCrLf
        End If

        If V = 0 Or O = 0 Then
            ispis = ispis & "Niste unijeli dovoljno podataka za izracunavanje volumena ili oplosja!" & vbCrLf
        End If

        ' Provjera

        If d > 0 And r > 0 And h > 0 Then
            If r * r + h * h < d * d - 0.00001 Or r * r + h * h > d * d + 0.00001 Then
                gispis = gispis & "r*r + h*h != d*d" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And r > 0 And d > 0 Then
            If r / d < Cos(alp) - 0.00001 Or r / d > Cos(alp) + 0.00001 Then
                gispis = gispis & "r / d != cos(α)" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And r > 0 And h > 0 Then
            If h / r < Tan(alp) - 0.00001 Or h / r > Tan(alp) + 0.00001 Then
                gispis = gispis & "h / r != tg(α)" & vbCrLf
                greska = True
            End If
        End If
        If alp > 0 And d > 0 And h > 0 Then
            If h / d < Sin(alp) - 0.00001 Or h / d > Sin(alp) + 0.00001 Then
                gispis = gispis & "h / d != sin(α)" & vbCrLf
                greska = True
            End If
        End If

        ' Posljednja provjera i ispis

        If greska = True Then
            ispis = pispis & "Unijeli ste neispravne podatke!" & vbCrLf & "Takav stozac ne postoji" & vbCrLf
            If gispis.Length > 0 Then
                ispis = ispis & "Razlog:" & vbCrLf & gispis
            End If
            MsgBox(ispis, MsgBoxStyle.Exclamation, "Stozac")
        Else
            MsgBox(ispis, MsgBoxStyle.Information, "Stozac")
        End If
prekid:
    End Sub

    Private Sub Button44_Click(sender As Object, e As EventArgs) Handles Button44.Click
        Dim a, b, c, d As Double
        a = 0
        b = 0
        c = 0
        d = 0

        For i = 0 To TextBox7.Text.Length - 1
            If TextBox7.Text(i) = "." Then
                Clear()
                GoTo kraj
            End If
        Next
        For i = 0 To TextBox8.Text.Length - 1
            If TextBox8.Text(i) = "." Then
                Clear()
                GoTo kraj
            End If
        Next
        For i = 0 To TextBox9.Text.Length - 1
            If TextBox9.Text(i) = "." Then
                Clear()
                GoTo kraj
            End If
        Next
        For i = 0 To TextBox10.Text.Length - 1
            If TextBox10.Text(i) = "." Then
                Clear()
                GoTo kraj
            End If
        Next

        If IsNumeric(TextBox7.Text) = False Or IsNumeric(TextBox8.Text) = False Or IsNumeric(TextBox9.Text) = False Or IsNumeric(TextBox10.Text) = False Then
            Clear()
        Else
            a = TextBox7.Text
            b = TextBox8.Text
            c = TextBox10.Text
            d = TextBox9.Text
            Label45.Visible = False

            TextBox11.Text = a + c & "i " & pz(b + d) & " " & Abs(b + d) & "j"
            TextBox12.Text = a - c & "i " & pz(b - d) & " " & Abs(b - d) & "j"
            TextBox13.Text = a * c + b * d
            TextBox14.Text = Kut_Ispis(Abs(Acos((a * c + b * d) / (Sqrt(a * a + b * b) * Sqrt(c * c + d * d)))), False)
            TextBox7.Focus()
        End If

kraj:
    End Sub

    Public Sub Clear()
        Label45.Visible = True
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox7.Focus()
        TextBox7.SelectionStart = TextBox7.TextLength
        TextBox7.ScrollToCaret()
    End Sub

    Public Function Trokut_Unos(ByRef a As Double, ByRef b As Double, ByRef c As Double, ByRef Va As Double, ByRef Vb As Double, ByRef Vc As Double, ByRef p As Double, ByRef q As Double, ByRef Ro As Double, ByRef ru As Double, ByRef ta As Double, ByRef tb As Double, ByRef tc As Double, ByRef alp As Double, ByRef bet As Double, ByRef gam As Double, ByRef greska As Boolean) As Integer
        Dim unos As String

        ' Unos
        If Pt = False Then

            If CheckBox1.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu stranice a:", "Unos podataka")
                    If IsNumeric(unos) = True Then a = unos
                    If unos = "" Then Return 0
                Loop While a <= 0
                rispis = rispis & "Stranica a = " & a & vbCrLf
            End If
            If CheckBox2.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu stranice b:", "Unos podataka")
                    If IsNumeric(unos) = True Then b = unos
                    If unos = "" Then Return 0
                Loop While b <= 0
                rispis = rispis & "Stranica b = " & b & vbCrLf
            End If
            If CheckBox3.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu stranice c:", "Unos podataka")
                    If IsNumeric(unos) = True Then c = unos
                    If unos = "" Then Return 0
                Loop While c <= 0
                rispis = rispis & "Stranica c = " & c & vbCrLf
            End If
            If CheckBox4.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu visine na stranicu a:", "Unos podataka")
                    If IsNumeric(unos) = True Then Va = unos
                    If unos = "" Then Return 0
                Loop While Va <= 0
                rispis = rispis & "Visina na stranicu a = " & Va & vbCrLf
            End If
            If CheckBox5.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu visine na stranicu b:", "Unos podataka")
                    If IsNumeric(unos) = True Then Vb = unos
                    If unos = "" Then Return 0
                Loop While Vb <= 0
                rispis = rispis & "Visina na stranicu b = " & Vb & vbCrLf
            End If
            If CheckBox6.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu visine na stranicu c:", "Unos podataka")
                    If IsNumeric(unos) = True Then Vc = unos
                    If unos = "" Then Return 0
                Loop While Vc <= 0
                rispis = rispis & "Visina na stranicu c = " & Vc & vbCrLf
            End If
            If CheckBox20.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu tezisnice ta u trokutu:", "Unos podataka")
                    If IsNumeric(unos) = True Then ta = unos
                    If unos = "" Then Return 0
                Loop While ta <= 0
                rispis = rispis & "Tezisnica ta = " & ta & vbCrLf
            End If
            If CheckBox21.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu tezisnice tb u trokutu:", "Unos podataka")
                    If IsNumeric(unos) = True Then tb = unos
                    If unos = "" Then Return 0
                Loop While tb <= 0
                rispis = rispis & "Tezisnica tb = " & tb & vbCrLf
            End If
            If CheckBox22.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu tezisnice tc u trokutu:", "Unos podataka")
                    If IsNumeric(unos) = True Then tc = unos
                    If unos = "" Then Return 0
                Loop While tc <= 0
                rispis = rispis & "Tezisnica tc = " & tc & vbCrLf
            End If
            If CheckBox7.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu polumjera opisane kruznice:", "Unos podataka")
                    If IsNumeric(unos) = True Then Ro = unos
                    If unos = "" Then Return 0
                Loop While Ro <= 0
                rispis = rispis & "Polumjer opisane kruznice = " & Ro & vbCrLf
            End If
            If CheckBox8.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu polumjera upisane kruznice:", "Unos podataka")
                    If IsNumeric(unos) = True Then ru = unos
                    If unos = "" Then Return 0
                Loop While ru <= 0
                rispis = rispis & "Polumjer upisane kruznice = " & ru & vbCrLf
            End If
            If CheckBox9.Checked = True Then
                Do
                    Form2.Label1.Text = "Unesi velicinu kuta α u stupnjevima:"
                    Kut.Reset()
                    Form2.ShowDialog()
                    alp = Kut.Izracunaj()
                    alp = (alp * System.Math.PI) / 180.0
                    If alp = 0 Then Return 0
                Loop While alp <= 0
                rispis = rispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
            End If
            If CheckBox10.Checked = True Then
                Do
                    Form2.Label1.Text = "Unesi velicinu kuta β u stupnjevima:"
                    Kut.Reset()
                    Form2.ShowDialog()
                    bet = Kut.Izracunaj()
                    bet = (bet * System.Math.PI) / 180.0
                    If bet = 0 Then Return 0
                Loop While bet <= 0
                rispis = rispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
            End If
            If CheckBox11.Checked = True Then
                Do
                    Form2.Label1.Text = "Unesi velicinu kuta ɣ u stupnjevima:"
                    Kut.Reset()
                    Form2.ShowDialog()
                    gam = Kut.Izracunaj()
                    gam = (gam * System.Math.PI) / 180.0
                    If gam = 0 Then Return 0
                Loop While gam <= 0
                rispis = rispis & "Kut ɣ = " & Kut_Ispis(gam, False) & vbCrLf
            End If
        End If

        ' Unos za pravokutan trokut
        If Pt = True Then

            If CheckBox12.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu stranice a:", "Unos podataka")
                    If IsNumeric(unos) = True Then a = unos
                    If unos = "" Then Return 0
                Loop While a <= 0
                rispis = rispis & "Stranica a = " & a & vbCrLf
            End If
            If CheckBox15.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu stranice b:", "Unos podataka")
                    If IsNumeric(unos) = True Then b = unos
                    If unos = "" Then Return 0
                Loop While b <= 0
                rispis = rispis & "Stranica b = " & b & vbCrLf
            End If
            If CheckBox13.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu stranice c:", "Unos podataka")
                    If IsNumeric(unos) = True Then c = unos
                    If unos = "" Then Return 0
                Loop While c <= 0
                rispis = rispis & "Stranica c = " & c & vbCrLf
            End If
            If CheckBox16.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu visine na stranicu c:", "Unos podataka")
                    If IsNumeric(unos) = True Then Vc = unos
                    If unos = "" Then Return 0
                Loop While Vc <= 0
                rispis = rispis & "Visin na stranicu c = " & Vc & vbCrLf
            End If
            If CheckBox23.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu tezisnice ta u trokutu:", "Unos podataka")
                    If IsNumeric(unos) = True Then ta = unos
                    If unos = "" Then Return 0
                Loop While ta <= 0
                rispis = rispis & "Tezisnice ta = " & ta & vbCrLf
            End If
            If CheckBox24.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu tezisnice tb u trokutu:", "Unos podataka")
                    If IsNumeric(unos) = True Then tb = unos
                    If unos = "" Then Return 0
                Loop While tb <= 0
                rispis = rispis & "Tezisnice tb = " & tb & vbCrLf
            End If
            If CheckBox25.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu tezisnice tc u trokutu:", "Unos podataka")
                    If IsNumeric(unos) = True Then tc = unos
                    If unos = "" Then Return 0
                Loop While tc <= 0
                rispis = rispis & "Tezisnice tc = " & tc & vbCrLf
            End If
            If CheckBox14.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu dijela hipotenuze p:", "Unos podataka")
                    If IsNumeric(unos) = True Then p = unos
                    If unos = "" Then Return 0
                Loop While p <= 0
                rispis = rispis & "Odjeljak p = " & p & vbCrLf
            End If
            If CheckBox17.Checked = True Then
                Do
                    unos = InputBox("Upisi duljinu dijela hipotenuze q:", "Unos podataka")
                    If IsNumeric(unos) = True Then q = unos
                    If unos = "" Then Return 0
                Loop While q <= 0
                rispis = rispis & "Odjeljak q = " & q & vbCrLf
            End If
            If CheckBox18.Checked = True Then
                Do
                    Form2.Label1.Text = "Unesi velicinu kuta α u stupnjevima:"
                    Kut.Reset()
                    Form2.ShowDialog()
                    alp = Kut.Izracunaj()
                    alp = (alp * System.Math.PI) / 180.0
                    If alp = 0 Then Return 0
                Loop While alp <= 0
                rispis = rispis & "Kut α = " & Kut_Ispis(alp, False) & vbCrLf
            End If
            If CheckBox19.Checked = True Then
                Do
                    Form2.Label1.Text = "Unesi velicinu kuta β u stupnjevima:"
                    Kut.Reset()
                    Form2.ShowDialog()
                    bet = Kut.Izracunaj()
                    bet = (bet * System.Math.PI) / 180.0
                    If bet = 0 Then Return 0
                Loop While bet <= 0
                rispis = rispis & "Kut β = " & Kut_Ispis(bet, False) & vbCrLf
            End If

        End If
        rispis = rispis & vbCrLf
        pispis = rispis
        Return 1

    End Function

    Public Function Trokut(ByRef a As Double, ByRef b As Double, ByRef c As Double, ByRef Va As Double, ByRef Vb As Double, ByRef Vc As Double, ByRef p As Double, ByRef q As Double, ByRef Ro As Double, ByRef ru As Double, ByRef ta As Double, ByRef tb As Double, ByRef tc As Double, ByRef alp As Double, ByRef bet As Double, ByRef gam As Double, ByRef Pov As Double, ByRef o As Double, ByRef s As Double, ByRef greska As Boolean)

        ' Racun

        For i = 1 To 3
            If a = 0 Then
                If c > 0 And b > 0 And alp > 0 Then
                    a = Sqrt((b * b) + (c * c) - 2 * b * c * Cos(alp))
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(b*b + c*c - 2*b*c*cos(α))" & vbCrLf
                ElseIf b > 0 And alp > 0 And bet > 0 Then
                    a = (b * Sin(alp)) / Sin(bet)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = (b*sin(α)) / sin(β)" & vbCrLf
                ElseIf c > 0 And alp > 0 And gam > 0 Then
                    a = (c * Sin(alp)) / Sin(gam)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = (c*sin(α)) / sin(γ)" & vbCrLf
                ElseIf Vb > 0 And gam > 0 Then
                    a = Vb / Sin(gam)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = Vb / sin(γ)" & vbCrLf
                ElseIf Vc > 0 And bet > 0 Then
                    a = Vc / Sin(bet)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = Vc / sin(β)" & vbCrLf
                ElseIf p > 0 And c > 0 Then
                    a = Sqrt(p * c)
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(p*c)" & vbCrLf
                ElseIf Ro > 0 And alp > 0 Then
                    a = 2 * Sin(alp) * Ro
                    rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = 2*sin(α)*Ro" & vbCrLf
                ElseIf c > 0 And ta > 0 And bet > 0 Then
                    Dim pom As Double = 4 * c * Cos(bet)
                    If pom * pom >= 16 * (c * c - ta * ta) Then
                        a = (pom - Sqrt(pom * pom - 16 * (c * c - ta * ta))) / 2
                        rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = 2*c*Cos(β) - 2*Sqrt(ta*ta - c*c*sin(β)*sin(β))" & vbCrLf
                    End If
                ElseIf b > 0 And ta > 0 And gam > 0 Then
                    Dim pom As Double = 4 * b * Cos(gam)
                    If pom * pom >= 16 * (b * b - ta * ta) Then
                        a = (pom - Sqrt(pom * pom - 16 * (b * b - ta * ta))) / 2
                        rispis = rispis & "a = " & FormatNumber(a, 4) & ", a =  2*b*Cos(γ) - 2*Sqrt(ta*ta - b*b*sin(γ)*sin(γ))" & vbCrLf
                    End If
                ElseIf Va > 0 And Vb > 0 And Vc > 0 Then
                    Dim k As Double = (1 + Va / Vb + Va / Vc) / 2
                    Dim raz As Double = (Va * Va) / (4 * (k * k * k * k - k * k * k * Va / Vc - k * k * k * Va / Vb + k * k * Va * Va / (Vb * Vc) - k * k * k + k * k * Va / Vc + k * k * Va / Vb - k * Va * Va / (Vb * Vc)))
                    If raz <= 0 Then
                        greska = True
                    Else
                        a = Sqrt(raz)
                        rispis = rispis & "a = " & FormatNumber(a, 4) & vbCrLf
                    End If
                ElseIf ta > 0 And tb > 0 And tc > 0 Then
                    Dim kor As Double = 2 * tb * tb + 2 * tc * tc - ta * ta
                    If kor <= 0 Then
                        greska = True
                        gispis = gispis & "2*tb*tb + 2*tc*tc - ta*ta <= 0" & vbCrLf
                    Else
                        a = (2 / 3) * Sqrt(kor)
                        rispis = rispis & "a = " & FormatNumber(a, 4) & ", a =  2/3 * sqrt(2*tb*tb + 2*tc*tc - ta*ta)" & vbCrLf
                    End If
                ElseIf b > 0 And c > 0 And ta > 0 Then
                    If 2 * (b * b + c * c) - 2 * ta * ta <= 0 Then
                        greska = True
                        gispis = gispis & "2*(b*b + c*c) - 2*ta*ta <= 0" & vbCrLf
                        a = -1
                    Else
                        a = Sqrt(2 * (b * b + c * c) - 2 * ta * ta)
                        rispis = rispis & "a = " & FormatNumber(a, 4) & ", a = sqrt(2*(b*b + c*c) - 2*ta*ta)" & vbCrLf
                    End If
                End If
            End If
            If b = 0 Then
                If c > 0 And a > 0 And bet > 0 Then
                    b = Sqrt((a * a) + (c * c) - 2 * a * c * Cos(bet))
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(a*a + c*c - 2*a*c*cos(β))" & vbCrLf
                ElseIf a > 0 And alp > 0 And bet > 0 Then
                    b = (a * Sin(bet)) / Sin(alp)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = (a*sin(β)) / sin(α)" & vbCrLf
                ElseIf c > 0 And bet > 0 And gam > 0 Then
                    b = (c * Sin(bet)) / Sin(gam)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = (c*sin(β)) / sin(γ)" & vbCrLf
                ElseIf Va > 0 And gam > 0 Then
                    b = Va / Sin(gam)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = Va / sin(γ)" & vbCrLf
                ElseIf Vc > 0 And alp > 0 Then
                    b = Vc / Sin(alp)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = Vc / sin(α)" & vbCrLf
                ElseIf q > 0 And c > 0 Then
                    b = Sqrt(q * c)
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(q*c)" & vbCrLf
                ElseIf Ro > 0 And bet > 0 Then
                    b = 2 * Sin(bet) * Ro
                    rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = 2*sin(β)*Ro" & vbCrLf
                ElseIf c > 0 And tb > 0 And alp > 0 Then
                    Dim pom As Double = 4 * c * Cos(alp)
                    If pom * pom >= 16 * (c * c - tb * tb) Then
                        b = (pom - Sqrt(pom * pom - 16 * (c * c - tb * tb))) / 2
                        rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = 2*c*Cos(α) - 2*Sqrt(tb*tb - c*c*sin(α)*sin(α))" & vbCrLf
                    End If
                ElseIf a > 0 And tb > 0 And gam > 0 Then
                    Dim pom As Double = 4 * a * Cos(gam)
                    If pom * pom >= 16 * (a * a - tb * tb) Then
                        b = (pom - Sqrt(pom * pom - 16 * (a * a - tb * tb))) / 2
                        rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = 2*a*Cos(γ) - 2*Sqrt(tb*tb - a*a*sin(γ)*sin(γ))" & vbCrLf
                    End If
                ElseIf Va > 0 And Vb > 0 And Vc > 0 Then
                    Dim k As Double = (1 + Vb / Va + Vb / Vc) / 2
                    Dim raz As Double = (Vb * Vb) / (4 * (k * k * k * k - k * k * k * Vb / Vc - k * k * k * Vb / Va + k * k * Vb * Vb / (Va * Vc) - k * k * k + k * k * Vb / Vc + k * k * Vb / Va - k * Vb * Vb / (Va * Vc)))
                    If raz <= 0 Then
                        greska = True
                    Else
                        b = Sqrt(raz)
                        rispis = rispis & "b = " & FormatNumber(b, 4) & vbCrLf
                    End If
                ElseIf ta > 0 And tb > 0 And tc > 0 Then
                    Dim kor As Double = 2 * ta * ta + 2 * tc * tc - tb * tb
                    If kor <= 0 Then
                        greska = True
                        gispis = gispis & "2*ta*ta + 2*tc*tc - tb*tb <= 0" & vbCrLf
                    Else
                        b = (2 / 3) * Sqrt(kor)
                        rispis = rispis & "b = " & FormatNumber(b, 4) & ", b =  2/3 * sqrt(2*ta*ta + 2*tc*tc - tb*tb)" & vbCrLf
                    End If
                ElseIf a > 0 And c > 0 And tb > 0 Then
                    If 2 * (a * a + c * c) - 2 * tb * tb <= 0 Then
                        greska = True
                        gispis = gispis & "2*(a*a + c*c) - 2*tb*tb <= 0" & vbCrLf
                        b = -1
                    Else
                        b = Sqrt(2 * (a * a + c * c) - 2 * tb * tb)
                        rispis = rispis & "b = " & FormatNumber(b, 4) & ", b = sqrt(2*(a*a + c*c) - 2*tb*tb)" & vbCrLf
                    End If
                End If
            End If
            If c = 0 Then
                If b > 0 And a > 0 And gam > 0 Then
                    c = Sqrt((a * a) + (b * b) - 2 * a * b * Cos(gam))
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = sqrt(a*a + b*b - 2*a*b*cos(γ))" & vbCrLf
                ElseIf a > 0 And alp > 0 And gam > 0 Then
                    c = (a * Sin(gam)) / Sin(alp)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = (a*sin(γ)) / sin(α)" & vbCrLf
                ElseIf b > 0 And bet > 0 And gam > 0 Then
                    c = (b * Sin(gam)) / Sin(bet)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = (b*sin(γ)) / sin(β)" & vbCrLf
                ElseIf Va > 0 And bet > 0 Then
                    c = Va / Sin(bet)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = Va / sin(β)" & vbCrLf
                ElseIf p > 0 And q > 0 Then
                    c = p + q
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = p + q" & vbCrLf
                ElseIf a > 0 And p > 0 Then
                    c = a * a / p
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = a*a / p" & vbCrLf
                ElseIf b > 0 And q > 0 Then
                    c = b * b / q
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = b*b / q" & vbCrLf
                ElseIf Vb > 0 And alp > 0 Then
                    c = Vb / Sin(alp)
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = Vb / sin(α)" & vbCrLf
                ElseIf Ro > 0 And gam > 0 Then
                    c = 2 * Sin(gam) * Ro
                    rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = 2*sin(γ)*Ro" & vbCrLf
                ElseIf a > 0 And tc > 0 And bet > 0 Then
                    Dim pom As Double = 4 * a * Cos(bet)
                    If pom * pom >= 16 * (a * a - tc * tc) Then
                        c = (pom - Sqrt(pom * pom - 16 * (a * a - tc * tc))) / 2
                        rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = 2*a*Cos(β) - 2*Sqrt(tc*tc - a*a*sin(β)*sin(β))" & vbCrLf
                    End If
                ElseIf b > 0 And tc > 0 And alp > 0 Then
                    Dim pom As Double = 4 * b * Cos(alp)
                    If pom * pom >= 16 * (b * b - tc * tc) Then
                        c = (pom - Sqrt(pom * pom - 16 * (b * b - tc * tc))) / 2
                        rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = 2*b*Cos(α) - 2*Sqrt(tc*tc - b*b*sin(α)*sin(α))" & vbCrLf
                    End If
                ElseIf Va > 0 And Vb > 0 And Vc > 0 Then
                    Dim k As Double = (1 + Vc / Vb + Vc / Va) / 2
                    Dim raz As Double = (Vc * Vc) / (4 * (k * k * k * k - k * k * k * Vc / Va - k * k * k * Vc / Vb + k * k * Vc * Vc / (Vb * Va) - k * k * k + k * k * Vc / Va + k * k * Vc / Vb - k * Vc * Vc / (Vb * Va)))
                    If raz <= 0 Then
                        greska = True
                    Else
                        c = Sqrt(raz)
                        rispis = rispis & "c = " & FormatNumber(c, 4) & vbCrLf
                    End If
                ElseIf ta > 0 And tb > 0 And tc > 0 Then
                    Dim kor As Double = 2 * tb * tb + 2 * ta * ta - tc * tc
                    If kor <= 0 Then
                        greska = True
                        gispis = gispis & "2*tb*tb + 2*ta*ta - tc*tc <= 0" & vbCrLf
                    Else
                        c = (2 / 3) * Sqrt(kor)
                        rispis = rispis & "c = " & FormatNumber(c, 4) & ", c =  2/3 * sqrt(2*ta*ta + 2*tb*tb - tc*tc)" & vbCrLf
                    End If
                ElseIf b > 0 And a > 0 And tc > 0 Then
                    If 2 * (b * b + a * a) - 2 * tc * tc <= 0 Then
                        greska = True
                        gispis = gispis & "2*(b*b + a*a) - 2*tc*tc <= 0" & vbCrLf
                        c = -1
                    Else
                        c = Sqrt(2 * (b * b + a * a) - 2 * tc * tc)
                        rispis = rispis & "c = " & FormatNumber(c, 4) & ", c = sqrt(2*(b*b +a*a) - 2*tc*tc)" & vbCrLf
                    End If
                End If
            End If
            If Va = 0 Then
                If c > 0 And bet > 0 Then
                    Va = c * Sin(bet)
                    rispis = rispis & "Va = " & FormatNumber(Va, 4) & ", Va = c*sin(β)" & vbCrLf
                ElseIf b > 0 And gam > 0 Then
                    Va = b * Sin(gam)
                    rispis = rispis & "Va = " & FormatNumber(Va, 4) & ", Va = b*sin(γ)" & vbCrLf
                End If
            End If

            If Vb = 0 Then
                If a > 0 And gam > 0 Then
                    Vb = a * Sin(gam)
                    rispis = rispis & "Vb = " & FormatNumber(Vb, 4) & ", Vb = a*sin(γ)" & vbCrLf
                ElseIf c > 0 And alp > 0 Then
                    Vb = c * Sin(alp)
                    rispis = rispis & "Vb = " & FormatNumber(Vb, 4) & ", Vb = c*sin(α)" & vbCrLf
                End If
            End If

            If Vc = 0 Then
                If a > 0 And bet > 0 Then
                    Vc = a * Sin(bet)
                    rispis = rispis & "Vc = " & FormatNumber(Vc, 4) & ", Vc = a*sin(β)" & vbCrLf
                ElseIf b > 0 And alp > 0 Then
                    Vc = b * Sin(alp)
                    rispis = rispis & "Vc = " & FormatNumber(Vc, 4) & ", Vc = b*sin(α)" & vbCrLf
                ElseIf p > 0 And bet > 0 Then
                    Vc = p * Tan(bet)
                    rispis = rispis & "Vc = " & FormatNumber(Vc, 4) & ", Vc = p*tg(β)" & vbCrLf
                ElseIf a > 0 And p > 0 Then
                    Vc = Sqrt(Abs(a * a - p * p))
                    rispis = rispis & "Vc = " & FormatNumber(Vc, 4) & ", Vc = sqrt(a*a -p*p)" & vbCrLf
                ElseIf q > 0 And alp > 0 Then
                    Vc = q * Tan(alp)
                    rispis = rispis & "Vc = " & FormatNumber(Vc, 4) & ", Vc = q*tg(α)" & vbCrLf
                ElseIf q > 0 And b > 0 Then
                    Vc = Sqrt(Abs(b * b - q * q))
                    rispis = rispis & "Vc = " & FormatNumber(Vc, 4) & ", Vc = sqrt(b*b -q*q)" & vbCrLf
                ElseIf p > 0 And q > 0 Then
                    Vc = Sqrt(p * q)
                    rispis = rispis & "Vc = " & FormatNumber(Vc, 4) & ", Vc = sqrt(p*q)" & vbCrLf
                End If
            End If

            If p = 0 Then
                If a > 0 And bet > 0 Then
                    p = a * Cos(bet)
                    rispis = rispis & "p = " & FormatNumber(p, 4) & ", p = a*cos(β)" & vbCrLf
                ElseIf Vc > 0 And bet > 0 Then
                    p = Vc / Tan(bet)
                    rispis = rispis & "p = " & FormatNumber(p, 4) & ", p = Vc*tg(β)" & vbCrLf
                ElseIf a > 0 And Vc > 0 Then
                    p = Sqrt(Abs(a * a - Vc * Vc))
                    rispis = rispis & "p = " & FormatNumber(p, 4) & ", p = sqrt(a*a -Vc*Vc)" & vbCrLf
                ElseIf a > 0 And c > 0 Then
                    p = a * a / c
                    rispis = rispis & "p = " & FormatNumber(p, 4) & ", p = a*a / c" & vbCrLf
                ElseIf Vc > 0 And q > 0 Then
                    p = Vc * Vc / q
                    rispis = rispis & "p = " & FormatNumber(p, 4) & ", p = Vc*Vc / q" & vbCrLf
                ElseIf c > 0 And q > 0 Then
                    p = c - q
                    rispis = rispis & "p = " & FormatNumber(p, 4) & ", p = c - q" & vbCrLf
                End If
            End If

            If q = 0 Then
                If b > 0 And alp > 0 Then
                    q = b * Cos(alp)
                    rispis = rispis & "q = " & FormatNumber(q, 4) & ", q = b*cos(α)" & vbCrLf
                ElseIf Vc > 0 And alp > 0 Then
                    q = Vc / Tan(alp)
                    rispis = rispis & "q = " & FormatNumber(q, 4) & ", q = Vc*tg(α)" & vbCrLf
                ElseIf b > 0 And Vc > 0 Then
                    q = Sqrt(Abs(b * b - Vc * Vc))
                    rispis = rispis & "q = " & FormatNumber(q, 4) & ", q = sqrt(b*b - Vc*Vc)" & vbCrLf
                ElseIf b > 0 And c > 0 Then
                    q = b * b / c
                    rispis = rispis & "q = " & FormatNumber(q, 4) & ", q = b*b / c" & vbCrLf
                ElseIf Vc > 0 And p > 0 Then
                    q = Vc * Vc / p
                    rispis = rispis & "q = " & FormatNumber(q, 4) & ", q = Vc*Vc / p" & vbCrLf
                ElseIf c > 0 And p > 0 Then
                    q = c - p
                    rispis = rispis & "q = " & FormatNumber(q, 4) & ", q = c - p" & vbCrLf
                End If
            End If

            If alp = -1 Then
                If a > 0 And b > 0 And bet > 0 Then
                    If a * Sin(bet) / b >= 1 Then
                        greska = True
                        gispis = gispis & "a*sin(β) / b >= 1" & vbCrLf
                    End If
                    alp = Asin(a * Sin(bet) / b)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(a*sin(β) / b)" & vbCrLf
                ElseIf Vc > 0 And b > 0 Then
                    If Vc / b >= 1 Then
                        greska = True
                        gispis = gispis & "Vc / b >= 1" & vbCrLf
                    End If
                    alp = Asin(Vc / b)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(Vc / b)" & vbCrLf
                ElseIf Vb > 0 And c > 0 Then
                    If Vb / c >= 1 Then
                        greska = True
                        gispis = gispis & "Vc / c >= 1" & vbCrLf
                    End If
                    alp = Asin(Vb / c)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(Vb / c)" & vbCrLf
                ElseIf a > 0 And c > 0 And gam > 0 Then
                    If a * Sin(gam) / c >= 1 Then
                        greska = True
                        gispis = gispis & "a*sin(γ) / c >= 1" & vbCrLf
                    End If
                    alp = Asin(a * Sin(gam) / c)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arcsin(a*sin(γ) / c)" & vbCrLf
                ElseIf a > 0 And b > 0 And c > 0 Then
                    If (b * b + c * c - a * a) / (2 * b * c) >= 1 Then
                        greska = True
                        gispis = gispis & "(b*b + c*c - a*a) / (2*b*c) >= 1" & vbCrLf
                    End If
                    alp = Acos((b * b + c * c - a * a) / (2 * b * c))
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos((b*b + c*c - a*a) / (2*b*c)" & vbCrLf
                ElseIf q > 0 And b > 0 And b > q Then
                    If q / b >= 1 Then
                        greska = True
                        gispis = gispis & "q / b >= 1" & vbCrLf
                    End If
                    alp = Acos(q / b)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arccos(q / b)" & vbCrLf
                ElseIf Vc > 0 And q > 0 Then
                    alp = Atan(Vc / q)
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = arctg(Vc / q)" & vbCrLf
                End If
            End If
            If bet = -1 Then
                If c > 0 And Va > 0 Then
                    If Va / c >= 1 Then
                        greska = True
                        gispis = gispis & "Va / c >= 1" & vbCrLf
                    End If
                    bet = Asin(Va / c)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(Va / c)" & vbCrLf
                ElseIf a > 0 And Vc > 0 Then
                    If Vc / a >= 1 Then
                        greska = True
                        gispis = gispis & "Vc / a >= 1" & vbCrLf
                    End If
                    bet = Asin(Vc / a)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(Vc / a)" & vbCrLf
                ElseIf b > 0 And a > 0 And alp > 0 Then
                    If b * Sin(alp) / a >= 1 Then
                        greska = True
                        gispis = gispis & "b*Sin(α) / a >= 1" & vbCrLf
                    End If
                    bet = Asin(b * Sin(alp) / a)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(b*sin(α) / a)" & vbCrLf
                ElseIf b > 0 And c > 0 And gam > 0 Then
                    If b * Sin(gam) / c >= 1 Then
                        greska = True
                        gispis = gispis & "b*Sin(γ) / c >= 1" & vbCrLf
                    End If
                    bet = Asin(b * Sin(gam) / c)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arcsin(b*sin(γ) / c)" & vbCrLf
                ElseIf a > 0 And b > 0 And c > 0 Then
                    If (a * a + c * c - b * b) / (2 * a * c) >= 1 Then
                        greska = True
                        gispis = gispis & "(a*a + c*c - b*b) / (2*a*c) >= 1" & vbCrLf
                    End If
                    bet = Acos((a * a + c * c - b * b) / (2 * a * c))
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos((a*a + c*c - b*b) / (2*a*c))" & vbCrLf
                ElseIf p > 0 And a > 0 And a > p Then
                    If p / a >= 1 Then
                        greska = True
                        gispis = gispis & "p / a >= 1" & vbCrLf
                    End If
                    bet = Acos(p / a)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arccos(p / a)" & vbCrLf
                ElseIf Vc > 0 And p > 0 Then
                    bet = Atan(Vc / p)
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = arctg(Vc / p)" & vbCrLf
                End If
            End If

            If gam = -1 Then
                If b > 0 And Va > 0 Then
                    If Va / b >= 1 Then
                        greska = True
                        gispis = gispis & "Va / b >= 1" & vbCrLf
                    End If
                    gam = Asin(Va / b)
                    If greska = False Then rispis = rispis & "γ = " & Kut_Ispis(gam, False) & ", γ = arcsin(Va / b)" & vbCrLf
                ElseIf a > 0 And Vb > 0 Then
                    If Vb / a >= 1 Then
                        greska = True
                        gispis = gispis & "Vb / a >= 1" & vbCrLf
                    End If
                    gam = Asin(Vb / a)
                    If greska = False Then rispis = rispis & "γ = " & Kut_Ispis(gam, False) & ", γ = arcsin(Vb / a)" & vbCrLf
                ElseIf a > 0 And b > 0 And c > 0 Then
                    If (b * b + a * a - c * c) / (2 * b * a) >= 1 Then
                        greska = True
                        gispis = gispis & "(b*b + a*a - c*c) / (2*b*a) >= 1" & vbCrLf
                    End If
                    gam = Acos((b * b + a * a - c * c) / (2 * b * a))
                    If greska = False Then rispis = rispis & "γ = " & Kut_Ispis(gam, False) & ", γ = arccos((b*b + a*a - c*c) / (2*b*a))" & vbCrLf
                End If
            End If

            If alp = -1 And bet > 0 And gam > 0 Then
                alp = System.Math.PI - bet - gam
                If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & ", α = 180° - β - γ" & vbCrLf
            ElseIf bet = -1 And alp > 0 And gam > 0 Then
                bet = System.Math.PI - alp - gam
                If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & ", β = 180° - α - γ" & vbCrLf
            ElseIf gam = -1 And alp > 0 And bet > 0 Then
                gam = System.Math.PI - alp - bet
                If greska = False Then rispis = rispis & "γ = " & Kut_Ispis(gam, False) & ", γ = 180° - α - β" & vbCrLf
            End If

            If alp > 0 And bet > 0 And gam > 0 And alp + bet + gam <> System.Math.PI Then
                If CheckBox9.Checked = False And System.Math.PI - alp + bet + gam >= System.Math.PI - 0.001 And System.Math.PI - alp + bet + gam <= System.Math.PI + 0.001 Then
                    alp = System.Math.PI - alp
                    If greska = False Then rispis = rispis & "α = " & Kut_Ispis(alp, False) & vbCrLf
                ElseIf CheckBox10.Checked = False And System.Math.PI - bet + alp + gam >= System.Math.PI - 0.001 And System.Math.PI - bet + alp + gam <= System.Math.PI + 0.001 Then
                    bet = System.Math.PI - bet
                    If greska = False Then rispis = rispis & "β = " & Kut_Ispis(bet, False) & vbCrLf
                ElseIf CheckBox11.Checked = False And System.Math.PI - gam + bet + alp >= System.Math.PI - 0.001 And System.Math.PI - gam + bet + alp <= System.Math.PI + 0.001 Then
                    gam = System.Math.PI - gam
                    If greska = False Then rispis = rispis & "γ = " & Kut_Ispis(gam, False) & vbCrLf
                End If
            End If

            If ta = 0 Then
                If a > 0 And b > 0 And c > 0 Then
                    ta = Sqrt(2 * (b * b + c * c) - a * a) / 2
                    rispis = rispis & "ta = " & FormatNumber(ta, 4) & ", ta = sqrt(2*(b*b + c*c) - a*a) / 2" & vbCrLf
                ElseIf c > 0 And a > 0 And bet > 0 Then
                    ta = Sqrt(c * c + (a * a) / 4 - a * c * Cos(bet))
                    rispis = rispis & "ta = " & FormatNumber(ta, 4) & ", ta = sqrt(c*c + (a*a)/4 - a*c*cos(β))" & vbCrLf
                ElseIf b > 0 And a > 0 And gam > 0 Then
                    ta = Sqrt(b * b + (a * a) / 4 - a * b * Cos(gam))
                    rispis = rispis & "ta = " & FormatNumber(ta, 4) & ", ta = sqrt(b*b + (a*a)/4 - a*b*cos(γ))" & vbCrLf
                ElseIf tb > 0 And a > 0 And tc > 0 Then
                    Dim k As Double
                    k = 2 * tb * tb + 2 * tc * tc - (9 / 4) * a * a
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "2*tb*tb + 2*tc*tc - (9 / 4)*a*a <= 0" & vbCrLf
                    Else
                        ta = Sqrt(k)
                        rispis = rispis & "ta = " & FormatNumber(ta, 4) & ", ta = sqrt(2*tb*tb + 2*tc*tc - 9*a*a/4)" & vbCrLf
                    End If
                ElseIf tb > 0 And b > 0 And tc > 0 Then
                    Dim k As Double
                    k = 0.5 * tb * tb + (9 / 8) * b * b - tc * tc
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "(1/2)*tb*tb + (9 / 8)*b*b - tc*tc <= 0" & vbCrLf
                    Else
                        ta = Sqrt(k)
                        rispis = rispis & "ta = " & FormatNumber(ta, 4) & ", ta = sqrt((1/2)*tb*tb + 9*b*b/8 - tc*tc)" & vbCrLf
                    End If
                ElseIf tb > 0 And c > 0 And tc > 0 Then
                    Dim k As Double
                    k = 0.5 * tc * tc + (9 / 8) * c * c - tb * tb
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "(1/2)*tc*tc + (9 / 8)*c*c - tb*tb <= 0" & vbCrLf
                    Else
                        ta = Sqrt(k)
                        rispis = rispis & "ta = " & FormatNumber(ta, 4) & ", ta = sqrt((1/2)*tc*tc + 9*c*c/8 - tb*tb)" & vbCrLf
                    End If
                End If
            End If
            If tb = 0 Then
                If a > 0 And b > 0 And c > 0 Then
                    tb = Sqrt(2 * (a * a + c * c) - b * b) / 2
                    rispis = rispis & "tb = " & FormatNumber(tb, 4) & ", tb = sqrt(2*(a*a + c*c) - b*b) / 2" & vbCrLf
                ElseIf c > 0 And b > 0 And alp > 0 Then
                    tb = Sqrt(c * c + (b * b) / 4 - b * c * Cos(alp))
                    rispis = rispis & "tb = " & FormatNumber(tb, 4) & ", tb = sqrt(c*c + (b*b)/4 - b*c*cos(α))" & vbCrLf
                ElseIf b > 0 And a > 0 And gam > 0 Then
                    tb = Sqrt(a * a + (b * b) / 4 - a * b * Cos(gam))
                    rispis = rispis & "tb = " & FormatNumber(tb, 4) & ", tb = sqrt(a*a + (b*b)/4 - b*a*cos(γ))" & vbCrLf
                ElseIf ta > 0 And a > 0 And tc > 0 Then
                    Dim k As Double
                    k = 0.5 * ta * ta + (9 / 8) * a * a - tc * tc
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "(1/2)*ta*ta + (9 / 8)*a*a - tc*tc <= 0" & vbCrLf
                    Else
                        tb = Sqrt(k)
                        rispis = rispis & "tb = " & FormatNumber(tb, 4) & ", tb = sqrt((1/2)*ta*ta + 9*a*a/8 - tc*tc)" & vbCrLf
                    End If
                ElseIf ta > 0 And b > 0 And tc > 0 Then
                    Dim k As Double
                    k = 2 * ta * ta + 2 * tc * tc - (9 / 4) * b * b
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "k = 2*ta*ta + 2*tc*tc - (9 / 4)*b*b <= 0" & vbCrLf
                    Else
                        tb = Sqrt(k)
                        rispis = rispis & "tb = " & FormatNumber(tb, 4) & ", tb = sqrt(2*ta*ta + 2*tc*tc - (9 / 4)*b*b)" & vbCrLf
                    End If
                ElseIf ta > 0 And c > 0 And tc > 0 Then
                    Dim k As Double
                    k = 0.5 * tc * tc + (9 / 8) * c * c - ta * ta
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "(1/2)*tc*tc + (9 / 8)*c*c - ta*ta <= 0" & vbCrLf
                    Else
                        tb = Sqrt(k)
                        rispis = rispis & "tb = " & FormatNumber(tb, 4) & ", tb = sqrt((1/2)*tc*tc + 9*c*c/8 - ta*ta)" & vbCrLf
                    End If
                End If
            End If
            If tc = 0 Then
                If a > 0 And b > 0 And c > 0 Then
                    tc = Sqrt(2 * (b * b + a * a) - c * c) / 2
                    rispis = rispis & "tc = " & FormatNumber(tc, 4) & ", tc = sqrt(2*(b*b + a*a) - c*c) / 2" & vbCrLf
                ElseIf c > 0 And a > 0 And bet > 0 Then
                    tc = Sqrt(a * a + (c * c) / 4 - a * c * Cos(bet))
                    rispis = rispis & "tc = " & FormatNumber(tc, 4) & ", tc = sqrt(a*a + (c*c)/4 - a*c*cos(β))" & vbCrLf
                ElseIf b > 0 And c > 0 And alp > 0 Then
                    tc = Sqrt(b * b + (c * c) / 4 - c * b * Cos(alp))
                    rispis = rispis & "tc = " & FormatNumber(tc, 4) & ", tc = sqrt(b*b + (c*c)/4 - b*c*cos(α))" & vbCrLf
                ElseIf ta > 0 And a > 0 And tb > 0 Then
                    Dim k As Double
                    k = 0.5 * ta * ta + (9 / 8) * a * a - tb * tb
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "(1/2)*ta*ta + (9 / 8)*a*a - tb*tb <= 0" & vbCrLf
                    Else
                        tc = Sqrt(k)
                        rispis = rispis & "tc = " & FormatNumber(tc, 4) & ", tc = sqrt((1/2)*ta*ta + 9*a*a/8 - tb*tb)" & vbCrLf
                    End If
                ElseIf tb > 0 And b > 0 And ta > 0 Then
                    Dim k As Double
                    k = 0.5 * tb * tb + (9 / 8) * b * b - ta * ta
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "(1/2)*tb*tb + (9 / 8)*b*b - ta*ta <= 0" & vbCrLf
                    Else
                        tc = Sqrt(k)
                        rispis = rispis & "tc = " & FormatNumber(tc, 4) & ", tc = sqrt((1/2)*tb*tb + 9*b*b/8 - ta*ta)" & vbCrLf
                    End If
                ElseIf ta > 0 And c > 0 And tb > 0 Then
                    Dim k As Double
                    k = 2 * ta * ta + 2 * tb * tb - (9 / 4) * c * c
                    If k <= 0 Then
                        greska = True
                        gispis = gispis & "2*ta*ta + 2*tb*tb - (9 / 4)*c*c <= 0" & vbCrLf
                    Else
                        tc = Sqrt(k)
                        rispis = rispis & "tc = " & FormatNumber(tc, 4) & ", tc = sqrt(2*ta*ta + 2*tb*tb - (9 / 4)*c*c)" & vbCrLf
                    End If
                End If
            End If

        Next i

        If alp = -1 Then
            alp = 0
        End If
        If bet = -1 Then
            bet = 0
        End If
        If gam = -1 Then
            gam = 0
        End If

        ' Provjera

        If alp < 0 Then
            greska = True
            gispis = gispis & "α < 0" & vbCrLf
        End If
        If bet < 0 Then
            greska = True
            gispis = gispis & "β < 0" & vbCrLf
        End If
        If gam < 0 Then
            greska = True
            gispis = gispis & "γ < 0" & vbCrLf
        End If

        'If a > 0 And ta > 0 And tb > 0 And tc > 0 Then
        '    If a * a < ((4 / 9) * (2 * tb * tb + 2 * tc * tc - ta * ta)) - 0.000001 Or a * a > ((4 / 9) * (2 * tb * tb + 2 * tc * tc - ta * ta)) + 0.000001 Then
        '        greska = True
        '        gispis = gispis & "a*a != ((4/9)*(2*tb*tb + 2*tc*tc - ta*ta))" & vbCrLf
        '    End If
        'End If
        'If b > 0 And ta > 0 And tb > 0 And tc > 0 Then
        '    If b * b < ((4 / 9) * (2 * ta * ta + 2 * tc * tc - tb * tb)) - 0.000001 Or b * b > ((4 / 9) * (2 * ta * ta + 2 * tc * tc - tb * tb)) + 0.000001 Then
        '        greska = True
        '        gispis = gispis & "b*b != ((4/9)*(2*ta*ta + 2*tc*tc - tb*tb))" & vbCrLf
        '    End If
        'End If
        'If c > 0 And ta > 0 And tb > 0 And tc > 0 Then
        '    If c * c < ((4 / 9) * (2 * ta * ta + 2 * tb * tb - tc * tc)) - 0.000001 Or c * c > ((4 / 9) * (2 * ta * ta + 2 * tb * tb - tc * tc)) + 0.000001 Then
        '        greska = True
        '        gispis = gispis & "c*c != ((4/9)*(2*ta*ta + 2*tb*tb - tc*tc))" & vbCrLf
        '    End If
        'End If

        If (alp = System.Math.PI / 2 And bet = System.Math.PI / 2) Or (alp = System.Math.PI / 2 And gam = System.Math.PI / 2) Or (bet = System.Math.PI / 2 And gam = System.Math.PI / 2) Then greska = True
        If alp > 0 And bet > 0 And gam > 0 Then
            If alp + bet + gam < 3.1415 Or alp + bet + gam > 3.1416 Then
                greska = True
                gispis = gispis & " α + β + γ != 180°" & vbCrLf
            End If
        End If

        If a > 0 And b > 0 And c > 0 Then
            If a + b <= c Then
                greska = True
                gispis = gispis & "a + b <= c" & vbCrLf
            End If
            If a + c <= b Then
                greska = True
                gispis = gispis & "a + c <= b" & vbCrLf
            End If
            If b + c <= a Then
                greska = True
                gispis = gispis & "b + c <= a" & vbCrLf
            End If
        End If

        If Vc > 0 And b > 0 And a > 0 Then
            If Vc > b Then
                greska = True
                gispis = gispis & "Vc > b" & vbCrLf
            End If
            If Vc > a Then
                greska = True
                gispis = gispis & "Vc > a" & vbCrLf
            End If
        End If
        If Vb > 0 And a > 0 And c > 0 Then
            If Vb > a Then
                greska = True
                gispis = gispis & "Vb > a" & vbCrLf
            End If
            If Vb > c Then
                greska = True
                gispis = gispis & "Vb > c" & vbCrLf
            End If
        End If
        If Va > 0 And b > 0 And c > 0 Then
            If Va > b Then
                greska = True
                gispis = gispis & "Va > b" & vbCrLf
            End If
            If Va > c Then
                greska = True
                gispis = gispis & "Va > c" & vbCrLf
            End If
        End If

        If a > 0 And b > 0 And c > 0 And alp > 0 And bet > 0 And gam > 0 Then
            If (a >= b And bet > alp) Then
                greska = True
                gispis = gispis & "a >= b, β > α" & vbCrLf
            End If
            If (b >= a And alp > bet) Then
                greska = True
                gispis = gispis & "b >= a, α > β" & vbCrLf
            End If
            If (a >= c And gam > alp) Then
                greska = True
                gispis = gispis & "a >= c, γ > α" & vbCrLf
            End If
            If (c >= a And alp > gam) Then
                greska = True
                gispis = gispis & "c >= a,  α > γ" & vbCrLf
            End If
            If (b >= c And gam > bet) Then
                greska = True
                gispis = gispis & "b >= c, γ > β" & vbCrLf
            End If
            If (c >= b And bet > gam) Then
                greska = True
                gispis = gispis & "c >= b, β > γ" & vbCrLf
            End If
        End If

        If a > 0 And b > 0 And Va > 0 And Vb > 0 Then
            If a / b < Vb / Va - 0.01 Or a / b > Vb / Va + 0.01 Then
                greska = True
                gispis = gispis & "a / b != Vb / Va" & vbCrLf
            End If
        ElseIf a > 0 And c > 0 And Va > 0 And Vc > 0 Then
            If a / c < Vc / Va - 0.01 Or a / c > Vc / Va + 0.01 Then
                greska = True
                gispis = gispis & "a / c != Vc / Va" & vbCrLf
            End If
        ElseIf b > 0 And c > 0 And Vb > 0 And Vc > 0 Then
            If b / c < Vc / Vb - 0.01 Or b / c > Vc / Vb + 0.01 Then
                greska = True
                gispis = gispis & "b / c != Vc / Vb" & vbCrLf
            End If
        End If

        ' Provjera za pravokutan trokut

        If p > 0 And q > 0 And c > 0 Then
            If a * a + b * b < c * c - 0.001 Or a * a + b * b > c * c + 0.001 Then
                greska = True
                gispis = gispis & "a*a + b*b != c*c" & vbCrLf
            ElseIf p + q < c - 0.001 Or p + q > c + 0.001 Then
                greska = True
                gispis = gispis & "p + q != c" & vbCrLf
            End If
        End If

        ' Povrsina i opseg
        rispis = rispis & vbCrLf

        If a > 0 And b > 0 And c > 0 Then
            s = (a + b + c) / 2
            rispis = rispis & "s = " & FormatNumber(s, 4) & ", s = (a + b + c) / 2" & vbCrLf
            o = a + b + c
            rispis = rispis & "O = " & FormatNumber(o, 4) & ", O = a + b + c" & vbCrLf
        End If

        If a > 0 And Va > 0 Then
            Pov = a * Va / 2
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = a*Va / 2" & vbCrLf & vbCrLf
        ElseIf b > 0 And Vb > 0 Then
            Pov = b * Vb / 2
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = b*Vb / 2" & vbCrLf & vbCrLf
        ElseIf c > 0 And Vc > 0 Then
            Pov = c * Vc / 2
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = c*Vc / 2" & vbCrLf & vbCrLf
        ElseIf s > 0 Then
            Pov = Sqrt(s * (s - a) * (s - b) * (s - c))
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = sqrt(s*(s-a)*(s-b)*(s-c))" & vbCrLf & vbCrLf
        ElseIf a > 0 And b > 0 And gam > 0 Then
            Pov = 0.5 * a * b * Sin(gam)
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = 0,5*a*b*sin(γ)" & vbCrLf & vbCrLf
        ElseIf a > 0 And c > 0 And bet > 0 Then
            Pov = 0.5 * a * c * Sin(bet)
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = 0,5*a*c*sin(β)" & vbCrLf & vbCrLf
        ElseIf b > 0 And c > 0 And alp > 0 Then
            Pov = 0.5 * b * c * Sin(alp)
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = 0,5*b*c*sin(α)" & vbCrLf & vbCrLf
        ElseIf a > 0 And alp > 0 And bet > 0 And gam > 0 Then
            Pov = (a * a * Sin(bet) * Sin(gam)) / (2 * Sin(alp))
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = (a*a*sin(β)*sin(γ)) / (2*sin(α))" & vbCrLf & vbCrLf
        ElseIf b > 0 And alp > 0 And bet > 0 And gam > 0 Then
            Pov = (b * b * Sin(alp) * Sin(gam)) / (2 * Sin(bet))
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = (b*b*sin(α)*sin(γ)) / (2*sin(β))" & vbCrLf & vbCrLf
        ElseIf c > 0 And alp > 0 And bet > 0 And gam > 0 Then
            Pov = (c * c * Sin(bet) * Sin(alp)) / (2 * Sin(gam))
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = (c*c*sin(α)*sin(β)) / (2*sin(γ))" & vbCrLf & vbCrLf
        ElseIf ru > 0 And s > 0 Then
            Pov = ru * s
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = ru * s" & vbCrLf & vbCrLf
        ElseIf a > 0 And b > 0 And c > 0 And Ro > 0 Then
            Pov = (a * b * c) / (4 * Ro)
            rispis = rispis & "P = " & FormatNumber(Pov, 4) & ", P = (a*b*c) / (4*Ro)" & vbCrLf & vbCrLf
        End If

        If Pov = 0 Or o = 0 Then
            rispis = rispis & "Niste unijeli dovoljno podataka za izracunavanje opsega ili povrsine!" & vbCrLf & vbCrLf
        End If

        If Pov > 0 And s > 0 Then
            Ro = (a * b * c) / (4 * Pov)
            rispis = rispis & "R = " & FormatNumber(Ro, 4) & ", R = (a*b*c) / (4*P)" & vbCrLf
        End If
        If Pov > 0 And s > 0 Then
            ru = Pov / s
            rispis = rispis & "r = " & FormatNumber(ru, 4) & ", r = P / s" & vbCrLf
        End If

        Return 1

    End Function

    Public Function Kut_Ispis(ByVal kut As Double, ByVal s As Boolean) As String
        Dim a As String
        Dim dec, sek As Double
        Dim st, min As Integer

        If s = True Then
            dec = kut - Int(kut)
            st = Int(kut)
            min = Int(dec * 60)
            sek = ((dec * 60) - Int(dec * 60)) * 60
            If FormatNumber(sek, 2) = 60 Then
                min = min + 1
                sek = 0
            End If
            If FormatNumber(min, 2) = 60 Then
                st = st + 1
                min = 0
            End If
            a = st & "° " & min & "' " & FormatNumber(sek, 2) & "''"
        Else
            dec = (kut * 180 / System.Math.PI) - Int(kut * 180 / System.Math.PI)
            st = Int(kut * 180 / System.Math.PI)
            min = Int(dec * 60)
            sek = ((dec * 60) - Int(dec * 60)) * 60
            If FormatNumber(sek, 2) = 60 Then
                min = min + 1
                sek = 0
            End If
            If FormatNumber(min, 2) = 60 Then
                st = st + 1
                min = 0
            End If

            a = st & "° " & min & "' " & FormatNumber(sek, 2) & "''"
        End If

        Return a

    End Function

    Public Function Graf_Ispis(ByVal a As Double, ByVal b As Double, ByVal c As Double) As String
        Dim text As String = Nothing
        If a = 0 And Abs(b) > 0 Then
            If c > 0 Then
                text = "Pravac y = " & b & "x + " & c
            ElseIf c < 0 Then
                text = "Pravac y = " & b & "x - " & Abs(c)
            ElseIf c = 0 Then
                text = "Pravac y = " & b & "x"
            End If

        ElseIf a = 0 And b = 0 Then
            text = "Pravac y = " & c
        ElseIf Abs(a) > 0 Then

            If c > 0 And b > 0 Then
                text = "Parabola y = " & a & "x^2 + " & b & "x + " & c
            ElseIf c > 0 And b < 0 Then
                text = "Parabola y = " & a & "x^2 - " & Abs(b) & "x + " & c
            ElseIf c < 0 And b > 0 Then
                text = "Parabola y = " & a & "x^2 + " & b & "x - " & Abs(c)
            ElseIf c < 0 And b < 0 Then
                text = "Parabola y = " & a & "x^2 - " & Abs(b) & "x - " & Abs(c)
            ElseIf c = 0 And b > 0 Then
                text = "Parabola y = " & a & "x^2 + " & b & "x"
            ElseIf c = 0 And b < 0 Then
                text = "Parabola y = " & a & "x^2 - " & Abs(b) & "x"
            ElseIf b = 0 And c > 0 Then
                text = "Parabola y = " & a & "x^2 + " & c
            ElseIf b = 0 And c < 0 Then
                text = "Parabola y = " & a & "x^2 - " & Abs(c)
            ElseIf c = 0 And b = 0 Then
                text = "Parabola y = " & a & "x^2"
            End If
        End If
        Return text
    End Function

    Public Function izdvoji_niz(ByVal ulaz As String, ByVal p As Integer, ByVal z As Integer) As String
        Dim a(50) As Char
        Dim j As Integer = 0
        For i = p To z
            a(j) = ulaz(i)
            j = j + 1
        Next
        Return a
    End Function

    Public Function zagrada(ByVal ulaz As String, ByVal p As Integer, ByVal z As Integer) As String
        Dim a As String = Nothing
        Dim pb As Integer = -1
        Dim zb As Integer = -1
        Dim broj As Double
        Dim k As Double = 0
        Dim za As Integer = -1
        Dim pom As Integer = 0
        Dim knep As Boolean = False
        Dim nep As Boolean = False
        Dim y As Boolean = False

        pb = p
        For i = p To z
            If ulaz(i) = "(" Then
                za = i + 1
                zb = i - 1
                GoTo racun
            ElseIf ulaz(i) = "x" Then
                nep = True
                za = i + 2
                zb = i - 1
                GoTo racun
            ElseIf ulaz(i) = "y" Then
                y = True
                za = i + 2
                zb = i - 1
                GoTo racun
            ElseIf ulaz(i) = "K" Then
                knep = True
                za = i + 2
                zb = i - 1
                GoTo racun
            End If
        Next
racun:
        Double.TryParse(izdvoji_niz(ulaz, pb, zb), k)
        pb = -1
        zb = -1
        For i = za To z

            If IsNumeric(ulaz(i)) Or ulaz(i) = "," Or ((ulaz(i) = "+" Or ulaz(i) = "-") And zb = -1) Then
                If pb = -1 Then pb = i
                zb = i

            Else

                If pom = 0 Then
                    Double.TryParse(izdvoji_niz(ulaz, pb, zb), broj)
                    pom = 1
                ElseIf pb > -1 Then
                    If ulaz(pb - 1) = "x" Or ulaz(pb - 1) = "K" Or ulaz(pb - 1) = "y" Then
                        Double.TryParse(izdvoji_niz(ulaz, pb, zb), broj)
                    Else
                        Double.TryParse(izdvoji_niz(ulaz, pb - 1, zb), broj)
                    End If

                End If

                If nep = True Then
                    If ulaz(i) = "x" Then
                        a = a & pz(k * broj) & Abs(k * broj) & "K"
                    ElseIf ulaz(i) = "K" Then
                        pStupanj = True
                    ElseIf ulaz(i) = "y" Then
                        pStupanj = True
                    ElseIf ulaz(i) = ")" And (ulaz(i - 1) = "K" Or ulaz(i - 1) = "x") Then
                        GoTo oz
                    ElseIf ulaz(i) = "+" Or ulaz(i) = "-" Or ulaz(i) = ")" Then
                        a = a & pz(k * broj) & Abs(k * broj) & "x"
                    End If
                    pb = -1
                    zb = -1

                ElseIf knep = True Then
                    If ulaz(i) = "x" Then
                        pStupanj = True
                    ElseIf ulaz(i) = "K" Then
                        pStupanj = True
                    ElseIf ulaz(i) = "y" Then
                        pStupanj = True
                    ElseIf ulaz(i) = ")" And (ulaz(i - 1) = "K" Or ulaz(i - 1) = "x") Then
                        GoTo oz
                    ElseIf ulaz(i) = "+" Or ulaz(i) = "-" Or ulaz(i) = ")" Then
                        a = a & pz(k * broj) & Abs(k * broj) & "K"
                    End If
                    pb = -1
                    zb = -1
                ElseIf y = True Then
                    If ulaz(i) = "x" Then
                        pStupanj = True
                    ElseIf ulaz(i) = "K" Then
                        pStupanj = True
                    ElseIf ulaz(i) = "y" Then
                        pStupanj = True
                    ElseIf ulaz(i) = ")" And (ulaz(i - 1) = "K" Or ulaz(i - 1) = "y" Or ulaz(i - 1) = "x") Then
                        GoTo oz
                    ElseIf ulaz(i) = "+" Or ulaz(i) = "-" Or ulaz(i) = ")" Then
                        a = a & pz(k * broj) & Abs(k * broj) & "y"
                    End If
                    pb = -1
                    zb = -1
                Else
                    If ulaz(i) = "x" Then
                        a = a & pz(k * broj) & Abs(k * broj) & "x"
                    ElseIf ulaz(i) = "K" Then
                        a = a & pz(k * broj) & Abs(k * broj) & "K"
                    ElseIf ulaz(i) = "y" Then
                        a = a & pz(k * broj) & Abs(k * broj) & "y"
                    ElseIf ulaz(i) = ")" And (ulaz(i - 1) = "K" Or ulaz(i - 1) = "x" Or ulaz(i - 1) = "y") Then
                        GoTo oz
                    ElseIf ulaz(i) = "+" Or ulaz(i) = "-" Or ulaz(i) = ")" Then
                        a = a & pz(k * broj) & Abs(k * broj)
                    End If
                    pb = -1
                    zb = -1
                End If
                promjena = True
            End If

        Next
oz:
        Return a
    End Function

    Public Function zamjeni(ByVal d As String, ByVal b As String, ByVal p As Integer, ByVal z As Integer) As String
        Dim na, nb As Integer
        Dim s() As Char = d
        nb = b.Length
        na = z - p + 1

        For i = z - (na - nb) + 1 To d.Length - 1
            If i + (na - nb) <= d.Length - 1 Then
                s(i) = s(i + (na - nb))
            Else
                s(i) = Nothing
            End If

        Next

        For i = 0 To b.Length - 1
            s(p + i) = b(i)
        Next

        Dim c(d.Length - (na - nb) - 1) As Char

        For i = 0 To d.Length - (na - nb) - 1
            c(i) = s(i)
        Next


        Return c
    End Function

    Public Function pz(ByVal a As Double) As String
        Dim s As Char
        If a < 0 Then
            s = "-"
        Else : s = "+"
        End If
        Return s
    End Function

    Public Function mnozi(ByVal a As String) As String
        Dim pb, zb, start, stp, k As Integer
        Dim na, nb As Integer
        Dim broj1, broj2, um As Double
        Dim nep As Boolean = False
        Dim knep As Boolean = False
        Dim y As Boolean = False
        Dim s() As Char = a
        Dim umnozak As String = Nothing
        pb = -1
        zb = -1
        k = 0
oznaka:
        nep = False
        knep = False
        y = False
        For i = 0 To a.Length - 1

            If IsNumeric(a(i)) Or a(i) = "," Or ((a(i) = "+" Or a(i) = "-") And zb = -1) Then
                If pb = -1 Then
                    pb = i
                    start = i
                End If
                zb = i

            ElseIf a(i) = "*" Then
                If k = 0 Then
                    Double.TryParse(izdvoji_niz(a, pb, zb), broj1)

                End If

                If k = 1 Then
                    Double.TryParse(izdvoji_niz(a, pb - 1, zb), broj1)
                    start = start - 1
                End If

                pb = -1
                zb = -1
                For j = i + 1 To a.Length - 1
                    If IsNumeric(a(j)) Or a(j) = "," Or ((a(j) = "+" Or a(j) = "-") And zb = -1) Then
                        If pb = -1 Then pb = j
                        zb = j
                        stp = j
                    Else
                        Double.TryParse(izdvoji_niz(a, pb, zb), broj2)
                        If a(j) = "x" Then
                            nep = True
                            stp = stp + 1
                        ElseIf a(j) = "K" Then
                            knep = True
                            stp = stp + 1
                        ElseIf a(j) = "y" Then
                            y = True
                            stp = stp + 1
                        End If
                        GoTo label
                    End If
                    If j = a.Length - 1 Then
                        Double.TryParse(izdvoji_niz(a, pb, zb), broj2)
                        If a(j) = "x" Then
                            nep = True
                            stp = stp + 1
                        ElseIf a(j) = "K" Then
                            knep = True
                            stp = stp + 1
                        ElseIf a(j) = "y" Then
                            y = True
                            stp = stp + 1
                        End If
                        GoTo label
                    End If
                Next

label:          um = broj1 * broj2

                If nep = True Then
                    umnozak = pz(um) & Abs(um) & "x"
                ElseIf knep = True Then
                    umnozak = pz(um) & Abs(um) & "K"
                ElseIf y = True Then
                    umnozak = pz(um) & Abs(um) & "y"
                Else : umnozak = pz(um) & Abs(um)
                End If

                nb = umnozak.Length
                na = stp - start + 1

                For j = stp - (na - nb) + 1 To a.Length - 1
                    If j + (na - nb) <= a.Length - 1 Then
                        s(j) = s(j + (na - nb))
                    Else
                        s(j) = Nothing
                    End If

                Next

                For j = 0 To nb - 1
                    s(start + j) = umnozak(j)
                Next

                Dim c(a.Length - (na - nb) - 1) As Char

                For j = 0 To a.Length - (na - nb) - 1
                    c(j) = s(j)
                Next
                a = c
                i = 0

                If susjed = True Then
                    ListBox2.Items.Add(a)
                    promjena = True
                Else
                    ListBox1.Items.Add(a)
                End If
                GoTo oznaka
            Else
                pb = -1
                zb = -1
                If a(i) = "(" Or a(i) = ")" Or a(i) = "x" Or a(i) = "K" Or a(i) = "y" Or a(i) = "=" Then
                    k = 0
                Else
                    k = 1
                End If
            End If
        Next
        Return a
    End Function

    Public Function dijeli(ByVal a As String) As String
        Dim pb, zb, start, stp, k As Integer
        Dim na, nb As Integer
        Dim broj1, broj2, um As Double
        Dim nep As Boolean = False
        Dim knep As Boolean = False
        Dim y As Boolean = False
        Dim s(1000) As Char
        Dim umnozak As String = Nothing
        pb = -1
        zb = -1
        k = 0
        For i = 0 To a.Length - 1
            s(i) = a(i)
        Next
oznaka:
        nep = False
        knep = False
        y = False
        For i = 0 To a.Length - 1

            If IsNumeric(a(i)) Or a(i) = "," Or ((a(i) = "+" Or a(i) = "-") And zb = -1) Then
                If pb = -1 Then
                    pb = i
                    start = i
                End If
                zb = i

            ElseIf a(i) = "/" Then
                If k = 0 Then
                    Double.TryParse(izdvoji_niz(a, pb, zb), broj1)

                End If

                If k = 1 Then
                    Double.TryParse(izdvoji_niz(a, pb - 1, zb), broj1)
                    start = start - 1
                End If

                pb = -1
                zb = -1
                For j = i + 1 To a.Length - 1
                    If IsNumeric(a(j)) Or a(j) = "," Or ((a(j) = "+" Or a(j) = "-") And zb = -1) Then
                        If pb = -1 Then pb = j
                        zb = j
                        stp = j
                    Else
                        Double.TryParse(izdvoji_niz(a, pb, zb), broj2)
                        If a(j) = "x" Then
                            nep = True
                            stp = stp + 1
                        ElseIf a(j) = "K" Then
                            knep = True
                            stp = stp + 1
                        ElseIf a(j) = "y" Then
                            y = True
                            stp = stp + 1
                        End If
                        GoTo label
                    End If
                    If j = a.Length - 1 Then
                        Double.TryParse(izdvoji_niz(a, pb, zb), broj2)
                        If a(j) = "x" Then
                            nep = True
                            stp = stp + 1
                        ElseIf a(j) = "K" Then
                            knep = True
                            stp = stp + 1
                        ElseIf a(j) = "y" Then
                            y = True
                            stp = stp + 1
                        End If
                        GoTo label
                    End If
                Next

label:          um = FormatNumber(broj1 / broj2, 3)

                If nep = True Then
                    umnozak = pz(um) & Abs(um) & "x"
                ElseIf knep = True Then
                    umnozak = pz(um) & Abs(um) & "K"
                ElseIf y = True Then
                    umnozak = pz(um) & Abs(um) & "y"
                Else : umnozak = pz(um) & Abs(um)
                End If

                nb = umnozak.Length
                na = stp - start + 1

                For j = a.Length + (nb - na) - 1 To stp + (nb - na) + 1 Step -1
                    s(j) = s(j + (na - nb))
                Next

                For j = 0 To nb - 1
                    s(start + j) = umnozak(j)
                Next

                Dim c(a.Length - (na - nb) - 1) As Char

                For j = 0 To a.Length - (na - nb) - 1
                    c(j) = s(j)
                Next
                a = c
                i = 0
                If susjed = True Then
                    ListBox2.Items.Add(a)
                    promjena = True
                Else
                    ListBox1.Items.Add(a)
                End If
                GoTo oznaka
            Else
                pb = -1
                zb = -1
                If a(i) = "(" Or a(i) = ")" Or a(i) = "x" Or a(i) = "K" Or a(i) = "y" Or a(i) = "=" Then
                    k = 0
                Else
                    k = 1
                End If
            End If
        Next
        Return a
    End Function

    Function jedispis(ByVal a As Double, ByVal b As Double, ByVal c As Double) As String
        Dim txt As String
        If Abs(a) > 0 Then
            txt = pz(a) & Abs(a) & "K" & pz(b) & Abs(b) & "x" & pz(c) & Abs(c)
        Else
            txt = pz(b) & Abs(b) & "x" & pz(c) & Abs(c)
        End If
        Return txt
    End Function

    Private Sub TextBox_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.Enter, TextBox6.Enter
        activeTextBox = CType(sender, TextBox)
    End Sub

    Private Sub CheckBox26_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox26.CheckedChanged
        If CheckBox26.Checked = True Then
            Sustav = True
        Else
            Sustav = False
        End If

        TextBox1.Focus()
        TextBox1.SelectionStart = TextBox1.TextLength
        TextBox1.ScrollToCaret()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox4.Text = TextBox4.Text & "x"
        TextBox4.Focus()
        TextBox4.SelectionStart = TextBox4.TextLength
        TextBox4.ScrollToCaret()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox4.Text = TextBox4.Text & "K"
        TextBox4.Focus()
        TextBox4.SelectionStart = TextBox4.TextLength
        TextBox4.ScrollToCaret()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox4.Text = TextBox4.Text & "("
        TextBox4.Focus()
        TextBox4.SelectionStart = TextBox4.TextLength
        TextBox4.ScrollToCaret()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        TextBox4.Text = TextBox4.Text & ")"
        TextBox4.Focus()
        TextBox4.SelectionStart = TextBox4.TextLength
        TextBox4.ScrollToCaret()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox4.Text = TextBox4.Text & "="
        TextBox4.Focus()
        TextBox4.SelectionStart = TextBox4.TextLength
        TextBox4.ScrollToCaret()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        pStupanj = False
        TextBox4.Text = ""
        ListBox1.Items.Clear()
        Label12.Text = "Rjesenja:"
        TextBox4.Focus()
        TextBox4.SelectionStart = TextBox4.TextLength
        TextBox4.ScrollToCaret()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        pStupanj = False
        susjed = False
        promjena = False
        TextBox5.Text = ""
        TextBox6.Text = ""
        ListBox2.Items.Clear()
        Label21.Text = "Rjesenja:"
        TextBox5.Focus()
        TextBox5.SelectionStart = TextBox5.TextLength
        TextBox5.ScrollToCaret()
        activeTextBox = TextBox5
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If (activeTextBox IsNot Nothing) Then
            activeTextBox.Text = activeTextBox.Text & "x"
            activeTextBox.Focus()
            activeTextBox.SelectionStart = activeTextBox.TextLength
            activeTextBox.ScrollToCaret()
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If (activeTextBox IsNot Nothing) Then
            activeTextBox.Text = activeTextBox.Text & "K"
            activeTextBox.Focus()
            activeTextBox.SelectionStart = activeTextBox.TextLength
            activeTextBox.ScrollToCaret()
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If (activeTextBox IsNot Nothing) Then
            activeTextBox.Text = activeTextBox.Text & "("
            activeTextBox.Focus()
            activeTextBox.SelectionStart = activeTextBox.TextLength
            activeTextBox.ScrollToCaret()
        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If (activeTextBox IsNot Nothing) Then
            activeTextBox.Text = activeTextBox.Text & ")"
            activeTextBox.Focus()
            activeTextBox.SelectionStart = activeTextBox.TextLength
            activeTextBox.ScrollToCaret()
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        If (activeTextBox IsNot Nothing) Then
            activeTextBox.Text = activeTextBox.Text & "="
            activeTextBox.Focus()
            activeTextBox.SelectionStart = activeTextBox.TextLength
            activeTextBox.ScrollToCaret()
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        If (activeTextBox IsNot Nothing) Then
            activeTextBox.Text = activeTextBox.Text & "y"
            activeTextBox.Focus()
            activeTextBox.SelectionStart = activeTextBox.TextLength
            activeTextBox.ScrollToCaret()
        End If
    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
        CheckBox6.Checked = False
        CheckBox7.Checked = False
        CheckBox8.Checked = False
        CheckBox9.Checked = False
        CheckBox10.Checked = False
        CheckBox11.Checked = False
        CheckBox20.Checked = False
        CheckBox21.Checked = False
        CheckBox22.Checked = False
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        CheckBox12.Checked = False
        CheckBox15.Checked = False
        CheckBox13.Checked = False
        CheckBox23.Checked = False
        CheckBox24.Checked = False
        CheckBox25.Checked = False
        CheckBox16.Checked = False
        CheckBox14.Checked = False
        CheckBox17.Checked = False
        CheckBox18.Checked = False
        CheckBox19.Checked = False
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        CheckBox27.Checked = False
        CheckBox28.Checked = False
        CheckBox31.Checked = False
        CheckBox33.Checked = False
        CheckBox32.Checked = False
    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        CheckBox40.Checked = False
        CheckBox39.Checked = False
        CheckBox35.Checked = False
        CheckBox34.Checked = False
        CheckBox36.Checked = False
        CheckBox41.Checked = False
        CheckBox42.Checked = False
        CheckBox43.Checked = False
    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        CheckBox53.Checked = False
        CheckBox52.Checked = False
        CheckBox55.Checked = False
        CheckBox54.Checked = False
        CheckBox48.Checked = False
        CheckBox47.Checked = False
        CheckBox57.Checked = False
        CheckBox56.Checked = False
        CheckBox49.Checked = False
        CheckBox46.Checked = False
        CheckBox29.Checked = False
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        CheckBox30.Checked = False
        CheckBox44.Checked = False
        CheckBox37.Checked = False
        CheckBox45.Checked = False
        CheckBox50.Checked = False
        CheckBox51.Checked = False
        CheckBox61.Checked = False
        CheckBox38.Checked = False
        CheckBox58.Checked = False
        CheckBox59.Checked = False
        CheckBox60.Checked = False
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        CheckBox72.Checked = False
        CheckBox68.Checked = False
        CheckBox69.Checked = False
        CheckBox62.Checked = False
        CheckBox65.Checked = False
        CheckBox70.Checked = False
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        CheckBox73.Checked = False
        CheckBox66.Checked = False
        CheckBox67.Checked = False
        CheckBox63.Checked = False
        CheckBox80.Checked = False
        CheckBox64.Checked = False
        CheckBox71.Checked = False
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        CheckBox79.Checked = False
        CheckBox77.Checked = False
        CheckBox76.Checked = False
        CheckBox74.Checked = False
        CheckBox81.Checked = False
        CheckBox82.Checked = False
        CheckBox75.Checked = False
        CheckBox78.Checked = False
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        CheckBox83.Checked = False
        CheckBox84.Checked = False
        CheckBox85.Checked = False
        CheckBox86.Checked = False
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        CheckBox90.Checked = False
        CheckBox89.Checked = False
        CheckBox88.Checked = False
        CheckBox87.Checked = False
    End Sub

    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox7.Focus()
        TextBox7.SelectionStart = TextBox7.TextLength
        TextBox7.ScrollToCaret()
    End Sub
End Class
