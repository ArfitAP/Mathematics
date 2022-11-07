
' Redaka: 33

Public Module Kut
    Public stupnjevi As Double = 0
    Public minute As Double = 0
    Public sekunde As Double = 0
    Public Kontrola As Boolean = True
    Public tekst As String
    Public mj As Integer
    Public ax, bx, cx, ax1, bx1, cx1 As Double


    Public Function Provjeri() As Boolean
        If stupnjevi < 0 Or (Kontrola = True And stupnjevi >= 180) Or minute < 0 Or minute >= 60 Or sekunde < 0 Or sekunde >= 60 Or (stupnjevi = 0 And minute = 0 And sekunde = 0) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Izracunaj() As Double
        Dim rez As Double = stupnjevi + minute / 60 + sekunde / 3600
        Return rez
    End Function

    Public Sub Reset()
        stupnjevi = 0
        minute = 0
        sekunde = 0
    End Sub

End Module
