
Option Explicit On
Option Strict On

' ===================================================================
' Easter Calculator 0.1
' -------------------------------------------------------------------
' Copyright (c) 2007 David Pinch.
' Download from http://www.thoughtproject.com/Snippets/Easter/
' ===================================================================

Friend NotInheritable Class EasterDate

    Friend Shared Function Feriados(Ano As Integer) As List(Of Tuple(Of Date, String))
        Dim Pascoa As Date = EasterDate.GetEasterDate(Ano)

        Dim F As New List(Of Tuple(Of Date, String))
        F.Add(New Tuple(Of Date, String)(New Date(Ano, 1, 1), "Ano Novo"))

        F.Add(New Tuple(Of Date, String)(DateAdd(DateInterval.Day, -48, Pascoa), "Segunda de Carnaval"))
        F.Add(New Tuple(Of Date, String)(DateAdd(DateInterval.Day, -47, Pascoa), "Terça de Carnaval"))
        F.Add(New Tuple(Of Date, String)(DateAdd(DateInterval.Day, -2, Pascoa), "Sexta Feira da Paixão"))
        F.Add(New Tuple(Of Date, String)(Pascoa, "Pascoa"))
        F.Add(New Tuple(Of Date, String)(DateAdd(DateInterval.Day, 60, Pascoa), "Corpus Christi"))

        F.Add(New Tuple(Of Date, String)(New Date(Ano, 4, 21), "Tiradentes"))
        F.Add(New Tuple(Of Date, String)(New Date(Ano, 5, 1), "Dia do Trabalhador"))
        F.Add(New Tuple(Of Date, String)(New Date(Ano, 9, 7), "Independência"))
        F.Add(New Tuple(Of Date, String)(New Date(Ano, 10, 12), "Nossa Senhora Aparecida"))
        F.Add(New Tuple(Of Date, String)(New Date(Ano, 11, 2), "Finados"))
        F.Add(New Tuple(Of Date, String)(New Date(Ano, 11, 15), "Proclamacão da Republica"))
        F.Add(New Tuple(Of Date, String)(New Date(Ano, 12, 25), "Natal"))

        Return F
    End Function


    ' ===============================================================
    ' New (default constructor)
    ' ---------------------------------------------------------------
    ' The default constructor is marked as private because the class
    ' contains only shared (static) methods.
    ' ===============================================================

    Private Sub New()
    End Sub

    ' ===============================================================
    ' Easter
    ' ---------------------------------------------------------------
    ' Calculates the date of Easter using an algorithm that was first
    ' published in Butcher's Ecclesiastical Calendar (1876).  It is
    ' valid for all years in the Gregorian calendar (1583+).  The
    ' code is based on an implementation by Peter Duffett-Smith in
    ' Practical Astronomy with your Calculator (3rd Edition).
    ' ===============================================================

    Public Shared Function GetEasterDate(ByVal Year As Integer) As Date

        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        Dim d As Integer
        Dim e As Integer
        Dim f As Integer
        Dim g As Integer
        Dim h As Integer
        Dim i As Integer
        Dim k As Integer
        Dim l As Integer
        Dim m As Integer
        Dim n As Integer
        Dim p As Integer

        If Year < 1583 Then

            Throw New ArgumentOutOfRangeException("year", "year must be after or equals 1583")

        Else

            ' Step 1: Divide the year by 19 and store the
            ' remainder in variable A.  Example: If the year
            ' is 2000, then A is initialized to 5.

            a = Year Mod 19

            ' Step 2: Divide the year by 100.  Store the integer
            ' result in B and the remainder in C.

            b = Year \ 100
            c = Year Mod 100

            ' Step 3: Divide B (calculated above).  Store the
            ' integer result in D and the remainder in E.

            d = b \ 4
            e = b Mod 4

            ' Step 4: Divide (b+8)/25 and store the integer
            ' portion of the result in F.

            f = (b + 8) \ 25

            ' Step 5: Divide (b-f+1)/3 and store the integer
            ' portion of the result in G.

            g = (b - f + 1) \ 3

            ' Step 6: Divide (19a+b-d-g+15)/30 and store the
            ' remainder of the result in H.

            h = (19 * a + b - d - g + 15) Mod 30

            ' Step 7: Divide C by 4.  Store the integer result
            ' in I and the remainder in K.

            i = c \ 4
            k = c Mod 4

            ' Step 8: Divide (32+2e+2i-h-k) by 7.  Store the
            ' remainder of the result in L.

            l = (32 + 2 * e + 2 * i - h - k) Mod 7

            ' Step 9: Divide (a + 11h + 22l) by 451 and
            ' store the integer portion of the result in M.

            m = (a + 11 * h + 22 * l) \ 451

            ' Step 10: Divide (h + l - 7m + 114) by 31.  Store
            ' the integer portion of the result in N and the
            ' remainder in P.

            n = (h + l - 7 * m + 114) \ 31
            p = (h + l - 7 * m + 114) Mod 31

            ' At this point p+1 is the day on which Easter falls.
            ' n is 3 for March or 4 for April.

            Return DateSerial(Year, n, p + 1)

        End If

    End Function

End Class
