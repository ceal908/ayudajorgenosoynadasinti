Option Strict Off
Imports System.Globalization

Public Class Form1

    Dim Fechasel As String = Now
    Dim pulsado As Boolean = False

    Dim Dia As Integer = Now.Day
    Dim Mes As Integer = Now.Month
    Dim Año As Integer = Now.Year

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        PasarFechaCalendario = Now

        If IsDate(PasarFechaCalendario) = True Then
            Lbmes.Text = UCase(Mid(Format(CDate(PasarFechaCalendario), "MMMM") & "de" & CDate(PasarFechaCalendario).Year, 1, 1)) & LCase(Mid(Format(CDate(PasarFechaCalendario), "MMMM") & "de" & CDate(PasarFechaCalendario).Year, 2, 100))
            Mes = CDate(PasarFechaCalendario).Month
            Fechasel = Format(CDate(PasarFechaCalendario), "dd/MM/yy")
            BotonesCalendario(PasarFechaCalendario)
        Else
            Lbmes.Text = UCase(Mid(Format(Now, "MMMM") & "de" & Now.Year, 1, 1)) & LCase(Mid(Format(Now, "MMMM") & "de" & Now.Year, 2, 100))
            Mes = Now.Month
            Fechasel = Format(Now, "dd/MM/yy")
            BotonesCalendario(Now)

        End If

    End Sub

    Sub BotonesCalendario(Fecha As String)

        Dim Dialni As Integer, UltimoDiaMes As Integer
        Dim h As Date = CDate("01/" & CDate(Fechasel).Month & "/" & CDate(Fechasel).Year)
        Dialni = h.DayOfWeek
        Dim F As Date
        F = DateAdd(DateInterval.Month, 1, h)
        UltimoDiaMes = DateAdd(DateInterval.Day, -1, F).Day
        If Dialni = 0 Then Dialni = 7

        For I As Integer = 1 To 42
            Me.Controls("label" & I).Visible = False
            Me.Controls("label" & I).ForeColor = Color.Black
            Me.Controls("label" & I).BackColor = Color.Ivory
        Next

        For I As Integer = 1 To Dialni - 1
            Me.Controls("label" & I).Visible = False
        Next

        For I As Integer = 1 To UltimoDiaMes
            Me.Controls("label" & I + Dialni - 1).BackColor = Color.Ivory
            Me.Controls("label" & I + Dialni - 1).ForeColor = Color.Black
            Me.Controls("label" & I + Dialni - 1).Text = I
            Me.Controls("label" & I + Dialni - 1).Visible = True
            Me.Controls("label" & I + Dialni - 1).Tag = CDate(I & "/" & CDate(Fechasel).Month & "/" & CDate(Fechasel).Year)

            If Format(CDate(I & "/" & CDate(Fechasel).Month & "/" & CDate(Fechasel).Year), "dd/MM/yyyy") = Format(CDate(Now), "dd/MM/yyyy") Then
                Me.Controls("label" & I + Dialni - 1).ForeColor = Color.Black
                Me.Controls("label" & I + Dialni - 1).BackColor = Color.LightCyan
            End If

            If IsDate(PasarFechaCalendario) = True Then
                If Format(CDate(I & "/" & CDate(Fechasel).Month & "/" & CDate(Fechasel).Year), "dd/MM/yyyy") = Format(CDate(PasarFechaCalendario), "dd/MM/yyyy") Then
                    Me.Controls("label" & I + Dialni - 1).ForeColor = Color.Black
                    Me.Controls("label" & I + Dialni - 1).BackColor = Color.Firebrick
                End If
            End If
        Next

    End Sub

End Class
