Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports SHDocVw
Imports System.Data.SqlClient
Imports ICSharpCode.SharpZipLib.Zip
Imports System.Web
Imports System.Data

Imports System
Imports System.Diagnostics
Imports System.Threading
Imports System.Threading.Tasks

Public Class Form1

    Public ftp As EnterpriseDT.Net.Ftp.FTPConnection


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim x As New AWebExplorer("dba", "PortalesQA", "aportales", "Aportales12")
        'x.getPH(New Date(2014, 2, 6), "I", 0)



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Try
        '    ftp = New EnterpriseDT.Net.Ftp.FTPConnection
        '    ftp.ServerAddress = "BWDES"
        '    ftp.ServerPort = "921"
        '    ftp.UserName = "ftp_piagui"
        '    ftp.Password = "ftppiagui"
        '    ftp.Connect()



        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Information)
        'End Try


    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.Timer1.Start() 'OJO, ESTE CÓDIGO ES PARA QUE SE EJCEUTE EN AUTOMATICO, MIENTRAS ESTE PROBANDO, NO DESCOMENTAR
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-ES")


        'MsgBox("Inicio")

        'Dim x As New AWebExplorer("dba", "PortalesQA", "aportales", "Aportales12")

        'x.getPH(New Date(2013, 12, 1), "V", 1)

        'MsgBox("FIN")
        'System.Threading.Thread.Sleep(2000)


    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Timer1.Stop()

        System.Threading.Thread.Sleep(2000)
        Me.Hide()
        Dim server As String = System.Configuration.ConfigurationSettings.AppSettings("server")
        Dim bd As String = System.Configuration.ConfigurationSettings.AppSettings("bd")
        Dim user As String = System.Configuration.ConfigurationSettings.AppSettings("user")
        Dim pwd As String = System.Configuration.ConfigurationSettings.AppSettings("pwd")

        Dim strCon As String = "Data Source=" & server & ";Initial Catalog=" & bd & ";User ID=" & user & ";Password=" & pwd & ""

        Dim con As New SqlClient.SqlConnection(strCon)
        Dim strSql As String = ""
        Dim strEstatus As String = ""
        Dim fecha As Date
        Dim archivo As String = ""
        Dim secuencia As Integer = 0
        Dim idDescarga As Integer = 0
        Dim ws As New AWebExplorer(server, bd, user, pwd)

        Try
            Me.TextBox1.Text = "Iniciando"
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-ES")


            strSql = "Select * from estatusPH where Estatus = 'EN_ESPERA'"
            Dim sqlCmd As New SqlClient.SqlCommand(strSql, con)
            Dim sqlRdr As SqlClient.SqlDataReader
            Me.TextBox1.Text &= vbCrLf & "Conectando a BD"


            con.Open()

            sqlRdr = sqlCmd.ExecuteReader()


            If sqlRdr.HasRows Then
                While sqlRdr.Read
                    strEstatus = sqlRdr.Item("Estatus")
                    archivo = sqlRdr.Item("TipoArchivo")
                    fecha = sqlRdr.Item("fecha")
                    secuencia = sqlRdr.Item("Secuencia")
                    idDescarga = sqlRdr.Item("idDescarga")
                End While
            End If
            con.Close()
            Me.TextBox1.Text &= vbCrLf & "cerrando BD"
            Me.TextBox1.Text &= vbCrLf & "Estatus: " & strEstatus
            Me.TextBox1.Text &= vbCrLf & "Estatus: " & archivo
            Me.TextBox1.Text &= vbCrLf & "Estatus: " & fecha


            If strEstatus = "EN_ESPERA" Then
                strSql = "UPDATE EstatusPH SET estatus='EN_PROCESO' WHERE fecha='" & fecha.ToString("dd/MM/yyyy") & "' AND TipoArchivo='" & archivo & "' and estatus ='" & strEstatus & "'"
                Dim sqlcmd2 As New SqlCommand(strSql, con)
                con.Open()
                sqlcmd2.ExecuteNonQuery()
                con.Close()

                Dim resultado As String = ""
                'Dim x As New AWebExplorer(server, bd, user, pwd)
                resultado = ws.getPH(fecha, archivo, secuencia)
                resultado = resultado.Replace("'", " ")

                strSql = "UPDATE EstatusPH SET Resultado='" & resultado & "' , Estatus ='TERMINADO'  WHERE fecha='" & fecha.ToString("dd/MM/yyyy") & "' AND TipoArchivo='" & archivo & "' and estatus ='EN_PROCESO'"
                Dim sqlcmd3 As New SqlCommand(strSql, con)
                con.Open()
                sqlcmd3.ExecuteNonQuery()
                con.Close()

            End If

            'x.getPH(New Date(2013, 12, 1), "T", 1)

        Catch ex As Exception
            'Dim ws As New AWebExplorer(server, bd, user, pwd)

            ws.setLog(idDescarga, "E", ex.Message.ToString())

            Dim xml As String = String.Empty
            xml = ws.GetXML(1, archivo, fecha)
            xml = xml.Replace("'", " ")

            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            Me.TextBox1.Text = ex.Message.ToString()
            strSql = "UPDATE EstatusPH SET Resultado='" & xml & "' , Estatus ='ERROR'  WHERE fecha='" & fecha.ToString("dd/MM/yyyy") & "' AND TipoArchivo='" & archivo & "' and estatus ='EN_PROCESO' and idDescarga= " & idDescarga & ""
            Dim sqlcmd4 As New SqlCommand(strSql, con)
            con.Open()
            sqlcmd4.ExecuteNonQuery()
            con.Close()
        Finally
            ws.terminarTodo()
        End Try

        Me.Close()

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim stopWatch As New Stopwatch()
        stopWatch.Start()

        Timer1.Stop()
        Dim a As New AWebExplorer("dba", "Portalesqa", "aportales", "Aportales12")

        If rdbtnventas.Checked Then
            Dim x As String = a.getPH("28/12/2015", "V", 0) 'descarga de ventas

            'MsgBox("ventas")
        End If

        If rdbtninventarios.Checked Then
            Dim x As String = a.getPH("29/12/2015", "I", 0) 'descarga de inventario
            'MsgBox("Inventarios")
        End If


        stopWatch.Stop()
        ' Get the elapsed time as a TimeSpan value.
        Dim ts As TimeSpan = stopWatch.Elapsed

        ' Format and display the TimeSpan value.
        Dim elapsedTime As String = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10)

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("C:\Codes\testing\test.txt", True)
        file.WriteLine("RunTime " + elapsedTime)
        file.Close()

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Timer1.Stop()
        Dim a As New AWebExplorer("dba", "PortalesQA", "aportales", "Aportales12")

        Dim x As String = a.getInventariosPH("31/05/2018", "Venta", "1")
    End Sub

End Class