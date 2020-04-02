Imports SHDocVw
Imports System.Data.SqlClient
Imports ICSharpCode.SharpZipLib.Zip
Imports System.Web
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Data


Public Class AWebExplorer
    Public ie As SHDocVw.InternetExplorer
    Public proceso As Long
    Public handler As Long
    Public strdirectorios() As String
    Public lista_archivos As System.Data.DataTable
    Public wb As IWebBrowserApp
    Public document As mshtml.HTMLDocument
    Public httpRequest As System.Net.HttpWebRequest
    Public httpResponse As System.Net.HttpWebResponse
    Public WebClient As Net.WebClient
    Public message As String
    Private conn As Data.SqlClient.SqlConnection
    Private sql As Data.SqlClient.SqlCommand
    Private rs As Data.SqlClient.SqlDataReader
    Private id_descarga As Integer
    Private header As Net.WebHeaderCollection
    Private cookiesContainer As Net.CookieContainer
    Public TArchivos As System.Data.DataTable
    Public Visible As Boolean
    Public ftp As EnterpriseDT.Net.Ftp.FTPConnection
    Public timeout As Integer

    Public hilo As Threading.Thread
    Public time As Integer
    Public sts_prc As String = ""

    Public Server As String
    Public bd As String
    Public user As String
    Public pwd As String

    Public server_ftp As String
    Public user_ftp As String
    Public pwd_ftp As String
    Public port_ftp As String

    Const WM_CLOSE = &H10
    Const INFINITE = &HFFFFFFFF
    Const SYNCHRONIZE = &H100000

    <System.Runtime.InteropServices.DllImport("user32.DLL")> _
    Private Shared Function SendMessage( _
            ByVal hWnd As System.IntPtr, ByVal wMsg As Integer, _
            ByVal wParam As Integer, ByVal lParam As Integer _
            ) As Integer
    End Function

    Private Declare Function GetWindowThreadProcessId Lib "user32" _
         (ByVal hwnd As Long, _
    ByVal lpdwProcessId As Long) As Long

    Private Declare Function GetPIDByHWnd Lib "user32" _
     (ByVal hwnd As Long) As Long

    Private Declare Function OpenProcess Lib "kernel32" _
         (ByVal dwDesiredAccess As Long, _
         ByVal bInheritHandle As Long, _
         ByVal dwProcessId As Long) As Long

    Private Declare Function FindWindow Lib "user32" _
   Alias "FindWindowA" _
   (ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long

    Public Sub New(ByVal xServer As String, ByVal xbd As String, ByVal xUser As String, ByVal xPwd As String)
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-ES")
        conn = New Data.SqlClient.SqlConnection("Data Source=dba;Initial Catalog=Portales;Integrated Security=No;User ID=aportales;Password=Aportales12;Persist Security Info=true")
        TArchivos = New System.Data.DataTable("Archivos")
        TArchivos.Columns.Add("Nombre")
        TArchivos.Columns.Add("Ruta")
        Visible = False
        ftp = New EnterpriseDT.Net.Ftp.FTPConnection
        Server = xServer
        bd = xbd
        user = xUser
        pwd = xPwd
        conn = New Data.SqlClient.SqlConnection("Data Source=" & Server & ";Initial Catalog=" & bd & ";Integrated Security=No;User ID=" & user & ";Password=" & pwd & ";Persist Security Info=true")
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("select * from config ", conn)
        rs = sql.ExecuteReader
        rs.Read()
        Visible = rs!visible
        timeout = rs!timeout
        rs.Close()
        conn.Close()
    End Sub
    Protected Overrides Sub Finalize()
        If Not ie Is Nothing Then
            ie.Quit()
            ie = Nothing
        End If
    End Sub
    Public Function getFTP(ByVal cliente As String) As Integer
        Dim res As Integer = 0
        Try
            conn.Open()
            sql = New Data.SqlClient.SqlCommand("select * from conf_ftp where cliente=@cliente", conn)
            sql.Parameters.AddWithValue("@cliente", cliente.ToUpper)
            rs = sql.ExecuteReader
            If rs.Read Then
                server_ftp = rs!servidor
                user_ftp = rs!usuario
                pwd_ftp = rs!pwd
                port_ftp = rs!puerto
            Else
                res = 1
                message = "Error: No se encontro configuración de FTP"
            End If
            rs.Close()
            conn.Close()
        Catch ex As Exception
            res = 1
            message = "Error: " & ex.Message
        End Try
        Return res
    End Function

    Public Sub iniciar()
        ie = New SHDocVw.InternetExplorerClass()
        wb = DirectCast(ie, IWebBrowserApp)
        ie.Visible = True
        handler = ie.HWND
        Dim procesos() As Process = Process.GetProcessesByName("iexplore")
        Dim x As Integer = 0
        While x < procesos.Length
            Dim nombre = procesos(x).ProcessName
            If procesos(x).MainWindowHandle = handler Then
                proceso = procesos(x).Id
            End If
            x = x + 1
        End While
        ie.Visible = False
    End Sub
    Public Function terminar() As Integer
        Try
            ie.Quit()
            SendMessage(proceso, 16, 0, 0)
            Dim procesos() As Process = Process.GetProcessesByName("iexplore")
            Dim x As Integer = 0
            While x < procesos.Length
                Dim nombre = procesos(x).ProcessName
                If procesos(x).Id = proceso Then
                    procesos(x).Kill()
                End If
                x = x + 1
            End While
        Catch ex As Exception
            message = "Error: " & ex.Message
            Return 1
        End Try

        Return 0
    End Function

    Public Sub buscaArchivos(ByVal ruta As String)
        lista_archivos = New System.Data.DataTable("archivos")
        lista_archivos.Columns.Add("archivo")
        lista_archivos.Columns.Add("ruta")
        ReDim strdirectorios(0)
        SearchDirectory(ruta)
    End Sub

    Public Function Browser(ByVal url As String) As Integer
        Try
            Dim o As Object = Nothing
            wb.Visible = Visible
            wb.Navigate(url, o, o, o, o)
            time = 0
            While wb.Busy And time < timeout
                System.Threading.Thread.Sleep("1000")
                setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                time = time + 1
            End While
            If time < timeout Then
                System.Threading.Thread.Sleep("3000")

                document = wb.Document
                Dim r() As String = url.Split("?")
                setLog(id_descarga, "I", "Pagina cargada: " & r(0))

                Dim cookies As String = document.cookie
                Dim domain As String = document.domain
                Dim c() As String = cookies.Split(";")
                Dim x As Integer = 0
                x = 0
                cookiesContainer = New Net.CookieContainer
                While x < c.Length
                    Dim val() As String = c(x).Split("=")
                    Dim c1 As New Net.Cookie
                    c1.Domain = domain
                    c1.Path = "/"
                    c1.Name = val(0).Trim
                    If val.Length > 1 Then
                        c1.Value = val(1).Trim
                    End If
                    cookiesContainer.Add(c1)
                    x = x + 1
                End While

                Return 0
            Else
                setLog(id_descarga, "E", "El sitio tardo demasiado en contestar. Intentar mas tarde...")
                Return 1
            End If
        Catch ex As Exception
            setLog(id_descarga, "E", "Error:" & ex.Message)
            Return 1
        End Try
    End Function

    Public Function setAttribute(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal valor As String) As Integer
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
        Dim iHTMLEle As mshtml.IHTMLElement
        Dim str As String = ""       
        For Each iHTMLEle In iHTMLCol
            If Not iHTMLEle.getAttribute(atributo) Is Nothing Then
                str = iHTMLEle.getAttribute(atributo).ToString
                If str.ToUpper.Trim.Equals(nombre.ToUpper.Trim) Then
                    iHTMLEle.setAttribute("value", valor)
                    setLog(id_descarga, "I", "Asignacion de valor [ " & nombre & " : " & valor & "]")
                    Return 0
                    Exit For
                ElseIf str.ToUpper.Contains(nombre.ToUpper.Trim) Then
                    iHTMLEle.setAttribute("value", valor)
                    setLog(id_descarga, "I", "Asignacion de valor [ " & nombre & " : " & valor & "]")
                    Return 0
                    Exit For               
                End If
            End If
        Next
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre)
        Return 1 
End Function

    Public Function getAttributeValue(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal AtributoValor As String, ByRef valor As String) As Integer
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
        Dim iHTMLEle As mshtml.IHTMLElement
        Dim str As String = ""
        For Each iHTMLEle In iHTMLCol
            If Not iHTMLEle.getAttribute(atributo) Is Nothing Then
                str = iHTMLEle.getAttribute(atributo).ToString
                If str.ToUpper.Equals(nombre.ToUpper) Then
                    valor = iHTMLEle.getAttribute(AtributoValor).ToString
                    setLog(id_descarga, "I", "Asignacion de valor [ " & nombre & " : " & valor & "]")
                    Return 0
                    Exit For
                End If
            End If
        Next
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre)
        Return 1
    End Function

    Public Function setCheck(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal valor As String) As Integer
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
        Dim iHTMLEle As mshtml.IHTMLElement
        Dim str As String = ""
        Dim val As String = ""
        For Each iHTMLEle In iHTMLCol
            If Not iHTMLEle.getAttribute(atributo) Is Nothing Then
                str = iHTMLEle.getAttribute(atributo).ToString
                If Not iHTMLEle.getAttribute("value") Is Nothing Then
                    val = iHTMLEle.getAttribute("value").ToString
                End If
                If str.ToUpper.Equals(nombre.ToUpper) And val.ToUpper.Equals(valor.ToUpper) Then
                    iHTMLEle.click()
                    setLog(id_descarga, "I", "Asignacion de valor [ " & nombre & " : " & valor & "]")
                    Return 0
                    Exit For
                End If
            End If
        Next
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre)
        Return 1
    End Function

    Public Function validaSitio(ByVal url As String) As Integer
        Dim http As Net.HttpWebRequest = Net.HttpWebRequest.Create(url)
        http.KeepAlive = False
        http.ConnectionGroupName = Guid.NewGuid.ToString
        Try
            Dim httpr As Net.HttpWebResponse = http.GetResponse
            header = New Net.WebHeaderCollection
            header = httpr.Headers
            Return 0
        Catch ex As Exception
            message = ex.Message
            Return 1
        End Try

    End Function

    Public Function sendClick(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String) As Integer
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
        Dim iHTMLEle As mshtml.IHTMLElement
        Dim str As String = ""
        For Each iHTMLEle In iHTMLCol
            If Not iHTMLEle.getAttribute(atributo) Is Nothing Then
                str = iHTMLEle.getAttribute(atributo).ToString
                If str.ToUpper.Equals(nombre.ToUpper) Then
                    setLog(id_descarga, "I", "Ejecucion de evento Click [ " & nombre & "]")
                    iHTMLEle.click()
                    time = 0
                    While wb.Busy And time < timeout
                        System.Threading.Thread.Sleep("1000")
                        setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                        time = time + 1
                    End While
                    If time < timeout Then
                        Return 0
                    Else
                        setLog(id_descarga, "E", "El servidor tardo mas de lo esperado en contestar...")
                        Return 1
                    End If
                    Exit For
                ElseIf str.ToUpper.Contains(nombre.ToUpper) Then
                    setLog(id_descarga, "I", "Ejecucion de evento Click [ " & nombre & "]")
                    iHTMLEle.click()
                    time = 0
                    While wb.Busy And time < timeout
                        System.Threading.Thread.Sleep("1000")
                        setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                        time = time + 1
                    End While
                    If time < timeout Then
                        Return 0
                    Else
                        setLog(id_descarga, "E", "El servidor tardo mas de lo esperado en contestar...")
                        Return 1
                    End If
                    Exit For
                End If
            End If
        Next
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre)
        Return 1
    End Function
    Public Function sendLink(ByVal valor As String) As Integer
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.links
        Dim iHTMLEle As mshtml.IHTMLElement
        For Each iHTMLEle In iHTMLCol

            If Not iHTMLEle.innerText Is Nothing AndAlso iHTMLEle.innerText = valor Then
                setLog(id_descarga, "I", "Ejecucion de evento Link [ " & iHTMLEle.innerText & "]")
                iHTMLEle.click()
                time = 0
                    While wb.Busy And time < 1800
                        System.Threading.Thread.Sleep("1000")
                        setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                        time = time + 1
                    End While
                    If time < 1800 Then
                        Return 0
                    Else
                        setLog(id_descarga, "E", "El servidor tardo mas de lo esperado en contestar...")
                        Return 1
                    End If
                Exit For
            ElseIf Not iHTMLEle.innerText Is Nothing AndAlso iHTMLEle.innerText.ToString.Contains(valor) Then
                setLog(id_descarga, "I", "Ejecucion de evento Link [ " & iHTMLEle.innerText & "]")
                iHTMLEle.click()
                time = 0
                    While wb.Busy And time < 1800
                        System.Threading.Thread.Sleep("1000")
                        setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                        time = time + 1
                    End While
                    If time < 1800 Then
                        Return 0
                    Else
                        setLog(id_descarga, "E", "El servidor tardo mas de lo esperado en contestar...")
                        Return 1
                    End If
                Exit For
            ElseIf Not iHTMLEle.innerHTML Is Nothing AndAlso iHTMLEle.innerHTML.ToString.Contains(valor) Then
                setLog(id_descarga, "I", "Ejecucion de evento Link [ " & iHTMLEle.innerText & "]")
                iHTMLEle.click()
                time = 0
                    While wb.Busy And time < 1800
                        System.Threading.Thread.Sleep("1000")
                        setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                        time = time + 1
                    End While
                    If time < 1800 Then
                        Return 0
                    Else
                        setLog(id_descarga, "E", "El servidor tardo mas de lo esperado en contestar...")
                        Return 1
                    End If
                Exit For
                'ElseIf Not iHTMLEle.id Is Nothing AndAlso iHTMLEle.innerHTML.ToString.Contains(valor) Then
                '    setLog(id_descarga, "I", "Ejecucion de evento Link [ " & iHTMLEle.innerText & "]")
                '    iHTMLEle.click()
                '    While wb.Busy
                '        System.Threading.Thread.Sleep("1000")
                '    End While
                '    Return 0
                '    Exit For
            End If


        Next
        setLog(id_descarga, "E", "Elemento no encontrado: " & valor)
        Return 1
    End Function


    Public Function buscaValor(ByVal campo As String, ByVal valor As String) As Boolean
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(campo)
        Dim iHTMLEle As mshtml.IHTMLElement
        Dim str As String = ""
        For Each iHTMLEle In iHTMLCol
            Dim html As String = iHTMLEle.innerHTML
            If Not html Is Nothing Then
                If html.Contains(valor) Then
                    Return True
                End If
            End If
        Next
        Return False
    End Function


    Public Function Sears(ByVal fecha As Date, ByVal Archivo As String) As String

        Dim err As Integer = 0
        Dim x As Integer = 0

        newDescarga("SEARS", fecha)
        err = getFTP("SEARS")
        If err = 0 Then

            conn.Open()
            sql = New Data.SqlClient.SqlCommand("select * from clogin where clogportal='SEARS' ", conn)
            Dim tabla As New System.Data.DataTable
            tabla.Load(sql.ExecuteReader)
            conn.Close()

            setLog(id_descarga, "I", "Iniciando proceso...")
            setLog(id_descarga, "I", "Tipo de descarga solicitada: " & Archivo)


            While x < tabla.Rows.Count And err = 0

                setLog(id_descarga, "I", "Procesando cliente: " & tabla.Rows(x).Item(3).ToString)
                If Archivo = "V" And err = 0 Then 'solo ventas
                    err = getVentasSears(fecha, tabla.Rows(x))
                End If
                If Archivo = "I" And err = 0 Then 'solo inventarios
                    err = getInventarioSears(fecha, tabla.Rows(x))
                End If
                If Archivo = "T" And err = 0 Then 'Archivo venas e inventarios
                    err = getVentasSears(fecha, tabla.Rows(x))
                    If err = 0 Then
                        err = getInventarioSears(fecha, tabla.Rows(x))
                    End If
                End If

                x = x + 1
            End While

            If err = 0 Then
                setLog(id_descarga, "I", "Proceso terminado correctamente...")
            Else
                setLog(id_descarga, "E", "Proceso terminado con errores, favor de revisar secuencia...")
                setLog(id_descarga, "I", "Eliminado archivos generados...")
                Dim ax As Integer = 0
                While ax < TArchivos.Rows.Count
                    'If System.IO.File.Exists(TArchivos.Rows(ax).Item(1).ToString() & TArchivos.Rows(ax).Item(0).ToString()) Then
                    '    System.IO.File.Delete(TArchivos.Rows(ax).Item(1).ToString() & TArchivos.Rows(ax).Item(0).ToString())
                    'End If
                    Try
                        ftp.ServerAddress = server_ftp
                        ftp.ServerPort = port_ftp
                        ftp.UserName = user_ftp
                        ftp.Password = pwd_ftp
                        ftp.Connect()
                        ftp.DeleteFile(TArchivos.Rows(ax).Item(0).ToString())
                        ftp.Close()
                        'System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                        setLog(id_descarga, "I", "Archivo eliminado: " & TArchivos.Rows(ax).Item(0).ToString())
                    Catch ex As Exception
                        setLog(id_descarga, "E", "Error:" & ex.Message)
                    End Try
                    ax = ax + 1
                End While
            End If

            conn.Open()
            sql = New Data.SqlClient.SqlCommand("update descargas set sts=@sts where id_descarga=@id_descarga", conn)
            sql.Parameters.AddWithValue("@id_descarga", id_descarga)
            sql.Parameters.AddWithValue("@sts", err)
            sql.ExecuteNonQuery()
            conn.Close()
        Else
            setLog(id_descarga, "E", message)
        End If

        Dim xml As String = ""
        xml = xml & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
        xml = xml & "<root>"
        xml = xml & "<resultado><valor>" & err & "</valor></resultado>"
        xml = xml & "<archivos>"
        x = 0
        While x < TArchivos.Rows.Count
            xml = xml & "<archivo>"
            xml = xml & "<nombre>" & TArchivos.Rows(x).Item(0).ToString() & "</nombre>"
            xml = xml & "<ruta>" & TArchivos.Rows(x).Item(1).ToString() & "</ruta>"
            xml = xml & "</archivo>"
            x = x + 1
        End While
        xml = xml & "</archivos>"
        xml = xml & "<idDescarga><id>" & id_descarga & "</id></idDescarga>"
        xml = xml & "<mensajes>"
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("select * from detdescargas where id_descarga=@id_descarga order by id_log", conn)
        sql.Parameters.AddWithValue("@id_descarga", id_descarga)
        rs = sql.ExecuteReader
        While rs.Read
            xml = xml & "<msg>"
            xml = xml & "<id>" & rs!id_log & "</id>"
            xml = xml & "<tipo>" & rs!tipomsg & "</tipo>"
            xml = xml & "<descripcion>" & rs!message & "</descripcion>"
            xml = xml & "</msg>"
        End While
        conn.Close()
        xml = xml & "</mensajes>"
        xml = xml & "</root>"
        Return xml

    End Function

    Private Function getVentasSears(ByVal fecha As Date, ByVal row As DataRow) As Integer
        iniciar()
        setLog(id_descarga, "I", "Iniciando descarga de Ventas...")
        Dim res As Integer = 0
        res = validaSitio(row.Item(5))
        If res = 0 Then
            res = Browser(row.Item(5))
            If res = 0 Then
                res = setAttribute("input", "name", "txtUsuarioBURN", row(3))
                If res = 0 Then
                    res = setAttribute("input", "name", "txtContrasenaBURN", row(4))
                    If res = 0 Then
                        res = sendClick("input", "name", "btnEntrar")
                        If res = 0 Then
                            If buscaValor("p", "Contraseña Incorrecta") Then
                                'Error contraseña incorrecta
                                setLog(id_descarga, "E", "Contraseña incorrecta, imposible accesar al portal")
                                res = 2
                            Else
                                setLog(id_descarga, "I", "Intentando descarga de archivos...")
                                Dim proveedor As String = row.Item(3).ToString
                                Dim hdnOpcion As String = "1" '1=ventas,2=inventarios
                                Dim hdhnAnio As String = fecha.Year
                                Dim hdnMes As String = fecha.Month
                                Dim hdnFiltroAux As String = "2" '2:rango de fechas para ventas,1:reporte por renglones inventarios
                                Dim hdnFechaIni As String = fecha.ToString("MM-dd-yyyy")
                                Dim hdnFechaFin As String = fecha.ToString("MM-dd-yyyy")
                                Dim optGenerar As String = "1" '1=ventas,2=Inventario
                                Dim optArts As String = "1" 'renglones
                                Dim optFechaPeriodo As String = "2" 'rango de fechas
                                Dim lstAnio As String = fecha.Year
                                Dim lstMes As String = fecha.Month
                                Dim txtFechaIni As String = fecha.ToString("MM-dd-yyyy")
                                Dim txtFechaFin As String = fecha.ToString("MM-dd-yyyy")
                                httpResponse = WRequest("http://proveedores.sears.com.mx/lib/main_xml.asp", "POST", "lstTiendas=3&hdnCodEmp=1&hdnCodFam2=9999-9999&hdnProvId=" & proveedor & "&hdnOpcion=" & hdnOpcion & "&hdhnAnio=" & hdhnAnio & "&hdnMes=" & hdnMes & "&hdnFiltroAux=" & hdnFiltroAux & "&hdnFechaIni=" & hdnFechaIni & "&hdnFechaFin=" & hdnFechaFin & "&optGenerar=" & optGenerar & "&optArts=" & optArts & "&optFechaPeriodo=" & optFechaPeriodo & "&lstAnio=" & lstAnio & "&lstMes=" & lstMes & "&optFechaPeriodo=" & optFechaPeriodo & "&txtFechaIni=" & txtFechaIni & "&txtFechaFin=" & txtFechaFin & "&optTelcel=1&cmdGenerar=Generar", cookiesContainer)
                                If Not httpResponse Is Nothing Then

                                    Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                    Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                    saveTo = saveTo.Replace(".tmp", ".zip")
                                    Dim writeStream As IO.FileStream = New IO.FileStream(saveTo, IO.FileMode.Create, IO.FileAccess.Write)
                                    ReadWriteStream(o, writeStream)
                                    httpResponse.Close()

                                    Dim fzip As FastZip = New FastZip()
                                    Dim a() = saveTo.Split("\")
                                    Dim archivo As String = a(a.Length - 1)
                                    Dim ruta As String = saveTo.Replace("\" & archivo, "")
                                    If My.Computer.FileSystem.DirectoryExists(ruta & "\tmp_piagui_" & id_descarga) Then
                                        My.Computer.FileSystem.DeleteDirectory(ruta & "\tmp_piagui_" & id_descarga, FileIO.DeleteDirectoryOption.DeleteAllContents)
                                    End If
                                    My.Computer.FileSystem.CreateDirectory(ruta & "\tmp_piagui_" & id_descarga)
                                    fzip.ExtractZip(saveTo, ruta & "\tmp_piagui_" & id_descarga, "")
                                    buscaArchivos(ruta & "\tmp_piagui_" & id_descarga)
                                    If lista_archivos.Rows.Count > 0 Then
                                        Dim xml As New Xml.XmlDocument
                                        xml.Load(lista_archivos.Rows(0).Item(1).ToString)
                                        Dim nodelist As Xml.XmlNodeList = xml.SelectNodes("/root/VTAS")
                                        If nodelist.Count = 0 Then
                                            res = 1
                                            setLog(id_descarga, "E", "No se encontro informacion para la fecha seleccionada. Intentar mas tarde.")
                                        End If

                                        'If Not System.IO.Directory.Exists(row.Item(6).ToString()) Then
                                        '    My.Computer.FileSystem.CreateDirectory(row.Item(6).ToString())
                                        'End If
                                        'If My.Computer.FileSystem.FileExists(row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml") Then
                                        '    My.Computer.FileSystem.DeleteFile(row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                        'End If
                                        'My.Computer.FileSystem.CopyFile(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                        If res = 0 Then
                                            Try
                                                ftp.ServerAddress = server_ftp
                                                ftp.ServerPort = port_ftp
                                                ftp.UserName = user_ftp
                                                ftp.Password = pwd_ftp
                                                ftp.Connect()
                                                ftp.UploadFile(lista_archivos.Rows(0).Item(1).ToString, "SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml", False)
                                                ftp.Close()
                                                'System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                                TArchivos.Rows.Add("SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml", row.Item(6).ToString() & "\")
                                                setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml")
                                                setLog(id_descarga, "I", "Fin de descarga de Ventas...")
                                            Catch ex As Exception
                                                setLog(id_descarga, "E", "Error:" & ex.Message)
                                                res = 1
                                            End Try
                                        End If
                                    End If
                                Else
                                    setLog(id_descarga, "E", message)
                                    res = 1
                                End If

                            End If
                        End If
                    End If
                End If
            End If
        Else
            setLog(id_descarga, "E", "Error de Sitio. " & message)
            res = 1
        End If
        terminar()
        Return res
    End Function

    Private Function getInventarioSears(ByVal fecha As Date, ByVal row As Data.DataRow) As Integer
        iniciar()
        setLog(id_descarga, "I", "Iniciando descarga de Inventarios...")
        Dim res As Integer = 0
        res = validaSitio(row.Item(5))
        If res = 0 Then
            res = Browser(row.Item(5))
            If res = 0 Then
                res = setAttribute("input", "name", "txtUsuarioBURN", row(3))
                If res = 0 Then
                    res = setAttribute("input", "name", "txtContrasenaBURN", row(4))
                    If res = 0 Then
                        res = sendClick("input", "name", "btnEntrar")
                        If res = 0 Then
                            If buscaValor("p", "Contraseña Incorrecta") Then
                                'Error contraseña incorrecta
                                setLog(id_descarga, "E", "Contraseña incorrecta, imposible accesar al portal")
                                res = 2
                            Else
                                setLog(id_descarga, "I", "Intentando descarga de archivos...")
                                Dim proveedor As String = row.Item(3).ToString
                                Dim hdnOpcion As String = "2" '1=ventas,2=inventarios
                                Dim hdhnAnio As String = fecha.Year
                                Dim hdnMes As String = fecha.Month
                                Dim hdnFiltroAux As String = "1" '2:rango de fechas para ventas,1:reporte por renglones inventarios
                                Dim hdnFechaIni As String = fecha.ToString("MM-dd-yyyy")
                                Dim hdnFechaFin As String = fecha.ToString("MM-dd-yyyy")
                                Dim optGenerar As String = "2" '1=ventas,2=Inventario
                                Dim optArts As String = "1" 'renglones
                                Dim optFechaPeriodo As String = "2" 'rango de fechas
                                Dim lstAnio As String = fecha.Year
                                Dim lstMes As String = fecha.Month
                                Dim txtFechaIni As String = fecha.ToString("MM-dd-yyyy")
                                Dim txtFechaFin As String = fecha.ToString("MM-dd-yyyy")
                                httpResponse = WRequest("http://proveedores.sears.com.mx/lib/main_xml.asp", "POST", "lstTiendas=3&hdnCodEmp=1&hdnCodFam2=9999-9999&hdnProvId=" & proveedor & "&hdnOpcion=" & hdnOpcion & "&hdhnAnio=" & hdhnAnio & "&hdnMes=" & hdnMes & "&hdnFiltroAux=" & hdnFiltroAux & "&hdnFechaIni=" & hdnFechaIni & "&hdnFechaFin=" & hdnFechaFin & "&optGenerar=" & optGenerar & "&optArts=" & optArts & "&optFechaPeriodo=" & optFechaPeriodo & "&lstAnio=" & lstAnio & "&lstMes=" & lstMes & "&optFechaPeriodo=" & optFechaPeriodo & "&txtFechaIni=" & txtFechaIni & "&txtFechaFin=" & txtFechaFin & "&optTelcel=1&cmdGenerar=Generar", cookiesContainer)
                                If Not httpResponse Is Nothing Then

                                    Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                    Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                    saveTo = saveTo.Replace(".tmp", ".zip")
                                    Dim writeStream As IO.FileStream = New IO.FileStream(saveTo, IO.FileMode.Create, IO.FileAccess.Write)
                                    ReadWriteStream(o, writeStream)
                                    httpResponse.Close()

                                    Dim fzip As FastZip = New FastZip()
                                    Dim a() = saveTo.Split("\")
                                    Dim archivo As String = a(a.Length - 1)
                                    Dim ruta As String = saveTo.Replace(archivo, "")
                                    If My.Computer.FileSystem.DirectoryExists(ruta & "\tmp_piagui_" & id_descarga) Then
                                        My.Computer.FileSystem.DeleteDirectory(ruta & "\tmp_piagui_" & id_descarga, FileIO.DeleteDirectoryOption.DeleteAllContents)
                                    End If
                                    My.Computer.FileSystem.CreateDirectory(ruta & "\tmp_piagui_" & id_descarga)
                                    fzip.ExtractZip(saveTo, ruta & "\tmp_piagui_" & id_descarga, "")
                                    buscaArchivos(ruta & "\tmp_piagui_" & id_descarga)
                                    If lista_archivos.Rows.Count > 0 Then

                                        Dim xml As New Xml.XmlDocument
                                        xml.Load(lista_archivos.Rows(0).Item(1).ToString)
                                        Dim nodelist As Xml.XmlNodeList = xml.SelectNodes("/root/row")
                                        If nodelist.Count = 0 Then
                                            res = 1
                                            setLog(id_descarga, "E", "No se encontro informacion para la fecha seleccionada. Intentar mas tarde.")
                                        End If
                                        'If Not System.IO.Directory.Exists(row.Item(6).ToString()) Then
                                        '    System.IO.Directory.CreateDirectory(row.Item(6).ToString())
                                        'End If
                                        'If System.IO.File.Exists(row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_INV_" & fecha.ToString("ddMMyyyy") & ".xml") Then
                                        '    System.IO.File.Delete(row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_INV_" & fecha.ToString("ddMMyyyy") & ".xml")
                                        'End If
                                        'System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_INV_" & fecha.ToString("ddMMyyyy") & ".xml")
                                        If res = 0 Then

                                            Try
                                                ftp.ServerAddress = server_ftp
                                                ftp.ServerPort = port_ftp
                                                ftp.UserName = user_ftp
                                                ftp.Password = pwd_ftp
                                                ftp.Connect()
                                                ftp.UploadFile(lista_archivos.Rows(0).Item(1).ToString, "SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml", False)
                                                ftp.Close()
                                                'System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                                TArchivos.Rows.Add("SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml", row.Item(6).ToString() & "\")
                                                setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml")
                                                setLog(id_descarga, "I", "Fin de descarga de Inventarios...")
                                            Catch ex As Exception
                                                setLog(id_descarga, "E", "Error:" & ex.Message)
                                                res = 1
                                            End Try

                                        End If
                                        'TArchivos.Rows.Add("SEARS_" & row(3) & "_INV_" & fecha.ToString("ddMMyyyy") & ".xml", row.Item(6).ToString() & "\")
                                        'setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_INV_" & fecha.ToString("ddMMyyyy") & ".xml")
                                        'setLog(id_descarga, "I", "Fin de descarga de Inventarios...")
                                    End If
                                Else
                                    setLog(id_descarga, "E", message)
                                    res = 1
                                End If

                                End If
                        End If
                    End If
                End If
            End If
        Else
            setLog(id_descarga, "E", "Error de Sitio. " & message)
            res = 1
        End If
        terminar()
        Return res

    End Function


    Private Sub newDescarga(ByVal cadena As String, ByVal fecha As Date)
        Dim conn1 As New Data.SqlClient.SqlConnection(conn.ConnectionString)
        conn1.Open()
        Dim sql1 As SqlClient.SqlCommand
        sql1 = New Data.SqlClient.SqlCommand("prc_newdescarga", conn1)
        sql1.CommandType = CommandType.StoredProcedure
        sql1.CommandTimeout = 0
        sql1.Parameters.AddWithValue("@cadena", cadena)
        sql1.Parameters.AddWithValue("@fecha", fecha)
        sql1.Parameters.AddWithValue("@id_descarga", 0)
        sql1.Parameters.Item(2).Direction = ParameterDirection.Output
        sql1.ExecuteNonQuery()
        id_descarga = sql1.Parameters.Item(2).Value
        conn1.Close()
    End Sub

    Private Sub setLog(ByVal id_descarga As Integer, ByVal tipo As String, ByVal msg As String)
        Dim conn1 As New Data.SqlClient.SqlConnection(conn.ConnectionString)
        conn1.Open()
        Dim sql1 As SqlClient.SqlCommand
        sql1 = New Data.SqlClient.SqlCommand("prc_setLog", conn1)
        sql1.CommandType = CommandType.StoredProcedure
        sql1.CommandTimeout = 0
        sql1.Parameters.AddWithValue("@id_descarga", id_descarga)
        sql1.Parameters.AddWithValue("@tipo", tipo)
        sql1.Parameters.AddWithValue("@message", msg)
        sql1.ExecuteNonQuery()
        conn1.Close()
    End Sub

    Function WRequest(ByVal URL As String, ByVal method As String, ByVal POSTdata As String, ByVal cookies As Net.CookieContainer) As Net.HttpWebResponse
        Dim responseData As String = ""
        Try
            Dim hwrequest As Net.HttpWebRequest = Net.WebRequest.Create(URL)
            Dim response As Net.HttpWebResponse
            hwrequest.Accept = "*/*"
            hwrequest.AllowAutoRedirect = True
            hwrequest.KeepAlive = False
            hwrequest.ConnectionGroupName = Guid.NewGuid.ToString
            hwrequest.UserAgent = "PiaguiResquest/1.0"
            hwrequest.Method = method
            hwrequest.CookieContainer = cookies
            hwrequest.Timeout = 180000
            If hwrequest.Method = "POST" Then
                hwrequest.ContentType = "application/x-www-form-urlencoded"
                Dim encoding As New Text.ASCIIEncoding()
                Dim postByteArray() As Byte = encoding.GetBytes(POSTdata)
                hwrequest.ContentLength = postByteArray.Length
                Dim postStream As IO.Stream = hwrequest.GetRequestStream()
                postStream.Write(postByteArray, 0, postByteArray.Length)
                postStream.Close()
            End If
            response = hwrequest.GetResponse()
            Return response
        Catch e As Exception
            message = "Error: " & e.Message
        End Try
        Return Nothing
    End Function

    Public Sub ReadWriteStream(ByVal readStream As IO.Stream, ByVal writeStream As IO.Stream)
        Dim Length As Integer = 1024
        Dim buffer As [Byte]() = New [Byte](Length - 1) {}
        Dim bytesRead As Integer = readStream.Read(buffer, 0, Length)
        ' write the required bytes
        While bytesRead > 0
            writeStream.Write(buffer, 0, bytesRead)
            bytesRead = readStream.Read(buffer, 0, Length)
        End While
        readStream.Close()
        writeStream.Close()
    End Sub

    Public Sub SearchDirectory(ByVal carpeta As String)
        Try
            Dim fso() As String
            Dim carpetax As String = ""
            Try
                fso = System.IO.Directory.GetDirectories(carpeta)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Dim i As Integer = 0
            While i < fso.Length
                carpetax = fso(i)
                'carpeta = fso(i)
                ReDim Preserve strdirectorios(UBound(strdirectorios) + 1)
                strdirectorios(UBound(strdirectorios)) = carpetax
                SearchDirectory(carpetax)
                i = i + 1
            End While

            Dim archivos(), nomarchivo() As String
            archivos = System.IO.Directory.GetFiles(carpeta)
            Dim x As Integer = 0
            While x < archivos.Length
                nomarchivo = archivos(x).Split("\")
                If nomarchivo.Length > 0 Then
                    Dim a As String = nomarchivo(nomarchivo.Length - 1)
                    nomarchivo = a.Split(".")
                    a = nomarchivo(0)
                    lista_archivos.Rows.Add(a.ToUpper, archivos(x))
                End If
                x = x + 1
            End While

        Catch ex As Exception
            setLog(id_descarga, "E", "Error: " & ex.Message)
        End Try
    End Sub

    Public Function Liverpool(ByVal fecha As Date, ByVal Archivo As String) As String
        Dim x As Integer = 0
        Dim c As Integer = 0
        Dim err As Integer = 0
        newDescarga("LIVERPOOL", fecha)

        err = getFTP("LIVERPOOL")
        If err = 0 Then

            conn.Open()
            sql = New Data.SqlClient.SqlCommand("select * from clogin where clogportal='LIVERPOOL' ", conn)
            Dim tabla As New System.Data.DataTable
            Dim conf As New System.Data.DataTable
            tabla.Load(sql.ExecuteReader)
            conn.Close()

            conn.Open()
            sql = New Data.SqlClient.SqlCommand("select * from conf_seccion_liverpool", conn)
            conf.Load(sql.ExecuteReader)
            conn.Close()

            setLog(id_descarga, "I", "Iniciando proceso...")
            setLog(id_descarga, "I", "Tipo de descarga solicitada: " & Archivo)

            While x < tabla.Rows.Count And err = 0
                setLog(id_descarga, "I", "Procesando cliente: " & tabla.Rows(x).Item(3).ToString)

                While c < conf.Rows.Count And err = 0

                    If Archivo = "V" And err = 0 And conf.Rows(c).Item(1).ToString = "VTA" Then 'solo ventas
                        err = getVentasLiverpool(fecha, tabla.Rows(x), conf.Rows(c))
                    End If
                    If Archivo = "I" And err = 0 And conf.Rows(c).Item(1).ToString = "INV" Then 'solo inventarios
                        err = getInventarioLiverpool(fecha, tabla.Rows(x), conf.Rows(c))
                    End If
                    If Archivo = "T" And err = 0 Then 'Archivo venas e inventarios
                        If conf.Rows(c).Item(1).ToString = "VTA" Then
                            err = getVentasLiverpool(fecha, tabla.Rows(x), conf.Rows(c))
                        End If
                        If err = 0 Then
                            If conf.Rows(c).Item(1).ToString = "INV" Then
                                err = getInventarioLiverpool(fecha, tabla.Rows(x), conf.Rows(c))
                            End If
                        End If
                    End If

                    System.Threading.Thread.Sleep(2000)
                    c = c + 1
                End While

                x = x + 1
            End While

            If err = 0 Then
                setLog(id_descarga, "I", "Proceso terminado correctamente...")
            Else
                setLog(id_descarga, "E", "Proceso terminado con errores, favor de revisar secuencia...")
                setLog(id_descarga, "I", "Eliminado archivos generados...")
                Dim ax As Integer = 0
                While ax < TArchivos.Rows.Count
                    Try
                        ftp.ServerAddress = server_ftp
                        ftp.ServerPort = port_ftp
                        ftp.UserName = user_ftp
                        ftp.Password = pwd_ftp
                        ftp.Connect()
                        ftp.DeleteFile(TArchivos.Rows(ax).Item(0).ToString())
                        ftp.Close()
                        setLog(id_descarga, "I", "Archivo eliminado: " & TArchivos.Rows(ax).Item(0).ToString())
                    Catch ex As Exception
                        setLog(id_descarga, "E", "Error:" & ex.Message)
                    End Try
                    ax = ax + 1
                End While
            End If
        Else
            setLog(id_descarga, "E", message)
        End If

        conn.Open()
        sql = New Data.SqlClient.SqlCommand("update descargas set sts=@sts where id_descarga=@id_descarga", conn)
        sql.Parameters.AddWithValue("@id_descarga", id_descarga)
        sql.Parameters.AddWithValue("@sts", err)
        sql.ExecuteNonQuery()
        conn.Close()

        Dim xml As String = ""
        xml = xml & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
        xml = xml & "<root>"
        xml = xml & "<resultado><valor>" & err & "</valor></resultado>"
        xml = xml & "<archivos>"
        x = 0
        While x < TArchivos.Rows.Count
            xml = xml & "<archivo>"
            xml = xml & "<nombre>" & TArchivos.Rows(x).Item(0).ToString() & "</nombre>"
            xml = xml & "<ruta>" & TArchivos.Rows(x).Item(1).ToString() & "</ruta>"
            xml = xml & "</archivo>"
            x = x + 1
        End While
        xml = xml & "</archivos>"
        xml = xml & "<idDescarga><id>" & id_descarga & "</id></idDescarga>"
        xml = xml & "<mensajes>"
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("select * from detdescargas where id_descarga=@id_descarga order by id_log", conn)
        sql.Parameters.AddWithValue("@id_descarga", id_descarga)
        rs = sql.ExecuteReader
        While rs.Read
            xml = xml & "<msg>"
            xml = xml & "<id>" & rs!id_log & "</id>"
            xml = xml & "<tipo>" & rs!tipomsg & "</tipo>"
            xml = xml & "<descripcion>" & rs!message & "</descripcion>"
            xml = xml & "</msg>"
        End While
        conn.Close()
        xml = xml & "</mensajes>"
        xml = xml & "</root>"
        Return xml
    End Function

    Public Function setSelect(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal valor As String, ByVal index As Integer, ByVal evento As String) As Integer
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
        Dim iHTMLEle As mshtml.IHTMLElement
        Dim str As String = ""
        For Each iHTMLEle In iHTMLCol
            If Not iHTMLEle.getAttribute(atributo) Is Nothing Then
                str = iHTMLEle.getAttribute(atributo).ToString
                If str.ToUpper.Equals(nombre.ToUpper) Then
                    iHTMLEle.setAttribute("value", valor)
                    If evento <> "" Then
                        Dim selectElem As mshtml.IHTMLSelectElement = DirectCast(iHTMLEle, mshtml.IHTMLSelectElement)
                        If Not selectElem Is Nothing Then
                            Dim tempElem As mshtml.HTMLSelectElement
                            tempElem = DirectCast(selectElem, mshtml.HTMLSelectElement)
                            tempElem.selectedIndex = index
                            Dim dummy As Object = Nothing
                            tempElem.FireEvent("onchange", dummy)
                        End If
                    End If

                    setLog(id_descarga, "I", "Asignacion de valor [ " & nombre & " : " & valor & "]")
                    Return 0
                    Exit For
                End If
            End If
        Next
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre)
        Return 1
    End Function

    Private Function getVentasLiverpool(ByVal fecha As Date, ByVal row As Data.DataRow, ByVal conf As Data.DataRow) As Integer
        iniciar()
        setLog(id_descarga, "I", "Iniciando descarga de Ventas...")
        Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true")
        Dim res As Integer = 0
        res = validaSitio(row.Item(5))
        If res = 0 Then
            res = Browser(row.Item(5))
            If res = 0 Then
                res = setAttribute("input", "name", "j_user", row(3).ToString)
                If res = 0 Then
                    res = setAttribute("input", "name", "j_password", row(4).ToString)
                    If res = 0 Then
                        res = sendClick("input", "name", "uidPasswordLogon")
                        If res = 0 Then
                            If buscaValor("span", "Autentificación de usuario fallida") Then
                                'Error contraseña incorrecta
                                setLog(id_descarga, "E", "Contraseña incorrecta, imposible accesar al portal")
                                res = 2
                            Else
                                res = Browser("https://bwsext.liverpool.com.mx/sap/bw/BEx?sap-language=es&sap-client=400&accessibility=&style_sheet=http%3A%2F%2Fproveedores.liverpool.com.mx%3A80%2Firj%2Fportalapps%2Fcom.sap.portal.design.portaldesigndata%2Fthemes%2Fportal%2Fcustomer%2FProveedores%2FBIReports30%2FBIReports30_nn7.css%3F6.0.16.0.1&TEMPLATE_ID=BWR_VTAS_POR_DIA_PROV")
                                'seleccionando las secciones 
                                If conf(2).ToString() <> "*" Then
                                    res = setSelect("select", "name", "VAR_OPERATOR_3", "BT", 1, "onchange")
                                    If res = 0 Then
                                        res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_3", conf(2).ToString)
                                        If res = 0 Then
                                            res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_3", conf(3).ToString)
                                        End If
                                    End If
                                End If
                                If res = 0 Then
                                    If conf(4).ToString <> "*" Then
                                        res = setSelect("select", "name", "VAR_OPERATOR_14", "BT", 1, "onchange")
                                        If res = 0 Then
                                            res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_14", conf(4).ToString)
                                            If res = 0 Then
                                                res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_14", conf(5).ToString)
                                            End If
                                        End If
                                    End If
                                    If res = 0 Then
                                        res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_15", fecha.ToString("dd.MM.yyyy"))
                                        Dim repeticiones As Integer = 0
                                        While repeticiones < 5 And res = 1
                                            System.Threading.Thread.Sleep("1000")
                                            res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_15", fecha.ToString("dd.MM.yyyy"))
                                        End While
                                        If res = 0 Then
                                            res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_15", fecha.ToString("dd.MM.yyyy"))
                                            If res = 0 Then
                                                res = sendClick("a", "href", "javascript:SAPBWBUTTON('PROCESS_VARIABLES','VAR_SUBMIT',' ')")
                                                If res = 0 Then
                                                    res = sendClick("a", "href", "JavaScript:SAPBW(1,'','','EXPAND','0MATERIAL','Y')")
                                                    If res = 0 Then
                                                        Dim url As String = wb.LocationURL.Replace("CMD=", "$")
                                                        Dim url1() As String = url.Split("$")
                                                        Dim url2() As String = url1(0).Split("?")

                                                        httpResponse = WRequest(url2(0), "POST", url2(1) & "CMD=EXPORT&DATA_PROVIDER=DATAPROVIDER_4&FORMAT=CSV&SEPARATOR=,&id=" & Guid.NewGuid.ToString, cookiesContainer)
                                                        If Not httpResponse Is Nothing Then
                                                            'web.Headers.Add("Server:es")
                                                            'web.Headers.Add("Server:" & Now.ToString("ddd", New System.Globalization.CultureInfo("en-US")) & ", " & Now.ToString("dd MMM yyyy hh:mm:ss", New System.Globalization.CultureInfo("en-US")) & " GMT")
                                                            'web.Headers.Add("Server:MYSAPSSO2=AjExMDAgABBwb3J0YWw6UDAwMDA3MzY5iAATYmFzaWNhdXRoZW50aWNhdGlvbgEACVAwMDAwNzM2OQIAAzAwMAMAA0VQUAQADDIwMTIwOTE4MDA0MQUABAAAAAgKAAlQMDAwMDczNjn%2FAW4wggFqBgkqhkiG9w0BBwKgggFbMIIBVwIBATELMAkGBSsOAwIaBQAwCwYJKoZIhvcNAQcBMYIBNjCCATICAQEwgYcwfjELMAkGA1UEBhMCTVgxDzANBgNVBAgTBk1leGljbzENMAsGA1UEBxMERC5GLjEtMCsGA1UEChMkRGlzdHJpYnVpZG9yYSBMaXZlcnBvb2wgUy5BLiBkZSBDLlYuMRIwEAYDVQQLEwlMaXZlcnBvb2wxDDAKBgNVBAMTA0VQUAIFAL9DfMkwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEyMDkxODAwNDE1N1owIwYJKoZIhvcNAQkEMRYEFE!HDMvoS3S3Yxc7mdUlWzVW0IEFMAkGByqGSM44BAMELjAsAhQchRgbu3!vjdC5STlbofIJfq%2FBNwIUbS3uzFtN%2Fk7tZoSMxjQhzznjt4c%3D;path=/;domain=liverpool.com.mx;HttpOnly,PortalAlias=portal; Path=/,saplb_*=(J2EE19460300)19460350; Version=1; Path=/," & c(1) & "; Version=1; Domain=.liverpool.com.mx; Path=/")
                                                            'web.Headers.Add("Server:SAP J2EE Engine/6.40")
                                                            'web.Headers.Add("Server:text/html; charset=UTF-8")
                                                            'web.Headers.Add("Server:1630")
                                                            'web.Headers.Add("Server:no-store,no-cache")
                                                            ' web.Headers.Add("Server:1.1 proveedores.liverpool.com.mx (Access Gateway 3.1.1-215)")
                                                            'MsgBox("Server:" & Now.ToString("ddd", New System.Globalization.CultureInfo("en-US")) & ", " & Now.ToString("dd MMM yyyy hh:mm:ss", New System.Globalization.CultureInfo("en-US")) & " GMT")
                                                            'web.CookieContainer = cookiesContainer

                                                            Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                                            Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                                            saveTo = saveTo.Replace(".tmp", ".csv")

                                                            Dim lineas As System.IO.StreamReader = New System.IO.StreamReader(o)
                                                            Dim w As IO.StreamWriter = New System.IO.StreamWriter(saveTo)
                                                            Dim ren As Integer = 0
                                                            While Not lineas.EndOfStream
                                                                Dim l As String = lineas.ReadLine
                                                                If Not l.Contains("Resultado") Then
                                                                    l = l.Replace("""", "|")
                                                                    l = l.Replace("|,|", "|")
                                                                    Dim cad() As String = l.Split("|")
                                                                    l = getCadena(cad)
                                                                    w.WriteLine(l)
                                                                    If ren = 0 Then
                                                                        If cad(1) <> "Centro" Then
                                                                            res = 1
                                                                            setLog(id_descarga, "E", "Archivo con errores. Formato desconocido.")
                                                                        End If
                                                                    End If
                                                                    ren = ren + 1
                                                                End If
                                                            End While
                                                            w.Flush()
                                                            w.Close()
                                                            If ren > 5 And res = 0 Then
                                                                Try
                                                                    ftp.ServerAddress = server_ftp
                                                                    ftp.ServerPort = port_ftp
                                                                    ftp.UserName = user_ftp
                                                                    ftp.Password = pwd_ftp
                                                                    ftp.Connect()
                                                                    Dim seccion As String = IIf(conf(2).ToString <> "*", "_" & conf(2).ToString, "")
                                                                    ftp.UploadFile(saveTo, "LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv", False)
                                                                    ftp.Close()
                                                                    'System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                                                    TArchivos.Rows.Add("LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv", row.Item(6).ToString() & "\")
                                                                    setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv")
                                                                    setLog(id_descarga, "I", "Fin de descarga de Ventas...")
                                                                Catch ex As Exception
                                                                    setLog(id_descarga, "E", "Error:" & ex.Message)
                                                                    res = 1
                                                                End Try
                                                            Else
                                                                res = 1
                                                                setLog(id_descarga, "E", "Informacion no encontrada. Intentar mas tarde.")
                                                            End If
                                                        End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                        End If
                    End If
                End If
            Else
                setLog(id_descarga, "E", "Error de Sitio. " & message)
                res = 1
            End If
        End If
        'WRequest("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent", "POST", "logout_submit=true", cookiesContainer)
        Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true")
        terminar()
        Return res
    End Function

    Private Function getInventarioLiverpool(ByVal fecha As Date, ByVal row As Data.DataRow, ByVal conf As Data.DataRow) As Integer
        iniciar()
        setLog(id_descarga, "I", "Iniciando descarga de Inventarios...")
        Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true")
        Dim res As Integer = 0
        res = validaSitio(row.Item(5))
        If res = 0 Then
            res = Browser(row.Item(5))
            If res = 0 Then
                res = setAttribute("input", "name", "j_user", row(3).ToString)
                If res = 0 Then
                    res = setAttribute("input", "name", "j_password", row(4).ToString)
                    If res = 0 Then
                        res = sendClick("input", "name", "uidPasswordLogon")
                        If res = 0 Then
                            If buscaValor("span", "Autentificación de usuario fallida") Then
                                'Error contraseña incorrecta
                                setLog(id_descarga, "E", "Contraseña incorrecta, imposible accesar al portal")
                                res = 2
                            Else
                                res = Browser("https://bwsext.liverpool.com.mx/sap/bw/BEx?sap-language=es&sap-client=400&accessibility=&style_sheet=http%3A%2F%2Fproveedores.liverpool.com.mx%3A80%2Firj%2Fportalapps%2Fcom.sap.portal.design.portaldesigndata%2Fthemes%2Fportal%2Fcustomer%2FProveedores%2FBIReports30%2FBIReports30_ie6.css%3F6.0.16.0.1&TEMPLATE_ID=BWR_VTAS_MENS_PROV_UNI_SHIST")
                                If res = 0 Then
                                    If conf(2).ToString <> "*" Then
                                        If conf(3).ToString = "" Then
                                            res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_3", conf(2).ToString)
                                        Else
                                            res = setSelect("select", "name", "VAR_OPERATOR_3", "BT", 1, "onchange")
                                            If res = 0 Then
                                                res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_3", conf(2).ToString)
                                                If res = 0 Then
                                                    res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_3", conf(3).ToString)
                                                End If
                                            End If
                                        End If
                                    End If
                                    If res = 0 Then
                                        If conf(4).ToString <> "*" Then
                                            If conf(5).ToString = "" Then
                                                res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_14", conf(4).ToString)
                                            Else
                                                res = setSelect("select", "name", "VAR_OPERATOR_14", "BT", 1, "onchange")
                                                If res = 0 Then
                                                    res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_14", conf(4).ToString)
                                                    If res = 0 Then
                                                        res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_14", conf(5).ToString)
                                                    End If
                                                End If
                                            End If
                                        End If

                                        If res = 0 Then
                                            res = sendClick("a", "href", "Javascript:SAPBWBUTTON('PROCESS_VARIABLES','VAR_SUBMIT',' ')")
                                            If res = 0 Then
                                                res = sendClick("a", "href", "JavaScript:SAPBW(1,'','','EXPAND','0MATERIAL','Y')")
                                                If res = 0 Then
                                                    Dim url As String = wb.LocationURL.Replace("CMD=", "$")
                                                    Dim url1() As String = url.Split("$")
                                                    Dim url2() As String = url1(0).Split("?")

                                                    httpResponse = WRequest(url2(0), "POST", url2(1) & "CMD=EXPORT&DATA_PROVIDER=DATAPROVIDER_4&FORMAT=CSV&SEPARATOR=,&NAME=PERRITO.CSV&id=" & Guid.NewGuid.ToString, cookiesContainer)

                                                    If Not httpResponse Is Nothing Then

                                                        Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                                        Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                                        saveTo = saveTo.Replace(".tmp", ".csv")

                                                        Dim lineas As System.IO.StreamReader = New System.IO.StreamReader(o)
                                                        Dim w As IO.StreamWriter = New System.IO.StreamWriter(saveTo)
                                                        Dim ren As Integer = 0
                                                        While Not lineas.EndOfStream
                                                            Dim l As String = lineas.ReadLine
                                                            If Not l.Contains("Resultado") Then

                                                                l = l.Replace("""", "|")
                                                                l = l.Replace("|,|", "|")
                                                                Dim cad() As String = l.Split("|")
                                                                l = getCadena(cad)
                                                                w.WriteLine(l)
                                                                If ren = 0 Then
                                                                    If cad(1) <> "Centro" Then
                                                                        res = 1
                                                                        setLog(id_descarga, "E", "Archivo con errores. Formato desconocido.")
                                                                    End If
                                                                End If
                                                                ren = ren + 1
                                                            End If
                                                        End While
                                                        w.Flush()
                                                        w.Close()

                                                        If ren > 5 And res = 0 Then
                                                            Try
                                                                ftp.ServerAddress = server_ftp
                                                                ftp.ServerPort = port_ftp
                                                                ftp.UserName = user_ftp
                                                                ftp.Password = pwd_ftp
                                                                ftp.Connect()
                                                                Dim seccion As String = IIf(conf(2).ToString <> "*", "_" & conf(2).ToString, "")
                                                                ftp.UploadFile(saveTo, "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv", False)
                                                                ftp.Close()
                                                                'System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                                                TArchivos.Rows.Add("LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv", row.Item(6).ToString() & "\")
                                                                setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv")
                                                                setLog(id_descarga, "I", "Fin de descarga de Inventarios...")
                                                            Catch ex As Exception
                                                                setLog(id_descarga, "E", "Error:" & ex.Message)
                                                                res = 1
                                                            End Try
                                                        Else
                                                            res = 1
                                                            setLog(id_descarga, "E", "Informacion no disponible. Intentar mas tarde.")
                                                        End If
                                                    End If
                                                End If
                                                End If
                                            End If

                                        End If
                                    End If
                                End If
                        End If
                    End If
                End If
            Else
                setLog(id_descarga, "E", "Error de Sitio. " & message)
                res = 1
            End If
        End If
        'WRequest("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent", "POST", "logout_submit=true", cookiesContainer)
        Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true")
        terminar()
        Return res
    End Function



    Private Function getCadena(ByVal cad() As String) As String
        Dim res As String = ""
        Dim x As Integer = 1
        While x < cad.Length - 1
            If x = 10 Then
                cad(x) = cad(x).Replace(",", "")
            Else
                cad(x) = cad(x).Replace(",", " ")
            End If
            If res = "" Then
                res = res & cad(x)
            Else
                res = res & "," & cad(x)
            End If
            x = x + 1
        End While
        Return res
    End Function

 



    Private Function getArchivosPH(ByVal fecha As Date, ByVal pstrTipoDocumento As String) As Integer
        ' iniciar()
        Dim dtProceso As New System.Data.DataTable
        Dim dtLogin As New System.Data.DataTable
        Dim sqlrdr As SqlDataReader
        Dim oExcel As Application
        Dim oBook As Workbook
        Dim strSQL As String = "SELECT * FROM cLogin where cLogPortal='PH'"
        Dim sqlCmdLogin As New SqlCommand(strSQL, conn)
        Dim iHTMLCol As mshtml.IHTMLElementCollection
        Dim iHTMLEle As mshtml.IHTMLElement
        Dim blnProcesoCorrecto = False
        Dim intResultado As Integer
        Dim strProveedor As String = String.Empty
        Dim sqlCmd As New SqlCommand()

        conn.Open()
        sqlrdr = sqlCmdLogin.ExecuteReader
        dtLogin.Load(sqlrdr)
        sqlrdr.Close()
        conn.Close()

        Dim strMensaje As String = ""
        Try
               For Each drLog As DataRow In dtLogin.Rows
                strSQL = "SELECT cProSecuencia, cProTipo, cProAtributo, cProNombreAtributo, cProValorAtributo, cProAccion FROM cProcesos where cProHabilitar=1 AND cProSociedad LIKE '%" & drLog.Item("cLogEmpresa") & "%' and cProProceso='" & pstrTipoDocumento & "' order by cProEmpresa, cProSociedad, cProProceso, cProSecuencia"
                'strSQL = "SELECT cProSecuencia, cProTipo, cProAtributo, cProNombreAtributo, cProValorAtributo, cProAccion FROM cProcesos where cProHabilitar=1 AND cProSociedad LIKE '%PIAGUI%' and cProProceso='" & pstrTipoDocumento & "' order by cProEmpresa, cProSociedad, cProProceso, cProSecuencia"
                sqlCmd.CommandText = strSQL
                sqlCmd.Connection = conn

                conn.Open()
                sqlrdr = sqlCmd.ExecuteReader
                dtProceso.Rows.Clear()
                dtProceso.Load(sqlrdr)
                sqlrdr.Close()
                conn.Close()
                intResultado = validaSitio(drLog.Item("cLogUrl"))


              
                If intResultado = 1 Then
                    setLog(id_descarga, "E", "Error: Validar Pagina " & drLog.Item("cLogUrl"))
                    intResultado = 1
                    Return intResultado
                    Exit Function
                End If
                setLog(id_descarga, "I", "Iniciando proceso para " & drLog.Item("cLogEmpresa") & "...")
                setLog(id_descarga, "I", "Tipo de descarga solicitada: " & pstrTipoDocumento)

                If Me.Browser(drLog.Item("cLogOutUrl")) = 1 Then
                    setLog(id_descarga, "E", "Error: Validar Pagina " & drLog.Item("cLogOutUrl"))
                    intResultado = 0
                    Return intResultado
                    Exit Function
                End If
                If Me.Browser(drLog.Item("cLogUrl")) = 1 Then
                    setLog(id_descarga, "E", "Error: Cargar Pagina " & drLog.Item("cLogUrl"))
                    intResultado = 1
                    Return intResultado
                    Exit Function
                End If
                strProveedor = drLog.Item("cLogUsuario")
                If Me.setAttribute("input", "name", "duser", drLog.Item("cLogUsuario")) = 1 Then
                    intResultado = 1
                    Return intResultado
                    Exit Function
                End If
                If Me.setAttribute("input", "name", "dpass", drLog.Item("cLogContrasenia")) = 1 Then
                    intResultado = 1
                    Return intResultado
                    Exit Function
                End If

                For drFila As Integer = 0 To dtProceso.Rows.Count - 1
                    strMensaje = dtProceso.Rows(drFila).Item("cProValorAtributo")
                    Select Case dtProceso.Rows(drFila).Item("cProAccion")
                        Case "Navegar"
                            intResultado = Browser(dtProceso.Rows(drFila).Item("cProValorAtributo"))
                        Case "Teclear"
                            If dtProceso.Rows(drFila).Item("cProNombreAtributo") = "lsS1) Fecha (mm/dd/aaaa)" Then
                                intResultado = Me.setAttribute(dtProceso.Rows(drFila).Item("cProTipo"), dtProceso.Rows(drFila).Item("cProAtributo"), dtProceso.Rows(drFila).Item("cProNombreAtributo"), fecha)
                            Else
                                intResultado = Me.setAttribute(dtProceso.Rows(drFila).Item("cProTipo"), dtProceso.Rows(drFila).Item("cProAtributo"), dtProceso.Rows(drFila).Item("cProNombreAtributo"), dtProceso.Rows(drFila).Item("cProValorAtributo"))
                            End If
                        Case "Click"
                            intResultado = Me.sendClick(dtProceso.Rows(drFila).Item("cProTipo"), dtProceso.Rows(drFila).Item("cProAtributo"), dtProceso.Rows(drFila).Item("cProNombreAtributo"))
                        Case "Frame"
                            Dim frm As mshtml.IHTMLWindow2
                            Dim ColFrms As mshtml.FramesCollection = CType(ie.Document, mshtml.HTMLDocument).frames
                            For i As Integer = 0 To ColFrms.length - 1
                                frm = ColFrms.item(i)
                                If frm.name.ToString.Trim = dtProceso.Rows(drFila).Item("cProNombreAtributo").ToString.Trim Then
                                    If dtProceso.Rows(drFila + 1).Item("cProTipo") = "Frame" Then
                                        ColFrms = frm.frames
                                        drFila = drFila + 1
                                        i = -1
                                    Else
                                        document = frm.document
                                        Exit For
                                    End If
                                End If
                            Next
                        Case "Link"
                            intResultado = Me.sendLink(dtProceso.Rows(drFila).Item("cProValorAtributo").ToString.Trim)
                        Case "Descarga"
                            httpResponse = Me.WRequest(dtProceso.Rows(drFila).Item("cProValorAtributo"), "get", "", cookiesContainer)
                            'httpResponse = Me.WRequest("http://www.phb2b.com.mx/wi/scripts/saveAsXls.asp", "post", "cmdBlock=all&cmd=asksave&cmdP1=%286.%202%29%20%20Inventarios%20Detalle%20Tienda%20-%20Hogar%20y%20L.%20Generales*1030*0*rep*wi00000001", cookiesContainer)
                            Dim o As System.IO.Stream = httpResponse.GetResponseStream
                            Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                            saveTo = saveTo.Replace(".tmp", ".xls")
                            Dim writeStream As IO.FileStream = New IO.FileStream(saveTo, IO.FileMode.Create, IO.FileAccess.Write)
                            Me.ReadWriteStream(o, writeStream)
                            oExcel = CreateObject("Excel.Application")
                            oBook = oExcel.Workbooks.Open(saveTo, , False)
                            If pstrTipoDocumento = "Ventas" Then
                                For Each sheet As Excel.Worksheet In oBook.Sheets
                                    Select Case sheet.Name
                                        Case "Depto_Sku_Tienda"
                                            sheet.Columns(14).NumberFormat = "#.##0,00"
                                        Case "Depto_Sku-Talla-Color_Tienda"
                                            sheet.Columns(18).NumberFormat = "#.##0,00"
                                        Case "Depto_Marca_Tda"
                                            sheet.Columns(7).NumberFormat = "#.##0,00"
                                        Case "Depto_Tda_Marca"
                                            sheet.Columns(8).NumberFormat = "#.##0,00"
                                    End Select
                                Next
                            End If
                            oBook.Save()
                            oBook.Close()
                            oExcel.Quit()
                            oBook = Nothing
                            oExcel = Nothing
                            ftp.ServerAddress = server_ftp
                            ftp.ServerPort = port_ftp
                            ftp.UserName = user_ftp
                            ftp.Password = pwd_ftp
                            ftp.Connect()
                            ftp.UploadFile(saveTo, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", False)
                            ftp.Close()
                            TArchivos.Rows.Add("PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", drLog("cLogRutaDescarga"))
                            setLog(id_descarga, "I", "Archivo creado: " & "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls")
                            setLog(id_descarga, "I", "Fin de descarga de Ventas...")
                            intResultado = 0
                    End Select
                    System.Threading.Thread.Sleep(1000)
                    If intResultado = 1 Then
                        Exit Try
                    End If
                Next
            Next
        Catch ex As Exception
            setLog(id_descarga, "E", "Error:" & ex.Message & " " & strMensaje & " " & ex.InnerException.ToString)
            intResultado = 1
            'Throw
        End Try
        'terminar()
        Return intResultado
    End Function



    Public Function getPH(ByVal fecha As Date, ByVal Archivo As String) As String
        Dim err As Integer = 0
        newDescarga("PH", fecha)
        err = getFTP("PH")
        If err = 0 Then
            iniciar()
            If Archivo = "V" Then 'solo ventas

                err = getArchivosPH(fecha, "Ventas")
            End If
            If Archivo = "I" Then 'solo inventarios

                err = getArchivosPH(fecha, "Inventarios")
            End If
            If Archivo = "T" And err = 0 Then 'Ventas e Inventarios
                err = getArchivosPH(fecha, "Ventas")
                If err = 0 Then
                    err = getArchivosPH(fecha, "Inventarios")
                End If
            End If
            terminar()
        End If


        If err = 0 Then
            setLog(id_descarga, "I", "Proceso terminado correctamente...")
        Else
            setLog(id_descarga, "E", "Proceso terminado con errores, favor de revisar secuencia...")
            setLog(id_descarga, "I", "Eliminando archivos generados...")
            Dim ax As Integer = 0
            While ax < TArchivos.Rows.Count               
                Try
                    ftp.ServerAddress = server_ftp
                    ftp.ServerPort = port_ftp
                    ftp.UserName = user_ftp
                    ftp.Password = pwd_ftp
                    ftp.Connect()
                    ftp.DeleteFile(TArchivos.Rows(ax).Item(0).ToString())
                    ftp.Close()
                    setLog(id_descarga, "I", "Archivo eliminado: " & TArchivos.Rows(ax).Item(0).ToString())
                Catch ex As Exception
                    setLog(id_descarga, "E", "Error:" & ex.Message)
                End Try
                ax = ax + 1
            End While
        End If

        conn.Open()
        sql = New Data.SqlClient.SqlCommand("update descargas set sts=@sts where id_descarga=@id_descarga", conn)
        sql.Parameters.AddWithValue("@id_descarga", id_descarga)
        sql.Parameters.AddWithValue("@sts", err)
        sql.ExecuteNonQuery()
        conn.Close()

        Dim xml As String = String.Empty
        xml = xml & "<?xml version=""1.0"" encoding=""UTF-8"" ?>" & vbCrLf
        xml = xml & "<root>" & vbCrLf
        xml = xml & "<resultado><valor>" & err & "</valor></resultado>" & vbCrLf
        xml = xml & "<archivos>" & vbCrLf
        Dim x As Integer = 0
        While x < TArchivos.Rows.Count
            xml = xml & "<archivo>" & vbCrLf
            xml = xml & "<nombre>" & TArchivos.Rows(x).Item(0).ToString() & "</nombre>" & vbCrLf
            xml = xml & "<ruta>" & TArchivos.Rows(x).Item(1).ToString() & "</ruta>" & vbCrLf
            xml = xml & "</archivo>" & vbCrLf
            x = x + 1
        End While
        xml = xml & "</archivos>" & vbCrLf
        xml = xml & "<idDescarga><id>" & id_descarga & "</id></idDescarga>"
        xml = xml & "<mensajes>" & vbCrLf
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("select * from detdescargas where id_descarga=@id_descarga order by id_log", conn)
        sql.Parameters.AddWithValue("@id_descarga", id_descarga)
        rs = sql.ExecuteReader
        While rs.Read
            xml = xml & "<msg>" & vbCrLf
            xml = xml & "<id>" & rs!id_log & "</id>" & vbCrLf
            xml = xml & "<tipo>" & rs!tipomsg & "</tipo>" & vbCrLf
            xml = xml & "<descripcion>" & rs!message & "</descripcion>" & vbCrLf
            xml = xml & "</msg>" & vbCrLf
        End While
        conn.Close()
        xml = xml & "</mensajes>" & vbCrLf
        xml = xml & "</root>"
        Return xml

    End Function

    Private Function busy() As Integer
        time = 0
        While wb.Busy And time < 300
            System.Threading.Thread.Sleep("1000")
            setLog(id_descarga, "I", "Esperando respuesta del servidor...OK")
            time = time + 1
        End While
        If time < 300 Then
            Return 0
        Else
            setLog(id_descarga, "E", "El servidor tardo mas de lo esperado en contestar...")
            Return 1
        End If
    End Function
End Class


