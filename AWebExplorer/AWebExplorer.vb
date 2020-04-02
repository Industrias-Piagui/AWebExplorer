Imports ICSharpCode.SharpZipLib.Zip
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports SHDocVw
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Threading.Tasks
Imports PortalPhRobot.BrowserRobot

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

    Public hwrequest As Net.HttpWebRequest
    Public response As Net.HttpWebResponse
    Public sts_hilo As Integer
    'Public timeout As Integer

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

    Public Enum Inv_PH
        COS1 = 1
        COS2 = 2
        COS3 = 3
        COS4 = 4
        IPI1 = 5
        IPI2 = 6
        IPI3 = 7
        IPI4 = 8
        IVK1 = 9
        LIFES = 10
    End Enum


    <System.Runtime.InteropServices.DllImport("user32.DLL")>
    Private Shared Function SendMessage(
            ByVal hWnd As System.IntPtr, ByVal wMsg As Integer,
            ByVal wParam As Integer, ByVal lParam As Integer
            ) As Integer
    End Function

    Private Declare Function GetWindowThreadProcessId Lib "user32" _
         (ByVal hwnd As Long,
    ByVal lpdwProcessId As Long) As Long

    Private Declare Function GetPIDByHWnd Lib "user32" _
     (ByVal hwnd As Long) As Long

    Private Declare Function OpenProcess Lib "kernel32" _
         (ByVal dwDesiredAccess As Long,
         ByVal bInheritHandle As Long,
         ByVal dwProcessId As Long) As Long

    Private Declare Function FindWindow Lib "user32" _
   Alias "FindWindowA" _
   (ByVal lpClassName As String,
   ByVal lpWindowName As String) As Long

    Public Sub New(ByVal xServer As String, ByVal xbd As String, ByVal xUser As String, ByVal xPwd As String)
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("es-ES")
        conn = New Data.SqlClient.SqlConnection("Data Source=dba;Initial Catalog=Portales;Integrated Security=No;User ID=aportales;Password=Aportales12;Persist Security Info=true")
        TArchivos = New System.Data.DataTable("Archivos")
        TArchivos.Columns.Add("Nombre")
        TArchivos.Columns.Add("Ruta")

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

        rs.Close()
        conn.Close()


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

    Public Function iniciar() As Integer
        Try
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

            Return 0
        Catch ex As Exception
            message = "Error al crear componente de IE. " & ex.Message
            Return 1
        End Try


    End Function
    Public Function terminar() As Integer
        Try

            Try
                ie.Quit()
                SendMessage(proceso, 16, 0, 0)
            Catch ex As Exception
            End Try

            Dim procesos() As Process = Process.GetProcessesByName("iexplore")
            Dim x As Integer = 0
            While x < procesos.Length
                Dim nombre = procesos(x).ProcessName
                If procesos(x).Id = proceso Then
                    procesos(x).Kill()
                End If
                x = x + 1
            End While

            Dim procesos_excel() As Process = Process.GetProcessesByName("EXCEL")
            x = 0
            While x < procesos_excel.Length
                Dim nombre = procesos_excel(x).ProcessName
                If procesos_excel(x).Id = proceso Then
                    procesos_excel(x).Kill()
                End If
                x = x + 1
            End While
        Catch ex As Exception
            message = "Error: " & ex.Message
            Return 1
        End Try
        System.Threading.Thread.Sleep(1000)
        Return 0
    End Function
    Public Function terminarTodo() As Integer
        Try

            If Not ie Is Nothing Then
                Try
                    ie.Quit()
                    SendMessage(proceso, 16, 0, 0)
                Catch ex As Exception
                End Try
            End If

            Dim procesos() As Process = Process.GetProcessesByName("iexplore")
            Dim x As Integer = 0
            While x < procesos.Length
                Dim nombre = procesos(x).ProcessName

                procesos(x).Kill()

                x = x + 1
            End While
        Catch ex As Exception
            message = "Error: " & ex.Message
            Return 1
        End Try
        System.Threading.Thread.Sleep(1000)
        Return 0
    End Function

    Public Function terminarIE() As Integer
        Try

            Dim procesos() As Process = Process.GetProcessesByName("iexplore")
            Dim x As Integer = 0
            While x < procesos.Length
                Dim nombre = procesos(x).ProcessName

                Try
                    procesos(x).Kill()
                Catch ex1 As Exception
                End Try

                x = x + 1
            End While
        Catch ex As Exception
            message = "Error: " & ex.Message
            Return 1
        End Try
        System.Threading.Thread.Sleep(1000)
        Return 0
    End Function
    Public Sub buscaArchivos(ByVal ruta As String)
        lista_archivos = New System.Data.DataTable("archivos")
        lista_archivos.Columns.Add("archivo")
        lista_archivos.Columns.Add("ruta")
        ReDim strdirectorios(0)
        SearchDirectory(ruta)
    End Sub

    Public Function Browser(ByVal url As String, ByVal timeout As Integer) As Integer
        Try

            Dim o As Object = Nothing
            wb.Visible = Visible
            wb.Navigate(url, o, o, o, o)

            setLog(id_descarga, "I", "Navegando: " & url)
            time = 0
            While wb.Busy And time < timeout
                System.Threading.Thread.Sleep("1000")
                setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                time = time + 1
            End While

            If time < timeout Then
                System.Threading.Thread.Sleep(3000)

                document = wb.Document
                Dim r() As String = url.Split("?")
                setLog(id_descarga, "I", "Pagina cargada: " & r(0))

                Dim cookies As String = document.cookie
                Dim domain As String = document.domain
                Dim c() As String = cookies.Split(";")
                Dim x As Integer = 0
                x = 0
                'If cookiesContainer Is Nothing Then
                cookiesContainer = New Net.CookieContainer
                'End If

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
            Else 'JG
                setLog(id_descarga, "E", "El sitio tardo demasiado en contestar. Intentar mas tarde...") 'JG
                Return 1
            End If
        Catch ex As Exception
            setLog(id_descarga, "E", "Error:" & ex.Message)
            Return 1
        End Try
    End Function

    Public Function setAttribute(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal valor As String) As Integer
        Dim rep As Integer = 0
        Dim str As String = ""
        Dim strElemento As String = String.Empty
        While rep < 5
            Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
            Dim iHTMLEle As mshtml.IHTMLElement
            For Each iHTMLEle In iHTMLCol
                If Not iHTMLEle.getAttribute(atributo) Is Nothing Then
                    strElemento &= "-" & iHTMLEle.getAttribute(atributo).ToString
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
                    ElseIf str.ToUpper.Replace(" ", "") = nombre.ToUpper.Replace(" ", "") Then
                        iHTMLEle.setAttribute("value", valor)
                        setLog(id_descarga, "I", "Asignacion de valor [ " & nombre & " : " & valor & "]")
                        Return 0
                        Exit For
                    End If
                End If
            Next
            setLog(id_descarga, "I", "Intento: " & rep + 1 & ". Intentando busqueda nuevamente...")
            wb.Refresh()
            While wb.Busy
                System.Threading.Thread.Sleep(1000)
            End While
            document = wb.Document
            System.Threading.Thread.Sleep(1000)
            rep = rep + 1
        End While
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre & " elementos: " & strElemento)
        Return 1
    End Function

    Public Function getAttributeValue(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal AtributoValor As String, ByRef valor As String) As Integer
        Dim str As String = ""
        Dim rep As Integer = 0
        While rep < 5
            Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
            Dim iHTMLEle As mshtml.IHTMLElement

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
            setLog(id_descarga, "I", "Intento: " & rep + 1 & ". Intentando busqueda nuevamente...")
            wb.Refresh()
            While wb.Busy
                System.Threading.Thread.Sleep(1000)
            End While
            document = wb.Document
            System.Threading.Thread.Sleep(1000)
            rep = rep + 1
        End While
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre)
        Return 1
    End Function

    Public Function setCheck(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal valor As String) As Integer
        Dim str As String = ""
        Dim val As String = ""
        Dim rep As Integer = 0
        While rep < 5
            Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
            Dim iHTMLEle As mshtml.IHTMLElement
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
            setLog(id_descarga, "I", "Intento: " & rep + 1 & ". Intentando busqueda nuevamente...")
            wb.Refresh()
            While wb.Busy
                System.Threading.Thread.Sleep(1000)
            End While
            document = wb.Document
            System.Threading.Thread.Sleep(1000)
            rep = rep + 1
        End While
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

    Public Function sendClick(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal timeout As Integer, Optional ByVal pdocument As mshtml.IHTMLDocument = Nothing) As Integer
        Dim str As String = ""
        Dim rep As Integer = 0
        While rep < 5
            Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
            Dim iHTMLEle As mshtml.IHTMLElement

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
                            System.Threading.Thread.Sleep(3000)
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
                            System.Threading.Thread.Sleep(3000)
                            Return 0
                        Else
                            setLog(id_descarga, "E", "El servidor tardo mas de lo esperado en contestar...")
                            Return 1
                        End If
                        Exit For
                    End If
                End If
            Next
            setLog(id_descarga, "I", "Intento: " & rep + 1 & ". Intentando busqueda nuevamente...")
            wb.Refresh()
            While wb.Busy
                System.Threading.Thread.Sleep(1000)
            End While
            document = wb.Document
            System.Threading.Thread.Sleep(1000)
            rep = rep + 1
        End While
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre)
        Return 1
    End Function

    Public Function sendClickPos(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal timeout As Integer, ByVal pos As Integer, Optional ByVal pdocument As mshtml.IHTMLDocument = Nothing) As Integer
        Dim str As String = ""
        Dim rep As Integer = 0
        While rep < 5


            Dim iHTMLCol As mshtml.IHTMLElementCollection

            If Not pdocument Is Nothing Then
                iHTMLCol = pdocument.getElementsByTagName(tipo)
            Else

                iHTMLCol = document.getElementsByTagName(tipo)
            End If


            Dim iHTMLEle As mshtml.IHTMLElement
            Dim i As Integer = 1
            For Each iHTMLEle In iHTMLCol
                If Not iHTMLEle.getAttribute(atributo) Is Nothing Then
                    str = iHTMLEle.getAttribute(atributo).ToString
                    If str.ToUpper.Equals(nombre.ToUpper) And i = pos Then
                        setLog(id_descarga, "I", "Ejecucion de evento Click [ " & nombre & "]")
                        iHTMLEle.click()
                        time = 0
                        While wb.Busy And time < timeout
                            System.Threading.Thread.Sleep("1000")
                            setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                            time = time + 1
                        End While
                        If time < timeout Then
                            System.Threading.Thread.Sleep(3000)
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
                            System.Threading.Thread.Sleep(3000)
                            Return 0
                        Else
                            setLog(id_descarga, "E", "El servidor tardo mas de lo esperado en contestar...")
                            Return 1
                        End If
                        Exit For
                    End If
                End If
                i = i + 1
            Next
            setLog(id_descarga, "I", "Intento: " & rep + 1 & ". Intentando busqueda nuevamente...")
            wb.Refresh()
            While wb.Busy
                System.Threading.Thread.Sleep(1000)
            End While
            document = wb.Document
            System.Threading.Thread.Sleep(1000)
            rep = rep + 1
        End While
        setLog(id_descarga, "E", "Elemento no encontrado: " & nombre)
        Return 1
    End Function

    Public Function sendLink(ByVal valor As String, ByVal timeout As Integer) As Integer
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.links
        Dim iHTMLEle As mshtml.IHTMLElement
        For Each iHTMLEle In iHTMLCol

            If Not iHTMLEle.innerText Is Nothing AndAlso iHTMLEle.innerText = valor Then
                setLog(id_descarga, "I", "Ejecucion de evento Link [ " & iHTMLEle.innerText & "]")
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
            ElseIf Not iHTMLEle.innerText Is Nothing AndAlso iHTMLEle.innerText.ToString.Contains(valor) Then
                setLog(id_descarga, "I", "Ejecucion de evento Link [ " & iHTMLEle.innerText & "]")
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
            ElseIf Not iHTMLEle.innerHTML Is Nothing AndAlso iHTMLEle.innerHTML.ToString.Contains(valor) Then
                setLog(id_descarga, "I", "Ejecucion de evento Link [ " & iHTMLEle.innerText & "]")
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


        Next
        setLog(id_descarga, "E", "Elemento no encontrado: " & valor)
        Return 1
    End Function

    Public Function buscaValor(ByVal campo As String, ByVal valor As String) As Boolean
        Dim str As String = ""
        Dim rep As Integer = 0
        'While rep < 5
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(campo)
        Dim iHTMLEle As mshtml.IHTMLElement
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

    Public Function Sears(ByVal fecha As Date, ByVal Archivo As String, ByVal secuencia As Integer) As String

        Dim err As Integer = 0
        Dim x As Integer = 0
        Dim ini_vta As Integer = 0
        Dim ini_inv As Integer = 0
        System.Net.ServicePointManager.Expect100Continue = False

        newDescarga("SEARS", fecha)
        If id_descarga > 0 Then
            setLog(id_descarga, "I", "Iniciando procesos Sears...")
            setLog(id_descarga, "I", "Numero de Intento: " & secuencia)
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
                        If secuencia = 0 And ini_vta = 0 Then
                            Elimina_Historial("4", fecha, "VTA")
                            ini_vta = 1
                        End If
                        err = getVentasSears(fecha, tabla.Rows(x))
                    End If
                    If Archivo = "I" And err = 0 Then 'solo inventarios
                        If secuencia = 0 And ini_inv = 0 Then
                            Elimina_Historial("4", fecha, "INV")
                            ini_inv = 1
                        End If
                        err = getInventarioSears(fecha, tabla.Rows(x))
                    End If
                    If Archivo = "T" And err = 0 Then 'Archivo venas e inventarios
                        If secuencia = 0 And ini_vta = 0 Then
                            Elimina_Historial("4", fecha, "VTA")
                            ini_vta = 1
                        End If
                        err = getVentasSears(fecha, tabla.Rows(x))
                        If err = 0 Then
                            If secuencia = 0 And ini_inv = 0 Then
                                Elimina_Historial("4", fecha, "INV")
                                ini_inv = 1
                            End If
                            err = getInventarioSears(fecha, tabla.Rows(x))
                        End If
                    End If

                    x = x + 1
                End While

                If err = 0 Then
                    setLog(id_descarga, "I", "Proceso terminado correctamente...")
                Else
                    setLog(id_descarga, "E", "Proceso terminado con errores, favor de revisar secuencia...")

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

            setLog(id_descarga, "I", "Generando XML de salida...")
            Dim xml As String = ""
            xml = xml & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
            xml = xml & "<root>"
            xml = xml & "<resultado><valor>" & err & "</valor></resultado>"
            xml = xml & "<archivos>"

            conn.Open()
            sql = New Data.SqlClient.SqlCommand("select * from archivos where fecha=@fecha and cliente=@cliente and tipo=@tipo", conn)
            sql.CommandTimeout = 0
            sql.Parameters.AddWithValue("@cliente", "4")
            sql.Parameters.AddWithValue("@tipo", IIf(Archivo = "V", "VTA", "INV"))
            sql.Parameters.AddWithValue("@fecha", fecha.ToString("ddMMyyyy"))
            rs = sql.ExecuteReader


            While rs.Read
                xml = xml & "<archivo>" & vbCrLf

                xml = xml & "<nombre>" & rs!archivo & "</nombre>" & vbCrLf

                xml = xml & "<ruta>" & rs!ruta & "</ruta>" & vbCrLf
                xml = xml & "</archivo>" & vbCrLf

            End While
            conn.Close()
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
        End If
        Return Nothing
    End Function

    Private Function getVentasSears(ByVal fecha As Date, ByVal row As DataRow) As Integer
        Dim res As Integer = 0
        Dim cont As Integer = 0

        If getArchivo("4", fecha, "SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml") = 1 Then
            setLog(id_descarga, "I", "Archivo SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml generado anteriormente...")
            Return 0
        End If

        res = iniciar()
        While res = 1 And cont < 10

            System.Threading.Thread.Sleep(1000)
            res = iniciar()
            cont = cont + 1
        End While
        If res = 1 Then
            setLog(id_descarga, "E", message)
            Return 1
        End If

        setLog(id_descarga, "I", "Iniciando descarga de Ventas...")

        res = validaSitio(row.Item(5))
        If res = 0 Then
            res = Browser(row.Item(5), 60)
            If res = 0 Then
                res = setAttribute("input", "name", "txtUsuarioBURN", row(3))
                If res = 0 Then
                    res = setAttribute("input", "name", "txtContrasenaBURN", row(4))
                    If res = 0 Then
                        res = sendClick("input", "name", "btnEntrar", 60)
                        If res = 0 Then
                            If buscaValor("p", "Contraseña Incorrecta") Or buscaValor("p", "Su contraseña ha expirado") Then
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
                                httpResponse = WRequest("http://proveedores.sears.com.mx/lib/main_xml.asp", "POST", "lstTiendas=3&hdnCodEmp=1&hdnCodFam2=9999-9999&hdnProvId=" & proveedor & "&hdnOpcion=" & hdnOpcion & "&hdhnAnio=" & hdhnAnio & "&hdnMes=" & hdnMes & "&hdnFiltroAux=" & hdnFiltroAux & "&hdnFechaIni=" & hdnFechaIni & "&hdnFechaFin=" & hdnFechaFin & "&optGenerar=" & optGenerar & "&optArts=" & optArts & "&optFechaPeriodo=" & optFechaPeriodo & "&lstAnio=" & lstAnio & "&lstMes=" & lstMes & "&optFechaPeriodo=" & optFechaPeriodo & "&txtFechaIni=" & txtFechaIni & "&txtFechaFin=" & txtFechaFin & "&optTelcel=1&cmdGenerar=Generar", cookiesContainer, 300)
                                If Not httpResponse Is Nothing Then
                                    setLog(id_descarga, "I", "Procesando datos descargados...")
                                    Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                    Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                    saveTo = saveTo.Replace(".tmp", ".zip")
                                    Dim tmp As String = saveTo
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
                                    Try
                                        fzip.ExtractZip(saveTo, ruta & "\tmp_piagui_" & id_descarga, "")
                                    Catch ex As Exception
                                        setLog(id_descarga, "E", "Error de descarga: " & ex.Message)
                                        res = 1
                                    End Try
                                    If res = 0 Then
                                        buscaArchivos(ruta & "\tmp_piagui_" & id_descarga)
                                        If lista_archivos.Rows.Count > 0 Then
                                            Dim xml As New Xml.XmlDocument
                                            xml.Load(lista_archivos.Rows(0).Item(1).ToString)
                                            Dim nodelist As Xml.XmlNodeList = xml.SelectNodes("/root/VTAS")
                                            If nodelist.Count = 0 Then
                                                res = 1
                                                setLog(id_descarga, "E", "No se encontro informacion para la fecha seleccionada. Intentar mas tarde.")
                                            End If

                                            If res = 0 Then

                                                Try
                                                    setLog(id_descarga, "I", "Enviando archivos a FTP...")
                                                    ftp.ServerAddress = server_ftp
                                                    ftp.ServerPort = port_ftp
                                                    ftp.UserName = user_ftp
                                                    ftp.Password = pwd_ftp
                                                    ftp.Connect()
                                                    ftp.UploadFile(lista_archivos.Rows(0).Item(1).ToString, "SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml", False)

                                                    Dim features As EnterpriseDT.Net.Ftp.FTPReply = ftp.InvokeFTPCommand("cwd /Respaldos")
                                                    ftp.UploadFile(lista_archivos.Rows(0).Item(1).ToString, "SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "-" & Now.ToString("ddMMyyyy_HHmmss") & ".xml", False)

                                                    ftp.Close()
                                                    'System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                                    setArchivo("4", fecha, "SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml", row.Item(6).ToString() & "\", "VTA")
                                                    TArchivos.Rows.Add("SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml", row.Item(6).ToString() & "\")
                                                    setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "SEARS_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & ".xml")
                                                    setLog(id_descarga, "I", "Fin de descarga de Ventas...")

                                                    If System.IO.File.Exists(lista_archivos.Rows(0).Item(1).ToString) Then
                                                        Try
                                                            System.IO.File.Delete(lista_archivos.Rows(0).Item(1).ToString)
                                                        Catch ex As Exception

                                                        End Try
                                                    End If

                                                Catch ex As Exception
                                                    setLog(id_descarga, "E", "Error:" & ex.Message)
                                                    res = 1
                                                End Try
                                            End If
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
        Dim res As Integer = 0
        Dim cont As Integer = 0

        If getArchivo("4", fecha, "SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml") = 1 Then
            setLog(id_descarga, "I", "Archivo SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml generado anteriormente...")
            Return 0
        End If

        res = iniciar()
        While res = 1 And cont < 10

            System.Threading.Thread.Sleep(1000)
            res = iniciar()
            cont = cont + 1
        End While
        If res = 1 Then
            setLog(id_descarga, "E", message)
            Return 1
        End If
        setLog(id_descarga, "I", "Iniciando descarga de Inventarios...")

        res = validaSitio(row.Item(5))
        If res = 0 Then
            res = Browser(row.Item(5), 60)
            If res = 0 Then
                res = setAttribute("input", "name", "txtUsuarioBURN", row(3))
                If res = 0 Then
                    res = setAttribute("input", "name", "txtContrasenaBURN", row(4))
                    If res = 0 Then
                        res = sendClick("input", "name", "btnEntrar", 60)
                        If res = 0 Then
                            If buscaValor("p", "Contraseña Incorrecta") Or buscaValor("p", "Su contraseña ha expirado") Then
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
                                httpResponse = WRequest("http://proveedores.sears.com.mx/lib/main_xml.asp", "POST", "lstTiendas=3&hdnCodEmp=1&hdnCodFam2=9999-9999&hdnProvId=" & proveedor & "&hdnOpcion=" & hdnOpcion & "&hdhnAnio=" & hdhnAnio & "&hdnMes=" & hdnMes & "&hdnFiltroAux=" & hdnFiltroAux & "&hdnFechaIni=" & hdnFechaIni & "&hdnFechaFin=" & hdnFechaFin & "&optGenerar=" & optGenerar & "&optArts=" & optArts & "&optFechaPeriodo=" & optFechaPeriodo & "&lstAnio=" & lstAnio & "&lstMes=" & lstMes & "&optFechaPeriodo=" & optFechaPeriodo & "&txtFechaIni=" & txtFechaIni & "&txtFechaFin=" & txtFechaFin & "&optTelcel=1&cmdGenerar=Generar", cookiesContainer, 500)
                                If Not httpResponse Is Nothing Then
                                    setLog(id_descarga, "I", "Procesando datos descargados...")
                                    Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                    Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                    Dim tmp As String = saveTo
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
                                    Try
                                        fzip.ExtractZip(saveTo, ruta & "\tmp_piagui_" & id_descarga, "")
                                    Catch ex As Exception
                                        setLog(id_descarga, "E", "Error de descarga: " & ex.Message)
                                        res = 1
                                    End Try
                                    If res = 0 Then
                                        buscaArchivos(ruta & "\tmp_piagui_" & id_descarga)
                                        If lista_archivos.Rows.Count > 0 Then

                                            Dim xml As New Xml.XmlDocument
                                            xml.Load(lista_archivos.Rows(0).Item(1).ToString)
                                            Dim nodelist As Xml.XmlNodeList = xml.SelectNodes("/root/row")
                                            If nodelist.Count = 0 Then
                                                res = 1
                                                setLog(id_descarga, "E", "No se encontro informacion para la fecha seleccionada. Intentar mas tarde.")
                                            End If
                                            If res = 0 Then

                                                Try
                                                    setLog(id_descarga, "I", "Enviando archivos a FTP...")
                                                    ftp.ServerAddress = server_ftp
                                                    ftp.ServerPort = port_ftp
                                                    ftp.UserName = user_ftp
                                                    ftp.Password = pwd_ftp
                                                    ftp.Connect()
                                                    ftp.UploadFile(lista_archivos.Rows(0).Item(1).ToString, "SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml", False)

                                                    Dim features As EnterpriseDT.Net.Ftp.FTPReply = ftp.InvokeFTPCommand("cwd /Respaldos")
                                                    ftp.UploadFile(lista_archivos.Rows(0).Item(1).ToString, "SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1-" & Now.ToString("ddMMyyyy_HHmmss") & ".xml", False)

                                                    ftp.Close()

                                                    setArchivo("4", fecha, "SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml", row.Item(6).ToString() & "\", "INV")
                                                    TArchivos.Rows.Add("SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml", row.Item(6).ToString() & "\")
                                                    setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "SEARS_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & "_" & "1.xml")
                                                    setLog(id_descarga, "I", "Fin de descarga de Inventarios...")

                                                    If System.IO.File.Exists(lista_archivos.Rows(0).Item(1).ToString) Then
                                                        Try
                                                            System.IO.File.Delete(lista_archivos.Rows(0).Item(1).ToString)
                                                        Catch ex As Exception

                                                        End Try
                                                    End If

                                                Catch ex As Exception
                                                    setLog(id_descarga, "E", "Error:" & ex.Message)
                                                    res = 1
                                                End Try

                                            End If

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


    Private Sub newDescarga(ByVal cadena As String, ByVal fecha As Date)
        Dim conn1 As New Data.SqlClient.SqlConnection(conn.ConnectionString)
        Dim tra As Data.SqlClient.SqlTransaction
        conn1.Open()
        Try
            tra = conn1.BeginTransaction
            Dim sql1 As SqlClient.SqlCommand
            sql1 = New Data.SqlClient.SqlCommand("prc_newdescarga", conn1, tra)
            sql1.CommandType = CommandType.StoredProcedure
            sql1.CommandTimeout = 0
            sql1.Parameters.AddWithValue("@cadena", cadena)
            sql1.Parameters.AddWithValue("@fecha", fecha)
            sql1.Parameters.AddWithValue("@id_descarga", 0)
            sql1.Parameters.Item(2).Direction = ParameterDirection.Output
            sql1.ExecuteNonQuery()
            id_descarga = sql1.Parameters.Item(2).Value
            tra.Commit()
        Catch ex As Exception
            If Not tra Is Nothing Then
                tra.Rollback()
            End If
            id_descarga = -1
        End Try
        conn1.Close()
    End Sub

    Private Sub getDescarga(ByVal cadena As String, ByVal fecha As Date)
        Dim conn1 As New Data.SqlClient.SqlConnection(conn.ConnectionString)

        Try
            Dim strSql As String = "Select * from estatusPH where fecha ='" & fecha.ToString("dd/MM/yyy") & "' and estatus ='EN_PROCESO'"
            Dim sqlCmd As New SqlClient.SqlCommand(strSql, conn1)
            Dim sqlRdr As SqlClient.SqlDataReader

            conn1.Open()

            sqlRdr = sqlCmd.ExecuteReader()

            If sqlRdr.HasRows Then
                While sqlRdr.Read
                    id_descarga = sqlRdr.Item("idDescarga")
                End While
            End If
            conn1.Close()
        Catch ex As Exception
            id_descarga = -1
        End Try

    End Sub

    Public Sub setLog(ByVal id_descarga As Integer, ByVal tipo As String, ByVal msg As String)
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

    Function WRequest(ByVal URL As String, ByVal method As String, ByVal POSTdata As String, ByVal cookies As Net.CookieContainer, ByVal timeout As Integer) As Net.HttpWebResponse
        Dim responseData As String = ""
        Try
            Dim r() As String = URL.Split("?")

            setLog(id_descarga, "I", "Conectandose via WebRequest: " & r(0))
            hwrequest = Net.WebRequest.Create(URL)
            'Dim response As Net.HttpWebResponse
            hwrequest.Accept = "*/*"
            hwrequest.AllowAutoRedirect = False
            hwrequest.KeepAlive = True
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
            If 1 = 0 Then
                response = hwrequest.GetResponse()
            End If
            'response = hwrequest.GetResponse()
            hilo = New Threading.Thread(AddressOf getResponse)
            sts_hilo = 1
            hilo.Start()
            Dim tiempo As Integer = 0
            While sts_hilo <> 0 And tiempo < timeout
                System.Threading.Thread.Sleep(1000)
                tiempo = tiempo + 1
                setLog(id_descarga, "I", "Esperando respuesta del servidor...")
                If sts_hilo = 2 Then
                    setLog(id_descarga, "E", message)
                    Exit While
                End If
            End While
            If sts_hilo = 0 And tiempo < timeout Then
                Return response
            Else
                If hilo.IsAlive Then
                    hilo.Abort()
                End If
                Return Nothing
            End If

            Return response
        Catch e As Exception
            message = "Error: " & e.Message
        End Try
        Return Nothing
    End Function

    Function WRequest2(ByVal URL As String, ByVal method As String, ByVal POSTdata As String, ByVal cookies As Net.CookieContainer, ByVal timeout As Integer) As Net.HttpWebResponse
        Dim responseData As String = ""
        Try
            Dim r() As String = URL.Split("?")

            setLog(id_descarga, "I", "Conectandose via WebRequest: " & r(0))
            hwrequest = Net.WebRequest.Create(URL)
            'Dim response As Net.HttpWebResponse
            hwrequest.Accept = "*/*"
            hwrequest.AllowAutoRedirect = False
            hwrequest.KeepAlive = True
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
            'hilo = New Threading.Thread(AddressOf getResponse)
            'sts_hilo = 1
            'hilo.Start()
            'Dim tiempo As Integer = 0
            'While sts_hilo <> 0 And tiempo < timeout
            '    System.Threading.Thread.Sleep(1000)
            '    tiempo = tiempo + 1
            '    setLog(id_descarga, "I", "Esperando respuesta del servidor...")
            'End While
            'If sts_hilo = 0 And tiempo < timeout Then
            Return response
            'Else
            'If hilo.IsAlive Then
            '    hilo.Abort()
            'End If
            'Return Nothing
            'End If

            Return response
        Catch e As Exception
            message = "Error: " & e.Message
        End Try
        Return Nothing
    End Function

    Public Sub getResponse()
        Try
            response = hwrequest.GetResponse
            sts_hilo = 0
        Catch ex As Exception
            sts_hilo = 2
            message = ex.Message
        End Try
    End Sub

    Public Sub ReadWriteStream(ByVal readStream As IO.Stream, ByVal writeStream As IO.Stream)
        Dim Length As Integer = 2048
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
    Public Function Elimina_Historial(ByVal cliente As String, ByVal fecha As Date, ByVal tipo As String) As Integer
        setLog(id_descarga, "I", "Eliminando historicos de descargas...")
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("select * from archivos where fecha=@fecha and cliente=@cliente and tipo=@tipo", conn)
        sql.CommandTimeout = 0
        sql.Parameters.AddWithValue("@fecha", fecha.ToString("ddMMyyyy"))
        sql.Parameters.AddWithValue("@cliente", cliente)
        sql.Parameters.AddWithValue("@tipo", tipo)
        rs = sql.ExecuteReader
        While rs.Read
            Try
                ftp.ServerAddress = server_ftp
                ftp.ServerPort = port_ftp
                ftp.UserName = user_ftp
                ftp.Password = pwd_ftp
                ftp.Connect()
                ftp.DeleteFile(rs!archivo)
                ftp.Close()
                setLog(id_descarga, "I", "Archivo eliminado: " & rs!archivo)
            Catch ex As Exception
                If ftp.IsConnected Then
                    ftp.Close()
                End If
                setLog(id_descarga, "E", "Error:" & ex.Message)
            End Try
        End While
        conn.Close()
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("delete from archivos where fecha=@fecha and cliente=@cliente and tipo=@tipo", conn)
        sql.CommandTimeout = 0
        sql.Parameters.AddWithValue("@fecha", fecha.ToString("ddMMyyyy"))
        sql.Parameters.AddWithValue("@cliente", cliente)
        sql.Parameters.AddWithValue("@tipo", tipo)
        sql.ExecuteNonQuery()
        conn.Close()
        Return 0

    End Function

    Public Function Liverpool(ByVal fecha As Date, ByVal Archivo As String, ByVal secuencia As Integer) As String
        Dim x As Integer = 0
        Dim c As Integer = 0
        Dim err As Integer = 0
        Dim ini_vta As Integer = 0
        Dim ini_inv As Integer = 0
        newDescarga("LIVERPOOL", fecha)
        System.Net.ServicePointManager.Expect100Continue = False
        If id_descarga > 0 Then
            setLog(id_descarga, "I", "Iniciando procesos Liverpool...")
            setLog(id_descarga, "I", "Numero de Intento: " & secuencia)
            err = getFTP("LIVERPOOL")

            If err = 0 Then
                conn.Open()
                sql = New Data.SqlClient.SqlCommand("select * from clogin where clogportal='LIVERPOOL' ", conn)
                Dim tabla As New System.Data.DataTable

                tabla.Load(sql.ExecuteReader)
                conn.Close()



                setLog(id_descarga, "I", "Iniciando proceso...")
                setLog(id_descarga, "I", "Tipo de descarga solicitada: " & Archivo)

                While x < tabla.Rows.Count And err = 0
                    c = 0
                    Dim conf As New System.Data.DataTable
                    conn.Open()
                    sql = New Data.SqlClient.SqlCommand("select * from conf_seccion_liverpool where proveedor = @proveedor", conn)
                    sql.Parameters.AddWithValue("@proveedor", tabla.Rows(x).Item(3).ToString)
                    conf.Load(sql.ExecuteReader)
                    conn.Close()


                    setLog(id_descarga, "I", "Procesando cliente: " & tabla.Rows(x).Item(3).ToString)

                    While c < conf.Rows.Count And err = 0

                        If Archivo = "V" And err = 0 And conf.Rows(c).Item(1).ToString = "VTA" Then 'solo ventas
                            If secuencia = 0 And ini_vta = 0 Then
                                'Elimina_Historial("2", fecha, "VTA")
                                ini_vta = 1
                            End If
                            err = getVentasLiverpool(fecha, tabla.Rows(x), conf.Rows(c))
                        End If
                        If Archivo = "I" And err = 0 And conf.Rows(c).Item(1).ToString = "INV" Then 'solo inventarios
                            If secuencia = 0 And ini_inv = 0 Then
                                'Elimina_Historial("2", fecha, "INV")
                                ini_inv = 1
                            End If
                            err = getInventarioLiverpool(fecha, tabla.Rows(x), conf.Rows(c))
                        End If
                        If Archivo = "T" And err = 0 Then 'Archivo venas e inventarios
                            If conf.Rows(c).Item(1).ToString = "VTA" Then
                                If secuencia = 0 And ini_vta = 0 Then
                                    'Elimina_Historial("2", fecha, "VTA")
                                    ini_vta = 1
                                End If
                                err = getVentasLiverpool(fecha, tabla.Rows(x), conf.Rows(c))
                            End If
                            If err = 0 Then
                                If conf.Rows(c).Item(1).ToString = "INV" Then
                                    If secuencia = 0 And ini_inv = 0 Then
                                        'Elimina_Historial("2", fecha, "INV")
                                        ini_inv = 1
                                    End If
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
                    'setLog(id_descarga, "I", "Eliminado archivos generados...")
                    'Dim ax As Integer = 0
                    'While ax < TArchivos.Rows.Count
                    '    'If System.IO.File.Exists(TArchivos.Rows(ax).Item(1).ToString() & TArchivos.Rows(ax).Item(0).ToString()) Then
                    '    '    System.IO.File.Delete(TArchivos.Rows(ax).Item(1).ToString() & TArchivos.Rows(ax).Item(0).ToString())
                    '    'End If

                    '    Try
                    '        ftp.ServerAddress = server_ftp
                    '        ftp.ServerPort = port_ftp
                    '        ftp.UserName = user_ftp
                    '        ftp.Password = pwd_ftp
                    '        ftp.Connect()
                    '        ftp.DeleteFile(TArchivos.Rows(ax).Item(0).ToString())
                    '        ftp.Close()
                    '        setLog(id_descarga, "I", "Archivo eliminado: " & TArchivos.Rows(ax).Item(0).ToString())
                    '    Catch ex As Exception
                    '        setLog(id_descarga, "E", "Error:" & ex.Message)
                    '    End Try
                    '    ax = ax + 1
                    'End While
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

            setLog(id_descarga, "I", "Generando XML de salida...")
            Dim xml As String = ""
            xml = xml & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
            xml = xml & "<root>"
            xml = xml & "<resultado><valor>" & err & "</valor></resultado>"
            xml = xml & "<archivos>"

            conn.Open()
            sql = New Data.SqlClient.SqlCommand("select * from archivos where fecha=@fecha and cliente=@cliente and tipo=@tipo", conn)
            sql.CommandTimeout = 0
            sql.Parameters.AddWithValue("@cliente", "2")
            sql.Parameters.AddWithValue("@tipo", IIf(Archivo = "V", "VTA", "INV"))
            sql.Parameters.AddWithValue("@fecha", fecha.ToString("ddMMyyyy"))
            rs = sql.ExecuteReader


            While rs.Read
                xml = xml & "<archivo>" & vbCrLf

                xml = xml & "<nombre>" & rs!archivo & "</nombre>" & vbCrLf

                xml = xml & "<ruta>" & rs!ruta & "</ruta>" & vbCrLf
                xml = xml & "</archivo>" & vbCrLf

            End While
            conn.Close()
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
        End If
        Return Nothing
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

    Public Function setText(ByVal tipo As String, ByVal atributo As String, ByVal nombre As String, ByVal valor As String, ByVal evento As String) As Integer
        Dim iHTMLCol As mshtml.IHTMLElementCollection = document.getElementsByTagName(tipo)
        Dim iHTMLEle As mshtml.IHTMLElement
        Dim str As String = ""
        For Each iHTMLEle In iHTMLCol
            If Not iHTMLEle.getAttribute(atributo) Is Nothing Then
                str = iHTMLEle.getAttribute(atributo).ToString
                If str.ToUpper.Equals(nombre.ToUpper) Then
                    iHTMLEle.setAttribute("value", valor)
                    If evento <> "" Then
                        Dim selectElem As mshtml.IHTMLTextElement = DirectCast(iHTMLEle, mshtml.IHTMLTextElement)
                        If Not selectElem Is Nothing Then
                            Dim tempElem As mshtml.HTMLTextElement
                            tempElem = DirectCast(selectElem, mshtml.HTMLTextElement)
                            'tempElem.selectedIndex = index
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

    Public Function OTBLiverpool(ByVal fecha As Date) As Integer
        Dim x As Integer = 0
        Dim c As Integer = 0
        Dim err As Integer = 0
        Dim ini_vta As Integer = 0
        Dim ini_inv As Integer = 0
        System.Net.ServicePointManager.Expect100Continue = False
        terminarIE()
        newDescarga("OTB_LIVERPOOL", fecha)
        If id_descarga > 0 Then
            setLog(id_descarga, "I", "Iniciando procesos OTB Liverpool...")
            err = getFTP("OTB_LIVERPOOL")
            If err = 0 Then
                conn.Open()
                sql = New Data.SqlClient.SqlCommand("select * from clogin where clogportal='LIVERPOOL' and isnull(rutaotb,'')<>''", conn)
                Dim tabla As New System.Data.DataTable
                Dim conf As New System.Data.DataTable
                tabla.Load(sql.ExecuteReader)
                conn.Close()

                setLog(id_descarga, "I", "Iniciando proceso...")
                'Elimina_Historial("2", fecha, "OTB")
                While x < tabla.Rows.Count And err = 0
                    setLog(id_descarga, "I", "Procesando cliente: " & tabla.Rows(x).Item(3).ToString)

                    conn.Open()
                    sql = New Data.SqlClient.SqlCommand("select * from conf_otb_liverpool where proveedor=@proveedor", conn)
                    sql.Parameters.AddWithValue("@proveedor", tabla.Rows(x).Item(3))
                    conf.Load(sql.ExecuteReader)
                    conn.Close()

                    While c < conf.Rows.Count And err = 0
                        err = getOTBLiverpool(fecha, tabla.Rows(x), conf.Rows(c), tabla.Rows(x).Item(3))
                        c = c + 1
                    End While
                    x = x + 1
                End While
            End If
        End If
        Return err
    End Function
    Private Function getOTBLiverpool(ByVal fecha As Date, ByVal row As Data.DataRow, ByVal conf As Data.DataRow, ByVal proveedor As String) As Integer
        Dim res As Integer = 0
        Try
            Dim cont As Integer = 0
            Dim seccion As String = IIf(conf(1).ToString <> "*", "_" & conf(1).ToString.Replace(",", "-"), "")
            If getArchivo("2", fecha, "LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & proveedor & "_" & conf(0) & seccion & ".xls") Then
                setLog(id_descarga, "I", "Archivo:" & "LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & proveedor & "_" & conf(0) & seccion & ".xls" & " generado anteriormente...")
                Return 0
            End If

            res = iniciar()
            While res = 1 And cont < 10
                System.Threading.Thread.Sleep(1000)
                res = iniciar()
                cont = cont + 1
            End While

            System.Net.ServicePointManager.Expect100Continue = False

            setLog(id_descarga, "I", "Iniciando descarga de OTB...")
            Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true", 60)


            res = validaSitio(row.Item(5))
            If res = 0 Then
                res = Browser(row.Item(5), 60)
                If res = 0 Then
                    res = setAttribute("input", "name", "j_user", row(3).ToString)
                    If res = 0 Then
                        res = setAttribute("input", "name", "j_password", row(4).ToString)
                        If res = 0 Then
                            res = sendClick("input", "name", "uidPasswordLogon", 60)
                            If res = 0 Then
                                If buscaValor("span", "Autentificación de usuario fallida") Then
                                    'Error contraseña incorrecta
                                    setLog(id_descarga, "E", "Contraseña incorrecta, imposible accesar al portal")
                                    message = "Error: Contraseña incorrecta, imposible accesar al portal."
                                    res = 2
                                Else
                                    'res = Browser("https://bwsext.liverpool.com.mx/sap/bw/BEx?sap-language=es&sap-client=400&accessibility=&style_sheet=http%3A%2F%2Fproveedores.liverpool.com.mx%3A80%2Firj%2Fportalapps%2Fcom.sap.portal.design.portaldesigndata%2Fthemes%2Fportal%2Flookfeel_portal%2FBIReports30%2FBIReports30_ie6.css%3F6.0.20.0.1&TEMPLATE_ID=ZMBPSMAP_Q002A_PR", 60)

                                    Dim numAleatorio As New Random(CInt(Date.Now.Ticks And Integer.MaxValue))

                                    Dim urlx As String = "https://bwsext.liverpool.com.mx/sap/bw/BEx?sap-client=400&sap-language=ES&accessibility=&style_sheet=http%3A%2F%2Fproveedores.liverpool.com.mx%3A80%2Fcom.sap.portal.design.portaldesigndata%2Fthemes%2Fportal%2Fprov_nuevo%2FBIReports30%2FBIReports30_ie6.css%3Fv%3D7.31.11.0.6&sap-tray-type=PLAIN&sap-tray-padding=X&TEMPLATE_ID=ZMBPSMAP_Q002A_PR&sapDocumentRenderingMode=EmulateIE8&NavMode=0&NavPathUpdate=true&sap-ie=EmulateIE8&idNum=" & numAleatorio.Next
                                    setLog(id_descarga, "I", urlx)
                                    res = Browser(urlx, 60)
                                    Dim wx1 As New System.IO.StreamWriter("Pagina0.html", False)
                                    wx1.WriteLine(document.documentElement.outerHTML)
                                    wx1.Flush()
                                    wx1.Close()
                                    If res <> 0 Then
                                        message = "Error: Reporte no disponible. Error de carga."
                                    End If
                                    Dim ini As Integer = 6
                                    If conf(0).ToString <> "*" Then
                                        'res = setSelect("select", "name", "VAR_OPERATOR_5", "BT", 0, "onchange")
                                        res = setAttribute("input", "name", "VAR_VALUE_EXT_4", conf(0).ToString)

                                        If res = 0 Then
                                            'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_5", conf(0).ToString)
                                        Else
                                            message = "Error: Reporte no disponible. Error de carga."
                                        End If
                                    End If
                                    Dim m As Integer = 0
                                    If res = 0 Then
                                        If conf(1) <> "*" Then
                                            Dim marcas() As String = conf(1).ToString.Split(",")
                                            While m < marcas.Length And res = 0
                                                If m > 0 Then
                                                    Dim pos As Integer = 10
                                                    res = sendClick("a", "href", "JavaScript:SAPBWBUTTON('PROCESS_VARIABLES','VAR_NEW_LINES', 'ZATTRPL2", 30)
                                                End If

                                                res = setSelect("select", "name", "VAR_OPERATOR_" & ini, "EQ", 0, "onchange")
                                                If res = 0 Then
                                                    res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_" & ini, marcas(m))
                                                Else
                                                    message = "Errro: Reporte no disponible. Error de carga."
                                                End If
                                                m = m + 1
                                                ini = ini + 1
                                            End While
                                        End If
                                    End If

                                    If res = 0 Then
                                        If m = 0 Then
                                            res = setAttribute("input", "name", "VAR_VALUE_EXT_12", fecha.ToString("yyyy"))
                                        Else
                                            res = setAttribute("input", "name", "VAR_VALUE_EXT_" & 12 + m - 1, fecha.ToString("yyyy"))
                                        End If

                                    End If
                                    Dim wx As New System.IO.StreamWriter("Pagina1.html", False)
                                    wx.WriteLine(document.documentElement.outerHTML)
                                    wx.Flush()
                                    wx.Close()
                                    System.Threading.Thread.Sleep(5000)
                                    If res = 0 Then
                                        res = sendClick("a", "href", "javascript:SAPBWBUTTON('PROCESS_VARIABLES','VAR_SUBMIT',' ')", 120)
                                        If res = 0 Then
                                            res = Browser(wb.LocationURL, 60)
                                            Dim url As String = wb.LocationURL.Replace("CMD=", "$")
                                            Dim url1() As String = url.Split("$")
                                            Dim url2() As String = url1(0).Split("?")

                                            httpResponse = WRequest2(url2(0), "POST", url2(1) & "DATA_PROVIDER=DATAPROVIDER_1&ENHANCED_MENU=&CMD=EXPORT&FORMAT=XLS&SUPPRESS_REPETITION_TEXTS=X", cookiesContainer, 100)
                                            If Not httpResponse Is Nothing Then
                                                setLog(id_descarga, "I", "Procesando datos descargados...")

                                                Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                                Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                                Dim tmp As String = saveTo
                                                saveTo = saveTo.Replace(".tmp", ".xls")
                                                If System.IO.File.Exists(tmp) Then
                                                    System.IO.File.Delete(tmp)
                                                End If

                                                Dim lineas As System.IO.StreamReader = New System.IO.StreamReader(o)
                                                Dim w As IO.StreamWriter = New System.IO.StreamWriter(saveTo)
                                                setLog(id_descarga, "I", "Archivo temporal generado: " & saveTo)
                                                Dim ren As Integer = 0
                                                While Not lineas.EndOfStream
                                                    Dim l As String = lineas.ReadLine
                                                    w.WriteLine(l)
                                                End While
                                                w.Flush()
                                                w.Close()

                                                Try
                                                    Dim meses() As String = {"ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"}
                                                    setLog(id_descarga, "I", "Enviando archivos a FTP...")
                                                    Dim anio As String = fecha.ToString("yyyy")
                                                    If Not System.IO.Directory.Exists(row(8).ToString & "\" & fecha.Year) Then
                                                        System.IO.Directory.CreateDirectory(row(8).ToString & "\" & fecha.Year)
                                                    End If
                                                    Dim mes As String = meses(fecha.Month - 1)
                                                    If Not System.IO.Directory.Exists(row(8).ToString & "\" & anio & "\" & meses(fecha.Month - 1)) Then
                                                        System.IO.Directory.CreateDirectory(row(8).ToString & "\" & anio & "\" & meses(fecha.Month - 1))
                                                    End If
                                                    Dim dia As String = fecha.ToString("dd")
                                                    If Not System.IO.Directory.Exists(row(8).ToString & "\" & anio & "\" & mes & "\" & dia) Then
                                                        System.IO.Directory.CreateDirectory(row(8).ToString & "\" & anio & "\" & mes & "\" & dia)
                                                    End If

                                                    Dim ruta As String = row(8).ToString & "\" & anio & "\" & mes & "\" & dia
                                                    'System.IO.File.Copy(saveTo, ruta & "\" & "LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & conf(0) & seccion & ".xls")

                                                    If System.IO.File.Exists(ruta & "\LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & proveedor & "_" & conf(0) & seccion & ".xls") Then
                                                        System.IO.File.Delete(ruta & "\LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & proveedor & "_" & conf(0) & seccion & ".xls")
                                                    End If
                                                    Dim xls As New Excel.Application
                                                    xls.DisplayAlerts = False
                                                    xls.Workbooks.Open(saveTo)
                                                    xls.Workbooks(1).SaveAs(ruta & "\LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & proveedor & "_" & conf(0) & seccion & ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal)

                                                    xls.Workbooks.Close()
                                                    xls.Quit()
                                                    xls = Nothing

                                                    'ftp.ServerAddress = server_ftp
                                                    'ftp.ServerPort = port_ftp
                                                    'ftp.UserName = user_ftp
                                                    'ftp.Password = pwd_ftp
                                                    'ftp.Connect()

                                                    'ftp.UploadFile(saveTo, "LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & conf(0) & seccion & ".xls", False)

                                                    ''Dim features As EnterpriseDT.Net.Ftp.FTPReply = ftp.InvokeFTPCommand("cwd /Respaldos")
                                                    ''ftp.UploadFile(saveTo, "LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "-" & Now.ToString("ddMMyyyy_HHmmss") & ".csv", False)

                                                    'ftp.Close()
                                                    ''System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                                    setArchivo("2", fecha, "LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & proveedor & "_" & conf(0) & seccion & ".xls", row.Item(6).ToString() & "\", "OTB")
                                                    TArchivos.Rows.Add("LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & proveedor & "_" & conf(0) & seccion & ".xls", row.Item(6).ToString() & "\")
                                                    setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "LIVERPOOL_OTB_" & fecha.ToString("ddMMyyyy") & "_" & proveedor & "_" & conf(0) & seccion & ".xls")
                                                    setLog(id_descarga, "I", "Fin de descarga de Ventas...")
                                                    If System.IO.File.Exists(saveTo) Then
                                                        Try
                                                            System.IO.File.Delete(saveTo)
                                                        Catch ex As Exception
                                                        End Try
                                                    End If
                                                Catch ex As Exception
                                                    setLog(id_descarga, "E", "Error:" & ex.Message)
                                                    message = "Excepcion: " & ex.Message
                                                    res = 1
                                                End Try

                                            End If
                                        End If

                                    End If

                                End If
                            End If
                        End If
                    End If

                End If
            End If
            Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true", 60)
            terminar()
        Catch ex As Exception
            res = 1
            setLog(id_descarga, "E", "Error:" & ex.Message)
            message = "Excepcion:" & ex.Message
        End Try
        terminar()
        Return res
    End Function
    Private Function getVentasLiverpool(ByVal fecha As Date, ByVal row As Data.DataRow, ByVal conf As Data.DataRow) As Integer
        Dim res As Integer = 0
        Dim cont As Integer = 0

        Dim seccion As String = IIf(conf(2).ToString <> "*", "_" & conf(2).ToString, "")
        If getArchivo("2", fecha, "LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv") = 1 Then
            setLog(id_descarga, "I", "Archivo: LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv generado anteriormente...")
            Return 0
        End If

        res = iniciar()
        While res = 1 And cont < 10
            'setLog(id_descarga, "E", message)
            System.Threading.Thread.Sleep(1000)
            res = iniciar()
            cont = cont + 1
        End While
        If res = 1 Then
            setLog(id_descarga, "E", message)
            Return 1
        End If
        setLog(id_descarga, "I", "Iniciando descarga de Ventas...")
        Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true", 60)

        res = validaSitio(row.Item(5))
        If res = 0 Then
            res = Browser(row.Item(5), 60)
            If res = 0 Then
                res = setAttribute("input", "name", "j_user", row(3).ToString)
                If res = 0 Then
                    res = setAttribute("input", "name", "j_password", row(4).ToString)
                    If res = 0 Then
                        res = sendClick("input", "name", "uidPasswordLogon", 60)
                        If res = 0 Then
                            If buscaValor("span", "Autentificación de usuario fallida") Then
                                'Error contraseña incorrecta
                                setLog(id_descarga, "E", "Contraseña incorrecta, imposible accesar al portal")
                                res = 2
                            Else
                                'res = Browser("https://bwsext.liverpool.com.mx/sap/bw/BEx?sap-language=es&sap-client=400&accessibility=&style_sheet=http%3A%2F%2Fproveedores.liverpool.com.mx%3A80%2Firj%2Fportalapps%2Fcom.sap.portal.design.portaldesigndata%2Fthemes%2Fportal%2Fcustomer%2FProveedores%2FBIReports30%2FBIReports30_nn7.css%3F6.0.16.0.1&TEMPLATE_ID=BWR_VTAS_POR_DIA_PROV", 60)
                                res = Browser("https://bwsext.liverpool.com.mx/sap/bw/BEx?sap-client=400&sap-language=ES&accessibility=&style_sheet=http%3A%2F%2Fproveedores.liverpool.com.mx%3A80%2Fcom.sap.portal.design.portaldesigndata%2Fthemes%2Fportal%2Fprov_nuevo%2FBIReports30%2FBIReports30_ie6.css%3Fv%3D7.31.11.0.6&sap-tray-type=PLAIN&sap-tray-padding=X&TEMPLATE_ID=BWR_VTAS_POR_DIA_PROV&sapDocumentRenderingMode=EmulateIE8&NavMode=0&NavPathUpdate=false&sap-ie=EmulateIE8", 60)
                                'seleccionando las secciones 
                                If conf(2).ToString() <> "*" Then
                                    'res = setSelect("select", "name", "VAR_OPERATOR_3", "BT", 1, "onchange")
                                    res = setAttribute("input", "name", "VAR_VALUE_EXT_2", conf(3).ToString)
                                    If res = 0 Then
                                        'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_3", conf(2).ToString)
                                        'If res = 0 Then
                                        ' res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_3", conf(3).ToString)
                                        'End If
                                    End If
                                Else
                                    res = setAttribute("input", "name", "VAR_VALUE_EXT_2", "*")
                                End If
                                If res = 0 Then
                                    If conf(4).ToString <> "*" Then
                                        'res = setSelect("select", "name", "VAR_OPERATOR_14", "BT", 1, "onchange")
                                        res = setSelect("select", "name", "VAR_OPERATOR_13", "BT", 1, "onchange")
                                        If res = 0 Then
                                            'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_14", conf(4).ToString)
                                            res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_13", conf(4).ToString)
                                            If res = 0 Then
                                                'res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_14", conf(5).ToString)
                                                res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_13", conf(5).ToString)
                                            End If
                                        End If
                                    Else
                                        'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_14", "")
                                        res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_13", "")
                                    End If
                                    If res = 0 Then
                                        'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_15", fecha.ToString("dd.MM.yyyy"))
                                        res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_14", fecha.ToString("dd.MM.yyyy"))
                                        Dim repeticiones As Integer = 0
                                        While repeticiones < 5 And res = 1
                                            System.Threading.Thread.Sleep("1000")
                                            'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_15", fecha.ToString("dd.MM.yyyy"))
                                            res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_14", fecha.ToString("dd.MM.yyyy"))
                                        End While
                                        If res = 0 Then
                                            'res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_15", fecha.ToString("dd.MM.yyyy"))
                                            res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_14", fecha.ToString("dd.MM.yyyy"))
                                            If res = 0 Then
                                                res = sendClick("a", "href", "javascript:SAPBWBUTTON('PROCESS_VARIABLES','VAR_SUBMIT',' ')", 120)
                                                If res = 0 Then
                                                    res = sendClick("a", "href", "JavaScript:SAPBW(1,'','','EXPAND','0MATERIAL','Y')", 300)
                                                    If res = 0 Then
                                                        Dim url As String = wb.LocationURL.Replace("CMD=", "$")
                                                        Dim url1() As String = url.Split("$")
                                                        Dim url2() As String = url1(0).Split("?")
                                                        Dim params() As String = url2(1).Split("&") 'JG

                                                        If row(3).ToString <> "P00048276" And row(3).ToString <> "P00141522" Then ' JG
                                                            '****************************************
                                                            Dim doc1 As mshtml.HTMLDocument = wb.Document 'JG

                                                            Dim form As mshtml.HTMLFormElement = CType(doc1.forms.item(0), mshtml.HTMLFormElement) 'JG
                                                            Dim REQUEST_NO As String = "REQUEST_NO=2"
                                                            Dim ruta_form As String = form.action 'JG
                                                            Dim param_form() As String = ruta_form.Split("&") 'JG

                                                            Dim parametro As String = params(0) & "&" & params(1) & "&" & param_form(3) &
                                                                           "&CMD=EXPORT&DATA_PROVIDER=DATAPROVIDER_4&FORMAT=CSV&SEPARATOR=,&id=" & Guid.NewGuid.ToString 'JG

                                                            httpResponse = WRequest(url2(0), "POST", parametro, cookiesContainer, 600) 'JG
                                                        Else
                                                            Dim parametro As String = params(0) & "&" & params(1) & "&" &
                                                                           "CMD=EXPORT&DATA_PROVIDER=DATAPROVIDER_4&FORMAT=CSV&SEPARATOR=,&id=" & Guid.NewGuid.ToString 'JG
                                                            httpResponse = WRequest(url2(0), "POST", parametro, cookiesContainer, 600) 'JG
                                                        End If ' JG
                                                        '****************************************
                                                        'httpResponse = WRequest(url2(0), "POST", url2(1) & "CMD=EXPORT&DATA_PROVIDER=DATAPROVIDER_4&FORMAT=CSV&SEPARATOR=,&id=" & Guid.NewGuid.ToString, cookiesContainer, 600) 'JG


                                                        If Not httpResponse Is Nothing Then
                                                            setLog(id_descarga, "I", "Procesando datos descargados...")


                                                            Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                                            Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                                            Dim tmp As String = saveTo
                                                            saveTo = saveTo.Replace(".tmp", ".csv")
                                                            If System.IO.File.Exists(tmp) Then
                                                                System.IO.File.Delete(tmp)
                                                            End If

                                                            Dim lineas As System.IO.StreamReader = New System.IO.StreamReader(o)
                                                            Dim w As IO.StreamWriter = New System.IO.StreamWriter(saveTo)
                                                            setLog(id_descarga, "I", "Archivo temporal generado: " & saveTo)
                                                            Dim ren As Integer = 0
                                                            Try

                                                                While Not lineas.EndOfStream
                                                                    Dim l As String = lineas.ReadLine
                                                                    setLog(id_descarga, "I", l)
                                                                    If Not l.Contains("Resultado") Then
                                                                        l = l.Replace("""", "|")
                                                                        l = l.Replace("|,|", "|")
                                                                        Dim cad() As String = l.Split("|")
                                                                        l = getCadena(cad)
                                                                        If l = "No existen datos adecuados" Then
                                                                            res = 1
                                                                            setLog(id_descarga, "E", "No se encontro informacion para descargar")
                                                                            Exit While
                                                                        End If
                                                                        w.WriteLine(l)
                                                                        If ren = 0 Then
                                                                            If cad(1) <> "Centro" Or cad(10) <> "Ventas $" Then
                                                                                res = 1
                                                                                setLog(id_descarga, "E", "Archivo con errores. Formato desconocido.")
                                                                            End If
                                                                        End If
                                                                        ren = ren + 1
                                                                    End If
                                                                End While
                                                                w.Flush()
                                                                w.Close()
                                                            Catch ex As Exception
                                                                setLog(id_descarga, "E", "Excepcion:" & ex.Message)
                                                                res = 1
                                                                w.Flush()
                                                                w.Close()
                                                            End Try


                                                            If ren > 0 And res = 0 Then

                                                                Try
                                                                    setLog(id_descarga, "I", "Enviando archivos a FTP...")
                                                                    ftp.ServerAddress = server_ftp
                                                                    ftp.ServerPort = port_ftp
                                                                    ftp.UserName = user_ftp
                                                                    ftp.Password = pwd_ftp
                                                                    ftp.Connect()

                                                                    ftp.UploadFile(saveTo, "LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv", False)

                                                                    Dim features As EnterpriseDT.Net.Ftp.FTPReply = ftp.InvokeFTPCommand("cwd /Respaldos")
                                                                    ftp.UploadFile(saveTo, "LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "-" & Now.ToString("ddMMyyyy_HHmmss") & ".csv", False)

                                                                    ftp.Close()

                                                                    setArchivo("2", fecha, "LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv", row.Item(6).ToString() & "\", "VTA")
                                                                    TArchivos.Rows.Add("LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv", row.Item(6).ToString() & "\")
                                                                    setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "LIVERPOOL_VTA_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & ".csv")
                                                                    setLog(id_descarga, "I", "Fin de descarga de Ventas...")
                                                                    Try
                                                                        If System.IO.File.Exists(saveTo) Then
                                                                            System.IO.File.Delete(saveTo)
                                                                        End If
                                                                    Catch ex As Exception
                                                                    End Try
                                                                Catch ex As Exception
                                                                    setLog(id_descarga, "E", "Error:" & ex.Message)
                                                                    res = 1
                                                                End Try
                                                            Else
                                                                res = 1
                                                                setLog(id_descarga, "E", "Informacion no encontrada. Intentar mas tarde.")
                                                            End If
                                                        Else
                                                            setLog(id_descarga, "E", "Error al intentar descarga...")
                                                            res = 1
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

        Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true", 60)
        terminar()
        Return res
    End Function

    Private Function getInventarioLiverpool(ByVal fecha As Date, ByVal row As Data.DataRow, ByVal conf As Data.DataRow) As Integer
        Dim res As Integer = 0
        Dim cont As Integer = 0

        Dim seccion As String = IIf(conf(2).ToString <> "*", "_" & conf(2).ToString, "")
        If getArchivo("2", fecha, "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv") = 1 Then
            setLog(id_descarga, "I", "Archivo: LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv generado anteriormente...")
            Return 0
        End If

        res = iniciar()
        While res = 1 And cont < 10

            System.Threading.Thread.Sleep(1000)
            res = iniciar()
            cont = cont + 1
        End While
        If res = 1 Then
            setLog(id_descarga, "E", message)
            Return 1
        End If
        setLog(id_descarga, "I", "Iniciando descarga de Inventarios...")
        Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true", 60)
        'System.Threading.Thread.Sleep(3000)

        res = validaSitio(row.Item(5))
        If res = 0 Then
            res = Browser(row.Item(5), 60)
            If res = 0 Then
                res = setAttribute("input", "name", "j_user", row(3).ToString)
                If res = 0 Then
                    res = setAttribute("input", "name", "j_password", row(4).ToString)
                    If res = 0 Then
                        res = sendClick("input", "name", "uidPasswordLogon", 60)
                        If res = 0 Then
                            If buscaValor("span", "Autentificación de usuario fallida") Then
                                'Error contraseña incorrecta
                                setLog(id_descarga, "E", "Contraseña incorrecta, imposible accesar al portal")
                                res = 2
                            Else
                                'res = Browser("https://bwsext.liverpool.com.mx/sap/bw/BEx?sap-language=es&sap-client=400&accessibility=&style_sheet=http%3A%2F%2Fproveedores.liverpool.com.mx%3A80%2Firj%2Fportalapps%2Fcom.sap.portal.design.portaldesigndata%2Fthemes%2Fportal%2Fcustomer%2FProveedores%2FBIReports30%2FBIReports30_ie6.css%3F6.0.16.0.1&TEMPLATE_ID=BWR_VTAS_MENS_PROV_UNI_SHIST", 60)
                                res = Browser("https://bwsext.liverpool.com.mx/sap/bw/BEx?sap-client=400&sap-language=ES&accessibility=&style_sheet=http%3A%2F%2Fproveedores.liverpool.com.mx%3A80%2Fcom.sap.portal.design.portaldesigndata%2Fthemes%2Fportal%2Fprov_nuevo%2FBIReports30%2FBIReports30_ie6.css%3Fv%3D7.31.11.0.6&sap-tray-type=PLAIN&sap-tray-padding=X&TEMPLATE_ID=BWR_VTAS_MENS_PROV_UNI_SHIST&sapDocumentRenderingMode=EmulateIE8&NavMode=0&NavPathUpdate=false&buildTree=false&sap-ie=EmulateIE8", 60)
                                'System.Threading.Thread.Sleep(2000)
                                If res = 0 Then
                                    'If conf(2).ToString <> "*" Then
                                    'If conf(3).ToString = "" Then
                                    res = setAttribute("input", "name", "VAR_VALUE_EXT_2", conf(2).ToString)
                                    'Else
                                    'res = setSelect("select", "name", "VAR_OPERATOR_3", "BT", 1, "onchange")
                                    'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_2", conf(2).ToString)
                                    'If res = 0 Then
                                    'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_3", conf(2).ToString)
                                    'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_3", conf(2).ToString)
                                    'If res = 0 Then
                                    'res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_3", conf(3).ToString)
                                    'res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_3", conf(3).ToString)
                                    'End If
                                    'End If
                                    'End If
                                    'Else
                                    'res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_2", "*")
                                    'End If
                                    If res = 0 Then
                                        If conf(4).ToString <> "*" Then
                                            If conf(5).ToString = "" Then
                                                res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_13", conf(4).ToString)
                                            Else
                                                res = setSelect("select", "name", "VAR_OPERATOR_13", "BT", 1, "onchange")
                                                If res = 0 Then
                                                    res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_13", conf(4).ToString)
                                                    If res = 0 Then
                                                        res = setAttribute("input", "name", "VAR_VALUE_HIGH_EXT_13", conf(5).ToString)
                                                    End If
                                                End If
                                            End If
                                        Else
                                            res = setAttribute("input", "name", "VAR_VALUE_LOW_EXT_13", "")
                                        End If

                                        If res = 0 Then
                                            res = sendClick("a", "href", "Javascript:SAPBWBUTTON('PROCESS_VARIABLES','VAR_SUBMIT',' ')", 60)
                                            If res = 0 Then
                                                res = sendClick("a", "href", "JavaScript:SAPBW(1,'','','EXPAND','0MATERIAL','Y')", 500)
                                                If res = 0 Then
                                                    Dim url As String = wb.LocationURL.Replace("CMD=", "$")
                                                    Dim url1() As String = url.Split("$")
                                                    Dim url2() As String = url1(0).Split("?")
                                                    Dim params() As String = url2(1).Split("&")

                                                    Dim doc1 As mshtml.HTMLDocument = wb.Document

                                                    Dim form As mshtml.HTMLFormElement = CType(doc1.forms.item(0), mshtml.HTMLFormElement)
                                                    Dim request_no As String = "REQUEST_NO=2"
                                                    If Not form Is Nothing Then
                                                        Dim ruta_form As String = form.action
                                                        Dim param_form() As String = ruta_form.Split("&")
                                                        request_no = param_form(3)
                                                    End If

                                                    'If row(3).ToString = "P00048276" Then
                                                    'httpResponse = WRequest("https://bwsext.liverpool.com.mx/sap/bw/BEx", "POST", "SAP-LANGUAGE=ES&PAGENO=1&REQUEST_NO=7&CMD=EXPORT&DATA_PROVIDER=DATAPROVIDER_4&FORMAT=CSV&SEPARATOR=,&NAME=PERRITO.CSV&id=" & Guid.NewGuid.ToString, cookiesContainer, 600)
                                                    'Else
                                                    Dim parametro As String = params(0) & "&" & params(1) & "&" & request_no &
                                                    "&CMD=EXPORT&DATA_PROVIDER=DATAPROVIDER_4&FORMAT=CSV&SEPARATOR=,&NAME=PERRITO.CSV&id=" & Guid.NewGuid.ToString

                                                    httpResponse = WRequest(url2(0), "POST", parametro, cookiesContainer, 600)
                                                    'End If


                                                    If Not httpResponse Is Nothing Then
                                                        setLog(id_descarga, "I", "Procesando datos descargados...")

                                                        Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                                        Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                                        Dim tmp As String = saveTo
                                                        saveTo = saveTo.Replace(".tmp", ".csv")

                                                        If System.IO.File.Exists(tmp) Then
                                                            System.IO.File.Delete(tmp)
                                                        End If


                                                        Dim lineas As System.IO.StreamReader = New System.IO.StreamReader(o)
                                                        Dim w As IO.StreamWriter = New System.IO.StreamWriter(saveTo)
                                                        setLog(id_descarga, "I", "Archivo temporal generado: " & saveTo)
                                                        Dim ren As Integer = 0
                                                        Try

                                                            While Not lineas.EndOfStream
                                                                Dim l As String = lineas.ReadLine
                                                                If Not l.Contains("Resultado") Then

                                                                    l = l.Replace("""", "|")
                                                                    l = l.Replace("|,|", "|")
                                                                    Dim cad() As String = l.Split("|")
                                                                    l = getCadena(cad)
                                                                    w.WriteLine(l)
                                                                    If ren = 0 Then
                                                                        If cad(1) <> "Centro" Or cad(17) <> "In Transfer" Then
                                                                            res = 1
                                                                            setLog(id_descarga, "E", "Archivo con errores. Formato desconocido.")
                                                                        End If
                                                                    End If
                                                                    ren = ren + 1
                                                                End If
                                                            End While
                                                            w.Flush()
                                                            w.Close()
                                                        Catch ex As Exception
                                                            setLog(id_descarga, "E", "Excepcion:" & ex.Message)
                                                            res = 1
                                                            w.Flush()
                                                            w.Close()
                                                        End Try


                                                        If ren > 1 And res = 0 Then
                                                            'Dim seccion As String = IIf(conf(2).ToString <> "*", "_" & conf(2).ToString, "")
                                                            'If Not System.IO.Directory.Exists(row.Item(6).ToString()) Then
                                                            '    My.Computer.FileSystem.CreateDirectory(row.Item(6).ToString())
                                                            'End If
                                                            'If My.Computer.FileSystem.FileExists(row.Item(6).ToString() & "\" & "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv") Then
                                                            '    My.Computer.FileSystem.DeleteFile(row.Item(6).ToString() & "\" & "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv")
                                                            'End If
                                                            'My.Computer.FileSystem.CopyFile(saveTo, row.Item(6).ToString() & "\" & "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv")
                                                            'TArchivos.Rows.Add("LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv", row.Item(6).ToString() & "\")
                                                            'setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv")
                                                            'setLog(id_descarga, "I", "Fin de descarga de Inventarios...")

                                                            Try
                                                                setLog(id_descarga, "I", "Enviando archivos a FTP...")
                                                                ftp.ServerAddress = server_ftp
                                                                ftp.ServerPort = port_ftp
                                                                ftp.UserName = user_ftp
                                                                ftp.Password = pwd_ftp
                                                                ftp.Connect()

                                                                ftp.UploadFile(saveTo, "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv", False)

                                                                Dim features As EnterpriseDT.Net.Ftp.FTPReply = ftp.InvokeFTPCommand("cwd /Respaldos")
                                                                ftp.UploadFile(saveTo, "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & "-" & Now.ToString("ddMMyyyy_HHmmss") & ".csv", False)

                                                                ftp.Close()
                                                                'System.IO.File.Move(lista_archivos.Rows(0).Item(1).ToString, row.Item(6).ToString() & "\" & "SEARS_" & row(3) & "_VTA_" & fecha.ToString("ddMMyyyy") & ".xml")
                                                                setArchivo("2", fecha, "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv", row.Item(6).ToString() & "\", "INV")
                                                                TArchivos.Rows.Add("LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv", row.Item(6).ToString() & "\")
                                                                setLog(id_descarga, "I", "Archivo creado: " & row.Item(6).ToString() & "\" & "LIVERPOOL_INV_" & fecha.ToString("ddMMyyyy") & "_" & row(3) & seccion & "_" & conf(6).ToString & ".csv")
                                                                setLog(id_descarga, "I", "Fin de descarga de Inventarios...")
                                                                Try
                                                                    If System.IO.File.Exists(saveTo) Then
                                                                        System.IO.File.Delete(saveTo)
                                                                    End If
                                                                Catch ex As Exception

                                                                End Try
                                                            Catch ex As Exception
                                                                setLog(id_descarga, "E", "Error:" & ex.Message)
                                                                res = 1
                                                            End Try
                                                        Else
                                                            res = 1
                                                            setLog(id_descarga, "E", "Informacion no disponible. Intentar mas tarde.")
                                                        End If
                                                    Else
                                                        setLog(id_descarga, "E", "Error al intentar descarga...")
                                                        res = 1
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
        Browser("https://proveedores.liverpool.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true", 60)
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

    Dim CadenaProveedor As String = String.Empty

    Public Function getVentasPH(ByVal fecha As Date, ByVal pstrTipoDocumento As String) As Integer
        Dim res As Integer = 0
        Dim cont As Integer = 0
        Dim intResultado As Integer

        res = iniciar()

        While res = 1 And cont < 10
            System.Threading.Thread.Sleep(1000)
            res = iniciar()
            cont = cont + 1
        End While
        If res = 1 Then
            setLog(id_descarga, "E", message)
            Return 1
        End If

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
        Dim saveTo1x As String

        Dim strProveedor As String = String.Empty
        Dim sqlCmd As New SqlCommand()

        conn.Open()
        sqlrdr = sqlCmdLogin.ExecuteReader
        dtLogin.Load(sqlrdr)
        sqlrdr.Close()
        conn.Close()


        Dim wait_time As Integer = 10000
        Dim loadcontrol_time As Integer = 3000

        strSQL = "SELECT  LOADPAGE_PH,LOADCONTROL_PH FROM CONFIG"
        sqlCmd.CommandText = strSQL
        sqlCmd.Connection = conn
        conn.Open()
        sqlrdr = sqlCmd.ExecuteReader
        If sqlrdr.HasRows Then
            While sqlrdr.Read
                wait_time = sqlrdr.Item(0)
                loadcontrol_time = sqlrdr.Item(1)
            End While
        End If
        sqlrdr.Close()
        conn.Close()
        sqlCmd.Dispose()
        sqlCmd = Nothing

        sqlCmd = New SqlCommand()

        Dim strMensaje As String = ""
        Try
            For Each drLog As DataRow In dtLogin.Rows
                strSQL = "SELECT cProSecuencia, cProTipo, cProAtributo, cProNombreAtributo, cProValorAtributo, cProAccion FROM cProcesos where cProHabilitar=1 AND cProSociedad LIKE '%" & drLog.Item("cLogEmpresa") & "%' and cProProceso='" & pstrTipoDocumento & "' order by cProEmpresa, cProSociedad, cProProceso, cProSecuencia"

                sqlCmd.CommandText = strSQL
                sqlCmd.Connection = conn

                conn.Open()
                sqlrdr = sqlCmd.ExecuteReader
                dtProceso.Rows.Clear()
                dtProceso.Load(sqlrdr)
                sqlrdr.Close()
                conn.Close()

                setLog(id_descarga, "I", "Iniciando proceso para " & drLog.Item("cLogEmpresa") & "...")
                setLog(id_descarga, "I", "Tipo de descarga solicitada: " & pstrTipoDocumento)

                strProveedor = drLog.Item("cLogUsuario")
                CadenaProveedor = strProveedor
                If getArchivo("3", fecha, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls") = 1 Then
                    setLog(id_descarga, "I", "Archivo PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls generado anteriormente...")

                    strSQL = "UPDATE EstatusPH SET estatus='TERMINADO' , Resultado='Archivo PH_" & IIf(pstrTipoDocumento = "Ventas", "V", "I") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls generado anteriormente...' WHERE fecha='" & fecha & "' AND TipoArchivo='" & pstrTipoDocumento & "' and estatus ='EN_PROCESO' "
                    Dim sqlcmd2 As New SqlCommand(strSQL, conn)
                    conn.Open()
                    sqlcmd2.ExecuteNonQuery()
                    conn.Close()


                    intResultado = 0
                Else

                    Browser("http://proveedores.palaciohierro.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true", 60)

                    If Me.Browser(drLog.Item("cLogUrl"), 60) = 1 Then
                        setLog(id_descarga, "E", "Error: Cargar Pagina " & drLog.Item("cLogUrl"))
                        intResultado = 1
                        Return intResultado
                        Exit Function
                    End If

                    'System.Threading.Thread.Sleep(15000) 'JG
                    While wb.Busy 'JG
                        System.Threading.Thread.Sleep(1000)
                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                    End While 'JG

                    strProveedor = drLog.Item("cLogUsuario")
                    If Me.setAttribute("input", "name", "j_user", drLog.Item("cLogUsuario")) = 1 Then
                        intResultado = 1
                        Return intResultado
                        Exit Function
                    End If

                    If Me.setAttribute("input", "name", "j_password", drLog.Item("cLogContrasenia")) = 1 Then
                        intResultado = 1
                        Return intResultado
                        Exit Function
                    End If

                    If sendClick("input", "name", "uidPasswordLogon", 60) = 1 Then
                        intResultado = 1
                        Return intResultado
                        Exit Function
                    End If


                    While wb.Busy = True
                        ' System.Threading.Thread.Sleep(wait_time) ' JG
                        System.Threading.Thread.Sleep(1000)
                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                    End While
                    System.Threading.Thread.Sleep(1000)

                    If Me.Browser("https://wdbop.palaciohierro.com.mx/BOE/OpenDocument/opendoc/openDocument.jsp?iDocID=FtkAyE9eHQMAfCgAAEB5hEQBAFBWvQBx&sIDType=CUID", 120) Then
                        setLog(id_descarga, "E", "Error: Cargar Pagina https://wdbop.palaciohierro.com.mx/BOE/OpenDocument/opendoc/openDocument.jsp?iDocID=FtkAyE9eHQMAfCgAAEB5hEQBAFBWvQBx&sIDType=CUID")

                        intResultado = 1
                        Return intResultado
                        Exit Function
                    End If

                    If document.body.innerHTML.Contains("Service Temporarily Unavailable") Then
                        setLog(id_descarga, "E", "Service Temporarily Unavailable. Portal fuera de linea.")
                        intResultado = 1
                    End If

                    Dim htmlDocParent As mshtml.HTMLDocument = document
                    Dim htmlDoc_openDocChildFrame As mshtml.HTMLDocument
                    Dim htmlDoc_SaveAsDlg As mshtml.HTMLDocument
                    Dim htmlDoc3 As mshtml.HTMLDocument
                    Dim htmlDoc4 As mshtml.HTMLDocument
                    Dim htmlDoc5 As mshtml.HTMLDocument

                    Dim htmlWnd As mshtml.IHTMLWindow2 = Nothing
                    Dim htmlWnd2 As mshtml.IHTMLWindow2 = Nothing
                    Dim htmlWnd3 As mshtml.IHTMLWindow2 = Nothing
                    Dim htmlWnd4 As mshtml.IHTMLWindow2 = Nothing
                    Dim frames As mshtml.FramesCollection


                    Dim htmlElement As mshtml.IHTMLElement

                    While wb.Busy 'JG
                        System.Threading.Thread.Sleep(1000)
                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                    End While 'JG
                    System.Threading.Thread.Sleep(5000)

                    frames = htmlDocParent.frames

                    For n As Integer = 0 To frames.length - 1
                        htmlWnd = CType(frames.item(n), mshtml.IHTMLWindow2)
                        If htmlWnd.name = "openDocChildFrame" Then
                            htmlDoc_openDocChildFrame = CType(htmlWnd.document, mshtml.HTMLDocument)

                            htmlDoc3 = CType(htmlDoc_openDocChildFrame.body.document, mshtml.HTMLDocument)
                            htmlWnd3 = CType(htmlDoc3.frames.item(0), mshtml.IHTMLWindow2)
                            htmlDoc4 = CType(htmlWnd3.document, mshtml.IHTMLDocument)

                            htmlElement = htmlDoc4.getElementById("RealBtn_CANCEL_BTN_promptsDlg")

                            htmlElement.click()
                            setLog(id_descarga, "I", "Click: RealBtn_CANCEL_BTN_promptsDlg")
                            While wb.Busy 'JG
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While 'JG
                            System.Threading.Thread.Sleep(1000)

                            htmlElement = htmlDoc4.getElementById("IconImg__dhtmlLib_240")

                            htmlElement.click()
                            setLog(id_descarga, "I", "Click: IconImg__dhtmlLib_240")
                            System.Threading.Thread.Sleep(1000)

                            Exit For
                        End If
                    Next


                    'While wb.Busy = True
                    '    System.Threading.Thread.Sleep(wait_time)
                    'End While
                    'System.Threading.Thread.Sleep(3000)
                    While wb.Busy 'JG
                        System.Threading.Thread.Sleep(1000)
                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                    End While 'JG
                    System.Threading.Thread.Sleep(2000)

                    For n As Integer = 0 To htmlDocParent.frames.length - 1
                        htmlWnd = htmlDocParent.frames.item(n)
                        If htmlWnd.name = "openDocChildFrame" Then
                            htmlDoc_openDocChildFrame = CType(htmlWnd.document, mshtml.HTMLDocument)


                            htmlDoc3 = CType(htmlDoc_openDocChildFrame.body.document, mshtml.HTMLDocument)
                            htmlWnd3 = CType(htmlDoc3.frames.item(0), mshtml.IHTMLWindow2)
                            htmlDoc4 = CType(htmlWnd3.document, mshtml.IHTMLDocument)

                            For n2 As Integer = 0 To htmlDoc4.frames.length - 1
                                htmlWnd4 = htmlDoc4.frames.item(n2)
                                If htmlWnd4.name = "saveAsDlg" Then
                                    htmlDoc_SaveAsDlg = CType(htmlWnd4.document, mshtml.HTMLDocument)
                                    htmlDoc5 = CType(htmlDoc_SaveAsDlg.body.document, mshtml.HTMLDocument)


                                    htmlElement = htmlDoc5.getElementById("yui-gen0-1-label")
                                    htmlElement.click()
                                    setLog(id_descarga, "I", "Click: yui-gen0-1-label")

                                    'System.Threading.Thread.Sleep(wait_time)
                                    While wb.Busy 'JG
                                        System.Threading.Thread.Sleep(1000)
                                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                                    End While 'JG
                                    System.Threading.Thread.Sleep(1000)

                                    htmlElement = htmlDoc5.getElementById("accordionNavigationView_drawer0_treeView_treeNode1_name")
                                    htmlElement.click()
                                    setLog(id_descarga, "I", "Click: accordionNavigationView_drawer0_treeView_treeNode1_name")

                                    While wb.Busy 'JG
                                        System.Threading.Thread.Sleep(1000)
                                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                                    End While 'JG
                                    System.Threading.Thread.Sleep(1000)

                                    'htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox1")
                                    'htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox8")
                                    Dim id As String = "openDocuments_DocumentSelector_detailView_checkbox8"
                                    htmlElement = htmlDoc5.getElementById(id)

                                    htmlElement.click()
                                    setLog(id_descarga, "I", "Click: " & id)

                                    Dim tempElem 'As mshtml.HTMLInputTextElement
                                    tempElem = DirectCast(htmlElement, mshtml.HTMLInputElement)

                                    Dim dummy As Object = Nothing
                                    tempElem.FireEvent("onclick", dummy)

                                    While wb.Busy 'JG
                                        System.Threading.Thread.Sleep(1000)
                                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                                    End While 'JG
                                    System.Threading.Thread.Sleep(1000)

                                    htmlElement = htmlDoc5.getElementById("openDocuments_OpenButton")

                                    htmlElement.click()
                                    setLog(id_descarga, "I", "Click: openDocuments_OpenButton")
                                    System.Threading.Thread.Sleep(1000)
                                    Exit For
                                End If
                            Next

                        End If
                    Next

                    While wb.Busy 'JG
                        System.Threading.Thread.Sleep(1000)
                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                    End While 'JG
                    System.Threading.Thread.Sleep(2000)

                    'System.Threading.Thread.Sleep(3000)

                    For n As Integer = 0 To htmlDocParent.frames.length - 1
                        htmlWnd = htmlDocParent.frames.item(n)
                        If htmlWnd.name = "openDocChildFrame" Then
                            htmlDoc_openDocChildFrame = CType(htmlWnd.document, mshtml.HTMLDocument)

                            htmlDoc3 = CType(htmlDoc_openDocChildFrame.body.document, mshtml.HTMLDocument)
                            htmlWnd3 = CType(htmlDoc3.frames.item(0), mshtml.IHTMLWindow2)
                            htmlDoc4 = CType(htmlWnd3.document, mshtml.IHTMLDocument)

                            'htmlElement = htmlDoc4.getElementById("_CWpromptstrLstElt0")
                            htmlElement = htmlDoc4.getElementById("_CWpromptstrLstElt_TWe_0")
                            htmlElement.click()

                            While wb.Busy 'JG
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While 'JG
                            System.Threading.Thread.Sleep(1000)

                            htmlElement = htmlDoc4.getElementById("text_promptLovZone_RightZone_oneTextField_date0")
                            'htmlElement.setAttribute("value",  "02/10/2013 12:00:00 a.m.", 1)
                            'htmlElement.setAttribute("value", fecha.ToString("dd/MM/yyyy") & " 12:00:00 AM", 1)
                            htmlElement.setAttribute("value", fecha.ToString("dd/MM/yyyy") & " 12:00:00 AM", 1)


                            Dim tempElem As mshtml.HTMLInputTextElement
                            tempElem = DirectCast(htmlElement, mshtml.HTMLInputTextElement)

                            Dim dummy As Object = Nothing
                            Dim evento As String = "onchange"
                            tempElem.FireEvent(evento, dummy)

                            While wb.Busy 'JG
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While 'JG
                            System.Threading.Thread.Sleep(1000)

                            htmlElement = htmlDoc4.getElementById("_CWpromptstrLstElt_TWe_1")
                            htmlElement.click()
                            setLog(id_descarga, "I", "Click: _CWpromptstrLstElt_TWe_1")

                            While wb.Busy 'JG
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While 'JG
                            System.Threading.Thread.Sleep(1000)

                            htmlElement = htmlDoc4.getElementById("text_promptLovZone_RightZone_oneTextField_date0")
                            'htmlElement.setAttribute("value", "03/10/2013 12:00:00 a.m.", 1)
                            'htmlElement.setAttribute("value", fecha.ToString("dd/MM/yyyy") & " 12:00:00 AM", 1)
                            htmlElement.setAttribute("value", fecha.ToString("dd/MM/yyyy") & " 12:00:00 AM", 1)


                            tempElem = DirectCast(htmlElement, mshtml.HTMLInputTextElement)
                            tempElem.FireEvent("onchange", dummy)

                            While wb.Busy 'JG
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While 'JG
                            System.Threading.Thread.Sleep(1000)

                            htmlElement = htmlDoc4.getElementById("RealBtn_OK_BTN_promptsDlg")

                            htmlElement.click()
                            setLog(id_descarga, "I", "Click: RealBtn_OK_BTN_promptsDlg")

                            While wb.Busy 'JG
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While 'JG
                            System.Threading.Thread.Sleep(1000)

                            Exit For

                        End If
                    Next

                    While wb.Busy 'JG
                        System.Threading.Thread.Sleep(1000)
                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                    End While 'JG
                    System.Threading.Thread.Sleep(4000)

                    Dim htmlElement_Inicio As mshtml.IHTMLElement
                    Dim htmlElement_Atras As mshtml.IHTMLElement
                    Dim htmlElement_Texto As mshtml.IHTMLElement
                    Dim htmlElement_Siguiente As mshtml.IHTMLElement
                    Dim htmlElement_Fin As mshtml.IHTMLElement
                    Dim htmlElement_Contenido As mshtml.IHTMLElement
                    Dim maxPaginas As Integer
                    Dim minPaginas As Integer
                    Dim strTexto As String

                    'System.Threading.Thread.Sleep(wait_time)
                    For n As Integer = 0 To htmlDocParent.frames.length - 1
                        htmlWnd = htmlDocParent.frames.item(n)
                        If htmlWnd.name = "openDocChildFrame" Then
                            htmlDoc_openDocChildFrame = CType(htmlWnd.document, mshtml.HTMLDocument)

                            'htmlElement = htmlDoc_openDocChildFrame.getElementById("")


                            Dim doc1x As mshtml.IHTMLWindow2 = CType(htmlDoc_openDocChildFrame.frames.item(0), mshtml.IHTMLWindow2)

                            Dim url1 As String = htmlDoc_openDocChildFrame.frames.item(0).location.href
                            Dim params1() As String = url1.Split("&")

                            Dim doc1Menu As mshtml.HTMLDocument = doc1x.document

                            doc1x = CType(doc1Menu.frames.item(3), mshtml.HTMLWindow2) 'sacando el frame con el texto
                            While doc1x.document.readyState <> "complete"
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While

                            While doc1Menu.readyState <> "complete"
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While

                            'Inicio IconImg__dhtmlLib_184
                            'Atras IconImg__dhtmlLib_185
                            'Texto IconImg_Txt__dhtmlLib_186
                            'Siguiente IconImg__dhtmlLib_188
                            'Fin IconImg__dhtmlLib_189

                            htmlElement_Inicio = doc1Menu.getElementById("IconImg__dhtmlLib_186")
                            htmlElement_Atras = doc1Menu.getElementById("IconImg__dhtmlLib_187")
                            htmlElement_Siguiente = doc1Menu.getElementById("IconImg__dhtmlLib_190")
                            htmlElement_Fin = doc1Menu.getElementById("IconImg__dhtmlLib_191")


                            htmlElement_Fin.click()
                            setLog(id_descarga, "I", "Click: htmlElement_Fin")
                            While wb.Busy 'JG
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While 'JG
                            System.Threading.Thread.Sleep(2000)
                            While doc1x.document.readyState <> "complete"
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While

                            htmlElement_Inicio.click()
                            setLog(id_descarga, "I", "Click: htmlElement_Inicio")

                            While wb.Busy 'JG
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While 'JG
                            System.Threading.Thread.Sleep(2000)
                            While doc1x.document.readyState <> "complete"
                                System.Threading.Thread.Sleep(1000)
                                setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                            End While

                            Dim url2 As String = doc1Menu.frames.item(3).location.href
                            Dim params2() As String = url2.Split("&")


                            Dim p1 As Integer = 0
                            maxPaginas = 0
                            While p1 < params2.Length
                                If params2(p1).Contains("nbPage") Then
                                    Dim val() As String = params2(p1).Split("=")
                                    maxPaginas = val(1)
                                End If
                                p1 = p1 + 1
                            End While

                            If maxPaginas = 0 Then
                                setLog(id_descarga, "E", "No se pudo determinar el numero de paginas")
                                intResultado = 1
                            End If

                            If maxPaginas > 0 Then

                                setLog(id_descarga, "I", "No Paginas:" & maxPaginas)
                                url2 = url2.Replace("iPage=first&", "")

                                Dim doc1Contenido = doc1x.document

                                saveTo1x = My.Computer.FileSystem.GetTempFileName.ToString
                                Dim tmp As String = saveTo1x

                                saveTo1x = saveTo1x.Replace(".tmp", ".html")
                                If System.IO.File.Exists(tmp) Then
                                    System.IO.File.Delete(tmp)
                                End If
                                Dim w As New System.IO.StreamWriter(saveTo1x, False, System.Text.Encoding.UTF8)
                                setLog(id_descarga, "I", "Archivo tempora:" & saveTo1x)

                                minPaginas = 1
                                While minPaginas <= maxPaginas

                                    setLog(id_descarga, "I", "Leyendo pagina:" & minPaginas)
                                    'doc1Contenido = doc1x.document

                                    w.WriteLine(doc1x.document.body.innerHTML)
                                    setLog(id_descarga, "I", "Escribiendo datos en archivo temporal")

                                    htmlElement_Siguiente.click()
                                    setLog(id_descarga, "I", "Click: htmlElement_Siguiente")
                                    ' System.Threading.Thread.Sleep(3000) 'JG
                                    While wb.Busy 'JG
                                        System.Threading.Thread.Sleep(1000)
                                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                                    End While 'JG
                                    System.Threading.Thread.Sleep(1000)

                                    While doc1x.document.readyState <> "complete"
                                        System.Threading.Thread.Sleep(1000)
                                        setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
                                    End While

                                    minPaginas = minPaginas + 1
                                End While

                                w.Flush()
                                w.Close()
                                setLog(id_descarga, "I", "Guardando archivos temporales...")

                                Dim excel As New Excel.Application
                                Dim sts_archivo As Boolean = True
                                'Dim book As Excel.Workbook = excel.Workbooks.Open("c:\Users\pablo_ovando\AppData\Local\Temp\tmp957A.xls")
                                Try
                                    Dim book As Excel.Workbook = excel.Workbooks.Open(saveTo1x)
                                    excel.Visible = False

                                    If book.Sheets.Count > 0 Then
                                        Dim sheet As Worksheet = book.Sheets(1)
                                        sheet.Name = "Depto_Sku_tienda"
                                        sheet.Columns(1).insert()
                                    End If

                                    tmp = saveTo1x
                                    saveTo1x = saveTo1x.Replace(".html", ".xls")
                                    book.SaveAs(saveTo1x, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal)
                                    setLog(id_descarga, "I", "Guardando archivos excel:" & saveTo1x)

                                    book.Close()
                                    excel.Quit()
                                    If System.IO.File.Exists(tmp) Then
                                        System.IO.File.Delete(tmp)
                                    End If
                                    book = Nothing
                                    excel = Nothing
                                Catch ex As Exception
                                    sts_archivo = False

                                End Try

                                ' Modificación: - 100914 LVH '''''''''''''''''''''''''''''''''''''''''''
                                ' Validar si tiene registros el archivo descargado '''''''''''''''''''''
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                If sts_archivo Then
                                    setLog(id_descarga, "I", "Verificando contenido archivo de Excel.")
                                    Dim excel2 As New Excel.Application
                                    Dim contador As Integer = 1
                                    Dim dTotal As Double = 0.0
                                    Try
                                        Dim book2 As Excel.Workbook = excel2.Workbooks.Open(saveTo1x) '"c:\TEMP\PH_VTA_09092014_PR0000184366.xls")
                                        excel2.Visible = False
                                        excel2.DisplayAlerts = False

                                        If book2.Sheets.Count > 0 Then
                                            Dim sheet2 As Worksheet = book2.Sheets(1)
                                            ' Revisar sólo 200 líneas para validar si tiene o no información
                                            While contador < 200 'sheet.Rows.Count
                                                Dim cel2 = sheet2.Cells(contador, 3)
                                                If cel2.value Is Nothing Then
                                                    sts_archivo = False
                                                Else
                                                    If cel2.value.ToString.Trim = "Total Marca:" Then
                                                        ' validar cantidad 13,14
                                                        cel2 = sheet2.Cells(contador, 14)
                                                        If cel2.value Is Nothing Then
                                                            sts_archivo = False
                                                        Else
                                                            If cel2.value.ToString.Trim = "" Then
                                                                sts_archivo = False
                                                            Else
                                                                Try
                                                                    dTotal += Convert.ToDouble(cel2.value.ToString.Trim)
                                                                Catch ex As Exception
                                                                    Continue While
                                                                End Try

                                                                sts_archivo = True
                                                            End If
                                                        End If
                                                    Else
                                                        sts_archivo = False
                                                    End If
                                                End If

                                                contador = contador + 1
                                            End While
                                        End If

                                        book2.Close()
                                        excel2.Quit()
                                        book2 = Nothing
                                        excel2 = Nothing

                                        If dTotal > 0 Then
                                            sts_archivo = True
                                            setLog(id_descarga, "I", "Archivo de Excel correcto con registros.")
                                        Else
                                            sts_archivo = False
                                            setLog(id_descarga, "E", "Error archivo de Excel sin Registros...")
                                            ' Depurar el archivo de Excel creado
                                            If System.IO.File.Exists(saveTo1x) Then
                                                System.IO.File.Delete(saveTo1x)
                                            End If
                                            setLog(id_descarga, "E", "Archivo de Excel eliminado...")
                                        End If
                                    Catch ex As Exception
                                        sts_archivo = False
                                        setLog(id_descarga, "E", "Error archivo de Excel sin Registros...")
                                    End Try
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                If sts_archivo Then
                                    setLog(id_descarga, "I", "Enviando archivos a FTP...")
                                    ftp.ServerAddress = server_ftp
                                    ftp.ServerPort = port_ftp
                                    ftp.UserName = user_ftp
                                    ftp.Password = pwd_ftp
                                    ftp.Connect()

                                    Dim strArchivoFtp As String = "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls"


                                    'ftp.UploadFile(saveTo1x, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", False)
                                    ftp.UploadFile(saveTo1x, strArchivoFtp, False)

                                    Dim features As EnterpriseDT.Net.Ftp.FTPReply = ftp.InvokeFTPCommand("cwd /Respaldos")
                                    'ftp.UploadFile(saveTo1x, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & "-" & Now.ToString("ddMMyyyy_HHmmss") & ".xls", False)
                                    ftp.UploadFile(saveTo1x, strArchivoFtp, False)
                                    ftp.Close()

                                    setArchivo("3", fecha, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", drLog("cLogRutaDescarga") & "\", IIf(pstrTipoDocumento = "Ventas", "VTA", "INV"))

                                    'setArchivo("3", fecha, strArchivo & "\", IIf(pstrTipoDocumento = "Ventas", "VTA", "INV"))

                                    TArchivos.Rows.Add("PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", drLog("cLogRutaDescarga") & "\")
                                    setLog(id_descarga, "I", "Archivo creado: " & "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls")
                                    setLog(id_descarga, "I", "Fin de descarga de archivos...")
                                    intResultado = 0

                                    If System.IO.File.Exists(saveTo1x) Then
                                        Try
                                            System.IO.File.Delete(saveTo1x)
                                        Catch ex As Exception
                                        End Try
                                    End If
                                    saveTo1x = saveTo1x.Replace(".xls", ".html")
                                    If System.IO.File.Exists(saveTo1x) Then
                                        Try
                                            System.IO.File.Delete(saveTo1x)
                                        Catch ex As Exception
                                        End Try
                                    End If

                                Else
                                    intResultado = 1
                                    setLog(id_descarga, "E", "Error en archivo. Formato desconocido o vacio...")
                                End If




                            End If


                            Exit For
                        End If
                    Next

                    If intResultado = 1 Then


                        Exit Try
                    End If


                End If
            Next
        Catch ex As Exception
            setLog(id_descarga, "E", "Error:" & ex.Message & " " & strMensaje)
            intResultado = 1
            'Throw
        End Try
        terminar()
        Return intResultado
    End Function

    Public Function getInventariosPH(ByVal fecha As Date, ByVal pstrTipoDocumento As String, ByVal reporte As Inv_PH) As Integer
        conn.Open()
        Dim sqlCmdLogin As New SqlCommand("SELECT * FROM cLogin where cLogPortal='PH'", conn)
        Dim sqlrdr = sqlCmdLogin.ExecuteReader
        Dim dtLogin As New System.Data.DataTable
        dtLogin.Load(sqlrdr)

        Dim drLogin As DataRow = dtLogin.Rows(0)
        Dim x1 As New PortalPhRobot.AwebExplorer()
        Dim tasks As Task(Of List(Of SalesVsInventories)) = x1.DownloadAsync(drLogin.Item("cLogUsuario"), drLogin.Item("cLogContrasenia"))

        sqlrdr.Close()
        conn.Close()
        Return 0
    End Function

    'Public Function getInventariosPH(ByVal fecha As Date, ByVal pstrTipoDocumento As String, ByVal reporte As Inv_PH) As Integer
    '    Dim res As Integer = 0
    '    Dim cont As Integer = 0
    '    Dim wait_time As Integer = 10000
    '    Dim loadcontrol_time As Integer = 3000
    '    res = iniciar()
    '    While res = 1 And cont < 10

    '        System.Threading.Thread.Sleep(1000)
    '        res = iniciar()
    '        cont = cont + 1
    '    End While
    '    If res = 1 Then
    '        setLog(id_descarga, "E", message)
    '        Return 1
    '    End If

    '    Dim dtProceso As New System.Data.DataTable
    '    Dim dtLogin As New System.Data.DataTable
    '    Dim sqlrdr As SqlDataReader
    '    Dim oExcel As Application
    '    Dim oBook As Workbook
    '    Dim strSQL As String = "SELECT * FROM cLogin where cLogPortal='PH'"
    '    Dim sqlCmdLogin As New SqlCommand(strSQL, conn)
    '    Dim iHTMLCol As mshtml.IHTMLElementCollection
    '    Dim iHTMLEle As mshtml.IHTMLElement
    '    Dim blnProcesoCorrecto = False
    '    Dim intResultado As Integer = 1
    '    Dim strProveedor As String = String.Empty
    '    Dim sqlCmd As New SqlCommand()
    '    Dim cx As Integer = 0

    '    conn.Open()
    '    sqlrdr = sqlCmdLogin.ExecuteReader
    '    dtLogin.Load(sqlrdr)
    '    sqlrdr.Close()
    '    conn.Close()


    '    strSQL = "SELECT  LOADPAGE_PH,LOADCONTROL_PH FROM CONFIG"
    '    sqlCmd.CommandText = strSQL
    '    sqlCmd.Connection = conn
    '    conn.Open()
    '    sqlrdr = sqlCmd.ExecuteReader
    '    If sqlrdr.HasRows Then
    '        While sqlrdr.Read
    '            wait_time = sqlrdr.Item(0)
    '            loadcontrol_time = sqlrdr.Item(1)
    '        End While
    '    End If
    '    sqlrdr.Close()
    '    conn.Close()
    '    sqlCmd.Dispose()
    '    sqlCmd = Nothing

    '    sqlCmd = New SqlCommand()


    '    Dim strMensaje As String = ""
    '    Try
    '        For Each drLog As DataRow In dtLogin.Rows

    '            strSQL = "SELECT cProSecuencia, cProTipo, cProAtributo, cProNombreAtributo, cProValorAtributo, cProAccion FROM cProcesos where cProHabilitar=1 AND cProSociedad LIKE '%" & drLog.Item("cLogEmpresa") & "%' and cProProceso='" & pstrTipoDocumento & "' order by cProEmpresa, cProSociedad, cProProceso, cProSecuencia"

    '            sqlCmd.CommandText = strSQL
    '            sqlCmd.Connection = conn

    '            conn.Open()
    '            sqlrdr = sqlCmd.ExecuteReader
    '            dtProceso.Rows.Clear()
    '            dtProceso.Load(sqlrdr)
    '            sqlrdr.Close()
    '            conn.Close()

    '            setLog(id_descarga, "I", "Iniciando proceso para " & drLog.Item("cLogEmpresa") & "...")
    '            setLog(id_descarga, "I", "Tipo de descarga solicitada: " & pstrTipoDocumento)

    '            strProveedor = drLog.Item("cLogUsuario")

    '            Dim strArchivo As String = "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & reporte.ToString & ".xls"


    '            If getArchivo("3", fecha, strArchivo) = 1 Then
    '                setLog(id_descarga, "I", "Archivo " & strArchivo & "  generado anteriormente...")

    '                strSQL = "UPDATE EstatusPH SET estatus='TERMINADO' , Resultado='Archivo " & strArchivo & "generado anteriormente...' WHERE fecha='" & fecha & "' AND TipoArchivo='" & pstrTipoDocumento & "' and estatus ='EN_PROCESO' "
    '                Dim sqlcmd2 As New SqlCommand(strSQL, conn)
    '                conn.Open()
    '                sqlcmd2.ExecuteNonQuery()
    '                conn.Close()

    '                intResultado = 0
    '            Else

    '                Browser("https://proveedores.palaciohierro.com.mx/irj/servlet/prt/portal/prtroot/com.sap.portal.navigation.masthead.LogOutComponent?logout_submit=true", 120)

    '                If Me.Browser(drLog.Item("cLogUrl"), 120) = 1 Then
    '                    setLog(id_descarga, "E", "Error: Cargar Pagina " & drLog.Item("cLogUrl"))
    '                    intResultado = 1
    '                    Return intResultado
    '                    Exit Function
    '                End If
    '                'System.Threading.Thread.Sleep(15000) ' JG
    '                While wb.Busy ' JG
    '                    setLog(id_descarga, "I", "Esperando respuesta de la pagina") ' JG
    '                End While ' JG

    '                strProveedor = drLog.Item("cLogUsuario")
    '                If Me.setAttribute("input", "name", "j_user", drLog.Item("cLogUsuario")) = 1 Then
    '                    intResultado = 1
    '                    Return intResultado
    '                    Exit Function
    '                End If

    '                If Me.setAttribute("input", "name", "j_password", drLog.Item("cLogContrasenia")) = 1 Then
    '                    intResultado = 1
    '                    Return intResultado
    '                    Exit Function
    '                End If

    '                If sendClick("input", "name", "uidPasswordLogon", 60) = 1 Then
    '                    intResultado = 1
    '                    Return intResultado
    '                    Exit Function
    '                End If

    '                If Me.Browser("https://wdbop.palaciohierro.com.mx/BOE/OpenDocument/opendoc/openDocument.jsp?iDocID=FtkAyE9eHQMAfCgAAEB5hEQBAFBWvQBx&sIDType=CUID", 120) Then
    '                    setLog(id_descarga, "E", "Error: Cargar Pagina https://wdbop.palaciohierro.com.mx/BOE/OpenDocument/opendoc/openDocument.jsp?iDocID=FtkAyE9eHQMAfCgAAEB5hEQBAFBWvQBx&sIDType=CUID")
    '                    intResultado = 1
    '                    Return intResultado
    '                    Exit Function
    '                End If

    '                If document.body.innerHTML.Contains("Service Temporarily Unavailable") Then
    '                    setLog(id_descarga, "E", "Service Temporarily Unavailable. Portal fuera de linea.")
    '                    intResultado = 1
    '                End If

    '                Dim htmlDocParent As mshtml.HTMLDocument = document
    '                Dim htmlDoc_openDocChildFrame As mshtml.HTMLDocument
    '                Dim htmlDoc_SaveAsDlg As mshtml.HTMLDocument
    '                Dim htmlDoc3 As mshtml.HTMLDocument
    '                Dim htmlDoc4 As mshtml.HTMLDocument
    '                Dim htmlDoc5 As mshtml.HTMLDocument

    '                Dim htmlWnd As mshtml.IHTMLWindow2 = Nothing
    '                Dim htmlWnd2 As mshtml.IHTMLWindow2 = Nothing
    '                Dim htmlWnd3 As mshtml.IHTMLWindow2 = Nothing
    '                Dim htmlWnd4 As mshtml.IHTMLWindow2 = Nothing

    '                Dim htmlElement As mshtml.IHTMLElement


    '                While wb.Busy 'JG
    '                    System.Threading.Thread.Sleep(1000)
    '                    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                End While 'JG
    '                System.Threading.Thread.Sleep(3000)

    '                For n As Integer = 0 To htmlDocParent.frames.length - 1
    '                    htmlWnd = htmlDocParent.frames.item(n)
    '                    If htmlWnd.name = "openDocChildFrame" Then
    '                        htmlDoc_openDocChildFrame = CType(htmlWnd.document, mshtml.HTMLDocument)


    '                        htmlDoc3 = CType(htmlDoc_openDocChildFrame.body.document, mshtml.HTMLDocument)
    '                        htmlWnd3 = CType(htmlDoc3.frames.item(0), mshtml.IHTMLWindow2)
    '                        htmlDoc4 = CType(htmlWnd3.document, mshtml.IHTMLDocument)

    '                        htmlElement = htmlDoc4.getElementById("RealBtn_CANCEL_BTN_promptsDlg")

    '                        cx = 0
    '                        While htmlElement Is Nothing And cx < 5
    '                            setLog(id_descarga, "I", "Esperando a pagina...")
    '                            System.Threading.Thread.Sleep(2000) ' JG
    '                            htmlElement = htmlDoc4.getElementById("RealBtn_CANCEL_BTN_promptsDlg")
    '                            cx = cx + 1
    '                        End While

    '                        htmlElement.click()
    '                        setLog(id_descarga, "I", "Click: RealBtn_CANCEL_BTN_promptsDlg")

    '                        While wb.Busy 'JG
    '                            System.Threading.Thread.Sleep(1000)
    '                            setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                        End While 'JG
    '                        System.Threading.Thread.Sleep(2000)

    '                        htmlElement = htmlDoc4.getElementById("IconImg__dhtmlLib_281")
    '                        cx = 0
    '                        While htmlElement Is Nothing And cx < 5
    '                            setLog(id_descarga, "I", "Esperando a pagina...")
    '                            System.Threading.Thread.Sleep(1000) ' JG
    '                            htmlElement = htmlDoc4.getElementById("IconImg__dhtmlLib_281")
    '                            cx = cx + 1
    '                        End While

    '                        htmlElement.click()
    '                        setLog(id_descarga, "I", "Click: IconImg__dhtmlLib_281")
    '                        System.Threading.Thread.Sleep(1000) ' JG
    '                        Exit For
    '                    End If
    '                Next

    '                'System.Threading.Thread.Sleep(loadcontrol_time) 'JG
    '                While wb.Busy 'JG
    '                    System.Threading.Thread.Sleep(1000)
    '                    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                End While 'JG
    '                System.Threading.Thread.Sleep(5000)

    '                For n As Integer = 0 To htmlDocParent.frames.length - 1
    '                    htmlWnd = htmlDocParent.frames.item(n)
    '                    If htmlWnd.name = "openDocChildFrame" Then
    '                        htmlDoc_openDocChildFrame = CType(htmlWnd.document, mshtml.HTMLDocument)

    '                        htmlDoc3 = CType(htmlDoc_openDocChildFrame.body.document, mshtml.HTMLDocument)
    '                        htmlWnd3 = CType(htmlDoc3.frames.item(0), mshtml.IHTMLWindow2)
    '                        htmlDoc4 = CType(htmlWnd3.document, mshtml.IHTMLDocument)

    '                        For n2 As Integer = 0 To htmlDoc4.frames.length - 1
    '                            htmlWnd4 = htmlDoc4.frames.item(n2)
    '                            If htmlWnd4.name = "saveAsDlg" Then
    '                                htmlDoc_SaveAsDlg = CType(htmlWnd4.document, mshtml.HTMLDocument)
    '                                htmlDoc5 = CType(htmlDoc_SaveAsDlg.body.document, mshtml.HTMLDocument)

    '                                htmlElement = htmlDoc5.getElementById("yui-gen0-1-label")
    '                                htmlElement.click()
    '                                setLog(id_descarga, "I", "Click: yui-gen0-1-label")
    '                                While wb.Busy 'JG
    '                                    System.Threading.Thread.Sleep(1000)
    '                                    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                                End While 'JG
    '                                System.Threading.Thread.Sleep(2000)

    '                                htmlElement = htmlDoc5.getElementById("accordionNavigationView_drawer0_treeView_treeNode1_name")
    '                                htmlElement.click()
    '                                setLog(id_descarga, "I", "Click: accordionNavigationView_drawer0_treeView_treeNode1_name")
    '                                ' System.Threading.Thread.Sleep(1000) 'JG


    '                                While wb.Busy 'JG
    '                                    System.Threading.Thread.Sleep(1000)
    '                                    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                                End While 'JG
    '                                System.Threading.Thread.Sleep(2000)

    '                                'Select Case reporte
    '                                '    Case Inv_PH.COS1

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) ' JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox3")

    '                                '    Case Inv_PH.COS2

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) ' JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox0")

    '                                '    Case Inv_PH.COS3

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) ' JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox1")

    '                                '    Case Inv_PH.COS4

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox2")

    '                                '    Case Inv_PH.IPI1

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox5")

    '                                '    Case Inv_PH.IPI2
    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG
    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox6")
    '                                '    Case Inv_PH.IPI3

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox7")

    '                                '    Case Inv_PH.IPI4

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox8")

    '                                '    Case Inv_PH.IVK1

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox1")

    '                                '    Case Inv_PH.LIFES

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG

    '                                '        htmlElement = htmlDoc5.getElementById("IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        htmlElement.click()
    '                                '        setLog(id_descarga, "I", "Click: IconImg_openDocuments_DocumentSelector_goForwardButton")
    '                                '        System.Threading.Thread.Sleep(1000) 'JG

    '                                '        htmlElement = htmlDoc5.getElementById("openDocuments_DocumentSelector_detailView_checkbox2")

    '                                'End Select


    '                                htmlElement.click()
    '                                setLog(id_descarga, "I", "Click: " & htmlElement.id)

    '                                While wb.Busy 'JG
    '                                    System.Threading.Thread.Sleep(1000)
    '                                    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                                End While 'JG
    '                                System.Threading.Thread.Sleep(1000)

    '                                htmlElement = htmlDoc5.getElementById("openDocuments_OpenButton")
    '                                cx = 0
    '                                While htmlElement Is Nothing And cx < 5
    '                                    setLog(id_descarga, "I", "Esperando a pagina...")
    '                                    System.Threading.Thread.Sleep(1000)
    '                                    htmlElement = htmlDoc4.getElementById("openDocuments_OpenButton")
    '                                    cx = cx + 1
    '                                End While

    '                                htmlElement.click()
    '                                setLog(id_descarga, "I", "Click: openDocuments_OpenButton")

    '                                Exit For
    '                            End If
    '                        Next

    '                    End If
    '                Next

    '                'System.Threading.Thread.Sleep(5000) ' JG
    '                While wb.Busy = True
    '                    'System.Threading.Thread.Sleep(wait_time) ' JG
    '                    System.Threading.Thread.Sleep(1000)
    '                    setLog(id_descarga, "I", "Click: Esperando carga de pagina") ' JG
    '                End While
    '                System.Threading.Thread.Sleep(3000)

    '                For n As Integer = 0 To htmlDocParent.frames.length - 1
    '                    htmlWnd = htmlDocParent.frames.item(n)
    '                    If htmlWnd.name = "openDocChildFrame" Then
    '                        htmlDoc_openDocChildFrame = CType(htmlWnd.document, mshtml.HTMLDocument)

    '                        While htmlDoc_openDocChildFrame.readyState <> "complete"
    '                            System.Threading.Thread.Sleep(1000)
    '                            setLog(id_descarga, "I", "Click: Esperando carga de pagina") ' JG
    '                        End While

    '                        htmlDoc3 = CType(htmlDoc_openDocChildFrame.body.document, mshtml.HTMLDocument)
    '                        htmlWnd3 = CType(htmlDoc3.frames.item(0), mshtml.IHTMLWindow2)
    '                        htmlDoc4 = CType(htmlWnd3.document, mshtml.IHTMLDocument)

    '                        htmlElement = htmlDoc4.getElementById("RealBtn_OK_BTN_promptsDlg")

    '                        cx = 0
    '                        While htmlElement Is Nothing And cx < 5
    '                            setLog(id_descarga, "I", "Esperando a pagina...")
    '                            System.Threading.Thread.Sleep(1000)
    '                            htmlElement = htmlDoc4.getElementById("RealBtn_OK_BTN_promptsDlg")
    '                            cx = cx + 1
    '                        End While

    '                        htmlElement.click()
    '                        setLog(id_descarga, "I", "Click: RealBtn_OK_BTN_promptsDlg")

    '                        While wb.Busy 'JG
    '                            System.Threading.Thread.Sleep(1000)
    '                            setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                        End While 'JG
    '                        System.Threading.Thread.Sleep(1000)
    '                        'System.Threading.Thread.Sleep(5000) ' JG

    '                        Exit For

    '                    End If
    '                Next

    '                'System.Threading.Thread.Sleep(30000) 'JG
    '                While wb.Busy 'JG
    '                    System.Threading.Thread.Sleep(1000)
    '                    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                End While 'JG
    '                System.Threading.Thread.Sleep(5000)

    '                Dim htmlElement_Inicio As mshtml.IHTMLElement
    '                Dim htmlElement_Atras As mshtml.IHTMLElement
    '                Dim htmlElement_Texto As mshtml.IHTMLElement
    '                Dim htmlElement_Siguiente As mshtml.IHTMLElement
    '                Dim htmlElement_Fin As mshtml.IHTMLElement
    '                Dim htmlElement_Contenido As mshtml.IHTMLElement
    '                Dim maxPaginas As Integer
    '                Dim minPaginas As Integer
    '                Dim strTexto As String


    '                For n As Integer = 0 To htmlDocParent.frames.length - 1
    '                    htmlWnd = htmlDocParent.frames.item(n)
    '                    If htmlWnd.name = "openDocChildFrame" Then
    '                        htmlDoc_openDocChildFrame = CType(htmlWnd.document, mshtml.HTMLDocument)

    '                        Dim doc1x As mshtml.IHTMLWindow2 = CType(htmlDoc_openDocChildFrame.frames.item(0), mshtml.IHTMLWindow2)
    '                        Dim doc1Menu As mshtml.HTMLDocument = doc1x.document

    '                        doc1x = CType(doc1Menu.frames.item(3), mshtml.HTMLWindow2) 'sacando el frame con el texto
    '                        While doc1x.document.readyState <> "complete"
    '                            System.Threading.Thread.Sleep(1000)
    '                            setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                        End While

    '                        While doc1Menu.readyState <> "complete"
    '                            System.Threading.Thread.Sleep(1000)
    '                            setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                        End While

    '                        'htmlElement_Inicio = doc1Menu.getElementById("IconImg__dhtmlLib_186")
    '                        'htmlElement_Atras = doc1Menu.getElementById("IconImg__dhtmlLib_187")
    '                        'htmlElement_Siguiente = doc1Menu.getElementById("IconImg__dhtmlLib_190")
    '                        'htmlElement_Fin = doc1Menu.getElementById("IconImg__dhtmlLib_191")

    '                        'htmlElement_Fin.click()
    '                        'setLog(id_descarga, "I", "Click: htmlElement_Fin")
    '                        ''System.Threading.Thread.Sleep(20000) ' JG
    '                        'While wb.Busy 'JG
    '                        '    System.Threading.Thread.Sleep(1000)
    '                        '    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                        'End While 'JG
    '                        'System.Threading.Thread.Sleep(2000)
    '                        'While doc1x.document.readyState <> "complete"
    '                        '    System.Threading.Thread.Sleep(1000)
    '                        '    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                        'End While

    '                        'htmlElement_Inicio.click()
    '                        'setLog(id_descarga, "I", "Click: htmlElement_Inicio")
    '                        ''System.Threading.Thread.Sleep(20000) ' JG
    '                        'While wb.Busy 'JG
    '                        '    System.Threading.Thread.Sleep(1000)
    '                        '    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                        'End While 'JG
    '                        'System.Threading.Thread.Sleep(2000)
    '                        'While doc1x.document.readyState <> "complete"
    '                        '    System.Threading.Thread.Sleep(1000)
    '                        '    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                        'End While

    '                        'doc1x = CType(doc1Menu.frames.item(3), mshtml.HTMLWindow2) 'sacando el frame con el texto

    '                        Dim url2 As String = doc1Menu.frames.item(3).location.href
    '                        Dim params2() As String = url2.Split("&")


    '                        Dim p1 As Integer = 0
    '                        maxPaginas = 0
    '                        While p1 < params2.Length
    '                            If params2(p1).Contains("nbPage") Then
    '                                Dim val() As String = params2(p1).Split("=")
    '                                maxPaginas = val(1)
    '                            End If
    '                            p1 = p1 + 1
    '                        End While
    '                        If maxPaginas = 0 Then
    '                            setLog(id_descarga, "E", "No se pudo determinar el numero de paginas.")
    '                            intResultado = 1
    '                        End If

    '                        If maxPaginas > 0 Then
    '                            setLog(id_descarga, "I", "No Paginas: " & maxPaginas)
    '                            url2 = url2.Replace("iPage=first&", "")

    '                            Dim doc1Contenido = doc1x.document


    '                            Dim saveTo1x As String = My.Computer.FileSystem.GetTempFileName.ToString
    '                            Dim filetmp As String = saveTo1x
    '                            saveTo1x = saveTo1x.Replace(".tmp", ".html")

    '                            If System.IO.File.Exists(filetmp) Then
    '                                System.IO.File.Delete(filetmp)
    '                            End If

    '                            Dim w As New System.IO.StreamWriter(saveTo1x, False, System.Text.Encoding.UTF8)
    '                            setLog(id_descarga, "I", "Generando archivo: " & saveTo1x)

    '                            minPaginas = 1
    '                            While minPaginas <= maxPaginas

    '                                'doc1Contenido = doc1x.document
    '                                setLog(id_descarga, "I", "Leyendo pagina:" & minPaginas)
    '                                'Try
    '                                'w.WriteLine(doc1Contenido.body.innerHTML)
    '                                w.WriteLine(doc1x.document.body.innerHTML)
    '                                'Catch ex As Exception
    '                                'setLog(id_descarga, "E", ex.Message)
    '                                'End Try

    '                                setLog(id_descarga, "I", "Escribiendo datos en archivo...")

    '                                htmlElement_Siguiente.click()
    '                                setLog(id_descarga, "I", "Click: htmlElement_Siguiente")
    '                                'System.Threading.Thread.Sleep(10000) 'JG
    '                                While wb.Busy 'JG
    '                                    System.Threading.Thread.Sleep(1000)
    '                                    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                                End While 'JG
    '                                System.Threading.Thread.Sleep(2000)
    '                                While doc1x.document.readyState <> "complete"
    '                                    System.Threading.Thread.Sleep(1000)
    '                                    setLog(id_descarga, "I", "Web Busy: pagina ocupada ") 'JG
    '                                End While

    '                                minPaginas = minPaginas + 1
    '                            End While


    '                            w.Flush()
    '                            w.Close()
    '                            setLog(id_descarga, "I", "Guardando archivos temporales...")

    '                            Dim excel As New Excel.Application
    '                            Dim sts_archivo As Boolean = False

    '                            Try
    '                                Dim book As Excel.Workbook = excel.Workbooks.Open(saveTo1x)
    '                                excel.Visible = False
    '                                excel.DisplayAlerts = False
    '                                If book.Sheets.Count > 0 Then
    '                                    Dim sheet As Worksheet = book.Sheets(1)
    '                                    Dim cel = sheet.Cells(9, 2)
    '                                    setLog(id_descarga, "I", "Cell(9,2):" & cel.value)
    '                                    If cel.value = "Resultado" Then
    '                                        sts_archivo = False
    '                                    Else
    '                                        cel = sheet.Cells(14, 1)
    '                                        setLog(id_descarga, "I", "Cell(15,1):" & cel.value)
    '                                        If cel.value.ToString = "" Then
    '                                            sts_archivo = False
    '                                        Else
    '                                            sts_archivo = True
    '                                        End If
    '                                    End If

    '                                    sheet.Name = "Sku"
    '                                End If
    '                                'Dim saveTo2x As String = My.Computer.FileSystem.GetTempFileName.ToString
    '                                filetmp = saveTo1x
    '                                saveTo1x = saveTo1x.Replace(".html", ".xls")

    '                                If sts_archivo Then
    '                                    book.SaveAs(saveTo1x, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal)
    '                                End If
    '                                'saveTo1x = saveTo2x 

    '                                'book.Save()
    '                                book.Close()
    '                                excel.Quit()
    '                                book = Nothing
    '                                excel = Nothing
    '                                If System.IO.File.Exists(filetmp) Then
    '                                    System.IO.File.Delete(filetmp)
    '                                End If
    '                            Catch ex As Exception
    '                                setLog(id_descarga, "E", "Excepcion:" & ex.Message)
    '                                sts_archivo = False
    '                            End Try

    '                            ' Modificación: - 100914 LVH '''''''''''''''''''''''''''''''''''''''''''
    '                            ' Validar si tiene registros el archivo descargado '''''''''''''''''''''
    '                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                            'If sts_archivo Then
    '                            '    setLog(id_descarga, "I", "Verificando contenido archivo de Excel.")
    '                            '    Dim excel2 As New Excel.Application
    '                            '    Dim contador As Integer = 1
    '                            '    Dim dTotal As Double = 0.0
    '                            '    Try
    '                            '        Dim book2 As Excel.Workbook = excel2.Workbooks.Open(saveTo1x) '"c:\TEMP\PH_VTA_09092014_PR0000184366.xls")
    '                            '        excel2.Visible = False
    '                            '        excel2.DisplayAlerts = False

    '                            '        If book2.Sheets.Count > 0 Then
    '                            '            Dim sheet2 As Worksheet = book2.Sheets(1)
    '                            '            ' Revisar sólo 200 líneas para validar si tiene o no información
    '                            '            While contador < 200 'sheet.Rows.Count
    '                            '                Dim cel2 = sheet2.Cells(contador, 3)
    '                            '                If cel2.value Is Nothing Then
    '                            '                    sts_archivo = False
    '                            '                Else
    '                            '                    If cel2.value.ToString.Trim = "Total Marca:" Then
    '                            '                        ' validar cantidad 13,14
    '                            '                        cel2 = sheet2.Cells(contador, 14)
    '                            '                        If cel2.value Is Nothing Then
    '                            '                            sts_archivo = False
    '                            '                        Else
    '                            '                            If cel2.value.ToString.Trim = "" Then
    '                            '                                sts_archivo = False
    '                            '                            Else
    '                            '                                Try
    '                            '                                    dTotal += Convert.ToDouble(cel2.value.ToString.Trim)
    '                            '                                Catch ex As Exception
    '                            '                                    Continue While
    '                            '                                End Try

    '                            '                                sts_archivo = True
    '                            '                            End If
    '                            '                        End If
    '                            '                    Else
    '                            '                        sts_archivo = False
    '                            '                    End If
    '                            '                End If

    '                            '                contador = contador + 1
    '                            '            End While
    '                            '        End If

    '                            '        book2.Close()
    '                            '        excel2.Quit()
    '                            '        book2 = Nothing
    '                            '        excel2 = Nothing

    '                            '        If dTotal > 0 Then
    '                            '            sts_archivo = True
    '                            '            setLog(id_descarga, "I", "Archivo de Excel correcto con registros.")
    '                            '        Else
    '                            '            sts_archivo = False
    '                            '            setLog(id_descarga, "E", "Error archivo de Excel sin Registros...")
    '                            '            ' Depurar el archivo de Excel creado
    '                            '            If System.IO.File.Exists(saveTo1x) Then
    '                            '                System.IO.File.Delete(saveTo1x)
    '                            '            End If
    '                            '            setLog(id_descarga, "E", "Archivo de Excel eliminado...")
    '                            '        End If
    '                            '    Catch ex As Exception
    '                            '        sts_archivo = False
    '                            '        setLog(id_descarga, "E", "Error archivo de Excel sin Registros...")
    '                            '    End Try
    '                            'End If
    '                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    '                            If sts_archivo Then
    '                                setLog(id_descarga, "I", "Enviando archivos a FTP...")
    '                                ftp.ServerAddress = server_ftp
    '                                ftp.ServerPort = port_ftp
    '                                ftp.UserName = user_ftp
    '                                ftp.Password = pwd_ftp
    '                                ftp.Connect()

    '                                Dim strArchivoFtp As String = "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & reporte.ToString & ".xls"

    '                                'ftp.UploadFile(saveTo1x, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", False)
    '                                ftp.UploadFile(saveTo1x, strArchivoFtp, False)

    '                                Dim features As EnterpriseDT.Net.Ftp.FTPReply = ftp.InvokeFTPCommand("cwd /Respaldos")
    '                                'ftp.UploadFile(saveTo1x, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & "-" & Now.ToString("ddMMyyyy_HHmmss") & ".xls", False)
    '                                ftp.UploadFile(saveTo1x, strArchivoFtp, False)

    '                                ftp.Close()

    '                                'setArchivo("3", fecha, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", drLog("cLogRutaDescarga") & "\", IIf(pstrTipoDocumento = "Ventas", "VTA", "INV"))
    '                                setArchivo("3", fecha, strArchivo, drLog("cLogRutaDescarga") & "\", IIf(pstrTipoDocumento = "Ventas", "VTA", "INV"))
    '                                'TArchivos.Rows.Add("PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", drLog("cLogRutaDescarga") & "\")
    '                                TArchivos.Rows.Add(strArchivo, drLog("cLogRutaDescarga") & "\")
    '                                'setLog(id_descarga, "I", "Archivo creado: " & "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls")
    '                                setLog(id_descarga, "I", "Archivo creado: " & strArchivo)
    '                                setLog(id_descarga, "I", "Fin de descarga de archivos...")
    '                                intResultado = 0

    '                                If System.IO.File.Exists(saveTo1x) Then
    '                                    Try
    '                                        System.IO.File.Delete(saveTo1x)
    '                                    Catch ex As Exception
    '                                    End Try
    '                                End If
    '                                saveTo1x = saveTo1x.Replace(".xls", ".html")
    '                                If System.IO.File.Exists(saveTo1x) Then
    '                                    Try
    '                                        System.IO.File.Delete(saveTo1x)
    '                                    Catch ex As Exception
    '                                    End Try
    '                                End If

    '                            Else
    '                                intResultado = 1
    '                                setLog(id_descarga, "E", "Error en archivo. Formato desconocido o vacio...")
    '                            End If






    '                        End If

    '                        Exit For
    '                    End If
    '                Next

    '                'If intResultado = 1 Then
    '                '    Exit Try
    '                'End If


    '            End If
    '        Next
    '    Catch ex As Exception
    '        setLog(id_descarga, "E", "Error:" & ex.Message & " " & strMensaje)
    '        intResultado = 1
    '        'Throw
    '    End Try
    '    terminar()
    '    Return intResultado
    'End Function

    Private Function getArchivosPH(ByVal fecha As Date, ByVal pstrTipoDocumento As String) As Integer
        Dim res As Integer = 0
        Dim cont As Integer = 0

        res = iniciar()
        While res = 1 And cont < 10
            'setLog(id_descarga, "E", message)
            System.Threading.Thread.Sleep(1000)
            res = iniciar()
            cont = cont + 1
        End While
        If res = 1 Then
            setLog(id_descarga, "E", message)
            Return 1
        End If

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

                setLog(id_descarga, "I", "Iniciando proceso para " & drLog.Item("cLogEmpresa") & "...")
                setLog(id_descarga, "I", "Tipo de descarga solicitada: " & pstrTipoDocumento)

                strProveedor = drLog.Item("cLogUsuario")
                If getArchivo("3", fecha, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls") = 1 Then
                    setLog(id_descarga, "I", "Archivo PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls generado anteriormente...")
                    intResultado = 0
                Else


                    If Me.Browser(drLog.Item("cLogUrl"), 120) = 1 Then
                        setLog(id_descarga, "E", "Error: Cargar Pagina " & drLog.Item("cLogUrl"))
                        intResultado = 1
                        Return intResultado
                        Exit Function
                    End If
                    strProveedor = drLog.Item("cLogUsuario")
                    If Me.setAttribute("input", "name", "j_user", drLog.Item("cLogUsuario")) = 1 Then
                        intResultado = 1
                        Return intResultado
                        Exit Function
                    End If
                    If Me.setAttribute("input", "name", "j_password", drLog.Item("cLogContrasenia")) = 1 Then
                        intResultado = 1
                        Return intResultado
                        Exit Function
                    End If

                    For drFila As Integer = 0 To dtProceso.Rows.Count - 1
                        strMensaje = dtProceso.Rows(drFila).Item("cProValorAtributo")
                        System.Threading.Thread.Sleep(1000)
                        Select Case dtProceso.Rows(drFila).Item("cProAccion")
                            Case "Navegar"
                                intResultado = Browser(dtProceso.Rows(drFila).Item("cProValorAtributo"), 120)
                            Case "Teclear"
                                If dtProceso.Rows(drFila).Item("cProNombreAtributo") = "lsS1) Fecha (mm/dd/aaaa)" Then
                                    'If drLog.Item("cLogUsuario") = "184366" Then
                                    intResultado = Me.setAttribute(dtProceso.Rows(drFila).Item("cProTipo"), dtProceso.Rows(drFila).Item("cProAtributo"), dtProceso.Rows(drFila).Item("cProNombreAtributo"), getMesEng(fecha.ToString("MM")) & fecha.ToString("/dd/yyyy"))
                                    'Else
                                    '   intResultado = Me.setAttribute(dtProceso.Rows(drFila).Item("cProTipo"), dtProceso.Rows(drFila).Item("cProAtributo"), dtProceso.Rows(drFila).Item("cProNombreAtributo"), fecha.ToString("dd/MM/yyyy"))
                                    'End If
                                Else
                                    intResultado = Me.setAttribute(dtProceso.Rows(drFila).Item("cProTipo"), dtProceso.Rows(drFila).Item("cProAtributo"), dtProceso.Rows(drFila).Item("cProNombreAtributo"), dtProceso.Rows(drFila).Item("cProValorAtributo"))
                                End If
                            Case "Click"
                                intResultado = Me.sendClick(dtProceso.Rows(drFila).Item("cProTipo"), dtProceso.Rows(drFila).Item("cProAtributo"), dtProceso.Rows(drFila).Item("cProNombreAtributo"), 480)
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
                                intResultado = Me.sendLink(dtProceso.Rows(drFila).Item("cProValorAtributo").ToString.Trim, 300)
                            Case "Descarga"
                                httpResponse = Me.WRequest(dtProceso.Rows(drFila).Item("cProValorAtributo"), "get", "", cookiesContainer, 300)
                                If Not httpResponse Is Nothing Then
                                    setLog(id_descarga, "I", "Procesando datos descargados...")
                                    'httpResponse = Me.WRequest("http://www.phb2b.com.mx/wi/scripts/saveAsXls.asp", "post", "cmdBlock=all&cmd=asksave&cmdP1=%286.%202%29%20%20Inventarios%20Detalle%20Tienda%20-%20Hogar%20y%20L.%20Generales*1030*0*rep*wi00000001", cookiesContainer)
                                    Dim o As System.IO.Stream = httpResponse.GetResponseStream
                                    Dim saveTo As String = My.Computer.FileSystem.GetTempFileName.ToString
                                    saveTo = saveTo.Replace(".tmp", ".xls")
                                    setLog(id_descarga, "I", "Archivo temporal generado: " & saveTo)
                                    Dim writeStream As IO.FileStream = New IO.FileStream(saveTo, IO.FileMode.Create, IO.FileAccess.Write)
                                    Me.ReadWriteStream(o, writeStream)

                                    Dim sts_archivo As Boolean = True
                                    Dim total As Integer = 0
                                    Try
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


                                        Dim sheet1 As Excel.Worksheet = oBook.Sheets(1)
                                        Dim activecell As Excel.Range = sheet1.Cells(1, 1)
                                        activecell = sheet1.Cells.Find(What:="Total Proveedor", After:=activecell, LookIn:=Excel.XlFindLookIn.xlFormulas _
                                        , LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=XlSearchDirection.xlNext,
                                        MatchCase:=False, SearchFormat:=False)
                                        If Not activecell Is Nothing Then
                                            total = 0
                                            If pstrTipoDocumento = "Ventas" Then
                                                total = sheet1.Cells(activecell.Row, 17).Value
                                            Else
                                                total = sheet1.Cells(activecell.Row, 10).Value
                                            End If
                                        End If
                                    Catch ex As Exception
                                        sts_archivo = False
                                    End Try
                                    oBook.Save()
                                    oBook.Close()
                                    oExcel.Quit()
                                    oBook = Nothing
                                    oExcel = Nothing
                                    If sts_archivo And total > 0 Then
                                        setLog(id_descarga, "I", "Enviando archivos a FTP...")
                                        ftp.ServerAddress = server_ftp
                                        ftp.ServerPort = port_ftp
                                        ftp.UserName = user_ftp
                                        ftp.Password = pwd_ftp
                                        ftp.Connect()

                                        ftp.UploadFile(saveTo, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", False)

                                        Dim features As EnterpriseDT.Net.Ftp.FTPReply = ftp.InvokeFTPCommand("cwd /Respaldos")
                                        ftp.UploadFile(saveTo, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & "-" & Now.ToString("ddMMyyyy_HHmmss") & ".xls", False)

                                        ftp.Close()

                                        'If Not System.IO.Directory.Exists(drLog.Item(6).ToString()) Then
                                        '    My.Computer.FileSystem.CreateDirectory(drLog.Item(6).ToString())
                                        'End If
                                        'If My.Computer.FileSystem.FileExists(drLog.Item(6).ToString() & "\PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls") Then
                                        '    My.Computer.FileSystem.DeleteFile(drLog.Item(6).ToString() & "\PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls")
                                        'End If
                                        'My.Computer.FileSystem.CopyFile(saveTo, drLog.Item(6).ToString() & "\PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls")
                                        setArchivo("3", fecha, "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", drLog("cLogRutaDescarga") & "\", IIf(pstrTipoDocumento = "Ventas", "VTA", "INV"))
                                        TArchivos.Rows.Add("PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls", drLog("cLogRutaDescarga") & "\")
                                        setLog(id_descarga, "I", "Archivo creado: " & "PH_" & IIf(pstrTipoDocumento = "Ventas", "VTA", "INV") & "_" & fecha.ToString("ddMMyyyy") & "_" & strProveedor & ".xls")
                                        setLog(id_descarga, "I", "Fin de descarga de archivos...")
                                        intResultado = 0
                                    Else
                                        intResultado = 1
                                        setLog(id_descarga, "E", "Error en archivo. Formato desconocido o vacio...")
                                    End If
                                Else
                                    intResultado = 1
                                    setLog(id_descarga, "E", "Error en la descarga. " & message)
                                End If
                        End Select
                        System.Threading.Thread.Sleep(3000)
                        If intResultado = 1 Then
                            Exit Try
                        End If
                    Next

                End If
            Next
        Catch ex As Exception
            setLog(id_descarga, "E", "Error:" & ex.Message & " " & strMensaje)
            intResultado = 1
            'Throw
        End Try
        terminar()
        Return intResultado
    End Function


    Public Function getPH(ByVal fecha As Date, ByVal Archivo As String, ByVal secuencia As Integer) As String
        Dim err As Integer = 0
        Dim ini_vta As Integer = 0
        Dim ini_inv As Integer = 0
        Dim contador As Integer = 1
        Dim xml As String = String.Empty

        getDescarga("PH", fecha)

        If id_descarga > 0 Then

            setLog(id_descarga, "I", "Iniciando procesos PH...")
            setLog(id_descarga, "I", "Numero de Intento: " & secuencia)
            err = getFTP("PH")
            If err = 0 Then

                If Archivo = "V" Then 'solo ventas
                    If secuencia = 0 And ini_vta = 0 Then

                        ini_vta = 1
                    End If

                    err = getVentasPH(fecha, "Ventas")
                End If


                If Archivo = "I" Then 'solo inventarios
                    If secuencia = 0 And ini_inv = 0 Then

                        ini_inv = 1
                    End If

                    'Inicia descarga de 1er archivo Costix
                    While contador <= 2
                        err = getInventariosPH(fecha, "Inventarios", Inv_PH.COS1)
                        If err = 0 Then
                            Exit While
                        End If
                        contador += 1
                    End While
                    contador = 1
                    If err = 0 Then

                        setLog(id_descarga, "I", "Invetario Costix Parte 1 terminado correctamente...")

                        'Inicia descarga de 2do archivo Costix
                        While contador <= 2
                            err = getInventariosPH(fecha, "Inventarios", Inv_PH.COS2)
                            If err = 0 Then
                                Exit While
                            End If
                            contador += 1
                        End While
                        contador = 1
                        If err = 0 Then


                            setLog(id_descarga, "I", "Invetario Costix Parte 2 terminado correctamente...")


                            'Inicia descarga de 3er archivo Costix
                            While contador <= 2
                                err = getInventariosPH(fecha, "Inventarios", Inv_PH.COS3)
                                If err = 0 Then
                                    Exit While
                                End If
                                contador += 1
                            End While
                            contador = 1

                            If err = 0 Then


                                setLog(id_descarga, "I", "Invetario Costix Parte 3 terminado correctamente...")


                                'Inicia descarga de 3er archivo Piagui
                                While contador <= 2

                                    err = getInventariosPH(fecha, "Inventarios", Inv_PH.COS4)
                                    If err = 0 Then
                                        Exit While
                                    End If
                                    contador += 1
                                End While
                                contador = 1

                                If err = 0 Then


                                    setLog(id_descarga, "I", "Invetario Costix Parte 4 terminado correctamente...")


                                    'Inicia descarga de 1er archivo Piagui
                                    While contador <= 2

                                        err = getInventariosPH(fecha, "Inventarios", Inv_PH.IPI1)
                                        If err = 0 Then
                                            Exit While
                                        End If
                                        contador += 1
                                    End While
                                    contador = 1

                                    If err = 0 Then

                                        setLog(id_descarga, "I", "Invetario Piagui Parte 1 terminado correctamente...")


                                        'Inicia descarga de 2do archivo Piagui
                                        While contador <= 2
                                            err = getInventariosPH(fecha, "Inventarios", Inv_PH.IPI2)
                                            If err = 0 Then
                                                Exit While
                                            End If
                                            contador += 1
                                        End While
                                        contador = 1
                                        If err = 0 Then

                                            setLog(id_descarga, "I", "Invetario Piagui Parte 2 terminado correctamente...")


                                            'Inicia descarga de 3er archivo Piagui
                                            While contador <= 2

                                                err = getInventariosPH(fecha, "Inventarios", Inv_PH.IPI3)
                                                If err = 0 Then
                                                    Exit While
                                                End If
                                                contador += 1
                                            End While
                                            contador = 1
                                            If err = 0 Then
                                                setLog(id_descarga, "I", "Invetario Piagui Parte 3 terminado correctamente...")

                                                'Inicia descarga de 4to archivo Piagui
                                                While contador <= 2

                                                    err = getInventariosPH(fecha, "Inventarios", Inv_PH.IPI4)
                                                    If err = 0 Then
                                                        Exit While
                                                    End If
                                                    contador += 1
                                                End While
                                                contador = 1

                                                If err = 0 Then
                                                    setLog(id_descarga, "I", "Invetario Ivanka 1 terminado correctamente...")

                                                    'Inicia descarga de 1er Ivanka
                                                    While contador <= 2
                                                        err = getInventariosPH(fecha, "Inventarios", Inv_PH.IVK1)
                                                        If err = 0 Then
                                                            Exit While
                                                        End If
                                                        contador += 1
                                                    End While

                                                    contador = 1
                                                    If err = 0 Then
                                                        setLog(id_descarga, "I", "Invetario Ivanka 1 terminado correctamente...")
                                                        'Inicia descarga de 1er CAT
                                                        While contador <= 2
                                                            err = getInventariosPH(fecha, "Inventarios", Inv_PH.LIFES)
                                                            If err = 0 Then
                                                                Exit While
                                                            End If
                                                            contador += 1
                                                        End While

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
                If Archivo = "T" And err = 0 Then 'Ventas e Inventarios
                    If secuencia = 0 And ini_vta = 0 Then
                        ' MODIFICACION 050315: Carlos Reyes solicitó quitar el proceso Elimina Historial

                        ini_vta = 1
                    End If

                    err = getVentasPH(fecha, "Ventas")
                    If err = 0 Then
                        If secuencia = 0 And ini_inv = 0 Then

                            ini_inv = 1
                        End If

                        'Inicia descarga de 1er archivo

                        While contador <= 2
                            err = getInventariosPH(fecha, "Inventarios", Inv_PH.COS1)
                            If err = 0 Then
                                Exit While
                            End If
                            contador += 1
                        End While
                        contador = 1
                        If err = 0 Then

                            setLog(id_descarga, "I", "Invetario Costix Parte 1 terminado correctamente...")

                            'Inicia descarga de 2do archivo
                            While contador <= 2
                                err = getInventariosPH(fecha, "Inventarios", Inv_PH.COS2)
                                If err = 0 Then
                                    Exit While
                                End If
                                contador += 1
                            End While
                            contador = 1
                            If err = 0 Then

                                setLog(id_descarga, "I", "Invetario Costix Parte 2 terminado correctamente...")

                                'Inicia descarga de 3er archivo
                                While contador <= 2
                                    err = getInventariosPH(fecha, "Inventarios", Inv_PH.COS3)
                                    If err = 0 Then
                                        Exit While
                                    End If
                                    contador += 1
                                End While
                                contador = 1
                                If err = 0 Then

                                    setLog(id_descarga, "I", "Invetario Costix Parte 3 terminado correctamente...")

                                    'Inicia descarga de 3er archivo
                                    While contador <= 2
                                        err = getInventariosPH(fecha, "Inventarios", Inv_PH.COS4)
                                        If err = 0 Then
                                            Exit While
                                        End If
                                        contador += 1
                                    End While
                                    contador = 1



                                    If err = 0 Then

                                        setLog(id_descarga, "I", "Invetario Costix Parte 4 terminado correctamente...")

                                        'Inicia descarga de 5to archivo
                                        While contador <= 2
                                            err = getInventariosPH(fecha, "Inventarios", Inv_PH.IPI1)
                                            If err = 0 Then
                                                Exit While
                                            End If
                                            contador += 1
                                        End While
                                        contador = 1


                                        If err = 0 Then

                                            setLog(id_descarga, "I", "Invetario Piagui Parte 1 terminado correctamente...")
                                            While contador <= 2
                                                err = getInventariosPH(fecha, "Inventarios", Inv_PH.IPI2)
                                                If err = 0 Then
                                                    Exit While
                                                End If
                                                contador += 1
                                            End While
                                            contador = 1
                                            If err = 0 Then

                                                setLog(id_descarga, "I", "Invetario Piagui Parte 2 terminado correctamente...")
                                                While contador <= 2
                                                    err = getInventariosPH(fecha, "Inventarios", Inv_PH.IPI3)
                                                    If err = 0 Then
                                                        Exit While
                                                    End If
                                                    contador += 1
                                                End While
                                                contador = 1
                                                If err = 0 Then
                                                    setLog(id_descarga, "I", "Invetario Piagui Parte 3 terminado correctamente...")

                                                    While contador <= 2
                                                        err = getInventariosPH(fecha, "Inventarios", Inv_PH.IPI4)
                                                        If err = 0 Then
                                                            Exit While
                                                        End If
                                                        contador += 1
                                                    End While
                                                    contador = 1
                                                    If err = 0 Then

                                                        setLog(id_descarga, "I", "Invetario Piagui Parte 4 terminado correctamente...")

                                                        While contador <= 2
                                                            err = getInventariosPH(fecha, "Inventarios", Inv_PH.IVK1)
                                                            If err = 0 Then
                                                                Exit While
                                                            End If
                                                            contador += 1
                                                        End While
                                                        contador = 1
                                                        If err = 0 Then
                                                            setLog(id_descarga, "I", "Invetario Ivanka terminado correctamente...")
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
                    terminar()
                End If


                If err = 0 Then
                    setLog(id_descarga, "I", "Proceso terminado correctamente...")
                Else
                    setLog(id_descarga, "E", "Proceso terminado con errores, favor de revisar secuencia...")

                    xml = GetXML(err, Archivo, fecha)
                    conn.Open()
                    sql = New Data.SqlClient.SqlCommand("update estatusPH set estatus='ERROR',resultado ='No termino la descarga',secuencia=" & secuencia & " where iddescarga=@id_descarga", conn)
                    sql.Parameters.AddWithValue("@id_descarga", id_descarga)
                    sql.Parameters.AddWithValue("@sts", err)
                    sql.ExecuteNonQuery()
                    conn.Close()

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


            setLog(id_descarga, "I", "Generando XML de salida...")

            xml = GetXML(err, Archivo, fecha)

            Return xml
        End If

        Return Nothing
    End Function
    Public Function GetXML(ByVal err As Integer, ByVal Archivo As String, ByVal fecha As Date) As String

        Dim xml As String = String.Empty
        xml = xml & "<?xml version=""1.0"" encoding=""UTF-8"" ?>" & vbCrLf
        xml = xml & "<root>" & vbCrLf
        xml = xml & "<resultado><valor>" & err & "</valor></resultado>" & vbCrLf
        xml = xml & "<archivos>" & vbCrLf

        conn.Open()
        sql = New Data.SqlClient.SqlCommand("select * from archivos where fecha=@fecha and cliente=@cliente and tipo=@tipo", conn)
        sql.CommandTimeout = 0
        sql.Parameters.AddWithValue("@cliente", "3")
        sql.Parameters.AddWithValue("@tipo", IIf(Archivo = "V", "VTA", "INV"))
        sql.Parameters.AddWithValue("@fecha", fecha.ToString("ddMMyyyy"))
        rs = sql.ExecuteReader


        While rs.Read
            xml = xml & "<archivo>" & vbCrLf

            xml = xml & "<nombre>" & rs!archivo & "</nombre>" & vbCrLf

            xml = xml & "<ruta>" & rs!ruta & "</ruta>" & vbCrLf
            xml = xml & "</archivo>" & vbCrLf

        End While
        conn.Close()
        xml = xml & "</archivos>" & vbCrLf
        xml = xml & "<idDescarga><id>" & id_descarga & "</id></idDescarga>"
        xml = xml & "<mensajes>" & vbCrLf
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("select * from detdescargas with(nolock) where id_descarga=@id_descarga and tipomsg='E' order by id_log", conn)
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






    Private Sub ValidaArchivos(ByVal fecha As Date, ByVal Archivo As String)

        Try

            setLog(id_descarga, "I", "Validando archivos en FTP...")
            ftp.ServerAddress = server_ftp
            ftp.ServerPort = port_ftp
            ftp.UserName = user_ftp
            ftp.Password = pwd_ftp
            ftp.Connect()
            Dim blnExiste As Boolean = True

            If Archivo = "I" Or Archivo = "T" Then



                Dim strArchivoFtp As String = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IPI1.ToString & ".xls"

                blnExiste = ftp.Exists(strArchivoFtp)
                If blnExiste = True Then
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IPI2.ToString & ".xls"
                    blnExiste = ftp.Exists(strArchivoFtp)
                    If blnExiste = True Then
                        strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IPI3.ToString & ".xls"
                        blnExiste = ftp.Exists(strArchivoFtp)
                        If blnExiste = True Then
                            strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IPI4.ToString & ".xls"
                            blnExiste = ftp.Exists(strArchivoFtp)
                            If blnExiste = True Then
                                strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.COS1.ToString & ".xls"
                                blnExiste = ftp.Exists(strArchivoFtp)
                                If blnExiste = True Then
                                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.COS2.ToString & ".xls"
                                    blnExiste = ftp.Exists(strArchivoFtp)
                                    If blnExiste = True Then
                                        strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.COS3.ToString & ".xls"
                                        blnExiste = ftp.Exists(strArchivoFtp)
                                        If blnExiste = True Then
                                            strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.COS4.ToString & ".xls"
                                            blnExiste = ftp.Exists(strArchivoFtp)
                                            If blnExiste = True Then
                                                strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IVK1.ToString & ".xls"
                                                blnExiste = ftp.Exists(strArchivoFtp)
                                                If blnExiste = True Then
                                                    setLog(id_descarga, "I", "si existen los archivos..." & strArchivoFtp)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If


                If blnExiste = False Then
                    setLog(id_descarga, "I", "No existe archivo..." & strArchivoFtp)
                    setLog(id_descarga, "I", "Borrando archivos...")
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IPI1.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IPI2.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IPI3.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IPI4.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.COS1.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.COS2.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.COS3.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.COS4.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    strArchivoFtp = "PH_INV_" & fecha.ToString("ddMMyyyy") & "_" & Inv_PH.IVK1.ToString() & ".xls"
                    If ftp.Exists(strArchivoFtp) Then
                        ftp.DeleteFile(strArchivoFtp)
                    End If
                    setLog(id_descarga, "I", "Borrado terminado..." & strArchivoFtp)
                End If
            End If



            If Archivo = "V" Or Archivo = "T" Then
                Dim strArchivoFtp As String = "PH_VTA_" & fecha.ToString("ddMMyyyy") & "_" & CadenaProveedor & ".xls"

                blnExiste = ftp.Exists(strArchivoFtp)

                If blnExiste = False Then
                    setLog(id_descarga, "I", "No existe archivo..." & strArchivoFtp)
                    setLog(id_descarga, "I", "Borrarndo archivos...")
                    ftp.DeleteFile(strArchivoFtp)
                Else
                    setLog(id_descarga, "I", "si existe archivo..." & strArchivoFtp)
                End If

            End If
            setLog(id_descarga, "I", "Fin de validacion de archivo...")
            ftp.Close()

        Catch ex As Exception
            setLog(id_descarga, "E", "Excepcion:" & ex.Message)
        End Try

    End Sub


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
    Public Sub setArchivo(ByVal cliente As String, ByVal fecha As Date, ByVal archivo As String, ByVal ruta As String, ByVal tipo As String)
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("insert into archivos(fecha,cliente,archivo,ruta,tipo)values(@fecha,@cliente,@archivo,@ruta,@tipo)", conn)
        sql.CommandTimeout = 0
        sql.Parameters.AddWithValue("@fecha", fecha.ToString("ddMMyyyy"))
        sql.Parameters.AddWithValue("@cliente", cliente)
        sql.Parameters.AddWithValue("@archivo", archivo)
        sql.Parameters.AddWithValue("@ruta", ruta)
        sql.Parameters.AddWithValue("@tipo", tipo)
        sql.ExecuteNonQuery()
        conn.Close()
    End Sub

    Public Function getArchivo(ByVal cliente As String, ByVal fecha As Date, ByVal archivo As String) As Integer
        conn.Open()
        sql = New Data.SqlClient.SqlCommand("select count(*) as cont from archivos where cliente=@cliente and fecha=@fecha and archivo=@archivo", conn)
        sql.CommandTimeout = 0
        sql.Parameters.AddWithValue("@cliente", cliente)
        sql.Parameters.AddWithValue("@fecha", fecha.ToString("ddMMyyyy"))
        sql.Parameters.AddWithValue("@archivo", archivo)
        rs = sql.ExecuteReader
        rs.Read()
        Dim resx As Integer = rs!cont
        conn.Close()
        Return resx
    End Function
    Public Function getMesEng(ByVal mes As Integer) As String
        Dim val As String = ""
        Select Case mes
            Case 1
                val = "jan"
            Case 2
                val = "feb"
            Case 3
                val = "mar"
            Case 4
                val = "apr"
            Case 5
                val = "may"
            Case 6
                val = "jun"
            Case 7
                val = "jul"
            Case 8
                val = "aug"
            Case 9
                val = "sep"
            Case 10
                val = "oct"
            Case 11
                val = "nov"
            Case 12
                val = "dec"
        End Select
        Return val
    End Function
End Class


