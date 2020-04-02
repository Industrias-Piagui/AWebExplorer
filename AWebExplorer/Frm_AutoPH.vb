Imports System.Data.SqlClient

Public Class Frm_AutoPH

    Private Sub Frm_AutoPH_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        descargaPH()
    End Sub

    Private Sub descargaPH()
        Timer1.Stop()
        Dim aIDs As String()
        Dim x As String


        Dim fecha As DateTime = DateTime.Now.AddDays(-1)



        Try

        aIDs = ObtenerEjecucionesETL().ToArray()

        Dim enproceso As Integer = 0
        Dim archivos As Integer = 0
        Dim inv As Integer = 0
        Dim vta As Integer = 0
        Dim stipoarchivo As String



        enproceso = Convert.ToInt16(aIDs(0).ToString())
        archivos = Convert.ToInt16(aIDs(1).ToString())
        inv = Convert.ToInt16(aIDs(2).ToString())
        vta = Convert.ToInt16(aIDs(3).ToString())

   
        Dim a As New AWebExplorer("dba", "PortalesPROD", "aportales", "Aportales12")
            a.getFTP("PH")
            a.Visible = True

        If inv < 11 Then
            stipoarchivo = "INVENTARIOS"


                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.COS1)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.COS2)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.COS3)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.COS4)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.IPI1)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.IPI2)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.IPI3)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.IPI4)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.IVK1)
                x = a.getInventariosPH(fecha, stipoarchivo, AWebExplorer.Inv_PH.LIFES)




        End If



            If vta < 1 Then
                stipoarchivo = "VENTAS"
                x = a.getVentasPH(fecha, stipoarchivo)
            End If

        Catch ex As Exception

            Me.Close()
        End Try

        Me.Close()
    End Sub



    Private Shared Function ObtenerEjecucionesETL() As List(Of String)
        Dim oDate As DateTime = DateTime.Now
        Dim aIDs As New List(Of String)()
        Dim dt As DataTable = ObtenerEjecucionETL()

        For Each row As DataRow In dt.Rows

            aIDs.Add(row("EN_PROCESO").ToString())

            aIDs.Add(row("ARCHIVOS").ToString())

            aIDs.Add(row("INV").ToString())

            aIDs.Add(row("VTA").ToString())

        Next

 

        Return aIDs
    End Function




    Private Shared Function ObtenerEjecucionETL() As DataTable

        Dim message As String
        Dim en_proceso As Int16
        Dim archivos As Int16
        Dim inv As Int16
        Dim vta As Int16


        Dim dt1 As New DataTable("Datos")
        Dim dc1 As New DataColumn("en_proceso", Type.GetType("System.Int32"))
        Dim dc2 As New DataColumn("archivos", Type.GetType("System.Int32"))
        Dim dc3 As New DataColumn("inv", Type.GetType("System.Int32"))
        Dim dc4 As New DataColumn("vta", Type.GetType("System.Int32"))

        dt1.Columns.Add(dc1)
        dt1.Columns.Add(dc2)
        dt1.Columns.Add(dc3)
        dt1.Columns.Add(dc4)


        Dim conn1 As Data.SqlClient.SqlConnection
        conn1 = New Data.SqlClient.SqlConnection("Data Source=DBA;Initial Catalog=PortalesPROD;Trusted_Connection=no;User ID=mp001;Password=mp001;Connect Timeout=36000")







        conn1.Open()
        'Try

        ' tra = conn1.BeginTransaction



        Dim sql As New SqlClient.SqlCommand("Prc_RevisarArchivosETL", conn1)
        sql.CommandType = CommandType.StoredProcedure
        Dim rs As SqlClient.SqlDataReader
        rs = sql.ExecuteReader






        If rs.Read Then
            'MsgBox(rs!En_proceso)
            'MsgBox(rs!archivos)
            'MsgBox(rs!inv)
            'MsgBox(rs!vta)


            en_proceso = rs!En_proceso
            archivos = rs!archivos
            inv = rs!inv
            vta = rs!vta
            Dim newCustomersRow As DataRow = dt1.NewRow()



            newCustomersRow("en_proceso") = en_proceso
            newCustomersRow("archivos") = archivos
            newCustomersRow("inv") = inv
            newCustomersRow("vta") = vta


            dt1.Rows.Add(newCustomersRow)


        End If





        'Catch e As Exception
        '    message = e.Message()
        'Finally
        '    conn1.Close()
        'End Try

        Return dt1
    End Function

End Class