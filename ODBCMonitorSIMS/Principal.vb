Imports System.Xml
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data
Imports System.Net.Mail
Imports System.Text.RegularExpressions

Module Principal

    Public conexionString As String = ""
    Public conexionStringPedidos As String = ""
    Public comando As New SqlClient.SqlCommand
    Public conexion As SqlClient.SqlConnection


    Public nombrePartner As String
    Dim logger As StreamWriter
    Dim lineaLogger As String
    Dim prefijo As String
    Public cadenaAS400 As String
    Public cadenaAS400CTL As String



    Sub Main()

        monitorear()

    End Sub


    Private Sub monitorear()

        Dim conEx As New SqlConnection
        Dim conIn As New SqlConnection
        Dim cmdEx As New SqlCommand
        Dim cmdIn As New SqlCommand



        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand


        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
                System.IO.File.Delete(file_log_path & "\log.txt")
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()
            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")
            prefijo = diccionario.Item("prefijo_archivo")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"
            nombrePartner = diccionario.Item("empresa")
            cadenaAS400 = diccionario.Item("DSN1")
            cadenaAS400CTL = diccionario.Item("DSN2")

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try

        ejecutarSQL("delete from personal_activo")

        Try
            ' Open the file using a stream reader.
            Using sr As New StreamReader("FNM079.txt")
                Dim line As String
                While sr.Read
                    'line = sr.ReadToEnd()
                    line = sr.ReadLine()
                    Dim s As String() = Regex.Split(line, ";")
                    ejecutarSQL("insert into personal_activo values('" + Trim(s(1)) + "') ")
                End While
            End Using
        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        End Try


        Try
            conEx.ConnectionString = conexionString
            conEx.Open()

            conIn.ConnectionString = conexionString
            conIn.Open()

            cmdEx.Connection = conEx
            cmdEx.CommandText = "select * from paciente order by id_paciente"

            cmdIn.Connection = conIn


            Dim lrdEx As SqlDataReader = cmdEx.ExecuteReader()
            Dim lrdIn As SqlDataReader
            Dim pedidos As String
            Dim cedula As String
            Dim cedulaAux As String
            pedidos = ""

            While lrdEx.Read()

                cedula = lrdEx.Item("id_paciente")
                cedulaAux = lrdEx.Item("id_paciente")

                If InStr(cedula, "-") > 0 Then
                    cedula = Mid(cedula, 1, InStr(cedula, "-") - 1)
                End If


                cmdIn.CommandText = "select * from personal_activo where cedpac='" + cedula + "'"
                lrdIn = cmdIn.ExecuteReader()
                If lrdIn.FieldCount > 0 Then
                    ejecutarSQL("UPDATE paciente SET estatus=1 where id_paciente = '" + Trim(cedula) + "'") ' COLOCA COMO ACTIVO AL TRABAJADOR POR DEFAULT, LOS FAMILIARES DEBEN ACTIVARSE UNO POR UNO POR SIMS
                    ejecutarSQL("UPDATE paciente SET estatus=1 where id_paciente like '%" + Trim(cedula) + "%' and estatus<>0  ") ' SI EL PACIENTE FUE INACTIVADO LA INTERFAZ NO LO ACTIVARA DE NUEVO
                    escribirLog("FECHA: " & DateTime.Now.ToString("dd/MM/YYYY") & ";HORA: " & DateTime.Now.ToString("HH:mm:ss") & ";" & cedulaAux & ";Estado:1", "")
                Else
                    ejecutarSQL("UPDATE paciente SET estatus=0 where id_paciente like '%" + Trim(cedula) + "%'")
                    escribirLog("FECHA: " & DateTime.Now.ToString("dd/MM/YYYY") & ";HORA: " & DateTime.Now.ToString("HH:mm:ss") & ";" & cedulaAux & ";Estado:0", "")
                End If

                If Not lrdIn.IsClosed Then
                    lrdIn.Close()
                End If


            End While
            conEx.Close()
            conIn.Close()



        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(PO) ")
        Finally
            conEx.Close()

        End Try

    End Sub

    Private Function ejecutarSQL(sql As String) As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As Integer
        Dim destinatarios As String
        destinatarios = ""


        Try

            conEx2.ConnectionString = conexionString
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = sql
            cmdEx2.ExecuteNonQuery()

        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try

        Return valor
    End Function




    Private Function buscarPedido(ByVal pedido As Integer) As Boolean

        Dim cnn As New Odbc.OdbcConnection(cadenaAS400)
        Dim rs As New Odbc.OdbcCommand("select ZFTRNK from F0041Z1 where ZFTRNK like '" & pedido & "%'    ", cnn)
        Dim reader As Odbc.OdbcDataReader
        Dim valor As Boolean
        valor = False
        cnn.Open()

        reader = rs.ExecuteReader
        While reader.Read()
            'valor = Trim(reader("IMUOM1"))
            valor = True
        End While

        cnn.Close()

        Return valor
    End Function


    Private Function obtenerDestinatarios() As String

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As Integer
        Dim sqlstring As String
        Dim destinatarios As String
        destinatarios = ""

        Try

            conEx2.ConnectionString = conexionString
            conEx2.Open()
            cmdEx2.Connection = conEx2
            cmdEx2.CommandText = "SELECT destinatarios FROM SYSTEM "
            Dim lrdEx2 As SqlDataReader = cmdEx2.ExecuteReader()

            While lrdEx2.Read()
                destinatarios = lrdEx2.GetString("0")
            End While

            Return destinatarios


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try

        Return valor
    End Function

    Private Sub EnvioCorreo(ByVal sender As String, ByVal password As String, ByVal recipients As String, ByVal pedidos As String)

        Dim correo As New MailMessage
        Dim smtp As New SmtpClient()

        Try
            correo.From = New MailAddress(sender, "monitorPedidos", System.Text.Encoding.UTF8)
            correo.To.Add(recipients)
            correo.SubjectEncoding = System.Text.Encoding.UTF8
            correo.Subject = "Tiene un nuevo Reporte de Pedidos no transferidos a JDE"
            correo.Body = "Los siguientes pedidos no estan en JDE y seran colocados como pendientes para ser sincronizados en la proxima ejecucion del proceso de interfaz:" & vbCrLf & Environment.NewLine() & pedidos & Environment.NewLine() & "Este mensaje fue enviado automaticamente solo con motivos de notificación, no es monitoreado ni acepta correos entrantes.Por favor, no responda este correo."
            correo.BodyEncoding = System.Text.Encoding.UTF8
            correo.IsBodyHtml = False
            correo.Priority = MailPriority.High
            smtp.Credentials = New System.Net.NetworkCredential(sender, password)
            smtp.Port = 2525
            smtp.Host = "172.23.20.66"
            'smtp.EnableSsl = True
            smtp.UseDefaultCredentials = False
            smtp.Send(correo)

        Catch ex As Exception

            Dim co As String
            co = ex.ToString

        End Try

    End Sub


    Private Function actualizarPedido(ByVal pedido As Integer) As Integer

        Dim conEx2 As New SqlConnection
        Dim cmdEx2 As New SqlCommand
        Dim valor As Integer
        Dim sqlstring As String

        Try

            conEx2.ConnectionString = conexionString
            conEx2.Open()

            sqlstring = ""
            sqlstring = "UPDATE ORDERS SET STATUS='PEN' WHERE ORDER_ID='" & pedido & "' "
            comando.Connection = conEx2
            comando.CommandText = " "
            comando.CommandText = sqlstring
            comando.ExecuteNonQuery()


        Catch ex As Exception
            escribirLog(ex.Message.ToString, "(OUTB) ")
        Finally
            conEx2.Close()

        End Try



        Return valor
    End Function





    Public Sub escribirLog(ByVal mensaje As String, ByVal proceso As String)

        Dim time As DateTime = DateTime.Now
        Dim format As String = "dd/MM/yyyy HH:mm "

        lineaLogger = proceso & time.ToString(format) & ":" & mensaje & vbNewLine
        logger.WriteLine(lineaLogger)
        logger.Flush()

    End Sub

    Private Function obtenerFecha() As String

        Dim time As DateTime = DateTime.Now
        Dim format As String = "yyyyMMdd"
        Return time.ToString(format)

    End Function

    Private Function obtenerFechaHora() As String

        Dim time As DateTime = DateTime.Now
        Dim format As String = "yyyyMMddHHmmss"
        Return time.ToString(format)

    End Function

    Private Function obtenerHora() As String

        Dim time As DateTime = DateTime.Now
        Dim format As String = "HHmmss"
        Return time.ToString(format)

    End Function

    Private Function obtenerHoraPedido() As String

        Dim time As DateTime = DateTime.Now
        Dim format As String = "HH:mm:ss"
        Return time.ToString(format)

    End Function




    Public Function obtenerNodosHijosDePadre(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As Dictionary(Of String, String)
        Dim diccionario As New Dictionary(Of String, String)
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                If Not diccionario.ContainsKey(nodoPadre(i).ChildNodes.Item(h).Name.Trim()) Then
                    diccionario.Add(nodoPadre(i).ChildNodes.Item(h).Name.Trim(), nodoPadre(i).ChildNodes.Item(h).InnerText.Trim())
                End If
            Next
        Next
        Return diccionario
    End Function


    Public Function obtenerNodosHijosDePadreLista(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As List(Of Dictionary(Of String, String))

        Dim listaDiccionario As New List(Of Dictionary(Of String, String))
        Dim diccionario As New Dictionary(Of String, String)
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            diccionario = New Dictionary(Of String, String)
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                If Not diccionario.ContainsKey(nodoPadre(i).ChildNodes.Item(h).Name.Trim()) Then
                    diccionario.Add(nodoPadre(i).ChildNodes.Item(h).Name.Trim(), nodoPadre(i).ChildNodes.Item(h).InnerText.Trim())
                End If
            Next
            listaDiccionario.Add(diccionario)
        Next
        Return listaDiccionario

    End Function

    Public Function obtenerListaNodosHijosDePadre(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As List(Of Nodo)

        Dim listaNodos As New List(Of Nodo)
        Dim nodo As Nodo
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                nodo = New Nodo()
                nodo.sName = nodoPadre(i).ChildNodes.Item(h).Name.Trim()
                nodo.sInner = nodoPadre(i).ChildNodes.Item(h).InnerText.Trim()
                listaNodos.Add(nodo)
            Next
        Next
        Return listaNodos

    End Function


    Private Function buscarNodo(ByVal name As String, ByVal listaNodos As List(Of Nodo)) As Nodo

        Dim nodo As Nodo = New Nodo()
        Dim encontrado As Boolean
        Dim nodos_enumerator As IEnumerator
        nodos_enumerator = listaNodos.GetEnumerator()
        encontrado = False

        nodo.sName = "NULL"

        Do While (nodos_enumerator.MoveNext) And Not encontrado
            nodo = CType(nodos_enumerator.Current, Nodo)
            If nodo.sName.CompareTo(name) = 0 Then
                encontrado = True
            Else
                nodo.sName = "NULL"
            End If

        Loop

        buscarNodo = nodo

    End Function

    Private Function obtenerNombreArchivo(ByVal directorio As String, ByVal nombreBase As String) As List(Of String)


        Dim listaArchivos As New List(Of String)
        Dim diccionario As New Dictionary(Of String, String)
        Dim strFileSize As String = ""
        Dim nombreArchivo As String
        nombreArchivo = ""

        Try
            Dim di As New IO.DirectoryInfo(directorio)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
            Dim fi As IO.FileInfo

            For Each fi In aryFi
                If InStr(fi.Name, nombreBase) > 0 Then
                    diccionario = New Dictionary(Of String, String)
                    nombreArchivo = fi.Name
                    nombreArchivo.Concat(".xml")

                    listaArchivos.Add(nombreArchivo)
                    'Exit For
                Else
                    nombreArchivo = ""
                End If

            Next

        Catch ex As Exception

            lineaLogger = "Línea de texto " & vbNewLine & "Otra linea de texto"
            logger.WriteLine(lineaLogger)
            logger.Flush()

        Finally

        End Try
        Return listaArchivos

    End Function






    Private Sub leerXML()
        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer
        Dim str As String
        Dim fs As New FileStream("products.xml", FileMode.Open, FileAccess.Read)

        xmldoc.Load(fs)
        xmlnode = xmldoc.GetElementsByTagName("Product")
        For i = 0 To xmlnode.Count - 1
            'xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
            str = xmlnode(i).ChildNodes.Item(0).Name.Trim() & " | " & xmlnode(i).ChildNodes.Item(1).InnerText.Trim() & " | " & xmlnode(i).ChildNodes.Item(2).InnerText.Trim()
            MsgBox(str)
        Next
    End Sub



    Private Sub crearXML()
        Dim writer As New XmlTextWriter("product.xml", System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("Table")
        createNode(1, "Product 1", "1000", writer)
        createNode(2, "Product 2", "2000", writer)
        createNode(3, "Product 3", "3000", writer)
        createNode(4, "Product 4", "4000", writer)
        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()
    End Sub


    Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal pPrice As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement("Product")
        writer.WriteStartElement("Product_id")
        writer.WriteString(pID)
        writer.WriteEndElement()
        writer.WriteStartElement("Product_name")
        writer.WriteString(pName)
        writer.WriteEndElement()
        writer.WriteStartElement("Product_price")
        writer.WriteString(pPrice)
        writer.WriteEndElement()
        writer.WriteEndElement()
    End Sub

    Private Sub buscar()
        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create("Product.xml", New XmlReaderSettings())
        Dim ds As New DataSet
        Dim dv As DataView
        ds.ReadXml(xmlFile)

        dv = New DataView(ds.Tables(0))
        dv.Sort = "Product_Name"
        Dim index As Integer = dv.Find("Product 2")

        If index = -1 Then
            MsgBox("Item Not Found")
        Else
            MsgBox(dv(index)("Product_Name").ToString() & "  " & dv(index)("Product_Price").ToString())
        End If
    End Sub

    Private Sub filtrar()

        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create("Product.xml", New XmlReaderSettings())
        Dim ds As New DataSet
        Dim dv As DataView
        ds.ReadXml(xmlFile)
        dv = New DataView(ds.Tables(0), "Product_price > = 3000", "Product_Name", DataViewRowState.CurrentRows)
        dv.ToTable().WriteXml("Result.xml")
        MsgBox("Done")
    End Sub

    Private Sub buscarPorTag()


        ' Open the XML file
        Dim xmlDocContinents As New XmlDocument
        xmlDocContinents.Load("Product.xml")

        ' Get a list of elements whose names are Continent
        Dim lstContinents As XmlNodeList = xmlDocContinents.GetElementsByTagName("Product")

        ' Retrieve the name of each continent and put it in the combo box
        Dim i As Integer
        For i = 0 To lstContinents.Count Step 1

            MsgBox(lstContinents(i).Attributes("Product_name").InnerText)

        Next
    End Sub

End Module
