Imports MySql.Data.MySqlClient
Public Class Resumenes
    Dim con As MySqlConnection
    Private Sub Resumenes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New MySqlConnection()
        con.ConnectionString = "server= localhost;user id= root; password=1111; database=tienda"
        CargarMeses()
        ListView1.Columns.Add("Fecha", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("FechaOrdn", 0, HorizontalAlignment.Left)
    End Sub
    Public Function CargarMeses() As Boolean
        ListView1.Items.Clear()
        'cargamos los proveedores
        Dim str(2) As String
        Dim itm As ListViewItem
        Dim proveedores As MySqlCommand = New MySqlCommand("select distinct mes,anio from fechas", con)
        proveedores.CommandType = CommandType.Text
        con.Open()
        Dim myReader As MySqlDataReader
        myReader = proveedores.ExecuteReader()
        Try
            While myReader.Read()
                str(0) = myReader.GetString(0) + "/" + myReader.GetString(1)
                str(1) = myReader.GetString(1) + rellenarCeros(myReader.GetString(0))
                itm = New ListViewItem(str)
                ListView1.Items.Add(itm)
                str(0) = ""
                str(1) = ""
            End While
        Finally
            myReader.Close()
            con.Close()
        End Try
        ListView1.ListViewItemSorter = New ListViewItemComparer(1)
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ListView1.FocusedItem Is Nothing Then
            MsgBox("Por favor seleccione una fecha")
        ElseIf ListView1.SelectedItems.Count = 0 Then
            MsgBox("Por favor seleccione una fecha")
        Else
            'variables para recoger el string
            Dim source As String = ListView1.SelectedItems.Item(0).Text.Trim()
            Dim Str() As String
            Str = source.Split("/")
            ' Set up all the parameters as generic objects so we can pass them in //
            Dim nombre As String
            nombre = "Estadistica" + Str(0) + "de" + Str(1) + ".doc"
            Dim wordApp As New Word.Application
            Dim fileName As Object = "normal.dot"
            ' template file name
            Dim newTemplate As Object = False
            Dim docType As Object = 0
            Dim isVisible As Object = True
            ' Create a new Document, by calling the Add function in the Documents collection
            Dim aDoc As Word.Document = wordApp.Documents.Add(fileName, newTemplate, docType, isVisible)
            wordApp.Visible = True
            'Insertamos encabezado
            '*********************
            wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            wordApp.Selection.Font.Bold = True
            wordApp.Selection.Font.Size = 15
            wordApp.Selection.TypeText(devolverMes(Str(0)) + " de " + Str(1))
            wordApp.Selection.TypeParagraph()
            wordApp.Selection.Font.Size = 12
            wordApp.Selection.Font.Bold = False
            wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            wordApp.Selection.TypeParagraph()
            wordApp.Selection.TypeParagraph()
            wordApp.Selection.TypeParagraph()


            'empezamos con las consultas!!
            '*********************
            con.Open()

            Dim i As Integer
            For i = 1 To DateTime.DaysInMonth(Str(1), Str(0))
                'consulta de dias
                '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                Dim cmd As MySqlCommand = New MySqlCommand("select * from fechas where dia=?val1 AND mes=?val2 AND anio =?val3 ;", con)
                cmd.CommandType = CommandType.Text
                Dim diasReader As MySqlDataReader
                Try
                    cmd.Parameters.AddWithValue("?val1", i)
                    cmd.Parameters.AddWithValue("?val2", Str(0))
                    cmd.Parameters.AddWithValue("?val3", Str(1))
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()
                    diasReader = cmd.ExecuteReader()
                    Dim dia As String = i
                    wordApp.Selection.TypeText("DIA " + dia + "  " + vbTab)
                    If diasReader.Read() Then

                        wordApp.Selection.TypeText("Venta: " + diasReader.GetString(3) + "  " + vbTab)
                        wordApp.Selection.TypeText("Caja: " + diasReader.GetString(4) + "  " + vbTab)
                        wordApp.Selection.TypeText("Cambio: " + diasReader.GetString(5) + "  " + vbTab)
                        'vbTab
                    End If
                    diasReader.Close()
                    'insertamos total de Pagos
                    '+++++++++++++++++++++++++++++++++++
                    Dim cmdPagos As MySqlCommand = New MySqlCommand("select round(sum(importe),2) from pagos where pendiente=?val3 AND dia=?val4 AND mes=?val5 AND anio=?val6 ;", con)
                    cmdPagos.CommandType = CommandType.Text
                    Dim pagosReader As MySqlDataReader
                    cmdPagos.Parameters.AddWithValue("?val3", "Pagado")
                    cmdPagos.Parameters.AddWithValue("?val4", i)
                    cmdPagos.Parameters.AddWithValue("?val5", Str(0))
                    cmdPagos.Parameters.AddWithValue("?val6", Str(1))
                    cmdPagos.Prepare()
                    cmdPagos.ExecuteNonQuery()
                    pagosReader = cmdPagos.ExecuteReader()
                    Dim pagos As Double = 0
                    Dim pagosCadena As String
                    If pagosReader.Read() Then
                        If pagosReader.IsDBNull(0) Then
                            pagos = 0
                        Else
                            pagos = pagosReader.GetString(0)
                            pagosCadena = pagos.ToString
                            wordApp.Selection.TypeText("Pagos: " + pagosCadena + "  " + vbTab)
                        End If
                    Else
                        pagos = 0
                    End If
                    pagosReader.Close()
                    'insertamos total de gastos
                    Dim cmdGastos As MySqlCommand = New MySqlCommand("select round(sum(importe),2) from gastos where dia=?val7 AND mes=?val8 AND anio=?val9 ;", con)
                    cmdGastos.CommandType = CommandType.Text
                    Dim gastosReader As MySqlDataReader
                    cmdGastos.Parameters.AddWithValue("?val7", i)
                    cmdGastos.Parameters.AddWithValue("?val8", Str(0))
                    cmdGastos.Parameters.AddWithValue("?val9", Str(1))
                    cmdGastos.Prepare()
                    cmdGastos.ExecuteNonQuery()
                    gastosReader = cmdGastos.ExecuteReader()
                    Dim gastos As Double = 0
                    Dim gastosCadena As String
                    If gastosReader.Read() Then
                        If gastosReader.IsDBNull(0) Then
                            gastos = 0
                        Else
                            gastos = gastosReader.GetString(0)
                            gastosCadena = gastos.ToString
                            wordApp.Selection.TypeText("Gastos: " + gastosCadena)

                        End If
                    Else
                        gastos = 0
                    End If
                    gastosReader.Close()
                Finally

                End Try
                wordApp.Selection.TypeText(vbNewLine)
                wordApp.Selection.TypeText(vbNewLine)
            Next
            wordApp.Selection.TypeText(vbNewLine)
            wordApp.Selection.Font.Bold = True
            'insertamos total de ventas
            Try
                Dim cmdTotalVentas As MySqlCommand = New MySqlCommand("select round(sum(venta),2) from fechas where mes=?val10 AND anio=?val11 ;", con)
                cmdTotalVentas.CommandType = CommandType.Text
                Dim TotalVentasReader As MySqlDataReader
                cmdTotalVentas.Parameters.AddWithValue("?val10", Str(0))
                cmdTotalVentas.Parameters.AddWithValue("?val11", Str(1))
                cmdTotalVentas.Prepare()
                cmdTotalVentas.ExecuteNonQuery()
                TotalVentasReader = cmdTotalVentas.ExecuteReader()
                Dim ventas As Double = 0
                Dim ventasCadena As String
                If TotalVentasReader.Read() Then
                    If TotalVentasReader.IsDBNull(0) Then
                        ventas = 0
                    Else
                        ventas = TotalVentasReader.GetString(0)
                        ventasCadena = ventas.ToString
                        wordApp.Selection.TypeText("TotalVentas: " + ventasCadena + "  " + vbTab)
                    End If
                Else
                    ventas = 0
                End If
                TotalVentasReader.Close()
                'insertamos total de Pagos
                Dim cmdTotalPagos As MySqlCommand = New MySqlCommand("select round(sum(importe),2) from pagos where mes=?val12 AND anio=?val13 ;", con)
                cmdTotalPagos.CommandType = CommandType.Text
                Dim TotalPagosReader As MySqlDataReader
                cmdTotalPagos.Parameters.AddWithValue("?val12", Str(0))
                cmdTotalPagos.Parameters.AddWithValue("?val13", Str(1))
                cmdTotalPagos.Prepare()
                cmdTotalPagos.ExecuteNonQuery()
                TotalPagosReader = cmdTotalPagos.ExecuteReader()
                Dim pagos As Double = 0
                Dim pagosCadena As String
                If TotalPagosReader.Read() Then
                    If TotalPagosReader.IsDBNull(0) Then
                        pagos = 0
                    Else
                        pagos = TotalPagosReader.GetString(0)
                        pagosCadena = pagos.ToString
                        wordApp.Selection.TypeText("TotalPagos: " + pagosCadena + "  " + vbTab)
                    End If
                Else
                    pagos = 0
                End If
                TotalPagosReader.Close()
                'insertamos total gastos

                Dim cmdTotalGastos As MySqlCommand = New MySqlCommand("select round(sum(importe),2) from gastos where mes=?val14 AND anio=?val15 ;", con)
                cmdTotalGastos.CommandType = CommandType.Text
                Dim TotalGastosReader As MySqlDataReader
                cmdTotalGastos.Parameters.AddWithValue("?val14", Str(0))
                cmdTotalGastos.Parameters.AddWithValue("?val15", Str(1))
                cmdTotalGastos.Prepare()
                cmdTotalGastos.ExecuteNonQuery()
                TotalGastosReader = cmdTotalGastos.ExecuteReader()
                Dim gastos As Double = 0
                Dim gastosCadena As String
                If TotalGastosReader.Read() Then
                    If TotalGastosReader.IsDBNull(0) Then
                        gastos = 0
                    Else
                        gastos = TotalGastosReader.GetString(0)
                        gastosCadena = gastos.ToString
                        wordApp.Selection.TypeText("TotalGastos: " + gastosCadena + "  " + vbTab)
                    End If
                Else
                    gastos = 0
                End If
                TotalGastosReader.Close()

            Catch ex As Exception
                MsgBox("error al calcular el computo total de ventas o pagos o gastos")
            End Try

            con.Close()

            '*********************
            'guardamos el archivo con el nombre del mes
            Try
                wordApp.ActiveDocument.SaveAs(FileName:=nombre)
            Catch
                MsgBox("error al guardar el archivo, compruebe que no se encuentre abierto otro con el mismo nombre")
            End Try

        End If

    End Sub
    Public Function devolverMes(ByVal numero As Integer) As String
        Dim mes As String = "??"
        If numero = 1 Then
            mes = "Enero"
        ElseIf numero = 2 Then
            mes = "Febrero"
        ElseIf numero = 3 Then
            mes = "Marzo"
        ElseIf numero = 4 Then
            mes = "Abril"
        ElseIf numero = 5 Then
            mes = "Mayo"
        ElseIf numero = 6 Then
            mes = "Junio"
        ElseIf numero = 7 Then
            mes = "Julio"
        ElseIf numero = 8 Then
            mes = "Agosto"
        ElseIf numero = 9 Then
            mes = "Septiembre"
        ElseIf numero = 10 Then
            mes = "Octubre"
        ElseIf numero = 11 Then
            mes = "Noviembre"
        ElseIf numero = 12 Then
            mes = "Diciembre"
        End If
        Return mes
    End Function
    Public Function rellenarCeros(ByVal numero As String) As String
        If numero.Length = 1 Then
            numero = "0" + numero
        End If
        Return numero
    End Function
End Class