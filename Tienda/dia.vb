Imports MySql.Data.MySqlClient
Public Class dia
    Public dia As Integer = 20
    Public mes As Integer = 12
    Public año As Integer = 1984
    Dim con As MySqlConnection
    Private Sub dia_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
        con = New MySqlConnection()
        con.ConnectionString = "server= localhost;user id= root; password=1111; database=tienda"
        Label8.Text = "0" + " €"
        mostrarVentas()
        mostrarCaja()
        mostrarCambio()
        CargarProveedores()
        CargarGastos()
        cargarPagosListView()
        ListView1.Columns.Add("Id", 30, HorizontalAlignment.Left)
        ListView1.Columns.Add("Proveedor", 155, HorizontalAlignment.Left)
        ListView1.Columns.Add("Importe", 60, HorizontalAlignment.Left)
        ListView1.Columns.Add("Situación", 80, HorizontalAlignment.Left)
        ListView1.Columns.Add("Fecha", 100, HorizontalAlignment.Left)
    End Sub

    Private Sub MaskedTextBox3_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs)

        MaskedTextBox3.Text = ""
    End Sub


    Private Sub MaskedTextBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MaskedTextBox3.Text = ""
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        InsertarPago()
        ComboBox1.SelectedIndex = "-1"
    End Sub
    Private Sub EnviarDia_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnviarDia.Click
        nuevaVenta()
        nuevoCambio()
        nuevaCaja()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If ListView1.FocusedItem Is Nothing Then
            MsgBox("Por favor seleccione un pago para borrar")
        Else
            If ListView1.SelectedItems.Count = 0 Then
                MsgBox("Por favor seleccione un pago para borrar")
            Else
                BorrarPago()
                ListView1.Items.Clear()
                cargarPagosListView()
            End If
        End If


    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Tienda.PagosPendientes.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()

    End Sub
    Public Function InsertarPago() As Boolean
        'insertar un nuevo pago
        con.Open()

        Dim cmd As MySqlCommand = New MySqlCommand("INSERT INTO pagos VALUES ( NULL, ?val2, ?val3, ?val4, ?val5 , ?val6,?val7);", con)
        Try

            If CheckBox2.Checked Then
                cmd.Parameters.AddWithValue("?val4", "Pendiente")
            Else
                cmd.Parameters.AddWithValue("?val4", "Pagado")
            End If
            CheckBox2.CheckState = 0
            cmd.Parameters.AddWithValue("?val2", MaskedTextBox1.Text)
            cmd.Parameters.AddWithValue("?val3", ComboBox1.Text)
            cmd.Parameters.AddWithValue("?val5", dia)
            cmd.Parameters.AddWithValue("?val6", mes)
            cmd.Parameters.AddWithValue("?val7", año)
            cmd.Prepare()
            cmd.Prepare()
            cmd.ExecuteNonQuery()

        Catch mierror As MySqlException
            MessageBox.Show("Problema!! ha introducido caracteres en la casilla importe")
        Finally
            con.Dispose()
        End Try
        con.Close()
        MaskedTextBox3.Text = ""
        MaskedTextBox1.Text = ""
        ComboBox1.Text = ""
        ListView1.Items.Clear()
        cargarPagosListView()
    End Function
    Public Function cargarPagosListView() As Boolean
        'cargamos los pagos
        Dim str(5) As String
        Dim itm As ListViewItem
        Dim pagos As MySqlCommand = New MySqlCommand("select * from pagos where dia=?val1 AND mes=?val2 AND anio=?val3 ORDER BY proveedor", con)
        pagos.CommandType = CommandType.Text
        con.Open()
        Dim myReader As MySqlDataReader

        Try
            pagos.Parameters.AddWithValue("?val1", dia)
            pagos.Parameters.AddWithValue("?val2", mes)
            pagos.Parameters.AddWithValue("?val3", año)
            pagos.Prepare()
            pagos.ExecuteNonQuery()
            myReader = pagos.ExecuteReader()
            While myReader.Read()
                str(0) = myReader.GetString(0)
                str(1) = myReader.GetString(2)
                str(2) = myReader.GetString(1)
                str(3) = myReader.GetString(3)
                str(4) = myReader.GetString(4) + "/" + myReader.GetString(5) + "/" + myReader.GetString(6)
                itm = New ListViewItem(str)
                ListView1.Items.Add(itm)
                str(0) = ""
                str(1) = ""
                str(2) = ""
                str(3) = ""
                str(4) = ""
            End While
            myReader.Close()
        Finally

            con.Close()
            'mostramos el total de pagos echo en el dia

            Dim cmd As MySqlCommand = New MySqlCommand("select round(sum(importe),2) from pagos where pendiente=?val1 AND dia=?val2 AND mes=?val3 AND anio=?val4 ;", con)
            Dim lector As MySqlDataReader
            con.Open()
            cmd.CommandType = CommandType.Text
            Try
                cmd.Parameters.AddWithValue("?val1", "Pagado")
                cmd.Parameters.AddWithValue("?val2", dia)
                cmd.Parameters.AddWithValue("?val3", mes)
                cmd.Parameters.AddWithValue("?val4", año)
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                lector = cmd.ExecuteReader()
                If lector.Read() Then
                    If lector.IsDBNull(0) Then
                        Label8.Text = "0" + " €"
                    Else
                        Label8.Text = lector.GetString(0) + " €"
                    End If
                Else
                    Label8.Text = "0" + " €"
                End If

            Catch mierror As MySqlException
                MessageBox.Show(" " & mierror.Message)
            Finally
                con.Dispose()
            End Try
            con.Close()
        End Try
    End Function
    Public Function nuevaVenta() As Boolean
        'insertamos venta
        If MaskedTextBox2.Text.Length = 0 Then
            'no hacer nada
        Else
            If Label9.Text.Trim = 0 Then
                con.Open()
                Dim cmd As MySqlCommand = New MySqlCommand("update fechas set venta=?val1 where dia =?val2 AND mes=?val3 AND anio = ?val4;", con)
                Try
                    cmd.Parameters.AddWithValue("?val1", MaskedTextBox2.Text)
                    cmd.Parameters.AddWithValue("?val2", dia)
                    cmd.Parameters.AddWithValue("?val3", mes)
                    cmd.Parameters.AddWithValue("?val4", año)
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()

                Catch mierror As MySqlException
                    MessageBox.Show(" " & mierror.Message)
                Finally
                    con.Dispose()
                End Try
                con.Close()
                Label9.Text = MaskedTextBox2.Text + " €"
                MaskedTextBox2.Text = ""
            Else
                con.Open()
                Dim cmd As MySqlCommand = New MySqlCommand("update fechas set venta=?val1 where dia=?val2 AND mes=?val3 AND anio=?val4;", con)
                Try
                    cmd.Parameters.AddWithValue("?val1", MaskedTextBox2.Text)
                    cmd.Parameters.AddWithValue("?val2", dia)
                    cmd.Parameters.AddWithValue("?val3", mes)
                    cmd.Parameters.AddWithValue("?val4", año)
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()

                Catch mierror As MySqlException
                    MessageBox.Show(" " & mierror.Message)
                Finally
                    con.Dispose()
                End Try
                con.Close()
                Label9.Text = MaskedTextBox2.Text + " €"
                MaskedTextBox2.Text = ""
            End If

        End If

    End Function
    Public Function nuevaCaja() As Boolean
        'insertamos caja
        If MaskedTextBox4.Text.Length = 0 Then
            'no hacer nada
        Else
            If Label12.Text.Trim = 0 Then
                con.Open()
                Dim cmd As MySqlCommand = New MySqlCommand("update fechas set caja=?val1 where dia =?val2 AND mes=?val3 AND anio = ?val4;", con)
                Try
                    cmd.Parameters.AddWithValue("?val1", MaskedTextBox4.Text)
                    cmd.Parameters.AddWithValue("?val2", dia)
                    cmd.Parameters.AddWithValue("?val3", mes)
                    cmd.Parameters.AddWithValue("?val4", año)
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()

                Catch mierror As MySqlException
                    MessageBox.Show(" " & mierror.Message)
                Finally
                    con.Dispose()
                End Try
                con.Close()
                Label12.Text = MaskedTextBox4.Text + " €"
                MaskedTextBox4.Text = ""
            Else
                con.Open()
                Dim cmd As MySqlCommand = New MySqlCommand("update fechas set caja=?val1 where dia=?val2 AND mes=?val3 AND anio=?val4;", con)
                Try
                    cmd.Parameters.AddWithValue("?val1", MaskedTextBox4.Text)
                    cmd.Parameters.AddWithValue("?val2", dia)
                    cmd.Parameters.AddWithValue("?val3", mes)
                    cmd.Parameters.AddWithValue("?val4", año)
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()

                Catch mierror As MySqlException
                    MessageBox.Show(" " & mierror.Message)
                Finally
                    con.Dispose()
                End Try
                con.Close()
                Label12.Text = MaskedTextBox4.Text + " €"
                MaskedTextBox4.Text = ""
            End If

        End If

    End Function
    Public Function nuevoCambio() As Boolean
        'insertamos cambio
        If MaskedTextBox5.Text.Length = 0 Then
            'no hacer nada
        Else
            If Label13.Text.Trim = 0 Then
                con.Open()
                Dim cmd As MySqlCommand = New MySqlCommand("update fechas set caja=?val1 where dia =?val2 AND mes=?val3 AND anio = ?val4;", con)
                Try
                    cmd.Parameters.AddWithValue("?val1", MaskedTextBox5.Text)
                    cmd.Parameters.AddWithValue("?val2", dia)
                    cmd.Parameters.AddWithValue("?val3", mes)
                    cmd.Parameters.AddWithValue("?val4", año)
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()

                Catch mierror As MySqlException
                    MessageBox.Show(" " & mierror.Message)
                Finally
                    con.Dispose()
                End Try
                con.Close()
                Label13.Text = MaskedTextBox5.Text + " €"
                MaskedTextBox5.Text = ""
            Else
                con.Open()
                Dim cmd As MySqlCommand = New MySqlCommand("update fechas set cambio=?val1 where dia=?val2 AND mes=?val3 AND anio=?val4;", con)
                Try
                    cmd.Parameters.AddWithValue("?val1", MaskedTextBox5.Text)
                    cmd.Parameters.AddWithValue("?val2", dia)
                    cmd.Parameters.AddWithValue("?val3", mes)
                    cmd.Parameters.AddWithValue("?val4", año)
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()

                Catch mierror As MySqlException
                    MessageBox.Show(" " & mierror.Message)
                Finally
                    con.Dispose()
                End Try
                con.Close()
                Label13.Text = MaskedTextBox5.Text + " €"
                MaskedTextBox5.Text = ""
            End If
        End If

    End Function

    Public Function BorrarPago() As Boolean
        'Borramos un pago
        con.Open()
        Dim cmd As MySqlCommand = New MySqlCommand("Delete from pagos where id_pago=(?val1);", con)
        Try
            cmd.Parameters.AddWithValue("?val1", ListView1.SelectedItems.Item(0).Text.Trim)
            cmd.Prepare()
            cmd.ExecuteNonQuery()

        Catch mierror As MySqlException
            MessageBox.Show(" " & mierror.Message)
        Finally
            con.Dispose()
        End Try
        con.Close()
    End Function

    Public Function mostrarVentas() As Boolean
        Dim cmd As MySqlCommand = New MySqlCommand("select venta from fechas where dia=?val1 AND mes=?val2 AND anio=?val3 ;", con)
        Dim lector As MySqlDataReader
        con.Open()
        cmd.CommandType = CommandType.Text
        Try
            cmd.Parameters.AddWithValue("?val1", dia)
            cmd.Parameters.AddWithValue("?val2", mes)
            cmd.Parameters.AddWithValue("?val3", año)
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            lector = cmd.ExecuteReader()


            If lector.Read() Then
                If lector.IsDBNull(0) Then
                    Label9.Text = "0 " + "€"
                Else
                    Label9.Text = lector.GetString(0) + " €"
                End If
            Else
                Label9.Text = "0 " + "€"
            End If
        Catch mierror As MySqlException
            MessageBox.Show("???? " & mierror.Message)
        Finally
            con.Dispose()
        End Try
        con.Close()
    End Function

    Function mostrarCambio() As Boolean
        Dim cmd As MySqlCommand = New MySqlCommand("select cambio from fechas where dia=?val1 AND mes=?val2 AND anio=?val3 ;", con)
        Dim lector As MySqlDataReader
        con.Open()
        cmd.CommandType = CommandType.Text
        Try
            cmd.Parameters.AddWithValue("?val1", dia)
            cmd.Parameters.AddWithValue("?val2", mes)
            cmd.Parameters.AddWithValue("?val3", año)
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            lector = cmd.ExecuteReader()


            If lector.Read() Then
                If lector.IsDBNull(0) Then
                    Label12.Text = "0" + " €"
                Else
                    Label12.Text = lector.GetString(0) + " €"
                End If

            Else
                Label12.Text = "0" + " €"
            End If
        Catch mierror As MySqlException
            MessageBox.Show(" " & mierror.Message)
        Finally
            con.Dispose()
        End Try
        con.Close()
    End Function
    Function mostrarCaja() As Boolean
        Dim cmd As MySqlCommand = New MySqlCommand("select caja from fechas where dia=?val1 AND mes=?val2 AND anio=?val3 ;", con)
        Dim lector As MySqlDataReader
        con.Open()
        cmd.CommandType = CommandType.Text
        Try
            cmd.Parameters.AddWithValue("?val1", dia)
            cmd.Parameters.AddWithValue("?val2", mes)
            cmd.Parameters.AddWithValue("?val3", año)
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            lector = cmd.ExecuteReader()


            If lector.Read() Then
                If lector.IsDBNull(0) Then
                    Label13.Text = "0" + " €"
                Else
                    Label13.Text = lector.GetString(0) + " €"
                End If

            Else
                Label13.Text = "0" + " €"
            End If
        Catch mierror As MySqlException
            MessageBox.Show(" " & mierror.Message)
        Finally
            con.Dispose()
        End Try
        con.Close()
    End Function
    Public Function CargarProveedores() As Boolean
        'cargamos los proveedores
        Dim proveedores As MySqlCommand = New MySqlCommand("select * from proveedores ORDER BY nombre", con)
        proveedores.CommandType = CommandType.Text
        con.Open()
        Dim myReader As MySqlDataReader
        myReader = proveedores.ExecuteReader()
        Try
            While myReader.Read()
                ComboBox1.Items.Add(myReader.GetString(0))
            End While
        Finally
            myReader.Close()
            con.Close()
        End Try
    End Function
    Public Function CargarGastos() As Boolean
        'cargamos los proveedores
        Dim gastos As MySqlCommand = New MySqlCommand("select * from tipoGastos ORDER BY nombre", con)
        gastos.CommandType = CommandType.Text
        con.Open()
        Dim myReader As MySqlDataReader
        myReader = gastos.ExecuteReader()
        Try
            While myReader.Read()
                ComboBox2.Items.Add(myReader.GetString(0))
            End While
        Finally
            myReader.Close()
            con.Close()
        End Try
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        InsertarGasto()
        ComboBox2.SelectedIndex = "-1"

    End Sub
    Public Function InsertarGasto() As Boolean
        'insertar un nuevo pago
        con.Open()
        Dim cmd As MySqlCommand = New MySqlCommand("INSERT INTO gastos VALUES ( NULL, ?val2, ?val3, ?val4, ?val5 , ?val6);", con)
        Try
            cmd.Parameters.AddWithValue("?val2", MaskedTextBox3.Text)
            cmd.Parameters.AddWithValue("?val3", ComboBox2.Text)
            cmd.Parameters.AddWithValue("?val4", dia)
            cmd.Parameters.AddWithValue("?val5", mes)
            cmd.Parameters.AddWithValue("?val6", año)
            cmd.Prepare()
            cmd.Prepare()
            cmd.ExecuteNonQuery()
        Catch mierror As MySqlException
            MessageBox.Show("Problema!! ha introducido caracteres en la casilla importe")
        Finally
            con.Dispose()
        End Try
        con.Close()
        MaskedTextBox3.Text = ""
        ComboBox2.Text = ""
    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Tienda.Gastos.Show()
    End Sub
    Public Function setConst(ByVal dia1 As Integer, ByVal mes1 As Integer, ByVal año1 As Integer) As Boolean
        dia = dia1
        mes = mes1
        año = año1
    End Function
End Class
