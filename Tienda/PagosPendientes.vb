Imports MySql.Data.MySqlClient
Public Class PagosPendientes
    Dim con As MySqlConnection
    Private Sub PagosPendientes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New MySqlConnection()
        con.ConnectionString = "server= localhost;user id= root; password=1111; database=tienda"
        ListView1.Items.Clear()
        cargarPagosListView()
        ListView1.Columns.Add("Id", 30, HorizontalAlignment.Left)
        ListView1.Columns.Add("Proveedor", 155, HorizontalAlignment.Left)
        ListView1.Columns.Add("Importe", 60, HorizontalAlignment.Left)
        ListView1.Columns.Add("Situación", 80, HorizontalAlignment.Left)
        ListView1.Columns.Add("Fecha", 100, HorizontalAlignment.Left)
    End Sub
    Public Function cargarPagosListView() As Boolean
        'cargamos los pagos
        Dim str(5) As String
        Dim itm As ListViewItem
        Dim pagos As MySqlCommand = New MySqlCommand("select * from pagos  Where pendiente=""Pendiente"" ", con)
        pagos.CommandType = CommandType.Text
        con.Open()
        Dim myReader As MySqlDataReader
        myReader = pagos.ExecuteReader()
        Try
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
        Finally
            myReader.Close()
            con.Close()
            'mostramos el total de pagos echo en el dia

            Dim cmd As MySqlCommand = New MySqlCommand("select sum(importe) from pagos where pendiente=?val1 ;", con)
            Dim lector As MySqlDataReader
            con.Open()
            cmd.CommandType = CommandType.Text
            Try
                cmd.Parameters.AddWithValue("?val1", "Pendiente")
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                lector = cmd.ExecuteReader()
                lector.Read()
                
                If lector.IsDBNull(0) Then
                    Label8.Text = "0"
                Else
                    Label8.Text = lector.GetString(0)
                End If

            Catch mierror As MySqlException
                MessageBox.Show(" " & mierror.Message)
            Finally
                con.Dispose()
            End Try
            con.Close()
        End Try
    End Function

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

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
        Tienda.dia.ListView1.Items.Clear()
        Tienda.dia.cargarPagosListView()
    End Sub

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
        ListView1.Items.Clear()
        cargarPagosListView()
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()


    End Sub
End Class