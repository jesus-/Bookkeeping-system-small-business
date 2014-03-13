Imports MySql.Data.MySqlClient
Public Class TipoGastos
    Dim con As MySqlConnection
    Private Sub Gastos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New MySqlConnection()
        con.ConnectionString = "server= localhost;user id= root; password=1111; database=tienda"
        ListView1.Columns.Add("Nombre Gasto", 250, HorizontalAlignment.Left)
        CargarGastos()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        nuevoGasto()
        ListView1.Items.Clear()
        CargarGastos()
    End Sub
    Public Function nuevoGasto() As Boolean
        'Insertar nuevo proveedor
        con.Open()
        Dim cmd As MySqlCommand = New MySqlCommand("INSERT INTO tipoGastos VALUES (?val1 );", con)
        Try
            cmd.Parameters.AddWithValue("?val1", MaskedTextBox3.Text)
            cmd.Prepare()
            cmd.ExecuteNonQuery()

        Catch mierror As MySqlException
            MessageBox.Show(" El gasto que intentas meter ya existe " & mierror.Message)
        Finally
            con.Dispose()
        End Try
        con.Close()
        MaskedTextBox3.Text = ""

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
                ListView1.Items.Add(myReader.GetString(0))
            End While
        Finally
            myReader.Close()
            con.Close()
        End Try
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ListView1.FocusedItem Is Nothing Then
            MsgBox("Por favor seleccione un gasto para borrar")
        Else
            If ListView1.SelectedItems.Count = 0 Then
                MsgBox("Por favor seleccione un gasto para borrar")
            Else
                BorrarGasto()
                ListView1.Items.Clear()
                CargarGastos()
            End If
        End If
    End Sub
    Public Function BorrarGasto() As Boolean
        'Borramos un pago
        con.Open()
        Dim cmd As MySqlCommand = New MySqlCommand("Delete from tipoGastos where nombre=(?val1);", con)
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub
End Class
