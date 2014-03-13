Imports MySql.Data.MySqlClient
Public Class proveedores
    Dim con As MySqlConnection
    Private Sub proveedores_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New MySqlConnection()
        con.ConnectionString = "server= localhost;user id= root; password=1111; database=tienda"
        ListView1.Columns.Add("Nombre Proveedor", 250, HorizontalAlignment.Left)
        CargarProveedores()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        nuevoProveedor()
        ListView1.Items.Clear()
        CargarProveedores()
    End Sub
    Public Function nuevoProveedor() As Boolean
        'Insertar nuevo proveedor
        con.Open()
        Dim cmd As MySqlCommand = New MySqlCommand("INSERT INTO proveedores VALUES (?val1 );", con)
        Try
            cmd.Parameters.AddWithValue("?val1", MaskedTextBox3.Text)
            cmd.Prepare()
            cmd.ExecuteNonQuery()

        Catch mierror As MySqlException
            MessageBox.Show(" El proveedor que intentas meter ya existe " & mierror.Message)
        Finally
            con.Dispose()
        End Try
        con.Close()
        MaskedTextBox3.Text = ""

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
                ListView1.Items.Add(myReader.GetString(0))
            End While
        Finally
            myReader.Close()
            con.Close()
        End Try
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ListView1.FocusedItem Is Nothing Then
            MsgBox("Por favor seleccione un proveedor para borrar")
        Else
            If ListView1.SelectedItems.Count = 0 Then
                MsgBox("Por favor seleccione un proveedor para borrar")
            Else
                BorrarPago()
                ListView1.Items.Clear()
                CargarProveedores()
            End If
        End If
    End Sub
    Public Function BorrarPago() As Boolean
        'Borramos un pago
        con.Open()
        Dim cmd As MySqlCommand = New MySqlCommand("Delete from proveedores where nombre=(?val1);", con)
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
