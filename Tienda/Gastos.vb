Imports MySql.Data.MySqlClient

Public Class Gastos
    Dim con As MySqlConnection
    Private Sub Gastos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        con = New MySqlConnection()
        con.ConnectionString = "server= localhost;user id= root; password=1111; database=tienda"
        ListView1.Columns.Add("Id", 30, HorizontalAlignment.Left)
        ListView1.Columns.Add("TipoGasto", 155, HorizontalAlignment.Left)
        ListView1.Columns.Add("Importe", 60, HorizontalAlignment.Left)
        ListView1.Columns.Add("Fecha", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("FechaOrdn", 0, HorizontalAlignment.Left)
        Label8.Text = "0" + " €"
        cargarGastosListView()

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Public Function cargarGastosListView() As Boolean
        'cargamos los gastos
        Dim str(5) As String
        Dim itm As ListViewItem
        Dim gastos As MySqlCommand = New MySqlCommand("select * from gastos ORDER BY tipoGasto", con)
        gastos.CommandType = CommandType.Text
        con.Open()
        Dim myReader As MySqlDataReader
        myReader = gastos.ExecuteReader()
        Try

            While myReader.Read()
                str(0) = myReader.GetString(0)
                str(1) = myReader.GetString(2)
                str(2) = myReader.GetString(1)
                str(3) = myReader.GetString(3) + "/" + myReader.GetString(4) + "/" + myReader.GetString(5)
                str(4) = myReader.GetString(5) + myReader.GetString(4) + myReader.GetString(3)
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

            Dim cmd As MySqlCommand = New MySqlCommand("select round(sum(importe),2) from gastos  ;", con)
            Dim lector As MySqlDataReader
            con.Open()
            cmd.CommandType = CommandType.Text
            Try
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
                    Label8.Text = lector.GetString(0) + " €"
                End If

            Catch mierror As MySqlException
                MessageBox.Show(" " & mierror.Message)
            Finally
                con.Dispose()
            End Try
            con.Close()
        End Try
        ListView1.ListViewItemSorter = New ListViewItemComparer(4)
   

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
                cargarGastosListView()
            End If
        End If
    End Sub
    Public Function BorrarGasto() As Boolean
        'Borramos un gasto
        con.Open()
        Dim cmd As MySqlCommand = New MySqlCommand("Delete from gastos where id_gasto=(?val1);", con)
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
End Class
Class ListViewItemComparer
    Implements IComparer

    Private col As Integer
    ' CONSTRUCTOR
    Public Sub New()
        col = 0
    End Sub
    ' CONSTRUCTOR
    Public Sub New(ByVal column As Integer)
        col = column
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim a As Integer
        Dim b As Integer
        Dim returnValue As Integer
        a = CType(x, ListViewItem).SubItems(col).Text
        b = CType(y, ListViewItem).SubItems(col).Text
        If a > b Then
            returnValue = -1
        ElseIf a = b Then
            returnValue = 0
        Else
            returnValue = 1
        End If
        Return returnValue
    End Function

End Class

