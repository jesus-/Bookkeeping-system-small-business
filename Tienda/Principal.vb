Imports MySql.Data.MySqlClient
Public Class Principal
    Dim con As MySqlConnection
    Private Sub Principal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New MySqlConnection()
        con.ConnectionString = "server= localhost;user id= root; password=1111; database=tienda"
        'comprobamos que la base de datos abre correctamente
        Try
            con.Open()
            con.Close()
        Catch mierror As MySqlException
            MessageBox.Show("Error de Conexión a la Base de Datos: " & mierror.Message)
        Finally
            con.Dispose()
        End Try
       
        CargarDias()
        ListView1.Columns.Add("Fecha", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("FechaOrdn", 0, HorizontalAlignment.Left)
        'cambiar la fecha actual del calendar Month
        ' MonthCalendar1.TodayDate = New Date(2006, 4, 22)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim stringSeparators As String = "/"
        Dim Str() As String
        If ListView1.FocusedItem Is Nothing Then
            MsgBox("Por favor seleccione una fecha")
        Else
            If ListView1.SelectedItems.Count = 0 Then
                MsgBox("Por favor seleccione una fecha")
            Else
                Dim source As String = ListView1.SelectedItems.Item(0).Text.Trim()
                Str = source.Split(stringSeparators)
                Tienda.dia.setConst(Str(0), Str(1), Str(2))
                Tienda.dia.Show()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Tienda.proveedores.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Tienda.TipoGastos.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dia As Integer
        Dim mes As Integer
        Dim año As Integer
        If MonthCalendar1.SelectionRange.Start = MonthCalendar1.SelectionRange.End Then

            dia = CStr(Me.MonthCalendar1.SelectionStart.Day)
            mes = CStr(Me.MonthCalendar1.SelectionStart.Month)
            año = CStr(Me.MonthCalendar1.SelectionStart.Year)
            InsertarDia(dia, mes, año)
            CargarDias()
        Else
            MsgBox("por favor, seleccione solo un dia")
        End If
    End Sub
    Public Function InsertarDia(ByVal dia As Integer, ByVal mes As Integer, ByVal año As Integer) As Boolean
        con.Open()
        Dim cmd As MySqlCommand = New MySqlCommand("INSERT INTO fechas VALUES (  ?val1, ?val2, ?val3, 0, 0, 0);", con)
        Try
            cmd.Parameters.AddWithValue("?val1", dia)
            cmd.Parameters.AddWithValue("?val2", mes)
            cmd.Parameters.AddWithValue("?val3", año)
            cmd.Prepare()
            cmd.ExecuteNonQuery()
        Catch mierror As MySqlException
            MessageBox.Show("Problema!! ya existia la fecha")
        Finally
            con.Dispose()
        End Try
        con.Close()
    End Function
    Public Function CargarDias() As Boolean
        ListView1.Items.Clear()
        'cargamos los proveedores
        Dim str(2) As String
        Dim itm As ListViewItem
        Dim cmd As MySqlCommand = New MySqlCommand("select * from fechas", con)
        cmd.CommandType = CommandType.Text
        con.Open()
        Dim myReader As MySqlDataReader
        myReader = cmd.ExecuteReader()
        Try
            While myReader.Read()
                str(0) = myReader.GetString(0) + "/" + myReader.GetString(1) + "/" + myReader.GetString(2)
                str(1) = myReader.GetString(2) + rellenarCeros(myReader.GetString(1)) + rellenarCeros(myReader.GetString(0))
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

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Tienda.Resumenes.Show()

    End Sub
    Public Function rellenarCeros(ByVal numero As String) As String
        If numero.Length = 1 Then
            numero = "0" + numero
        End If
        Return numero
    End Function
End Class
