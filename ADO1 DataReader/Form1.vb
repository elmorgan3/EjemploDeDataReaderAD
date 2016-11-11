Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class Form1
    Private con As SqlConnection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con = New SqlConnection
        con.ConnectionString = "Data Source=.\SQLEXPRESS;Initial Catalog=MAGATZEM;Trusted_Connection=True;"
        Try
            con.Open()
        Finally

            If con.State = ConnectionState.Closed Then
                MsgBox("No se ha podido establecer la comunicacion. Houston tenemos un problema")
                Close()
            End If

        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim comanda As SqlCommand
        Dim reader As SqlDataReader


        comanda = New SqlCommand
        comanda.Connection = con
        comanda.CommandText = "select idproducto, nombreproducto, idcategoria from productos"
        reader = comanda.ExecuteReader()

        OmplirLlista(reader)
    End Sub

    Private Sub OmplirLlista(reader As SqlDataReader)
        Dim NombreProducto As String
        ListBox1.Items.Clear()

        While reader.Read
            NombreProducto = reader("NombreProducto").ToString()
            NombreProducto = reader("IdProducto").ToString() + " -- " + NombreProducto + " -- " + reader("idCategoria").ToString()
            ListBox1.Items.Add(NombreProducto)
        End While

        reader.Close()

    End Sub

    'Busca los por categoria que haya puesto en el textbox categoria
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim comanda As SqlCommand
        Dim reader As SqlDataReader

        comanda = New SqlCommand
        comanda.Connection = con
        comanda.CommandText = "select idproducto, nombreproducto, idcategoria from productos"
        comanda.CommandText = comanda.CommandText + " where idcategoria = " + txbCategoria.Text

        reader = comanda.ExecuteReader()
        OmplirLlista(reader)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim comanda As SqlCommand
        Dim reader As SqlDataReader

        comanda = New SqlCommand
        comanda.Connection = con

        'comanda.CommandText = "select idproducto, nombreproducto, idcategoria from productos"
        'comanda.CommandText = comanda.CommandText + " where nombreproducto = '" + txbNom.Text + "'"

        comanda.CommandText = "select idproducto, nombreproducto, idcategoria from productos"
        comanda.CommandText = comanda.CommandText + " where nombreproducto like '%{0}%' and idcategoria={1}"
        comanda.CommandText = String.Format(comanda.CommandText, txbNom.Text, txbCategoria.Text)

        reader = comanda.ExecuteReader()
        OmplirLlista(reader)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim comanda As SqlCommand
        Dim Update As String
        Dim registres As Integer

        comanda = New SqlCommand
        comanda.Connection = con

        Update = "update productos set nombreproducto = '{0}' where idproducto = {1}"
        Update = String.Format(Update, txbDescripcion.Text, txbIdProducto.Text)
        comanda.CommandText = Update

        registres = comanda.ExecuteNonQuery()

        If registres = 0 Then
            MsgBox("No s'ha actualitzat cap registre")
        End If

    End Sub
End Class
