Imports System.Data.OleDb
Imports System.Data
Public Class Form1
    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application.StartupPath & "\BDCOCHES.mdb;")
    Dim odataset As New DataSet
    Dim consulta As New OleDbDataAdapter("select * from TBMODELOS", con)
    Dim consulta2 As New OleDbDataAdapter("select * from TBVENDIDOS", con)
    Dim builder As New OleDbCommandBuilder(consulta)
    Dim builder2 As New OleDbCommandBuilder(consulta2)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con.Open()
        consulta.Fill(odataset, "TBMODELOS")
        consulta2.Fill(odataset, "TBVENDIDOS")
        con.Close()
        Dim fila As DataRow
        For Each fila In odataset.Tables("TBMODELOS").Rows
            ComboBox1.Items.Add(fila.Item("MODELO"))
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim fila As DataRow
        For Each fila In odataset.Tables("TBMODELOS").Rows
            If ComboBox1.SelectedItem = fila.Item("MODELO") Then
                TextBox1.Text = fila.Item("MODELO")
                TextBox2.Text = fila.Item("CILINDRADA")
                TextBox3.Text = fila.Item("MOTOR")
                TextBox4.Text = fila.Item("UNIDADES")
                PictureBox1.Image = Image.FromFile(Application.StartupPath & "/" & fila.Item("FOTO"))
            End If

        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim FILA As DataRow
        For Each FILA In odataset.Tables("TBVENDIDOS").Rows
            ListBox1.Items.Add(FILA.Item("TELEFONO") & " - " & FILA.Item("MODELO"))
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView1.DataSource = odataset
        DataGridView1.DataMember = "TBMODELOS"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        consulta.Update(odataset, "TBMODELOS")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim telefono As String
        telefono = InputBox("Dime el número de telefono del comprador")

        Dim fila As DataRow
        fila = odataset.Tables("TBVENDIDOS").NewRow
        fila("TELEFONO") = telefono
        fila("MODELO") = ComboBox1.SelectedItem
        odataset.Tables("TBVENDIDOS").Rows.Add(fila)
        consulta2.Update(odataset, "TBVENDIDOS")

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim fila As DataRow
        fila = odataset.Tables("TBMODELOS").NewRow
        fila("MODELO") = TextBox6.Text
        fila("CILINDRADA") = TextBox7.Text
        fila("MOTOR") = TextBox8.Text
        fila("UNIDADES") = TextBox9.Text
        odataset.Tables("TBMODELOS").Rows.Add(fila)
        consulta.Update(odataset, "TBMODELOS")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.Rows.Count > 0 Then
            For Each fila As DataGridViewRow In DataGridView1.Rows
                If Not fila Is Nothing Then
                    For Each celda As DataGridViewCell In fila.Cells
                        If Not celda.Value Is Nothing Then
                            MsgBox(celda.Value)
                        End If
                    Next
                End If
            Next
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        TextBox1.Text = odataset.Tables("TBMODELOS").Rows(e.RowIndex).Item("MODELO")
        TextBox2.Text = odataset.Tables("TBMODELOS").Rows(e.RowIndex).Item("CILINDRADA")
        TextBox3.Text = odataset.Tables("TBMODELOS").Rows(e.RowIndex).Item("MOTOR")
        TextBox4.Text = odataset.Tables("TBMODELOS").Rows(e.RowIndex).Item("UNIDADES")
    End Sub

End Class
