Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadMain_Table()

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        AddIndicator = True
        AddForm.ShowDialog()
    End Sub
    Public Sub LoadMain_Table()
        Dim UserDS As New DataSet
        LoadData("SELECT * FROM MAIN_TABLE", UserDS)
        Me.DataGridView1.DataSource = UserDS.Tables(0)
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If DataGridView1.SelectedRows.Count <> 0 Then
            AddIndicator = False
            AddForm.ShowDialog()
        Else
            MsgBox("Please! Select row you want to update.")
        End If
    End Sub

    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        If DataGridView1.SelectedRows.Count <> 0 Then
            ExecuteSqlStmt("Delete From Main_Table Where Id =" & Val(DataGridView1.SelectedRows(0).Cells(0).Value))
            LoadMain_Table()
        Else
            MsgBox("Please! Select row you want to delete.")
        End If
    End Sub
End Class
