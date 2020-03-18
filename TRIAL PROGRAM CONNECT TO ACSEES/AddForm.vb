Public Class AddForm
    Private Sub AddForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If AddIndicator = True Then
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
            TextBox5.Clear()
            TextBox6.Clear()
            TextBox7.Clear()
            TextBox8.Clear()
            TextBox9.Clear()
            TextBox10.Clear()
        Else
            TextBox1.Text = Form1.DataGridView1.SelectedRows(0).Cells(1).Value.ToString
            TextBox2.Text = Form1.DataGridView1.SelectedRows(0).Cells(2).Value.ToString
            TextBox3.Text = Form1.DataGridView1.SelectedRows(0).Cells(3).Value.ToString
            TextBox4.Text = Form1.DataGridView1.SelectedRows(0).Cells(4).Value.ToString
            TextBox5.Text = Form1.DataGridView1.SelectedRows(0).Cells(5).Value.ToString
            TextBox6.Text = Form1.DataGridView1.SelectedRows(0).Cells(6).Value.ToString
            TextBox7.Text = Form1.DataGridView1.SelectedRows(0).Cells(7).Value.ToString
            TextBox8.Text = Form1.DataGridView1.SelectedRows(0).Cells(8).Value.ToString
            TextBox9.Text = Form1.DataGridView1.SelectedRows(0).Cells(9).Value.ToString
            TextBox10.Text = Form1.DataGridView1.SelectedRows(0).Cells(10).Value.ToString
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If AddIndicator = True Then
            ExecuteSqlStmt("Insert Into Main_Table([שם משפחה],[שם פרטי],[כתובת],[טלפון בית],[מס' נייד],[תאריך לידה],
                           [ת זהות],[בדיקה חוזרת],[E-MAIL],[מס' לקוח])
                           Values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" _
                           & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" _
                           & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" _
                           & TextBox9.Text & "','" & TextBox10.Text & "')")
            Form1.LoadMain_Table()
        Else
            ExecuteSqlStmt("Update Main_Table Set [שם משפחה] = '" & TextBox1.Text & "',[שם פרטי]='" _
                            & TextBox2.Text & "',[כתובת]='" & TextBox3.Text & "',[טלפון בית]='" _
                             & TextBox4.Text & "',[מס' נייד]='" & TextBox5.Text & "',[תאריך לידה]='" _
                              & TextBox6.Text & "',[ת זהות]='" & TextBox7.Text & "',[בדיקה חוזרת]='" _
                               & TextBox8.Text & "',[E-MAIL]='" & TextBox9.Text & "',[מס' לקוח]='" _
                                & TextBox10.Text & "' where Id =" & Val(Form1.DataGridView1.SelectedRows(0).Cells(0).Value))
            Form1.LoadMain_Table()
            Me.Close()
        End If

    End Sub
End Class