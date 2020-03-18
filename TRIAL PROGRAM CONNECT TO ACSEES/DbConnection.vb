Imports System.Data
Imports System.Data.OleDb


Module DbConnection
    Public AddIndicator As Boolean

    Dim dbProvider As String = "PROVIDER=Microsoft.Ace.OLEDB.12.0;"
    Dim dbSource As String = "Data Source="
    Dim dbPathAndFilename As String = Application.StartupPath & "\OPTICA DATA.accdb"
    Public conAn As New OleDb.OleDbConnection(dbProvider & dbSource & dbPathAndFilename & " ;Jet OLEDB:Database Password=fales")
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim cmd As OleDbCommand
    Dim sql As String

    Public Sub LoadData(ByVal sql As String, ByVal dataset As DataSet)

        If conAn.State = ConnectionState.Closed Then
            conAn.Open()
        End If
        Try

            da = New OleDbDataAdapter(sql, conAn)
            da.Fill(dataset)
            conAn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub ExecuteSqlStmt(ByVal sql As String)
        If conAn.State = ConnectionState.Closed Then
            conAn.Open()
        End If
        Try
            'conAn.Open()
            cmd = New OleDbCommand(sql, conAn)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conAn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function LoadDataRetrievId(ByVal sql As String) As Integer
        Dim query2 As String = "Select @@Identity"
        Dim ID As Integer
        Using cmd As New OleDbCommand(sql, conAn)
            conAn.Open()
            cmd.ExecuteNonQuery()
            cmd.CommandText = query2
            ID = cmd.ExecuteScalar()
            conAn.Close()
        End Using
        Return ID
    End Function

    Public Sub LoadDataTable(ByVal sql As String, ByVal dataset As DataTable)
        If conAn.State = ConnectionState.Closed Then
            conAn.Open()
        End If
        Try
            'conAn.Open()
            da = New OleDbDataAdapter(sql, conAn)
            da.Fill(dataset)
            conAn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


End Module
