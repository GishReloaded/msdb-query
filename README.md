# msdb-query
.NET Library for help of queries to database

### Examples
```vb
    Public Function InsertUser(ByVal email As String, ByVal name As String) As Boolean
        Using connection As SqlConnection = SQL.sqlConnection(connectionString)

            If SQL.Connect(connection) Then
                Dim cmd As SqlCommand = New SqlCommand("[dbo].[InsertUser]", connection)
                cmd.Parameters.AddWithValue("Email", email)
                cmd.Parameters.AddWithValue("Name", name)
                Dim success As Boolean = SQL.ExecuteNonQuery(cmd)
                Return success
            End If
        End Using
    End Function

    Public Class UserModel
        <SQLParameter("Id")>
        Public Property Id As Integer
        <SQLParameter("RoleId")>
        Public Property RoleId As Integer
        <SQLParameter("Email")>
        Public Property Email As String
        <SQLParameter("Name")>
        Public Property Name As String
    End Class

    Public Function GetUsers() As List(Of UserModel)
        Dim list As List(Of UserModel) = New List(Of UserModel)()

        Using connection As SqlConnection = SQL.sqlConnection(connectionString)

            If SQL.Connect(connection) Then
                Dim cmd As SqlCommand = New SqlCommand("[dbo].[GetUsers]", connection)
                Dim success As List(Of UserModel) = SQL.DataReader(Of UserModel)(cmd)
                If success IsNot Nothing Then list.AddRange(success)
            End If
        End Using

        Return list
    End Function
```
