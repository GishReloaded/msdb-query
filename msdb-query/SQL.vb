Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Runtime.Serialization
Imports System.Runtime.CompilerServices

Namespace MSDBQuery
    Public Enum DataOperation
        none = 0
        Read = 1
        Save = 2
        Delete = 4
        All = Read Or Save Or Delete
    End Enum

    Public Class SQLParameterAttribute
        Inherits Attribute

        Public Property Name As String
        Public Property Read As ParameterDirection = ParameterDirection.Input
        Public Property Save As ParameterDirection = ParameterDirection.Input
        Public Property Delete As ParameterDirection = 0
        Public Property size As Integer = 0
        Public Property type As SqlDbType = CType((-1), SqlDbType)

        Public Sub New(ByVal Optional Name As String = Nothing, ByVal Optional Read As ParameterDirection = ParameterDirection.Input, ByVal Optional Save As ParameterDirection = ParameterDirection.Input, ByVal Optional Delete As ParameterDirection = 0)
            Me.Name = Name
            Me.Read = Read
            Me.Save = Save
            Me.Delete = Delete
        End Sub

        Friend Function Parameter(ByVal value As Object, ByVal dir As ParameterDirection) As SqlParameter
            Return Parameter(Name, value, dir)
        End Function

        Friend Function Parameter(ByVal parameterName As String, ByVal value As Object, ByVal dir As ParameterDirection) As SqlParameter
            Dim par As SqlParameter = New SqlParameter(parameterName, value)
            par.Direction = dir
            If size > 0 Then par.Size = size
            If CInt(type) <> -1 Then par.SqlDbType = type
            Return par
        End Function
    End Class

    Public Module SQL
        <Extension()>
        Function sqlConnection(ByVal connectionString As String) As SqlConnection
            Return New SqlConnection(connectionString)
        End Function

        Function Connect(ByVal conn As SqlConnection) As Boolean
            If conn.State = System.Data.ConnectionState.Open Then Return True
            conn.Open()
            Return True
        End Function

        Function Close(ByVal conn As SqlConnection) As Boolean
            If conn.State = System.Data.ConnectionState.Closed Then Return True
            conn.Close()
            Return True
        End Function

        Function ExecuteNonQuery(ByVal cmd As SqlCommand) As Boolean
            Dim wasClosed As Boolean = cmd.Connection.State = System.Data.ConnectionState.Closed

            Try
                Connect(cmd.Connection)
                cmd.CommandType = System.Data.CommandType.StoredProcedure
                cmd.ExecuteNonQuery()
                Return True
            Finally
                If wasClosed Then Close(cmd.Connection)
            End Try
        End Function

        Function DataReader(Of T As Class)(ByVal cmd As SqlCommand) As List(Of T)
            Dim wasClosed As Boolean = cmd.Connection.State = System.Data.ConnectionState.Closed

            Try
                Connect(cmd.Connection)
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                Using dr As SqlDataReader = cmd.ExecuteReader()
                    Return dr.MapList(Of T)()
                End Using

            Finally
                If wasClosed Then Close(cmd.Connection)
            End Try
        End Function

        Function DataReader(Of T1 As Class, T2 As Class)(ByVal cmd As SqlCommand) As List(Of Object)
            Dim wasClosed As Boolean = cmd.Connection.State = System.Data.ConnectionState.Closed

            Try
                Dim ret As List(Of Object) = New List(Of Object)()
                Connect(cmd.Connection)
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                Using dr As SqlDataReader = cmd.ExecuteReader()
                    ret.Add(dr.MapList(Of T1)())

                    If dr.NextResult() Then
                        ret.Add(dr.MapList(Of T2)())
                    Else
                        ret.Add(New List(Of T2)())
                    End If
                End Using

                Return ret
            Finally
                If wasClosed Then Close(cmd.Connection)
            End Try
        End Function

        Function DataReader(Of T1 As Class, T2 As Class, T3 As Class)(ByVal cmd As SqlCommand) As List(Of Object)
            Dim wasClosed As Boolean = cmd.Connection.State = System.Data.ConnectionState.Closed

            Try
                Dim ret As List(Of Object) = New List(Of Object)()
                Connect(cmd.Connection)
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                Using dr As SqlDataReader = cmd.ExecuteReader()
                    ret.Add(dr.MapList(Of T1)())

                    If dr.NextResult() Then
                        ret.Add(dr.MapList(Of T2)())
                    Else
                        ret.Add(New List(Of T2)())
                    End If

                    If dr.NextResult() Then
                        ret.Add(dr.MapList(Of T3)())
                    Else
                        ret.Add(New List(Of T3)())
                    End If
                End Using

                Return ret
            Finally
                If wasClosed Then Close(cmd.Connection)
            End Try
        End Function

        Function DataRead(Of T As Class)(ByVal cmd As SqlCommand) As T
            Dim wasClosed As Boolean = cmd.Connection.State = System.Data.ConnectionState.Closed

            Try
                Connect(cmd.Connection)
                cmd.CommandType = System.Data.CommandType.StoredProcedure

                Using dr As SqlDataReader = cmd.ExecuteReader()

                    If dr.Read() Then
                        Return dr.MapData(Of T)()
                    End If
                End Using

            Finally
                If wasClosed Then Close(cmd.Connection)
            End Try

            Return Nothing
        End Function

        <Extension()>
        Function MapData(Of T As Class)(ByVal dr As SqlDataReader) As T
            Dim targetType As Type = GetType(T)
            Dim ret As T = Activator.CreateInstance(Of T)()

            For i As Integer = 0 To dr.VisibleFieldCount - 1
                Dim name As String = dr.GetName(i).ToLower()

                For Each targetPropertyInfo In targetType.GetProperties()

                    If targetPropertyInfo.CanWrite Then
                        Dim attr As DataMemberAttribute = targetPropertyInfo.GetCustomAttribute(Of DataMemberAttribute)()

                        If attr IsNot Nothing Then

                            If String.IsNullOrEmpty(attr.Name) AndAlso targetPropertyInfo.Name.ToLower() = name Then
                                targetPropertyInfo.SetValue(ret, convertedValue(dr.GetValue(i), targetPropertyInfo.PropertyType))
                            ElseIf attr.Name?.ToLower() = name Then
                                targetPropertyInfo.SetValue(ret, convertedValue(dr.GetValue(i), targetPropertyInfo.PropertyType))
                            End If
                        Else
                            Dim sqlAttribute = targetPropertyInfo.GetCustomAttribute(Of SQLParameterAttribute)()

                            If sqlAttribute IsNot Nothing Then

                                If String.IsNullOrEmpty(sqlAttribute.Name) AndAlso targetPropertyInfo.Name.ToLower() = name Then
                                    targetPropertyInfo.SetValue(ret, convertedValue(dr.GetValue(i), targetPropertyInfo.PropertyType))
                                ElseIf sqlAttribute.Name?.ToLower() = name Then
                                    targetPropertyInfo.SetValue(ret, convertedValue(dr.GetValue(i), targetPropertyInfo.PropertyType))
                                End If
                            End If
                        End If
                    End If
                Next
            Next

            Return ret
        End Function

        <Extension()>
        Function MapList(Of T As Class)(ByVal dr As SqlDataReader) As List(Of T)
            Dim ret As List(Of T) = New List(Of T)()

            While dr.Read()
                ret.Add(MapData(Of T)(dr))
            End While

            Return ret
        End Function

        <Extension()>
        Sub AddParametersFromObject(ByVal cmd As SqlCommand, ByVal source As Object, ByVal Optional dataOperation As DataOperation = DataOperation.All)
            Dim sourceType As Type = source.[GetType]()

            For Each sourcePropertyInfo In sourceType.GetProperties()

                If dataOperation <> 0 Then
                    Dim attr As DataMemberAttribute = sourcePropertyInfo.GetCustomAttribute(Of DataMemberAttribute)()

                    If attr IsNot Nothing Then

                        If Not String.IsNullOrEmpty(attr.Name) Then
                            cmd.Parameters.AddWithValue(attr.Name.ToLower(), escapeNull(sourcePropertyInfo.GetValue(source)))
                        Else
                            cmd.Parameters.AddWithValue(sourcePropertyInfo.Name.ToLower(), escapeNull(sourcePropertyInfo.GetValue(source)))
                        End If
                    Else
                        Dim sqlAttr As SQLParameterAttribute = sourcePropertyInfo.GetCustomAttribute(Of SQLParameterAttribute)()

                        If sqlAttr IsNot Nothing Then
                            Dim dir As ParameterDirection = 0

                            If (dataOperation And DataOperation.Read) <> 0 Then
                                dir = sqlAttr.Read
                            ElseIf (dataOperation And DataOperation.Delete) <> 0 Then
                                dir = sqlAttr.Delete
                            ElseIf (dataOperation And DataOperation.Save) <> 0 Then
                                dir = sqlAttr.Save
                            End If

                            If dir <> 0 Then

                                If Not String.IsNullOrEmpty(sqlAttr.Name) Then
                                    cmd.Parameters.Add(sqlAttr.Parameter(escapeNull(sourcePropertyInfo.GetValue(source)), dir))
                                Else
                                    cmd.Parameters.Add(sqlAttr.Parameter(sourcePropertyInfo.Name.ToLower(), escapeNull(sourcePropertyInfo.GetValue(source)), dir))
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End Sub

        <Extension()>
        Sub FillObjectFromParameters(ByVal cmd As SqlCommand, ByVal target As Object, ByVal Optional dataOperation As DataOperation = DataOperation.All)
            Dim targetType As Type = target.[GetType]()

            For Each targetPropertyInfo In targetType.GetProperties()

                If dataOperation <> 0 Then
                    Dim sqlAttr As SQLParameterAttribute = targetPropertyInfo.GetCustomAttribute(Of SQLParameterAttribute)()

                    If sqlAttr IsNot Nothing Then
                        Dim dir As ParameterDirection = 0

                        If (dataOperation And DataOperation.Read) <> 0 Then
                            dir = sqlAttr.Read
                        ElseIf (dataOperation And DataOperation.Delete) <> 0 Then
                            dir = sqlAttr.Delete
                        ElseIf (dataOperation And DataOperation.Save) <> 0 Then
                            dir = sqlAttr.Save
                        End If

                        If dir = ParameterDirection.InputOutput OrElse dir = ParameterDirection.Output Then

                            If Not String.IsNullOrEmpty(sqlAttr.Name) Then
                                If cmd.Parameters.Contains(sqlAttr.Name) Then targetPropertyInfo.SetValue(target, convertedValue(cmd.Parameters(sqlAttr.Name).Value, targetPropertyInfo.PropertyType))
                            ElseIf cmd.Parameters.Contains(targetPropertyInfo.Name.ToLower()) Then
                                targetPropertyInfo.SetValue(target, convertedValue(cmd.Parameters(targetPropertyInfo.Name.ToLower()).Value, targetPropertyInfo.PropertyType))
                            End If
                        End If
                    End If
                End If
            Next
        End Sub

        Private Function escapeDBNull(ByVal v As Object) As Object
            If v = DBNull.Value Then Return Nothing

            If v.[GetType]() = GetType(String) Then
                If TryCast(v, String) = "null" Then Return Nothing
            End If

            Return v
        End Function

        Private Function convertedValue(ByVal v As Object, ByVal targetType As Type) As Object
            If v = DBNull.Value Then Return Nothing

            If v.[GetType]() = GetType(String) Then
                If TryCast(v, String) = "null" Then Return Nothing
            End If

            If targetType.IsGenericType Then
                Return Convert.ChangeType(v, targetType.GetGenericArguments()(0))
            End If

            Return Convert.ChangeType(v, targetType)
        End Function

        Private Function escapeNull(ByVal v As Object) As Object
            If v Is Nothing Then Return DBNull.Value

            If v.[GetType]() = GetType(String) Then
                If TryCast(v, String) = "null" Then Return DBNull.Value
            End If

            Return v
        End Function

        Function isNull(Of T As Class)(ByVal value As T, ByVal nullValue As T) As T
            If value Is Nothing Then Return nullValue
            If value Is DBNull.Value Then Return nullValue

            If value.[GetType]() = GetType(String) Then
                If (TryCast(value, String)).ToLower() = "null" Then Return nullValue
            End If

            Return value
        End Function

        <Extension()>
        Function getDateTime(ByVal cmd As SqlCommand, ByVal prm As String) As DateTime
            Dim o As Object = cmd.Parameters(prm).Value
            If o = DBNull.Value Then Return DateTime.MinValue
            Return Convert.ToDateTime(o)
        End Function

        <Extension()>
        Function getInt(ByVal cmd As SqlCommand, ByVal prm As String) As Integer
            Dim o As Object = cmd.Parameters(prm).Value
            If o = DBNull.Value Then Return 0
            Return Convert.ToInt32(o)
        End Function

        <Extension()>
        Function getString(ByVal cmd As SqlCommand, ByVal prm As String) As String
            Dim o As Object = cmd.Parameters(prm).Value
            If o = DBNull.Value Then Return ""
            Return Convert.ToString(o)
        End Function

        <Extension()>
        Function returnValue(ByVal cmd As SqlCommand) As Integer
            Return cmd.getInt("ReturnValue")
        End Function

        <Extension()>
        Function addReturnValueParameter(ByVal cmd As SqlCommand) As SqlParameter
            Dim prm = cmd.Parameters.Add("ReturnValue", SqlDbType.Int)
            prm.Direction = ParameterDirection.ReturnValue
            Return prm
        End Function
    End Module
End Namespace
