Public Class SQLConnection
    Private xCmdDestination As SqlClient.SqlCommand
    Private xDaDestination As SqlClient.SqlDataAdapter
    Private xCnDestination As SqlClient.SqlConnection

    Private xCmdOrigin As SqlClient.SqlCommand
    Private xDaOrigin As SqlClient.SqlDataAdapter
    Private xCnOrigin As SqlClient.SqlConnection
    Private xDtOrigin, dtDestination As DataTable

    Public ReadOnly Property CmdDestination As SqlClient.SqlCommand
        Get
            Return xCmdDestination
        End Get
    End Property

    Public ReadOnly Property DaDestination() As SqlClient.SqlDataAdapter
        Get
            Return xDaDestination
        End Get
    End Property

    Public ReadOnly Property CnDestination As SqlClient.SqlConnection
        Get
            Return xCnDestination
        End Get
    End Property

    Public ReadOnly Property CmdOrigin() As SqlClient.SqlCommand
        Get
            Return xCmdOrigin
        End Get
    End Property

    Public ReadOnly Property DaOrigin() As SqlClient.SqlDataAdapter
        Get
            Return xDaOrigin
        End Get
    End Property

    Public ReadOnly Property CnOrigin As SqlClient.SqlConnection
        Get
            Return xCnOrigin
        End Get
    End Property

    Public Sub New()
        xCnOrigin = New SqlClient.SqlConnection()
        xCnDestination = New SqlClient.SqlConnection()
        xCmdOrigin = New SqlClient.SqlCommand()
        xCmdDestination = New SqlClient.SqlCommand()
        xDaOrigin = New SqlClient.SqlDataAdapter(xCmdOrigin)
        xDaDestination = New SqlClient.SqlDataAdapter(xCmdDestination)
    End Sub

    Public Function IsOpen() As Boolean
        Return xCnOrigin.State = ConnectionState.Open Or xCnDestination.State = ConnectionState.Open
    End Function

    Public Sub Open(dataSource1 As String, initialCatalog1 As String, user1 As String, password1 As String, dataSource2 As String, initialCatalog2 As String, user2 As String, password2 As String)
        xCnOrigin.ConnectionString = $"Data Source={dataSource1}; Initial Catalog={initialCatalog1}; User ID={user1}; Password={password1}"
        xCnDestination.ConnectionString = $"Data Source={dataSource2}; Initial Catalog={initialCatalog2}; User ID={user2}; Password={password2}"

        xCnOrigin.Open()
        xCnDestination.Open()

        xCmdOrigin.Connection = xCnOrigin
        xCmdDestination.Connection = xCnDestination
    End Sub
End Class
