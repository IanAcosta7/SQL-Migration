Imports System.Threading
Imports System.Globalization
Imports System.ComponentModel
Imports System.Linq
Imports Microsoft.Office.Interop
Imports System.IO

Public Class frmMigration

    Dim cmdOrigin As SqlClient.SqlCommand
    Dim cmdDestination As SqlClient.SqlCommand
    Dim daOrigin As SqlClient.SqlDataAdapter
    Dim daDestination As SqlClient.SqlDataAdapter
    Dim dtOrigin, dtDestination As DataTable

    Dim cnOrigin As SqlClient.SqlConnection
    Dim cnDestination As SqlClient.SqlConnection
    Dim firstTable As String

    Dim analyzedTables As List(Of String)
    Dim diffs As List(Of String)
    Dim deletedTables As New List(Of String)
    Dim insertedTables As New List(Of String)
    Private Sub frmMigration_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Solo para testing
        Me.txtDB1.Text = "BizOneFashionMarketOld"
        Me.txtDB2.Text = "BOFM"
        Me.txtUser1.Text = "sa"
        Me.txtUser2.Text = "sa"
        Me.txtFirstTable.Text = "article"
        ' Solo para testing

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-BZ", False)
        Thread.CurrentThread.CurrentCulture.ClearCachedData()

        cnOrigin = New SqlClient.SqlConnection()
        cnDestination = New SqlClient.SqlConnection()
    End Sub

    Private Sub connectToDatabase()
        If cnOrigin.State = ConnectionState.Closed Or cnDestination.State = ConnectionState.Closed Then
            Me.validateFields()

            cnOrigin.ConnectionString = $"Data Source={Me.txtServer1.Text}; Initial Catalog={Me.txtDB1.Text}; User ID={Me.txtUser1.Text}; Password={Me.txtPass1.Text}"
            cnDestination.ConnectionString = $"Data Source={Me.txtServer2.Text}; Initial Catalog={Me.txtDB2.Text}; User ID={Me.txtUser2.Text}; Password={Me.txtPass2.Text}"

            cnOrigin.Open()
            cnDestination.Open()
        End If
    End Sub

    Private Sub validateFields()
        ' Throw New NotImplementedException()
    End Sub

    Private Sub Analyze(Optional inverse As Boolean = False)
        Dim trans As SqlClient.SqlTransaction

        Try
            Me.connectToDatabase()

            cmdOrigin = New SqlClient.SqlCommand
            cmdDestination = New SqlClient.SqlCommand
            daOrigin = New SqlClient.SqlDataAdapter(cmdOrigin)
            daDestination = New SqlClient.SqlDataAdapter(cmdDestination)

            cmdOrigin.Connection = cnOrigin
            cmdDestination.Connection = cnDestination

            Me.analyzedTables = New List(Of String)
            Me.diffs = Me.getDiffs()

            Me.GetRelations(Me.firstTable, inverse)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Throw ex
        End Try
    End Sub

    Private Sub GetRelations(tableName As String, Optional inverse As Boolean = False)
        Dim analyzedTablesVal As List(Of String) = analyzedTables

        Dim depTables = Me.getDepTables(tableName)
        Dim refTables = Me.getRefTables(tableName)

        depTables.RemoveAll(Function(str) analyzedTablesVal.Contains(str))
        refTables.RemoveAll(Function(str) analyzedTablesVal.Contains(str))

        If inverse Then
            For Each depTable As String In depTables
                Me.GetRelations(depTable, inverse)
            Next
        Else
            For Each refTable As String In refTables
                Me.GetRelations(refTable, inverse)
            Next
        End If

        If Not analyzedTables.Contains(tableName) Then
            analyzedTables.Add(tableName)
        End If

        If inverse Then
            For Each refTable As String In refTables
                Me.GetRelations(refTable, inverse)
            Next
        Else
            For Each depTable As String In depTables
                Me.GetRelations(depTable, inverse)
            Next
        End If
    End Sub

    Private Sub Migrate(sender As Object, reseedAndDelete As Boolean)
        Dim trans As SqlClient.SqlTransaction

        Try
            Me.connectToDatabase()

            cmdOrigin = New SqlClient.SqlCommand
            cmdDestination = New SqlClient.SqlCommand
            daOrigin = New SqlClient.SqlDataAdapter(cmdOrigin)
            daDestination = New SqlClient.SqlDataAdapter(cmdDestination)

            cmdOrigin.Connection = cnOrigin
            cmdDestination.Connection = cnDestination

            Dim progress = 0

            ' Reseed and Delete
            If reseedAndDelete Then
                For Each tableName In analyzedTables
                    If clbAnalyzedTables.CheckedItems.Contains(tableName) Then
                        Me.ReseedAndDelete(tableName)
                        progress += 1
                        DirectCast(sender, BackgroundWorker).ReportProgress(progress * 100 / Me.analyzedTables.Count / 2)
                    End If
                Next
            End If

            ' Se vuelven a analizar las tablas pero de forma inversa
            Me.Analyze(True)

            trans = cnDestination.BeginTransaction("TRANSFER")
            cmdDestination.Transaction = trans

            ' Inserts
            For Each tableName In analyzedTables
                If Not Me.diffs.Contains(tableName) Then
                    If clbAnalyzedTables.CheckedItems.Contains(tableName) Then
                        Me.Insert(tableName)
                        progress += 1
                        DirectCast(sender, BackgroundWorker).ReportProgress((progress * 100 / Me.analyzedTables.Count) / IIf(reseedAndDelete, 2, 1))
                    End If
                End If
            Next

            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error de Migración")
        End Try
    End Sub

    Private Function getDiffs() As List(Of String)
        Dim tablesOrigin As New DataTable
        Dim tablesDestination As New DataTable
        Dim tablesOriginStr As New List(Of String)
        Dim tablesDestinationStr As New List(Of String)

        cmdOrigin.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
        cmdDestination.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"

        daOrigin.Fill(tablesOrigin)
        daDestination.Fill(tablesDestination)

        For i As Integer = 0 To tablesOrigin.Rows.Count - 1
            tablesOriginStr.Add(tablesOrigin.Rows(i).Item("TABLE_NAME"))
        Next

        For i As Integer = 0 To tablesDestination.Rows.Count - 1
            tablesDestinationStr.Add(tablesDestination.Rows(i).Item("TABLE_NAME"))
        Next

        Return tablesDestinationStr.Except(tablesOriginStr).ToList()
    End Function

    Private Sub ReseedAndDelete(tableName As String)
        Try
            Dim tableIsIdentity As Boolean

            cmdDestination.CommandText = $"select * from {tableName}"

            dtDestination = New DataTable
            daDestination.Fill(dtDestination)

            tableIsIdentity = Me.checkIdentity(cmdDestination, dtDestination.Clone(), tableName)

            ' Se borra la tabla
            cmdDestination.CommandText = $"delete from {tableName}"
            cmdDestination.ExecuteNonQuery()

            ' Se hace reseed
            If tableIsIdentity Then
                cmdDestination.CommandText = $"dbcc checkident({tableName}, reseed, 1)"
                cmdDestination.ExecuteNonQuery()
            End If

            deletedTables.Add(tableName)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Insert(tableName As String)
        Try
            Dim tableIsIdentity As Boolean

            cmdOrigin.CommandText = $"select * from {tableName}"
            cmdDestination.CommandText = $"select * from {tableName}"

            dtOrigin = New DataTable
            dtDestination = New DataTable

            daOrigin.Fill(dtOrigin)
            daDestination.Fill(dtDestination)

            tableIsIdentity = Me.checkIdentity(cmdDestination, dtDestination.Clone(), tableName)

            For Each row As DataRow In dtOrigin.Rows
                ' Si es identidad se pone en on
                If tableIsIdentity Then
                    Me.changeIdentity(cmdDestination, True, tableName)
                End If

                ' Se migran los datos
                cmdDestination.CommandText = Me.generateInsertQuery(tableName, dtOrigin.Columns, dtDestination.Columns, row)
                cmdDestination.ExecuteNonQuery()

                If tableIsIdentity Then
                    Me.changeIdentity(cmdDestination, False, tableName)
                End If
            Next

            insertedTables.Add(tableName)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function getDepTables(tableName As String) As List(Of String)
        Dim tables As New DataTable
        Dim tableNames As New List(Of String)

        cmdDestination.CommandText = $"SELECT OBJECT_NAME(f.referenced_object_id) TableName FROM sys.foreign_keys AS f WHERE OBJECT_NAME (f.parent_object_id) = '{tableName}'"
        daDestination.Fill(tables)

        For i As Integer = 0 To tables.Rows.Count - 1
            tableNames.Add(tables.Rows(i).Item("TableName"))
        Next

        Return tableNames
    End Function

    Private Function getRefTables(tableName As String) As List(Of String)
        Dim tables As New DataTable
        Dim tableNames As New List(Of String)

        cmdDestination.CommandText = $"SELECT OBJECT_NAME(f.parent_object_id) TableName FROM sys.foreign_keys AS f WHERE OBJECT_NAME (f.referenced_object_id) = '{tableName}'"
        daDestination.Fill(tables)

        For i As Integer = 0 To tables.Rows.Count - 1
            tableNames.Add(tables.Rows(i).Item("TableName"))
        Next

        Return tableNames
    End Function

    Private Function generateInsertQuery(tableName As String, columnsOrigin As DataColumnCollection, columnsDestination As DataColumnCollection, row As DataRow) As String
        Dim values As String = "("
        Dim columnsNames As String = "("

        For col As Int16 = 0 To columnsOrigin.Count - 1
            Dim value As String

            ' Si es string se agregan comillas
            If row.Item(col).GetType = GetType(String) Then
                value = $"'{row.Item(col).ToString.Replace("'", "''")}'"
            Else
                value = row.Item(col).ToString()
            End If

            ' Si es null se escribe null
            If row.IsNull(col) Then
                value = "null"
            End If

            ' Si es boolean se escribe 0 o 1
            If row.Item(col).GetType = GetType(Boolean) Then
                value = IIf(row.Item(col), 1, 0)
            End If

            If columnsDestination.Contains(columnsOrigin.Item(col).ColumnName) Then
                columnsNames += $"{columnsOrigin.Item(col).ColumnName}, "
                values += $"{value}, "
            End If
        Next
        columnsNames = $"{columnsNames.Remove(columnsNames.LastIndexOf(","))})"
        values = $"{values.Remove(values.LastIndexOf(","))})"

        Return $"insert into {tableName} {columnsNames} values {values}"
    End Function

    Private Sub changeIdentity(cmd As SqlClient.SqlCommand, active As Boolean, table As String)
        Try
            cmd.Connection = cnDestination

            cmd.CommandText = $"SET IDENTITY_INSERT {table} {IIf(active, "ON", "OFF")}"

            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function checkIdentity(cmd As SqlClient.SqlCommand, dt As DataTable, tableName As String) As Boolean
        Dim isIdentity As Boolean = False

        Try
            cmd.Connection = cnDestination

            For Each column As DataColumn In dt.Columns

                cmd.CommandText = $"Select is_identity From sys.columns Where Name = '{column.ColumnName}' AND object_id = OBJECT_ID('{tableName}')"

                If cmd.ExecuteScalar() Then
                    isIdentity = True
                    Exit For
                End If

            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return isIdentity
    End Function

    Private Sub bgwMigrate_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bgwMigrate.DoWork
        Me.Migrate(sender, cbReseedAndDelete.Checked)
    End Sub

    Private Sub bgwMigrate_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwMigrate.ProgressChanged
        lbInsertedTables.DataSource = insertedTables
        pbMigration.Value = e.ProgressPercentage
    End Sub

    Private Sub bgwMigrate_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgwMigrate.RunWorkerCompleted
        lbInsertedTables.DataSource = insertedTables
        pbMigration.Value = 0
    End Sub

    Private Sub btnAnalyze_Click(sender As Object, e As EventArgs) Handles btnAnalyze.Click
        Try
            'Me.tables = New String() {"article"}
            If txtFirstTable.Text <> String.Empty Then
                Me.firstTable = txtFirstTable.Text
            End If

            pbMigration.Style = ProgressBarStyle.Marquee
            bgwAnalyze.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub bgwAnalyze_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwAnalyze.DoWork
        Try
            Me.Analyze()
        Catch ex As Exception
            e.Cancel = True
        End Try
    End Sub

    Private Sub bgwAnalyze_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwAnalyze.ProgressChanged
        Me.lbInsertedTables.DataSource = Me.insertedTables
        pbMigration.Value = e.ProgressPercentage
    End Sub

    Private Sub bgwAnalyze_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgwAnalyze.RunWorkerCompleted
        pbMigration.Style = ProgressBarStyle.Blocks
        cbReseedAndDelete.Enabled = True
        btnMigrate.Enabled = True
        lblAnalyze.Text = "Seleccione las tablas que desea migrar."
        clbAnalyzedTables.DataSource = analyzedTables

        For i As Int64 = 0 To clbAnalyzedTables.Items.Count - 1
            If Not Me.diffs.Contains(clbAnalyzedTables.Items(i)) Then
                clbAnalyzedTables.SetItemCheckState(i, CheckState.Checked)
            End If
        Next
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim sfdExcelExport As New SaveFileDialog()

        sfdExcelExport.Filter = "xlsx files (*.xlsx)|All files (*.*)"

        If sfdExcelExport.ShowDialog() = DialogResult.OK Then
            Dim excel As New Excel.Application
            excel.Visible = True

            excel.Workbooks.Add()

            Dim worksheet As Excel.Worksheet = DirectCast(excel.ActiveSheet, Excel.Worksheet)

            worksheet.Cells(1, "A") = "Tablas Analizadas"
            worksheet.Cells(1, "B") = "Tablas Migradas"

            For i As Int64 = 1 To clbAnalyzedTables.Items.Count
                worksheet.Cells(i + 1, "A") = clbAnalyzedTables.Items(i - 1)
            Next

            For i As Int64 = 1 To lbInsertedTables.Items.Count
                worksheet.Cells(i + 1, "B") = lbInsertedTables.Items(i - 1)
            Next

            worksheet.Columns(1).AutoFit()
            worksheet.Columns(2).AutoFit()

            worksheet.SaveAs(sfdExcelExport.FileName)
        End If
    End Sub

    Private Sub btnMigrate_Click(sender As Object, e As EventArgs) Handles btnMigrate.Click
        Try
            lblAnalyze.Text = ""
            bgwMigrate.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
