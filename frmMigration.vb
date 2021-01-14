Imports System.Threading
Imports System.Globalization
Imports System.ComponentModel
Imports System.Linq
Imports Microsoft.Office.Interop
Imports System.IO

Public Class frmMigration
    Dim firstTable As String

    Dim analyzedTables As List(Of String)
    Dim diffs As List(Of String)
    Dim deletedTables As New List(Of String)
    Dim insertedTables As New List(Of String)

    Dim sqlConn As New SQLConnection()
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
    End Sub

    Private Sub connectToDatabase()
        If Not Me.sqlConn.IsOpen() Then
            Me.sqlConn.Open(Me.txtServer1.Text, Me.txtDB1.Text, Me.txtUser1.Text, Me.txtPass1.Text, Me.txtServer2.Text, Me.txtDB2.Text, Me.txtUser2.Text, Me.txtPass2.Text)
        End If
    End Sub

    Private Sub Analyze(Optional inverse As Boolean = False)
        Try
            Me.connectToDatabase()

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

            trans = sqlConn.CnDestination.BeginTransaction("TRANSFER")
            sqlConn.CmdDestination.Transaction = trans

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

        sqlConn.CmdOrigin.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
        sqlConn.CmdDestination.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"

        sqlConn.DaOrigin.Fill(tablesOrigin)
        sqlConn.DaDestination.Fill(tablesDestination)

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
            Dim dtDestination As New DataTable

            sqlConn.CmdDestination.CommandText = $"select * from {tableName}"
            sqlConn.DaDestination.Fill(dtDestination)

            tableIsIdentity = Me.checkIdentity(dtDestination.Clone(), tableName)

            ' Se borra la tabla
            sqlConn.CmdDestination.CommandText = $"delete from {tableName}"
            sqlConn.CmdDestination.ExecuteNonQuery()

            ' Se hace reseed
            If tableIsIdentity Then
                sqlConn.CmdDestination.CommandText = $"dbcc checkident({tableName}, reseed, 1)"
                sqlConn.CmdDestination.ExecuteNonQuery()
            End If

            deletedTables.Add(tableName)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Insert(tableName As String)
        Try
            Dim tableIsIdentity As Boolean
            Dim dtOrigin As New DataTable
            Dim dtDestination As New DataTable

            sqlConn.CmdOrigin.CommandText = $"select * from {tableName}"
            sqlConn.CmdDestination.CommandText = $"select * from {tableName}"

            sqlConn.DaOrigin.Fill(dtOrigin)
            sqlConn.DaDestination.Fill(dtDestination)

            tableIsIdentity = Me.checkIdentity(dtDestination.Clone(), tableName)

            For Each row As DataRow In dtOrigin.Rows
                ' Si es identidad se pone en on
                If tableIsIdentity Then
                    Me.changeIdentity(True, tableName)
                End If

                ' Se migran los datos
                sqlConn.CmdDestination.CommandText = Me.generateInsertQuery(tableName, dtOrigin.Columns, dtDestination.Columns, row)
                sqlConn.CmdDestination.ExecuteNonQuery()

                If tableIsIdentity Then
                    Me.changeIdentity(False, tableName)
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

        sqlConn.CmdDestination.CommandText = $"SELECT OBJECT_NAME(f.referenced_object_id) TableName FROM sys.foreign_keys AS f WHERE OBJECT_NAME (f.parent_object_id) = '{tableName}'"
        sqlConn.DaDestination.Fill(tables)

        For i As Integer = 0 To tables.Rows.Count - 1
            tableNames.Add(tables.Rows(i).Item("TableName"))
        Next

        Return tableNames
    End Function

    Private Function getRefTables(tableName As String) As List(Of String)
        Dim tables As New DataTable
        Dim tableNames As New List(Of String)

        sqlConn.CmdDestination.CommandText = $"SELECT OBJECT_NAME(f.parent_object_id) TableName FROM sys.foreign_keys AS f WHERE OBJECT_NAME (f.referenced_object_id) = '{tableName}'"
        sqlConn.DaDestination.Fill(tables)

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

    Private Sub changeIdentity(active As Boolean, table As String)
        Try
            sqlConn.CmdDestination.CommandText = $"SET IDENTITY_INSERT {table} {IIf(active, "ON", "OFF")}"

            sqlConn.CmdDestination.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function checkIdentity(dt As DataTable, tableName As String) As Boolean
        Dim isIdentity As Boolean = False

        Try
            For Each column As DataColumn In dt.Columns

                sqlConn.CmdDestination.CommandText = $"Select is_identity From sys.columns Where Name = '{column.ColumnName}' AND object_id = OBJECT_ID('{tableName}')"

                If sqlConn.CmdDestination.ExecuteScalar() Then
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
        lblAmountInserted.Text = $"Cantidad: {insertedTables.Count()}"
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
        lblAmountAnalyzed.Text = $"Cantidad: {analyzedTables.Count()}"
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
