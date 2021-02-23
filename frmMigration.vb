Imports System.Threading
Imports System.Globalization
Imports System.ComponentModel
Imports System.Linq
Imports Microsoft.Office.Interop
Imports System.IO

Public Class frmMigration
    Dim tables() As String

    Dim analyzedTables As List(Of String)
    Dim diffs As List(Of String)
    Dim deletedTables As New List(Of String)
    Dim insertedTables As New List(Of String)
    Dim originColumns As New List(Of Integer)
    Dim destinationColumns As New List(Of Integer)
    Dim insertedColumns As New List(Of Integer)
    Dim notMigratedTables As New List(Of String)
    Dim progressText As String = ""

    Dim sqlConn As New SQLConnection()
    Private Sub frmMigration_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Solo para testing
        Me.txtDB1.Text = "BizOneFashionMarketOld"
        Me.txtDB2.Text = "BOFM"
        Me.txtUser1.Text = "sa"
        Me.txtUser2.Text = "sa"
        ' Solo para testing

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-BZ", False)
        Thread.CurrentThread.CurrentCulture.ClearCachedData()
    End Sub

    Private Sub ConnectToDatabase()
        If Not Me.sqlConn.IsOpen() Then
            Me.sqlConn.Open(Me.txtServer1.Text, Me.txtDB1.Text, Me.txtUser1.Text, Me.txtPass1.Text, Me.txtServer2.Text, Me.txtDB2.Text, Me.txtUser2.Text, Me.txtPass2.Text)
        End If
    End Sub

    Private Sub Analyze(Optional inverse As Boolean = False)
        Try
            Me.tables = GetTopLevelParents().ToArray()
            Me.analyzedTables = New List(Of String)

            For Each table As String In Me.tables
                Me.GetRelations(table, inverse)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Throw ex
        End Try
    End Sub

    Private Sub GetRelations(tableName As String, Optional inverse As Boolean = False)
        If Not diffs.Contains(tableName) Then
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
        End If
    End Sub

    Private Sub Migrate(sender As Object, reseedAndDelete As Boolean)
        Try
            Dim progress = 0

            ' Se conecta a la base de datos
            ConnectToDatabase()

            ' Reseed and Delete
            If reseedAndDelete Then
                For Each tableName In analyzedTables
                    progressText = $"Borrando: {tableName}"
                    If clbAnalyzedTables.CheckedItems.Contains(tableName) Then
                        Me.ReseedAndDelete(tableName)
                        progress += 1
                        DirectCast(sender, BackgroundWorker).ReportProgress(progress * 100 / Me.clbAnalyzedTables.CheckedItems.Count)
                    End If
                Next
            End If

            progressText = "Analizando..."

            ' Se vuelven a analizar las tablas pero de forma inversa
            Me.diffs = GetDiffs()
            Me.Analyze(True)

            progress = 0
            ' Inserts
            For Each tableName In analyzedTables
                If clbAnalyzedTables.CheckedItems.Contains(tableName) Then
                    Dim trans As SqlClient.SqlTransaction
                    trans = sqlConn.CnDestination.BeginTransaction("TRANSFER")
                    sqlConn.CmdDestination.Transaction = trans

                    progressText = $"Migrando: {tableName}"
                    DirectCast(sender, BackgroundWorker).ReportProgress(progress * 100 / Me.clbAnalyzedTables.CheckedItems.Count)

                    Try
                        Me.Insert(tableName)
                        progress += 1
                        DirectCast(sender, BackgroundWorker).ReportProgress(progress * 100 / Me.clbAnalyzedTables.CheckedItems.Count)
                        trans.Commit()
                    Catch ex As Exception
                        trans.Rollback()

                        If MsgBox($"No se ha podido insertar la tabla {tableName}. ¿Desea continuar  la migración de todas formas?", vbYesNo + vbExclamation, "Error") = MsgBoxResult.Yes Then
                            ' Reconexion a la base de datos
                            ConnectToDatabase()

                            If Me.HasIdentity(tableName) Then
                                Me.SetIdentityInsert(False, tableName)
                            End If
                        Else
                            Throw ex
                        End If
                    End Try
                End If
            Next

            notMigratedTables = getNotMigratedTables()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error en la Migración")
            Throw ex
        End Try
    End Sub

    Private Function GetTopLevelParents() As List(Of String)
        Dim topLevelParents As New List(Of String)
        Dim dtParents As New DataTable()

        sqlConn.CmdOrigin.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES LEFT JOIN sys.foreign_keys AS f ON TABLE_NAME = OBJECT_NAME(f.parent_object_id) WHERE TABLE_TYPE = 'BASE TABLE' AND OBJECT_NAME(f.parent_object_id) IS NULL"

        sqlConn.DaOrigin.Fill(dtParents)

        For i As Integer = 0 To dtParents.Rows.Count - 1
            topLevelParents.Add(dtParents.Rows(i).Item("TABLE_NAME"))
        Next

        Return topLevelParents
    End Function

    Private Function getNotMigratedTables() As List(Of String)
        Dim tablesOrigin As New DataTable
        Dim tablesOriginStr As New List(Of String)

        sqlConn.CmdOrigin.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"

        sqlConn.DaOrigin.Fill(tablesOrigin)

        For i As Integer = 0 To tablesOrigin.Rows.Count - 1
            tablesOriginStr.Add(tablesOrigin.Rows(i).Item("TABLE_NAME"))
        Next

        Return tablesOriginStr.Except(insertedTables).ToList()
    End Function

    Private Function GetDiffs() As List(Of String)
        Dim dtDiffs As New DataTable
        Dim diffs As New List(Of String)

        sqlConn.CmdOrigin.CommandText = "SELECT DB_NAME()"
        Dim originTable As String = sqlConn.CmdOrigin.ExecuteScalar()
        sqlConn.CmdDestination.CommandText = $"SELECT ISNULL(origin.TABLE_NAME, destination.TABLE_NAME) TABLE_NAME FROM {originTable}.INFORMATION_SCHEMA.TABLES origin FULL OUTER JOIN INFORMATION_SCHEMA.TABLES destination ON origin.TABLE_NAME = destination.TABLE_NAME WHERE (origin.TABLE_TYPE = 'BASE TABLE' OR destination.TABLE_TYPE = 'BASE TABLE') AND (origin.TABLE_NAME IS NULL OR destination.TABLE_NAME IS NULL)"

        sqlConn.DaDestination.Fill(dtDiffs)

        For i As Integer = 0 To dtDiffs.Rows.Count - 1
            diffs.Add(dtDiffs.Rows(i).Item("TABLE_NAME"))
        Next

        Return diffs
    End Function

    Private Function GetOriginDiffs() As List(Of String)
        Dim dtDiffs As New DataTable
        Dim diffs As New List(Of String)

        sqlConn.CmdOrigin.CommandText = "SELECT DB_NAME()"
        Dim originTable As String = sqlConn.CmdOrigin.ExecuteScalar()
        sqlConn.CmdDestination.CommandText = $"SELECT *, ISNULL(origin.TABLE_NAME, destination.TABLE_NAME) TABLE_NAME FROM {originTable}.INFORMATION_SCHEMA.TABLES origin LEFT JOIN INFORMATION_SCHEMA.TABLES destination ON origin.TABLE_NAME = destination.TABLE_NAME WHERE (origin.TABLE_TYPE = 'BASE TABLE' OR destination.TABLE_TYPE = 'BASE TABLE') AND destination.TABLE_NAME IS NULL "

        sqlConn.DaDestination.Fill(dtDiffs)

        For i As Integer = 0 To dtDiffs.Rows.Count - 1
            diffs.Add(dtDiffs.Rows(i).Item("TABLE_NAME"))
        Next

        Return diffs
    End Function

    Private Sub ReseedAndDelete(tableName As String)
        Try
            Dim tableIsIdentity As Boolean = Me.HasIdentity(tableName)

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
            Dim dtOrigin As DataTable
            Dim dtDestination As DataTable
            Dim page As Int64 = 0

            Do
                dtOrigin = SelectPage(tableName, page, "origin")
                dtDestination = SelectPage(tableName, page, "destination")

                If tableIsIdentity = Nothing Then
                    tableIsIdentity = Me.HasIdentity(tableName)
                End If

                ' Si es identidad se pone en on
                If tableIsIdentity Then
                    Me.SetIdentityInsert(True, tableName)
                End If

                For Each row As DataRow In dtOrigin.Rows

                    ' Se migran los datos
                    sqlConn.CmdDestination.CommandText = Me.generateInsertQuery(tableName, dtOrigin.Columns, dtDestination.Columns, row)
                    If sqlConn.CmdDestination.CommandText <> String.Empty Then
                        sqlConn.CmdDestination.ExecuteNonQuery()

                        Console.WriteLine(sqlConn.CmdDestination.CommandText)

                        If sqlConn.CmdDestination.Parameters.Count > 0 Then
                            sqlConn.CmdDestination.Parameters.Clear()
                        End If
                    Else
                        Exit For
                    End If
                Next

                If tableIsIdentity Then
                    Me.SetIdentityInsert(False, tableName)
                End If

                page += 1
            Loop While dtOrigin.Rows.Count > 0 Or dtDestination.Rows.Count > 0

            insertedTables.Add(tableName)
            originColumns.Add(dtOrigin.Columns.Count)
            destinationColumns.Add(dtDestination.Columns.Count)

            Dim insertedColumnsCount As Integer
            For Each col As DataColumn In dtOrigin.Columns
                If dtDestination.Columns.Contains(col.ColumnName) Then
                    insertedColumnsCount += 1
                End If
            Next
            insertedColumns.Add(insertedColumnsCount)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function SelectPage(tableName As String, page As Long, db As String) As DataTable
        ' Consulta Select de una tabla realizada por páginas, el "ORDER BY GETDATE()" solamente se utiliza como constante
        ' Significa que la consulta no se debe ordenar. Ya que SQL 2008 requiere un orden en esta consulta, pero no
        ' permite utilizar ORDER BY NULL.

        Dim dt As New DataTable
        Dim pageSize As Integer = 10000
        Dim cmd As String = $"SELECT * FROM (SELECT ROW_NUMBER() OVER (ORDER BY GETDATE()) AS RowNum, * FROM {tableName}) AS RowConstrainedResult WHERE RowNum >= {1 + pageSize * page} AND RowNum <= {pageSize + pageSize * page} ORDER BY RowNum"

        If db = "origin" Then
            sqlConn.CmdOrigin.CommandText = cmd
            sqlConn.DaOrigin.Fill(dt)
        Else
            sqlConn.CmdDestination.CommandText = cmd
            sqlConn.DaDestination.Fill(dt)
        End If

        dt.Columns.Remove("RowNum")

        Return dt
    End Function

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
        Dim originColumnNames As New List(Of String)
        Dim destinationColumnNames As New List(Of String)

        Dim values As New List(Of String)
        Dim columns As New List(Of String)

        For Each colOrigin As DataColumn In columnsOrigin
            originColumnNames.Add(colOrigin.ColumnName)
        Next

        For Each colDestination As DataColumn In columnsDestination
            destinationColumnNames.Add(colDestination.ColumnName)
        Next

        columns = originColumnNames.Intersect(destinationColumnNames).ToList()

        For Each col As String In columns
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

            If row.Item(col).GetType = GetType(DateTime) Then
                Dim dateVal As DateTime = DirectCast(row.Item(col), DateTime)
                value = $"CONVERT(DATETIME, '{dateVal.Year}-{dateVal.Month}-{dateVal.Day} {dateVal.Hour}:{dateVal.Minute}:{dateVal.Second}.{dateVal.Millisecond}')"
            End If

            ' Si es array de bytes
            If row.Item(col).GetType = GetType(Byte()) Then
                Dim param As SqlClient.SqlParameter = sqlConn.CmdDestination.Parameters.Add($"@Content{columns.IndexOf(col)}", SqlDbType.VarBinary)
                param.Value = CType(row.Item(col), Byte())
                param.Size = CType(row.Item(col), Byte()).Count()
                value = $"@Content{columns.IndexOf(col)}"
            End If

            values.Add(value)
        Next

        Return IIf(columns.Count() > 0, $"insert into {tableName} ({String.Join(",", columns)}) values ({String.Join(",", values)})", Nothing)
    End Function

    Private Sub SetIdentityInsert(active As Boolean, table As String)
        Try
            sqlConn.CmdDestination.CommandText = $"SET IDENTITY_INSERT {table} {IIf(active, "ON", "OFF")}"

            sqlConn.CmdDestination.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Function HasIdentity(tableName As String) As Boolean
        Dim isIdentity As Boolean = False

        Try
            sqlConn.CmdDestination.CommandText = $"SELECT COUNT(*) 'HAS_IDENTITY' FROM SYS.IDENTITY_COLUMNS WHERE OBJECT_NAME(OBJECT_ID) = '{tableName}' AND OBJECT_SCHEMA_NAME(object_id) = 'dbo'"

            If sqlConn.CmdDestination.ExecuteScalar() > 0 Then
                isIdentity = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return isIdentity
    End Function

    Private Sub bgwMigrate_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bgwMigrate.DoWork
        Try
            Me.Migrate(sender, cbReseedAndDelete.Checked)
        Catch ex As Exception
            e.Cancel = True
        End Try
    End Sub

    Private Sub bgwMigrate_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwMigrate.ProgressChanged
        Try
            lbInsertedTables.DataSource = Nothing
            lbInsertedTables.DataSource = insertedTables
            lblProgressText.Text = progressText
            lblAmountInserted.Text = $"Cantidad: {insertedTables.Count()}"
            pbMigration.Value = e.ProgressPercentage
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub bgwMigrate_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgwMigrate.RunWorkerCompleted
        gbOrigin.Enabled = True
        gbDestination.Enabled = True
        If Not e.Cancelled Then
            lbInsertedTables.DataSource = Nothing
            lbInsertedTables.DataSource = insertedTables
            lblAmountInserted.Text = $"Cantidad: {insertedTables.Count()}"
            pbMigration.Value = 0
            lblProgressText.Text = ""

            MsgBox("Migración realizada con éxito.", MsgBoxStyle.Information, "Finalizado")
        End If
    End Sub

    Private Sub btnAnalyze_Click(sender As Object, e As EventArgs) Handles btnAnalyze.Click
        Try
            gbOrigin.Enabled = False
            gbDestination.Enabled = False
            pbMigration.Style = ProgressBarStyle.Marquee
            bgwAnalyze.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub bgwAnalyze_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgwAnalyze.DoWork
        Try
            ConnectToDatabase()
            diffs = Me.GetOriginDiffs()
            Analyze()
        Catch ex As Exception
            e.Cancel = True
        End Try
    End Sub

    Private Sub bgwAnalyze_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles bgwAnalyze.ProgressChanged
        Me.lbInsertedTables.DataSource = Me.insertedTables
        pbMigration.Value = e.ProgressPercentage
    End Sub

    Private Sub bgwAnalyze_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bgwAnalyze.RunWorkerCompleted
        gbOrigin.Enabled = True
        gbDestination.Enabled = True
        pbMigration.Style = ProgressBarStyle.Blocks

        If Not e.Cancelled Then
            cbReseedAndDelete.Enabled = True
            btnMigrate.Enabled = True
            lblAnalyze.Text = "Seleccione las tablas que desea migrar."
            lblAmountAnalyzed.Text = $"Cantidad: {analyzedTables.Count()}"
            Dim source = analyzedTables.Select(Function(item) item.Clone).ToList()
            source.Sort()
            clbAnalyzedTables.DataSource = source

            For i As Int64 = 0 To clbAnalyzedTables.Items.Count - 1
                clbAnalyzedTables.SetItemCheckState(i, CheckState.Checked)
            Next
        End If
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim sfdExcelExport As New SaveFileDialog()

        sfdExcelExport.Filter = "xlsx files (*.xlsx)|All files (*.*)"

        If sfdExcelExport.ShowDialog() = DialogResult.OK Then
            Dim excel As New Excel.Application

            excel.Workbooks.Add()

            Dim worksheet As Excel.Worksheet = DirectCast(excel.ActiveSheet, Excel.Worksheet)

            worksheet.Cells(1, "A") = "Tablas Analizadas"
            worksheet.Cells(1, "B") = "Tablas No Migradas"
            worksheet.Cells(1, "D") = "Tablas Migradas"
            worksheet.Cells(1, "E") = "Columnas Origen"
            worksheet.Cells(1, "F") = "Columnas Destino"
            worksheet.Cells(1, "G") = "Columnas migradas"

            For i As Int64 = 1 To clbAnalyzedTables.Items.Count
                worksheet.Cells(i + 1, "A") = clbAnalyzedTables.Items(i - 1)
            Next

            For i As Int64 = 1 To notMigratedTables.Count
                worksheet.Cells(i + 1, "B") = notMigratedTables.ElementAt(i - 1)
            Next

            For i As Int64 = 1 To lbInsertedTables.Items.Count
                worksheet.Cells(i + 1, "D") = insertedTables.ElementAt(i - 1)
            Next

            For i As Int64 = 1 To originColumns.Count
                worksheet.Cells(i + 1, "E") = originColumns.ElementAt(i - 1)
            Next

            For i As Int64 = 1 To destinationColumns.Count
                worksheet.Cells(i + 1, "F") = destinationColumns.ElementAt(i - 1)
            Next

            For i As Int64 = 1 To insertedColumns.Count
                worksheet.Cells(i + 1, "G") = insertedColumns.ElementAt(i - 1)
            Next


            worksheet.Columns(1).AutoFit()
            worksheet.Columns(2).AutoFit()

            worksheet.SaveAs(sfdExcelExport.FileName)
        End If
    End Sub

    Private Sub btnMigrate_Click(sender As Object, e As EventArgs) Handles btnMigrate.Click
        Try
            gbOrigin.Enabled = False
            gbDestination.Enabled = False
            lblAnalyze.Text = ""
            bgwMigrate.RunWorkerAsync()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnSelectAll_Click(sender As Object, e As EventArgs) Handles btnSelectAll.Click
        For i As Integer = 0 To clbAnalyzedTables.Items.Count - 1
            clbAnalyzedTables.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub btnUnselectAll_Click(sender As Object, e As EventArgs) Handles btnUnselectAll.Click
        For i As Integer = 0 To clbAnalyzedTables.Items.Count - 1
            clbAnalyzedTables.SetItemChecked(i, False)
        Next
    End Sub
End Class
