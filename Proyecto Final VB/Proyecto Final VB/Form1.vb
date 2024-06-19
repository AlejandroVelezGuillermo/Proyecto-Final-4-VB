Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.IO
Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Namespace Bases_Persona
    Partial Public Class Datos
        Inherits Form

        Shared ConexionString As String = "server = localhost\SQLEXPRESS; database = Musica; integrated security = true"
        Private Conexion As SqlConnection = New SqlConnection(ConexionString)
        Private Adaptador As SqlDataAdapter
        Private TablaDatos As DataTable

        Public Sub New()
            InitializeComponent()
            Dim consulta As String = "select * from dbo.Spotify_Songs"
            Adaptador = New SqlDataAdapter(consulta, Conexion)
            TablaDatos = New DataTable()
            Conexion.Open()
            Adaptador.Fill(TablaDatos)
            dataGridView1.DataSource = TablaDatos
            Dim btnExportar As Button = New Button With {
                .Text = "Exportar a Excel",
                .Location = New Point(10, 300)
            }
            btnExportar.Click += AddressOf BtnExportar_Click
            Me.Controls.Add(btnExportar)
            Dim btnImportar As Button = New Button With {
                .Text = "Importar desde Excel",
                .Location = New Point(150, 300)
            }
            btnImportar.Click += AddressOf BtnImportar_Click
            Me.Controls.Add(btnImportar)
        End Sub

        Private Sub BtnBuscar_Click(ByVal sender As Object, ByVal e As EventArgs)
            Busqueda1()
        End Sub

        Private Sub BtnRefrescar_Click(ByVal sender As Object, ByVal e As EventArgs)
            Recargar()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            Busqueda2()
        End Sub

        Private Sub BtnExportar_Click(ByVal sender As Object, ByVal e As EventArgs)
            ExportarExcel()
        End Sub

        Private Sub BtnImportar_Click(ByVal sender As Object, ByVal e As EventArgs)
            ImportarDesdeExcel()
        End Sub
        Private Sub Busqueda1()
            If radioButton1.Checked = True Then
                Dim consulta As String = "select * from dbo.Spotify_Songs where track_name =" & "'" & textBox1.Text & "'" & ""
                Adaptador = New SqlDataAdapter(consulta, Conexion)
                TablaDatos = New DataTable()
                Adaptador.Fill(TablaDatos)
                dataGridView1.DataSource = TablaDatos
            ElseIf radioButton2.Checked = True Then
                Dim consulta As String = "select * from dbo.Spotify_Songs where artist_s_name =" & "'" & textBox1.Text & "'" & ""
                Adaptador = New SqlDataAdapter(consulta, Conexion)
                TablaDatos = New DataTable()
                Adaptador.Fill(TablaDatos)
                dataGridView1.DataSource = TablaDatos
            ElseIf radioButton3.Checked = True Then
                Dim consulta As String = "select * from dbo.Spotify_Songs where released_year =" & textBox1.Text & ""
                Adaptador = New SqlDataAdapter(consulta, Conexion)
                TablaDatos = New DataTable()
                Adaptador.Fill(TablaDatos)
                dataGridView1.DataSource = TablaDatos
            Else
                MessageBox.Show("Porfavor Seleccione Una Opbcion y llene el campo.")
            End If
        End Sub
        Private Sub Busqueda2()
            If Not String.IsNullOrWhiteSpace(textBox2.Text) AndAlso Not String.IsNullOrWhiteSpace(textBox3.Text) Then
                Dim consulta As String = "select * from dbo.Spotify_Songs where artist_s_name = '" & textBox2.Text & "' and released_year = " + textBox3.Text & " ORDER BY track_name ASC"
                Adaptador = New SqlDataAdapter(consulta, Conexion)
                TablaDatos = New DataTable()
                Adaptador.Fill(TablaDatos)
                dataGridView1.DataSource = TablaDatos
            Else
                MessageBox.Show("Por favor, asegúrate de llenar ambos campos.")
            End If
        End Sub

        Private Sub Recargar()
            Dim consulta As String = "select * from dbo.Spotify_Songs"
            Adaptador = New SqlDataAdapter(consulta, Conexion)
            TablaDatos = New DataTable()
            Adaptador.Fill(TablaDatos)
            dataGridView1.DataSource = TablaDatos
            MessageBox.Show("Recargando Los Datos.")
        End Sub

        Private Sub ExportarExcel()
            Using sfd As SaveFileDialog = New SaveFileDialog() With {
                .Filter = "Excel Workbook|*.xlsx"
            }

                If sfd.ShowDialog() = DialogResult.OK Then

                    Using pck As ExcelPackage = New ExcelPackage()
                        Dim ws As ExcelWorksheet = pck.Workbook.Worksheets.Add("Sheet1")

                        For i As Integer = 0 To dataGridView1.Columns.Count - 1
                            ws.Cells(1, i + 1).Value = dataGridView1.Columns(i).HeaderText
                        Next

                        For i As Integer = 0 To dataGridView1.Rows.Count - 1

                            For j As Integer = 0 To dataGridView1.Columns.Count - 1
                                ws.Cells(i + 2, j + 1).Value = dataGridView1.Rows(i).Cells(j).Value?.ToString()
                            Next
                        Next

                        Dim bin = pck.GetAsByteArray()
                        File.WriteAllBytes(sfd.FileName, bin)
                    End Using

                    MessageBox.Show("Datos exportados exitosamente.")
                End If
            End Using
        End Sub

        Private Sub ImportarDesdeExcel()
            Using ofd As OpenFileDialog = New OpenFileDialog() With {
                .Filter = "Excel Workbook|*.xlsx"
            }

                If ofd.ShowDialog() = DialogResult.OK Then
                    Dim path As String = ofd.FileName
                    Dim excelConnectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};Extended Properties='Excel 12.0 Xml;HDR=YES;'"

                    Using excelConnection As OleDbConnection = New OleDbConnection(excelConnectionString)
                        excelConnection.Open()
                        Dim dtExcelSchema As DataTable
                        dtExcelSchema = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                        Dim sheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
                        Dim dataAdapter As OleDbDataAdapter = New OleDbDataAdapter($"SELECT * FROM [{sheetName}]", excelConnection)
                        Dim dt As DataTable = New DataTable()
                        dataAdapter.Fill(dt)

                        Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Conexion)
                            bulkCopy.DestinationTableName = "dbo.Spotify_Songs"
                            bulkCopy.WriteToServer(dt)
                        End Using

                        MessageBox.Show("Datos importados exitosamente.")
                        Recargar()
                    End Using
                End If
            End Using
        End Sub
    End Class
End Namespace
