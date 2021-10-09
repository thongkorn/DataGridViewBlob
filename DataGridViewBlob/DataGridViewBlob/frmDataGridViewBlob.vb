#Region "ABOUT"
' / ---------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose: Sample code to retrieve data to show in DataGridView.
' / Microsoft Visual Basic .NET (2010) + MS Access 2010+
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / ---------------------------------------------------------------
#End Region

Imports System.Data.OleDb

Public Class frmDataGridViewBlob
    Dim Conn As OleDb.OleDbConnection
    Dim DA As New System.Data.OleDb.OleDbDataAdapter()
    Dim DR As OleDbDataReader
    Dim Cmd As New System.Data.OleDb.OleDbCommand
    Dim DT As New DataTable
    Dim strSQL As String

    '// Connect MS Access DataBase
    Function ConnectDataBase() As OleDb.OleDbConnection
        Return New OleDb.OleDbConnection( _
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
            MyPath(Application.StartupPath) & "data\" & "Countries.accdb;Persist Security Info=True")
    End Function

    ' / --------------------------------------------------------------------------------
    ' / Get my project path
    Function MyPath(ByVal AppPath As String) As String
        '/ MessageBox.Show(AppPath);
        MyPath = AppPath.ToLower.Replace("\bin\debug", "\").Replace("\bin\release", "\").Replace("\bin\x86\debug", "\")
        '/ Return Value
        '// If not found folder then put the \ (BackSlash) at the end.
        If Microsoft.VisualBasic.Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
    End Function

    Private Sub frmDataGridViewBlob_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Conn = ConnectDataBase()
        Call SetupDataGridView(dgvData)
        Call GetDataTable()
    End Sub

    Private Sub SetupDataGridView(ByRef DGV As DataGridView)
        With DGV
            .Columns.Clear()
            .RowTemplate.Height = 26
            .AllowUserToOrderColumns = True
            .AllowUserToDeleteRows = False
            .AllowUserToAddRows = False
            .ReadOnly = True
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .Font = New Font("Tahoma", 9)
            '// Even-Odd Color
            .AlternatingRowsDefaultCellStyle.BackColor = Color.OldLace
            .DefaultCellStyle.SelectionBackColor = Color.SeaGreen
            .ReadOnly = True
            .MultiSelect = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            '// Header Styles
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.Navy
                .ForeColor = Color.Black
                .Font = New Font("Tahoma", 10, FontStyle.Bold)
            End With
            '//
            Dim CountryPK As New DataGridViewTextBoxColumn
            With CountryPK
                .DataPropertyName = "CountryPK"
                .HeaderText = "CountryPK"
                '// Hidden Column.
                .Visible = False
            End With
            .Columns.Add(CountryPK)
            '//
            Dim Flag As New DataGridViewImageColumn
            With Flag
                .DataPropertyName = "Flag"
                .HeaderText = "Flag"
                .ImageLayout = DataGridViewImageCellLayout.Stretch
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            .Columns.Add(Flag)
            '//
            Dim A2 As New DataGridViewTextBoxColumn
            With A2
                .DataPropertyName = "A2"
                .HeaderText = "A2"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            .Columns.Add(A2)
            '//
            Dim Country As New DataGridViewTextBoxColumn
            With Country
                .DataPropertyName = "Country"
                .HeaderText = "Country"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            .Columns.Add(Country)
            '//
            Dim Capital As New DataGridViewTextBoxColumn
            With Capital
                .DataPropertyName = "Capital"
                .HeaderText = "Capital"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            .Columns.Add(Capital)
            '//
            Dim ZoneName As New DataGridViewTextBoxColumn
            With ZoneName
                .DataPropertyName = "ZoneName"
                .HeaderText = "ZoneName"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            .Columns.Add(ZoneName)

            Dim Population As New DataGridViewTextBoxColumn
            With Population
                .DataPropertyName = "Population"
                .HeaderText = "Population"
                .DefaultCellStyle.Format = "N0"
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            .Columns.Add(Population)
        End With
    End Sub

    '// การค้นหาข้อมูล หรือแสดงผลข้อมูลทั้งหมด จะใช้เพียงโปรแกรมย่อยแบบ Sub Program ชุดเดียว
    '// หากค่าที่ส่งมาเป็น False (หรือไม่ส่งมา จะถือเป็นค่า Default) นั่นคือให้แสดงผลข้อมูลทั้งหมด
    '// หากค่าที่ส่งมาเป็น True จะเป็นการค้นหาข้อมูล
    Private Sub GetDataTable(Optional ByVal blnSearch As Boolean = False)
        strSQL = _
            " SELECT CountryPK, Flag, A2, Country, Capital, Population, Zones.ZoneName " & _
            " FROM Countries INNER JOIN Zones ON Countries.ZoneFK = Zones.ZonePK "
        '// หากส่งค่า True มาจะเป็นการค้นหา
        If blnSearch Then
            strSQL = strSQL & _
                    " WHERE " & _
                    " [A2] " & " Like '%" & Trim(txtSearch.Text) & "%'" & " OR " & _
                    " [Country] " & " Like '%" & Trim(txtSearch.Text) & "%'" & " OR " & _
                    " [Capital] " & " Like '%" & Trim(txtSearch.Text) & "%'" & " OR " & _
                    " [ZoneName] " & " Like '%" & Trim(txtSearch.Text) & "%'"
            '// Else ไม่ต้องมี
        End If
        '// เอา strSQL มาเรียงต่อกัน
        strSQL = strSQL & " ORDER BY Countries.A2 "
        '/
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Try
            Cmd = Conn.CreateCommand
            Cmd.CommandText = strSQL
            DT = New DataTable
            DA = New OleDbDataAdapter(Cmd)
            DA.Fill(DT)
            dgvData.DataSource = DT
            lblRecordCount.Text = "[จำนวน : " & DT.Rows.Count & " รายการ.]"
            DT.Dispose()
            DA.Dispose()
            txtSearch.Text = ""
            txtSearch.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtSearch_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
        If Trim(txtSearch.Text.Trim) = "" Or Len(Trim(txtSearch.Text.Trim)) = 0 Then Return
        '// RetrieveData(True) It means searching for information.
        If e.KeyChar = Chr(13) Then '// Press Enter
            '// No beep.
            e.Handled = True
            '// Undesirable characters for the database.
            txtSearch.Text = txtSearch.Text.Trim.Replace("'", "").Replace("%", "").Replace("*", "")
            '// Send True Value is Search Data.
            Call GetDataTable(True)
        End If
    End Sub

    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        Call GetDataTable(True)
    End Sub

    Private Sub dgvData_CellDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvData.CellDoubleClick
        If dgvData.RowCount = 0 Then Return
        '// Read the value of the focus row.
        Dim iRow As Integer = dgvData.CurrentRow.Index
        MessageBox.Show( _
            "Primary key : " & dgvData.Item(0, iRow).Value & vbCrLf & _
            "Country : " & dgvData.Item(3, iRow).Value & vbCrLf & _
            "Capital is : " & dgvData.Item(4, iRow).Value)
    End Sub

    Private Sub frmDataGridViewBlob_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Conn.State = ConnectionState.Open Then Conn.Close()
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

End Class
