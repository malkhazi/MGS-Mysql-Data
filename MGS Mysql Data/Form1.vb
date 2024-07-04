'Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Diagnostics.Eventing
Imports MySql.Data.MySqlClient
Imports Mysqlx
Public Class Form1
    Private cDt_Tables As New DataTable
    Private cDV_Tables As New DataView
    Private cDt_Procedures As New DataTable
    Private cDV_Procedures As New DataView
    Private cDt_Functions As New DataTable
    Private cDv_Functions As New DataView
    Private cDt_TableData As DataTable
    Private cflg_Det As Boolean = False
    Private cfli_indx1 As Integer = -1
    Private cfli_indx2 As Integer = -1
    Private cfli_indx3 As Integer = -1
    Private cfli_RowsCount As Integer = 10
    Private my_DBName As String = ""
    Private sql_txt As String
    Private dCd_MySQL As MySqlCommand
    Private dAd_MySQL As MySqlDataAdapter

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.AppIcon
        Me.TB_IP.Text = My_Get_Setting(Me.TB_IP.Name, "")
        Me.Tb_Port.Text = My_Get_Setting(Me.Tb_Port.Name, "")
        Me.TB_User.Text = My_Get_Setting(Me.TB_User.Name, "")
        Gl_Pass = Wrapp_My.DecryptData(My_Get_Setting(Me.Tb_Pass.Name, ""))
        Me.Tb_DBName.Text = My_Get_Setting(Me.Tb_DBName.Name, "")
        Me.Nud_LimitData.Value = CInt(My_Get_Setting(Me.Nud_LimitData.Name, "10"))
        cDt_Tables.Columns.Add("Id", GetType(Int32))
        cDt_Tables.Columns.Add("name", GetType(String))
        cDt_Tables.Columns.Add("count", GetType(Int64))

        cDV_Tables = cDt_Tables.DefaultView
        cDV_Procedures = cDt_Procedures.DefaultView
        cDv_Functions = cDt_Functions.DefaultView

        Dim fnt_name As String = My_Get_Setting("FontName", "").Trim
        Dim fnt_size As Single = Convert.ToSingle(My_Get_Setting("FontSize", "0"))
        Me.Tb_FontName.Text = Me.DGV_TableData.DefaultCellStyle.Font.Name
        Me.Tb_FontSize.Text = Me.DGV_TableData.DefaultCellStyle.Font.Size
        If fnt_name.Length > 0 And fnt_size > 0 Then
            Me.Tb_FontName.Text = fnt_name
            Me.Tb_FontSize.Text = fnt_size
            Try
                Dim newFont As New Font(Me.Tb_FontName.Text, Convert.ToSingle(Me.Tb_FontSize.Text), FontStyle.Regular)
                Me.DGV_TableData.DefaultCellStyle.Font = newFont
            Catch ex As Exception
                MsgBox(ex.Message,, "Error")
            End Try
        End If
    End Sub

    Private Sub Bt_Conn_Click(sender As Object, e As EventArgs) Handles Bt_Conn.Click
        Me.SplitContainer3.Panel1.Enabled = False
        Conn_str = String.Format("server={0};port={1};userid={2};password={3};database={4}",
                                 Me.TB_IP.Text.Trim,
                                 Me.Tb_Port.Text.Trim,
                                 Me.TB_User.Text.Trim,
                                 IIf(Me.Tb_Pass.Text.Trim.Length > 0, Me.Tb_Pass.Text.Trim, Gl_Pass),
                                 Me.Tb_DBName.Text.Trim)
        my_DBName = Me.Tb_DBName.Text.Trim
        Dim flg_conn As Boolean = False
        Using connection As New MySqlConnection(Conn_str)
            Try
                connection.Open()
                connection.Close()
                flg_conn = True
            Catch ex As MySqlException
                MsgBox(ex.Message,, "MySQL Error")
                Me.SplitContainer3.Panel1.Enabled = True
            Catch ex As Exception
                MsgBox(ex.Message,, "Error")
                Me.SplitContainer3.Panel1.Enabled = True
            End Try
        End Using
        If flg_conn Then
            My_Save_Setting(Me.TB_IP.Name, Me.TB_IP.Text.Trim)
            My_Save_Setting(Me.Tb_Port.Name, Me.Tb_Port.Text.Trim)
            My_Save_Setting(Me.TB_User.Name, Me.TB_User.Text.Trim)
            My_Save_Setting(Me.Tb_DBName.Name, Me.Tb_DBName.Text.Trim)
            If Me.Tb_Pass.Text.Trim.Length > 0 Then
                My_Save_Setting(Me.Tb_Pass.Name, Wrapp_My.EncryptData(Me.Tb_Pass.Text.Trim))
            End If
            My_Save_Setting(Me.Nud_LimitData.Name, Me.Nud_LimitData.Value)
            cfli_RowsCount = Me.Nud_LimitData.Value
            Pb_Download1.Visible = True
            Me.DGV_Tables.DataSource = Nothing
            Me.DGV_Procedures.DataSource = Nothing
            Me.DGV_Functions.DataSource = Nothing
            cDt_Tables.Clear()
            cDt_Procedures.Clear()
            cDt_Functions.Clear()
            Me.BW_01.RunWorkerAsync()
        End If
    End Sub

    Private Sub BW_01_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BW_01.DoWork
        Dim Err_txt As String = ""
        Dim i_ID As Integer = 0
        Dim tml_name As String = ""
        Dim ConnMysql As New MySqlConnection(Conn_str)
        ConnMysql.Open()
        Try
            dAd_MySQL = New MySqlDataAdapter("SHOW TABLES", ConnMysql)
            Dim tmp_DT As New DataTable
            dAd_MySQL.Fill(tmp_DT)
            For Each dr1 As DataRow In tmp_DT.Rows
                i_ID = +1
                tml_name = dr1(0)
                dCd_MySQL = New MySqlCommand($"SELECT COUNT(*) FROM {tml_name}", ConnMysql)
                Dim rt_Count As Integer = Convert.ToInt32(dCd_MySQL.ExecuteScalar)
                Dim dr2 As DataRow = cDt_Tables.NewRow
                dr2("id") = i_ID
                dr2("name") = tml_name
                dr2("count") = rt_Count
                cDt_Tables.Rows.Add(dr2)
            Next
            sql_txt = "SELECT ROUTINE_NAME name FROM information_schema.ROUTINES " &
              $"WHERE ROUTINE_TYPE='PROCEDURE' AND ROUTINE_SCHEMA='{my_DBName}'"
            dAd_MySQL = New MySqlDataAdapter(sql_txt, ConnMysql)
            dAd_MySQL.Fill(cDt_Procedures)
            sql_txt = "SELECT ROUTINE_NAME name FROM information_schema.ROUTINES " &
                $"WHERE ROUTINE_TYPE='FUNCTION' AND ROUTINE_SCHEMA='{my_DBName}'"
            dAd_MySQL = New MySqlDataAdapter(sql_txt, ConnMysql)
            dAd_MySQL.Fill(cDt_Functions)
        Catch ex As MySqlException
            Err_txt = "MySQL Error|" & ex.Message
        Catch ex As Exception
            Err_txt = "Error|" & ex.Message
        Finally
            ConnMysql.Close()
        End Try
        e.Result = Err_txt
    End Sub

    Private Sub BW_01_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BW_01.RunWorkerCompleted
        If e.Result.ToString.Trim.Length > 0 Then
            Dim err1() As String = e.Result.ToString.Split("|")
            MsgBox(err1(1),, err1(0))
        Else
            Me.DGV_Tables.AutoGenerateColumns = False
            Me.DGV_Tables.DataSource = cDV_Tables
            Me.DGV_Procedures.AutoGenerateColumns = False
            Me.DGV_Procedures.DataSource = cDV_Procedures
            Me.DGV_Functions.AutoGenerateColumns = False
            Me.DGV_Functions.DataSource = cDv_Functions
        End If
        Me.Pb_Download1.Visible = False
        Me.Bt_Disconnect.Enabled = True
        cflg_Det = True
    End Sub

    Private Sub DGV_Tables_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV_Tables.CellMouseDown
        cfli_indx1 = e.RowIndex
        Down_Rows()
    End Sub

    Private Sub DGV_Tables_SelectionChanged(sender As Object, e As EventArgs) Handles DGV_Tables.SelectionChanged
        If Me.DGV_Tables.SelectedCells.Count > 0 Then
            cfli_indx1 = Me.DGV_Tables.SelectedCells(0).RowIndex
            Down_Rows()
        End If
    End Sub

    Private Sub Down_Rows()
        If cfli_indx1 >= 0 And cflg_Det Then
            cflg_Det = False
            Me.DGV_TableData.DataSource = Nothing
            cDt_TableData = New DataTable
            Me.Pb_Download2.Visible = True
            BW_Det.RunWorkerAsync(Me.DGV_Tables.Rows(cfli_indx1).Cells("name_1").Value)
        End If
    End Sub

    Private Sub BW_Det_DoWork(sender As Object, e As DoWorkEventArgs) Handles BW_Det.DoWork
        Dim tbl_name As String = e.Argument
        Dim Err_txt As String = ""
        Dim ConnMysql As New MySqlConnection(Conn_str)
        ConnMysql.Open()
        Try
            sql_txt = "SELECT COLUMN_NAME FROM information_schema.COLUMNS " &
                $"WHERE TABLE_SCHEMA = '{my_DBName}' AND TABLE_NAME = '{tbl_name}' AND EXTRA LIKE '%auto_increment%'"
            dCd_MySQL = New MySqlCommand(sql_txt, ConnMysql)
            Dim AI_ColName As String = Convert.ToString(dCd_MySQL.ExecuteScalar())
            If Not String.IsNullOrEmpty(AI_ColName) Then
                dAd_MySQL = New MySqlDataAdapter($"SELECT * FROM {tbl_name} ORDER BY {AI_ColName} DESC LIMIT {cfli_RowsCount}", ConnMysql)
            Else
                dAd_MySQL = New MySqlDataAdapter($"SELECT * FROM {tbl_name}  LIMIT {cfli_RowsCount}", ConnMysql)
            End If
            dAd_MySQL.Fill(cDt_TableData)
        Catch ex As MySqlException
            Err_txt = "MySQL Error|" & ex.Message
        Catch ex As Exception
            Err_txt = "Error|" & ex.Message
        Finally
            ConnMysql.Close()
        End Try
        e.Result = Err_txt
    End Sub
    Private Sub BW_Det_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BW_Det.RunWorkerCompleted
        If e.Result.ToString.Trim.Length > 0 Then
            Dim err1() As String = e.Result.ToString.Split("|")
            MsgBox(err1(1),, err1(0))
        Else
            Me.DGV_TableData.DataSource = cDt_TableData
        End If
        Me.Pb_Download2.Visible = False
        cflg_Det = True
    End Sub

    Private Sub Bt_Disconnect_Click(sender As Object, e As EventArgs) Handles Bt_Disconnect.Click
        Me.Bt_Disconnect.Enabled = False
        Me.DGV_Tables.DataSource = Nothing
        Me.DGV_TableData.DataSource = Nothing
        cDt_Tables.Clear()
        cDt_TableData = Nothing
        Me.SplitContainer3.Panel1.Enabled = True
    End Sub

    Private Sub Bt_GetFontNames_Click(sender As Object, e As EventArgs) Handles Bt_GetFontNames.Click
        Dim FntDialog As New FontDialog()
        If FntDialog.ShowDialog() = DialogResult.OK Then
            Me.Tb_FontName.Text = FntDialog.Font.Name
            Me.Tb_FontSize.Text = FntDialog.Font.Size
            My_Save_Setting("FontName", Me.Tb_FontName.Text)
            My_Save_Setting("FontSize", Me.Tb_FontSize.Text)
            Try
                Dim newFont As New Font(Me.Tb_FontName.Text.Trim, Convert.ToSingle(Me.Tb_FontSize.Text), FontStyle.Regular)
                Me.DGV_TableData.DefaultCellStyle.Font = newFont
            Catch ex As Exception
                MsgBox(ex.Message,, "Error")
            End Try
        End If
    End Sub


    Private Sub DGV_Procedures_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV_Procedures.CellMouseDown
        cfli_indx2 = e.RowIndex
        Load_Procedure()
    End Sub

    Private Sub DGV_Procedures_SelectionChanged(sender As Object, e As EventArgs) Handles DGV_Procedures.SelectionChanged
        If Me.DGV_Procedures.SelectedCells.Count > 0 Then
            cfli_indx2 = Me.DGV_Procedures.SelectedCells(0).RowIndex
            Load_Procedure()
        End If
    End Sub

    Private Sub Load_Procedure()
        If cfli_indx2 >= 0 Then
            Dim cr_proc1 As String = Me.DGV_Procedures.Rows(cfli_indx2).Cells(0).Value
            Dim ConnMysql As New MySqlConnection(Conn_str)
            ConnMysql.Open()
            Try
                sql_txt = $"SHOW CREATE PROCEDURE {cr_proc1}"
                dCd_MySQL = New MySqlCommand(sql_txt, ConnMysql)
                Using Rd_MySQL As MySqlDataReader = dCd_MySQL.ExecuteReader()
                    If Rd_MySQL.Read() Then
                        Dim ProcTxt As String = Rd_MySQL.GetString(2)
                        Me.Rtb_ProcedureText.Text = $"Procedure name {cr_proc1}:" & vbCrLf & vbCrLf
                        Me.Rtb_ProcedureText.Text &= ProcTxt
                    Else
                        Me.Rtb_ProcedureText.Text = $"Procedure {cr_proc1} not found."
                    End If
                End Using
            Catch ex As MySqlException
                MsgBox(ex.Message,, "MySQL Error")
            Catch ex As Exception
                MsgBox(ex.Message,, "Error")
            Finally
                ConnMysql.Close()
            End Try
        End If
    End Sub

    Private Sub DGV_Functions_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Functions.CellContentClick

    End Sub

    Private Sub DGV_Functions_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGV_Functions.CellMouseDown
        cfli_indx3 = e.RowIndex
        Load_Function()
    End Sub

    Private Sub DGV_Functions_SelectionChanged(sender As Object, e As EventArgs) Handles DGV_Functions.SelectionChanged
        If Me.DGV_Functions.SelectedCells.Count > 0 Then
            cfli_indx3 = Me.DGV_Functions.SelectedCells(0).RowIndex
            Load_Function()
        End If
    End Sub

    Private Sub Load_Function()
        If cfli_indx3 >= 0 Then
            Dim cr_Func1 As String = Me.DGV_Functions.Rows(cfli_indx3).Cells(0).Value
            Dim ConnMysql As New MySqlConnection(Conn_str)
            ConnMysql.Open()
            Try
                sql_txt = $"SHOW CREATE FUNCTION {cr_Func1}"
                dCd_MySQL = New MySqlCommand(sql_txt, ConnMysql)
                Using Rd_MySQL As MySqlDataReader = dCd_MySQL.ExecuteReader()
                    If Rd_MySQL.Read() Then
                        Dim FuncTxt As String = Rd_MySQL.GetString(2)
                        Me.Rtb_Functions.Text = $"Function name {cr_Func1}:" & vbCrLf & vbCrLf
                        Me.Rtb_Functions.Text &= FuncTxt
                    Else
                        Me.Rtb_Functions.Text = $"Function {cr_Func1} not found."
                    End If
                End Using
            Catch ex As MySqlException
                MsgBox(ex.Message,, "MySQL Error")
            Catch ex As Exception
                MsgBox(ex.Message,, "Error")
            Finally
                ConnMysql.Close()
            End Try
        End If
    End Sub

    Private Sub Tb_Filtr1_TextChanged(sender As Object, e As EventArgs) Handles Tb_Filtr1.TextChanged
        Filtr_1()
    End Sub

    Private Sub Bt_FltrClear1_Click(sender As Object, e As EventArgs) Handles Bt_FltrClear1.Click
        Me.Tb_Filtr1.Text = ""
        Filtr_1()
    End Sub

    Private Sub Filtr_1()
        Try
            Dim Rt_Fltr As String = ""
            Dim Fltr_Add As String = ""
            Dim Fnd_txts() As String = Me.Tb_Filtr1.Text.Trim.Split(" ")
            For Each stxt1 As String In Fnd_txts
                If stxt1.Trim.Length > 0 Then
                    Rt_Fltr &= Fltr_Add & String.Format("(name like '%{0}%')", stxt1)
                    Fltr_Add = " AND "
                End If
            Next
            cDV_Tables.RowFilter = Rt_Fltr
        Catch ex As Exception
            cDV_Tables.RowFilter = ""
        End Try
    End Sub

    Private Sub Tb_Filtr2_TextChanged(sender As Object, e As EventArgs) Handles Tb_Filtr2.TextChanged
        Filtr_2()
    End Sub

    Private Sub Bt_FltrClear2_Click(sender As Object, e As EventArgs) Handles Bt_FltrClear2.Click
        Me.Tb_Filtr2.Text = ""
        Filtr_2()
    End Sub

    Private Sub Filtr_2()
        Try
            Dim Rt_Fltr As String = ""
            Dim Fltr_Add As String = ""
            Dim Fnd_txts() As String = Me.Tb_Filtr2.Text.Trim.Split(" ")
            For Each stxt2 As String In Fnd_txts
                If stxt2.Trim.Length > 0 Then
                    Rt_Fltr &= Fltr_Add & String.Format("(name like '%{0}%')", stxt2)
                    Fltr_Add = " AND "
                End If
            Next
            cDV_Procedures.RowFilter = Rt_Fltr
        Catch ex As Exception
            cDV_Procedures.RowFilter = ""
        End Try
    End Sub

    Private Sub Tb_Filtr3_TextChanged(sender As Object, e As EventArgs) Handles Tb_Filtr3.TextChanged
        Filtr_3()
    End Sub

    Private Sub Bt_FltrClear3_Click(sender As Object, e As EventArgs) Handles Bt_FltrClear3.Click
        Me.Tb_Filtr3.Text = ""
        Filtr_3()
    End Sub

    Private Sub Filtr_3()
        Try
            Dim Rt_Fltr As String = ""
            Dim Fltr_Add As String = ""
            Dim Fnd_txts() As String = Me.Tb_Filtr3.Text.Trim.Split(" ")
            For Each stxt3 As String In Fnd_txts
                If stxt3.Trim.Length > 0 Then
                    Rt_Fltr &= Fltr_Add & String.Format("(name like '%{0}%')", stxt3)
                    Fltr_Add = " AND "
                End If
            Next
            cDv_Functions.RowFilter = Rt_Fltr
        Catch ex As Exception
            cDv_Functions.RowFilter = ""
        End Try
    End Sub
End Class
