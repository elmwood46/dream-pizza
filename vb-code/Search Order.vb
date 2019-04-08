Public Class frm_order_search
    '=====================================
    ' DISPLAY MEMBER DETAILS
    '=====================================
    Private Sub lb_memberID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_memberID.SelectedIndexChanged
        ' clear textbox
        txt_display.Clear()

        ' check for existing textbox selection - do nothing if a NULL box is clicked
        If Not lb_memberID.SelectedItem = 0 Then
            ' get id
            Dim selectedID = lb_memberID.SelectedItem '.ToString()
            currentSelected = selectedID

            'establish database variables
            Dim con As New OleDb.OleDbConnection
            Dim da As OleDb.OleDbDataAdapter
            Dim ds As DataSet = New DataSet
            Dim sql As String

            'Open connection
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\dreamPizza.accdb; Jet OLEDB:Database Password = ;"
            con.Open()

            'fill dataset with all data
            sql = "SELECT title, firstName, lastName, address, phoneNo, membership, delivery FROM tbl_members WHERE memberNo = " & selectedID

            da = New OleDb.OleDbDataAdapter(sql, con)
            da.Fill(ds, "Anything")

            ' process data & display
            Dim tabsiz As Integer = 8
            Dim boldFont As New Font(txt_display.Font.FontFamily, 10, FontStyle.Bold)
            Dim normFont As New Font(txt_display.Font.FontFamily, 10, FontStyle.Regular)

            'define headers
            Dim header(4) As String
            header(0) = "Name: "
            header(1) = "Address: "
            header(2) = "Phone: "
            header(3) = "Membership Status: "
            header(4) = "Delivery Instructions: "

            'define text
            Dim content(4) As String
            content(0) = ds.Tables(0).Rows(0).ItemArray(0) & " " & ds.Tables(0).Rows(0).ItemArray(1) & " " & ds.Tables(0).Rows(0).ItemArray(2)
            content(1) = ds.Tables(0).Rows(0).ItemArray(3)
            content(2) = ds.Tables(0).Rows(0).ItemArray(4)
            content(3) = ds.Tables(0).Rows(0).ItemArray(5)
            content(4) = """" & ds.Tables(0).Rows(0).ItemArray(6) & """"

            'draw details
            For n = 0 To 4
                txt_display.Font = boldFont
                txt_display.Text = txt_display.Text & header(n)
                txt_display.Font = normFont
                txt_display.Text = txt_display.Text & Environment.NewLine
                txt_display.SelectionIndent = tabsiz
                txt_display.Text = txt_display.Text & content(n)
                txt_display.Text = txt_display.Text & Environment.NewLine
                txt_display.Text = txt_display.Text & Environment.NewLine
                txt_display.SelectionIndent = 0
            Next

            'close connection
            con.Close()
        End If

    End Sub

    '=====================================
    ' SEARCH BAR
    '=====================================
    Private Sub txt_search_TextChanged(sender As Object, e As EventArgs) Handles txt_search.TextChanged

        ' clear selected item
        lb_memberID.SelectedItem = 0
        txt_display.Clear()

        If txt_search.Text IsNot "" Then

            ' clear current listbox contents
            lb_members.Items.Clear()
            lb_dots.Items.Clear()
            lb_memberID.Items.Clear()

            ' establish database variables
            Dim mySearch As String = txt_search.Text.ToLower()
            Dim con As New OleDb.OleDbConnection
            Dim da As OleDb.OleDbDataAdapter
            Dim ds As DataSet = New DataSet
            Dim sql As String

            'Open connection
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\dreamPizza.accdb; Jet OLEDB:Database Password = ;"
            con.Open()

            'fill dataset with all data
            sql = "SELECT [memberNo], [firstName], [lastName], [phoneNo] FROM tbl_members"
            da = New OleDb.OleDbDataAdapter(sql, con)
            da.Fill(ds, "Anything")

            'process results
            Dim intRowCount, intColumnCount As Integer
            intRowCount = ds.Tables(0).Rows.Count - 1
            intColumnCount = ds.Tables(0).Columns.Count - 1

            ' check for search match
            For row = 0 To intRowCount
                For col = 0 To intColumnCount
                    If InStr(ds.Tables(0).Rows(row).ItemArray(col).ToString.ToLower, mySearch, ) > 0 Then
                        If Not col = 0 Then
                            lb_members.Items.Add(ds.Tables(0).Rows(row).ItemArray(col).ToString)
                            If Not lb_dots.Items.Count >= 13 Then
                                lb_dots.Items.Add(".......")
                            End If
                            lb_memberID.Items.Add(ds.Tables(0).Rows(row).ItemArray(0).ToString)
                        End If
                    End If
                Next col
            Next row

            'Close connection
            da.Dispose()
            ds.Dispose()
            con.Close()
            con.Dispose()

        Else
            ' check if no text
            ' clear current listbox contents
            lb_members.Items.Clear()
            lb_dots.Items.Clear()
            lb_memberID.Items.Clear()
        End If

        ' set current selected to nothing
        currentSelected = ""
    End Sub

    '=====================================
    ' SCROLL LB_MEMBERS WITH TIMER
    '=====================================
    Private Sub tim_updatelistbox_Tick(sender As Object, e As EventArgs) Handles tim_updatelistbox.Tick
        lb_members.TopIndex = lb_memberID.TopIndex
    End Sub

    '=====================================
    ' EXIT BUTTON
    '=====================================
    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        Me.Hide()
        frm_homescreen.Show()
        'frm_homescreen.Location = Me.Location
        txt_search.Text = ""
        currentSelected = ""
        txt_display.Text = ""
    End Sub

    '=====================================
    ' LOAD FORM
    '=====================================
    Private Sub frm_members_Load(sender As Object, e As EventArgs) Handles Me.Load
        'CenterToScreen()
    End Sub

    '=====================================
    ' CONFIRM ENABLED TIMER
    '=====================================
    Private Sub tim_updateConfirm_Tick(sender As Object, e As EventArgs) Handles tim_updateConfirm.Tick
        If Not currentSelected = "" Then
            btn_confirm.Enabled = True
        Else
            btn_confirm.Enabled = False
        End If
    End Sub

    '=====================================
    ' CONFIRM BUTTON
    '=====================================
    Private Sub btn_confirm_Click(sender As Object, e As EventArgs) Handles btn_confirm.Click
        Me.Hide()
        frm_pizzaList.Show()
        frm_pizzaList.Location = Me.Location
        txt_search.Text = ""
        txt_display.Text = ""
    End Sub
End Class