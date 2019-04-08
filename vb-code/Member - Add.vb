Public Class frm_addMember

    '=====================================
    ' FORM LOAD
    '=====================================
    Private Sub frm_addMember_ImeModeChanged(sender As Object, e As EventArgs) Handles Me.ImeModeChanged
        'start timer
        tim_blinktext.Enabled = False
    End Sub

    '=====================================
    ' CANCEL BUTTON
    '=====================================
    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click
        Dim i As Integer = 0

        'check to see if they started creating a member
        For Each cntrl As Control In Me.Controls
            If TypeOf cntrl Is TextBox Then
                If Not cntrl.Text.Replace(" ", "") = "" Then
                    i = MsgBox("Cancel member creation?", MsgBoxStyle.YesNo)
                    Exit For
                End If
            End If
        Next

        If i = vbYes Or i = 0 Then
            Me.Hide()
            Me.Controls.Clear()
            Me.InitializeComponent()
            frm_members.Show()
        End If
    End Sub

    '=====================================
    ' CONFIRM BUTTON
    '=====================================
    Private Sub btn_confim_Click(sender As Object, e As EventArgs) Handles btn_confim.Click

        ' check user input for speech marks
        If checkSpeechMarks() = False Then
            Exit Sub
        End If

        'CHECK FOR ALL INPUT
        For Each cntrl As Control In Me.Controls
            If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
                If cntrl.Text.Replace(" ", "") = "" Then
                    ' exit the for loop on no value
                    If lbl_nullMsg.Visible = True Then
                        ' start flashing timer
                        tim_blinktext.Enabled = True
                    Else
                        lbl_nullMsg.Visible = True
                    End If
                    Exit Sub
                End If
            End If
        Next

        'IF ALL INPUTS ARE CONFIRMED:---------------------------
        ' check for correct fields
        If Not IsNumeric(txt_phone.Text.Replace(" ", "")) Then
            MsgBox("Please enter a valid phone number.")
            Exit Sub
        End If

        ' establish database variables
        Dim con As New OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        Dim ds As DataSet = New DataSet
        Dim sql As String

        'Open connection

        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\dreamPizza.accdb; Jet OLEDB:Database Password = ;"
        con.Open()

        'fill dataset with all data
        sql = "SELECT * FROM tbl_members"

        da = New OleDb.OleDbDataAdapter(sql, con)
        da.Fill(ds, "Anything")

        ' check for maximum members
        If ds.Tables(0).Columns(0).ExtendedProperties.Count >= 99998 Then
            MsgBox("ERROR: Data capacity reached. Cannot add more members.")
            Exit Sub
        End If

        ' define loop vars
        Dim intRowCount, intColumnCount As Integer
        intRowCount = ds.Tables(0).Rows.Count - 1
        intColumnCount = ds.Tables(0).Columns.Count - 1

        ' define checking vars
        Dim fnameCheck = False
        Dim lnameCheck = False
        Dim phCheck = False

        ' check input with existing database
        For row = 0 To intRowCount
            For col = 0 To intColumnCount
                If ds.Tables(0).Rows(row).ItemArray(col).ToString.ToLower.Replace(" ", "") = txt_firstName.Text.ToLower.Replace(" ", "") Then
                    fnameCheck = True
                End If
                If ds.Tables(0).Rows(row).ItemArray(col).ToString.ToLower.Replace(" ", "") = txt_lastName.Text.ToLower.Replace(" ", "") Then
                    lnameCheck = True
                End If
                If ds.Tables(0).Rows(row).ItemArray(col) = txt_phone.Text.Replace(" ", "") Then
                    phCheck = True
                End If
            Next col
        Next row

        ' display confirmation message if member exists
        Dim i As Integer
        If phCheck = True Then
            ' check for phone and name
            If fnameCheck = True AndAlso lnameCheck = True Then
                i = MsgBox("A member with this name and phone number already exists. Continue anyway?", MsgBoxStyle.YesNo)
                ' check for just phone
            Else
                i = MsgBox("A member with this phone number already exists. Continue anyway?", MsgBoxStyle.YesNo)
            End If
        Else
            ' check for no phone but name
            If fnameCheck = True AndAlso lnameCheck = True Then
                i = MsgBox("A member with this name already exists. Continue anyway?", MsgBoxStyle.YesNo)
                ' no data match
            Else
                i = vbYes
            End If
        End If

        ' CONFIRMED MEMBER ADDED
        If i = vbYes Then

            ' GENERATE UNIQUE MEMBER ID
            Dim randomClass As New Random()
            Dim memval As New Integer
            memval = randomClass.Next(1, 99999)

            For row = 0 To intRowCount
                If memval = ds.Tables(0).Rows(row).ItemArray(0) Then
                    Do While memval = ds.Tables(0).Rows(row).ItemArray(0)
                        memval = randomClass.Next(1, 99999)
                    Loop
                    row = 0
                End If
            Next row

            'set sqlarray(0)
            Dim sqlArray(7)
            sqlArray(0) = memval.ToString
            sqlArray(1) = cbox_title.Text
            sqlArray(2) = txt_firstName.Text
            sqlArray(3) = txt_lastName.Text
            sqlArray(4) = txt_address.Text
            sqlArray(5) = txt_phone.Text.Replace(" ", "")
            sqlArray(6) = cbox_member.Text
            sqlArray(7) = txt_delivery.Text

            ' repair text
            For n = 1 To 7
                sqlArray(n).Replace("'", """")
            Next

            ' sql insert
            Dim sqlInsertQuery As New OleDb.OleDbCommand
            sql = "INSERT INTO tbl_members (memberNo, title, firstName, lastName, address, phoneNo, membership, delivery) VALUES ("

            ' add to the sql 
            For n = 0 To 7
                sql = sql & "'" & sqlArray(n) & "'"
                If Not n = 7 Then
                    sql = sql + ","
                Else
                    sql = sql + ");"
                End If
            Next

            ' insert into database & close connection
            ' check state
            Console.Write("State is " & con.State)

            sqlInsertQuery.Connection = con
            sqlInsertQuery.CommandText = sql
            sqlInsertQuery.ExecuteNonQuery()
            sqlInsertQuery.Dispose()
            da.Dispose()
            ds.Dispose()
            con.Close()
            con.Dispose()
            memval = 0
            sql = ""

            ' clear fields & wrap up
            MsgBox("Member Added")
            For Each cntrl As Control In Me.Controls
                If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
                    cntrl.ResetText()
                End If
            Next

            ' exit form
            Me.Hide()
            Me.Controls.Clear()
            Me.InitializeComponent()
            frm_members.Show()

        Else
            ' if they don't want to add a member with same name/phone, end
            da.Dispose()
            ds.Dispose()
            con.Close()
            con.Dispose()
        End If
    End Sub

    '=====================================
    ' TIMER TICK
    '=====================================
    Dim timer As Integer = 0
    Private Sub tim_blinktext_Tick(sender As Object, e As EventArgs) Handles tim_blinktext.Tick
        'toggle timer switch
        timer += 1

        If lbl_nullMsg.Visible = False Then
            lbl_nullMsg.Visible = True
        Else
            lbl_nullMsg.Visible = False
        End If

        If timer = 10 Then
            tim_blinktext.Stop()
            timer = 0
        End If
    End Sub

    '=====================================
    ' DELIVERY TEXT CHANGED
    '=====================================
    Private Sub txt_delivery_TextChanged(sender As Object, e As EventArgs) Handles txt_delivery.TextChanged
        lab_deliveryText.Text = txt_delivery.TextLength & "/" & txt_delivery.MaxLength
    End Sub

    '=====================================
    ' VISIBLE CHANGED
    '=====================================
    Private Sub frm_addMember_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Me.Visible Then
            ' if I am visible, set the initial value of the character count label
            lab_deliveryText.Text = "0/" & txt_delivery.MaxLength
        End If
    End Sub
End Class