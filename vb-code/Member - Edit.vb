Public Class frm_editMember

    ' create initial saving values
    Dim orderedControls(6)
    Dim initialValue(6)
    Dim selectedMember = currentSelected

    '=====================================
    ' FORM VISIBLE
    '=====================================
    Private Sub frm_editMember_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Me.Visible = True Then
            ' set the edited member
            selectedMember = currentSelected
            currentSelected = ""
            ' populate form with data
            'establish database variables
            Dim con As New OleDb.OleDbConnection
            Dim da As OleDb.OleDbDataAdapter
            Dim ds As DataSet = New DataSet
            Dim sql As String

            'Open connection
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|\dreamPizza.accdb; Jet OLEDB:Database Password = ;"
            con.Open()

            'fill dataset with all data
            sql = "SELECT [title], [firstName], [lastName], [address], [phoneNo], [membership], [delivery] FROM [tbl_members] WHERE memberNo = " & selectedMember

            da = New OleDb.OleDbDataAdapter(sql, con)
            da.Fill(ds, "Anything")

            ' record the controls in order
            orderedControls(0) = cbox_title_edit
            orderedControls(1) = txt_firstName_edit
            orderedControls(2) = txt_lastName_edit
            orderedControls(3) = txt_address_edit
            orderedControls(4) = txt_phone_edit
            orderedControls(5) = cbox_member_edit
            orderedControls(6) = txt_delivery_edit

            ' fill template with data & record the initial values of the data
            lab_memberName.Text = selectedMember
            For n = 0 To 6
                orderedControls(n).text = ds.Tables(0).Rows(0).ItemArray(n)
                initialValue(n) = ds.Tables(0).Rows(0).ItemArray(n)
            Next
            con.Close()

        End If
    End Sub

    '=====================================
    ' CONFIRM BUTTON
    '=====================================
    Private Sub btn_confirm_Click(sender As Object, e As EventArgs) Handles btn_confirm.Click

        ' check user input for speech marks
        If checkSpeechMarks() = False Then
            Exit Sub
        End If

        'CHECK FOR BOXES FILLED
        For n = 0 To 6
            If orderedControls(n).text = "" Then
                ' exit the for loop on no value
                If lab_nullMsg.Visible = True Then
                    ' start flashing timer
                    tim_blinktext.Enabled = True
                Else
                    lab_nullMsg.Visible = True
                End If
                Exit Sub
            End If
        Next

        'IF ALL INPUTS ARE CONFIRMED:---------------------------
        ' CHECK FOR VALID PHONE
        If Not IsNumeric(txt_phone_edit.Text.Replace(" ", "")) Then
            MsgBox("Please enter a valid phone number.")
            Exit Sub
        End If
        ' -------------------------------------------------------

        'UPDATE DATABASE: 
        ' establish database variables
        Dim con As New OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        Dim ds As DataSet = New DataSet
        Dim sql As String = ""

        'Open connection
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|dreamPizza.accdb; Jet OLEDB:Database Password = ;"
        con.Open()

        'fill dataset with all data
        Dim short_table(6)
        short_table(0) = "[title]"
        short_table(1) = "[firstName]"
        short_table(2) = "[lastName]"
        short_table(3) = "[address]"
        short_table(4) = "[phoneNo]"
        short_table(5) = "[membership]"
        short_table(6) = "[delivery]"

        Dim sqlInsertQuery As New OleDb.OleDbCommand(sql, con)

        sql = "UPDATE tbl_members SET "
        For n = 0 To 6
            sql = sql & short_table(n) & " = " & "'" & orderedControls(n).text & "'"
            If Not n = 6 Then
                sql = sql & ", "
            End If
        Next
        sql = sql & " WHERE memberNo = " & selectedMember

        ' update database and close connection
        sqlInsertQuery.Connection = con
        sqlInsertQuery.CommandText = sql
        sqlInsertQuery.ExecuteNonQuery()
        sqlInsertQuery.Dispose()
        ds.Dispose()
        con.Close()
        con.Dispose()

        MsgBox("Member #" & selectedMember & " was updated successfully.")
        exit_form()
    End Sub

    '=====================================
    ' EXIT FORM SUB
    '=====================================
    Sub exit_form()
        selectedMember = ""
        Me.Hide()
        'CenterToScreen()
        frm_members.Show()
    End Sub

    '=====================================
    ' EXIT BUTTON
    '=====================================
    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        Dim i = vbNo
        For n = 0 To 6
            If Not orderedControls(n).text = initialValue(n) Then
                i = MsgBox("Changes have been made. Exit anyway?", MsgBoxStyle.YesNo)
                Exit For
            ElseIf n = 6 Then
                i = vbYes
            End If
        Next
        If i = vbYes Then
            exit_form()
        Else
            Exit Sub
        End If
    End Sub

    '=====================================
    ' RESET BUTTON
    '=====================================
    Private Sub btn_reset_Click(sender As Object, e As EventArgs) Handles btn_reset.Click
        Dim i = MsgBox("Reset values?", MsgBoxStyle.YesNo)
        If i = vbYes Then
            For n = 0 To 6
                orderedControls(n).text = initialValue(n)
            Next
        End If
    End Sub

    '=====================================
    ' LABEL TIMER TICK
    '=====================================
    Dim timer As Integer = 0
    Private Sub tim_blinktext_Tick(sender As Object, e As EventArgs) Handles tim_blinktext.Tick
        'toggle timer switch
        Timer += 1

        If lab_nullMsg.Visible = False Then
            lab_nullMsg.Visible = True
        Else
            lab_nullMsg.Visible = False
        End If

        If Timer = 10 Then
            tim_blinktext.Stop()
            Timer = 0
        End If
    End Sub

    '=====================================
    ' DELETE MEMBER
    '=====================================
    Private Sub btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        Dim r = MsgBox("Are you sure you want to delete this member?", MsgBoxStyle.YesNo)
        If r = MsgBoxResult.Yes Then

            ' establish database variables
            Dim con As New OleDb.OleDbConnection
            'Dim da As OleDb.OleDbDataAdapter
            Dim ds As DataSet = New DataSet
            Dim sql As String = ""

            'Open connection
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|dreamPizza.accdb; Jet OLEDB:Database Password = ;"
            con.Open()
            sql = "DELETE * FROM tbl_members WHERE memberNo = " & selectedMember
            Dim sqlDeleteQuery As New OleDb.OleDbCommand(sql, con)

            ' delete from database and close connection
            sqlDeleteQuery.Connection = con
            sqlDeleteQuery.CommandText = sql
            sqlDeleteQuery.ExecuteNonQuery()
            sqlDeleteQuery.Dispose()
            ds.Dispose()
            con.Close()
            con.Dispose()

            'exit form
            MsgBox("Member #" & selectedMember & " was deleted.")
            exit_form()
        End If
    End Sub

    '=====================================
    ' ENABLE BUTTONS TIMER
    '=====================================
    Private Sub tim_checkControls_Tick(sender As Object, e As EventArgs) Handles tim_checkControls.Tick
        ' CHECK FOR NO CHANGES AT ALL
        Dim changed As Boolean = False
        ' turn buttons off if no edits were made
        For n = 0 To 6
            If Not orderedControls(n).text = initialValue(n) Then
                changed = True
            End If
        Next
        If changed = False Then
            btn_confirm.Enabled = False
            btn_reset.Enabled = False
        Else
            btn_confirm.Enabled = True
            btn_reset.Enabled = True
        End If
    End Sub

    '=====================================
    ' CHANGE "delivery" textbox text
    '=====================================
    Private Sub txt_delivery_TextChanged(sender As Object, e As EventArgs) Handles txt_delivery_edit.TextChanged
        lab_directionNumber.Text = txt_delivery_edit.TextLength & "/" & txt_delivery_edit.MaxLength
    End Sub

    '=====================================
    ' Change "address" textbox text
    '=====================================
    Private Sub txt_address_TextChanged(sender As Object, e As EventArgs) Handles txt_address_edit.TextChanged
        lab_addressNumber.Text = txt_address_edit.TextLength & "/" & txt_address_edit.MaxLength
    End Sub
End Class