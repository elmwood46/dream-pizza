Public Class frm_homescreen
    '=====================================
    ' MEMBERS BUTTON CLICK
    '=====================================
    Private Sub btn_members_Click(sender As Object, e As EventArgs) Handles btn_members.Click
        Me.Hide()
        frm_members.Show()
    End Sub

    '=====================================
    ' EXIT BUTTON
    '=====================================
    Private Sub btn_exit_Click(sender As Object, e As EventArgs) Handles btn_exit.Click
        End
    End Sub

    '=====================================
    ' ORDER BUTTON
    '=====================================
    Private Sub btn_order_Click(sender As Object, e As EventArgs) Handles btn_order.Click
        Me.Hide()
        frm_order_search.Show()
    End Sub

    '=====================================
    ' FORM BECAME VISIBLE/INVISIBLE
    '=====================================
    Private Sub frm_homescreen_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Me.Visible = True Then
            CenterToScreen()
        End If
    End Sub
End Class