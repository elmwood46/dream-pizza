Public Class frm_pizzaList

    ' form vars
    Dim dvSelected
    Dim newline = Environment.NewLine
    Public selected_member
    Public memberValue(4)

    '=====================================
    ' FORM BECAME VISIBLE/INVISIBLE
    '=====================================
    Private Sub frm_pizzaList_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Visible = True Then

            'initialize -  reset these variables now that form is visible (they are used to add text to the reciept, and need to be reset here)
            selected_member = currentSelected
            For k = 0 To 4
                memberValue(k) = ""
            Next
            Dim newline = Environment.NewLine

            'establish database vars:
            Dim con = New OleDb.OleDbConnection
            Dim da As OleDb.OleDbDataAdapter
            Dim ds As DataSet = New DataSet
            Dim sql As String
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|dreamPizza.accdb; Jet OLEDB:Database Password = ;"
            con.Open()

            sql = "SELECT [title], [firstName], [lastName], [address], [membership] FROM [tbl_members] WHERE memberNo = " & currentSelected

            da = New OleDb.OleDbDataAdapter(sql, con)
            da.Fill(ds, "Anything")

            For n = 0 To 4
                memberValue(n) = ds.Tables(0).Rows(0).ItemArray(n)
            Next n

            ' populate member data label
            lab_memberData.Text = memberValue(0) & " " & memberValue(1) & " " & memberValue(2) & newline & memberValue(3) & newline & "Membership Status: " & memberValue(4)

            'refresh dataset
            ds.Dispose()
            ds = New DataSet

            ' POPULATE DATA GRID VIEW
            Try
                'delete current datagridview
                dgv_pizzas.Rows.Clear()

                'fill dataset with all data
                sql = "SELECT [pizzaName], [costSmall], [costLarge] FROM [tbl_pizza]"

                da = New OleDb.OleDbDataAdapter(sql, con)
                da.Fill(ds, "Anything")

                'process results
                Dim intRowCount, intColumnCount As Integer
                intRowCount = ds.Tables(0).Rows.Count - 1
                intColumnCount = ds.Tables(0).Columns.Count - 1

                ' populate table
                For row = 0 To intRowCount
                    dgv_pizzas.Rows.Add(ds.Tables(0).Rows(row).ItemArray(0), ds.Tables(0).Rows(row).ItemArray(1), ds.Tables(0).Rows(row).ItemArray(2))
                Next row

                ' set all rows and columns visible in the dataGridView
                For row = 0 To intRowCount
                    dgv_pizzas.Rows(row).Visible = True
                Next
                For col = 0 To intColumnCount
                    dgv_pizzas.Columns(col).Visible = True
                Next

                'close conection
                con.Close()
                con.Dispose()
                da.Dispose()
                ds.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, Me.Text)
            End Try

            'Set Labels
            getprice()
            lab_orderTitle.Text = "Order for Member: " & currentSelected & ", to deliver."
        End If
    End Sub

    '=====================================
    ' ENABLE BUTTONS
    '=====================================
    Private Sub dgv_pizzas_SelectionChanged(sender As Object, e As EventArgs) Handles dgv_pizzas.SelectionChanged

        If Not dgv_pizzas.SelectedRows.Count = Nothing Then
            ' update dvSelected
            dvSelected = dgv_pizzas.SelectedRows.Item(0).Cells(0).Value

            ' change button text
            btn_addLarge.Text = "Add Large " & newline & """" & dvSelected & """"
            btn_addSmall.Text = "Add Small " & newline & """" & dvSelected & """"

            'change label
            change_pizzanum_lab()
        End If
    End Sub

    '=====================================
    ' CHANGE PIZZA COUNTER LABEL
    '=====================================
    Sub change_pizzanum_lab()
        Dim sum As Integer = 0
        For n = 0 To mlb_quantity.Items.Count - 1
            sum += CInt(mlb_quantity.Items.Item(n))
        Next
        If sum = 1 Then
            lab_numPizzas.Text = sum.ToString & " pizza in order."
        Else
            lab_numPizzas.Text = sum.ToString & " pizzas in order."
        End If
    End Sub

    '=====================================
    ' ADD PIZZA
    '=====================================
    Sub add_pizza(source, size, amount)
        ' get size/price
        Dim sizePrice As Double
        If size = " (Small)" Then
            sizePrice = Convert.ToDouble(dgv_pizzas.SelectedRows.Item(0).Cells(1).Value)
        Else
            sizePrice = Convert.ToDouble(dgv_pizzas.SelectedRows.Item(0).Cells(2).Value)
        End If

        ' add to list box
        If lb_order.Items.Contains(source & size) Then
            Dim curr_amount As Integer = mlb_quantity.Items.Item(lb_order.Items.IndexOf(source & size))
            mlb_quantity.Items.RemoveAt(lb_order.Items.IndexOf(source & size))
            mlb_quantity.Items.Insert(lb_order.Items.IndexOf(source & size), curr_amount + amount)
            mlb_price.Items.RemoveAt(lb_order.Items.IndexOf(source & size))
            mlb_price.Items.Insert(lb_order.Items.IndexOf(source & size), (CDbl(mlb_quantity.Items.Item(lb_order.Items.IndexOf(source & size))) * sizePrice))
        Else
            lb_order.Items.Add(source.ToString & size)
            mlb_quantity.Items.Add(1)
            mlb_price.Items.Add((CDbl(mlb_quantity.Items.Item(lb_order.Items.IndexOf(source & size))) * sizePrice))
        End If

        'change labels
        change_pizzanum_lab()
        getprice()

    End Sub

    '=====================================
    ' GET DISCOUNT
    '=====================================
    Private Function getDiscount()
        ' this returns the discount value as a multiplier, 1 being no discount, 0.95 being 5% discount, etc.
        Dim myVal = memberValue(4).ToString.ToLower
        If chk_takeaway.Checked Then
            If myVal = "none" Then
                lab_discount.Text = "Discount: N/A"
                Return 1
            ElseIf myVal = "basic" Then
                lab_discount.Text = "Discount: 5%"
                Return 0.95
            ElseIf myVal = "premium" Then
                lab_discount.Text = "Discount: 15%"
                Return 0.85
            Else
                MsgBox("Error: Impossible Membership")
                End
            End If
        Else
            If myVal = "none" Then
                lab_discount.Text = "Extra $5 Delivery Fee"
                Return -1
            ElseIf myVal = "basic" Then
                lab_discount.Text = "Discount: N/A"
                Return 1
            ElseIf myVal = "premium" Then
                lab_discount.Text = "Discount: 10%"
                Return 0.9
            Else
                MsgBox("Error: Impossible Membership")
                End
            End If
        End If
    End Function

    '=====================================
    ' GET RAW PRICE
    '=====================================
    Private Function getRawPrice()
        Dim p As Double = 0
        If mlb_price.Items.Count > 0 Then
            For n = 0 To mlb_price.Items.Count - 1
                p += CDbl(mlb_price.Items.Item(n))
            Next
        End If
        Return p
    End Function

    '=====================================
    ' CACLULATE PRICE
    '=====================================
    Private Sub getprice()
        Dim sum As Double = getRawPrice()
        lab_price.Text = "Price: $" + sum.ToString

        If getDiscount() = -1 Then
            sum += 5
        Else
            sum *= getDiscount()
        End If

        ' alter price for gst
        sum *= gst_value

        lab_totalPrice.Text = "Total Price: $" & sum.ToString
    End Sub

    '=====================================
    ' BUTTON ADD SMALL
    '=====================================
    Private Sub btn_addLarge_Click(sender As Object, e As EventArgs) Handles btn_addSmall.Click
        add_pizza(dvSelected, " (Small)", 1)
    End Sub

    '=====================================
    ' ADD LARGE
    '=====================================
    Private Sub btn_addSmall_Click(sender As Object, e As EventArgs) Handles btn_addLarge.Click
        add_pizza(dvSelected, " (Large)", 1)
    End Sub

    '=====================================
    ' CHANGE LISTBOX_ORDER SELECTION
    '=====================================
    Private Sub lb_order_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lb_order.SelectedIndexChanged
        If Not lb_order.SelectedItem = Nothing Then
            btn_remove_one.Enabled = True
            btn_removeAll.Enabled = True
            btn_remove_one.Text = "Remove One " & newline & """" & lb_order.SelectedItem & """"
            btn_removeAll.Text = "Remove All " & newline & """" & lb_order.SelectedItem & """"
        End If
    End Sub

    '=====================================
    ' REMOVE SINGLE
    '=====================================
    Private Sub btn_remove_one_Click(sender As Object, e As EventArgs) Handles btn_remove_one.Click

        If Not lb_order.SelectedItem = Nothing Then
            ' dim curr_amount
            Dim curr_amount As Integer = CInt(mlb_quantity.Items.Item(lb_order.SelectedIndex))
            Dim curr_cost As Double = CDbl(mlb_price.Items.Item(lb_order.SelectedIndex))
            Dim sizePrice As Double = curr_cost / curr_amount

            ' remove pizza from listbox if quantity <= 0, else just reduce by one
            If ((curr_amount - 1) <= 0) Then
                mlb_quantity.Items.RemoveAt(lb_order.SelectedIndex)
                mlb_price.Items.RemoveAt(lb_order.SelectedIndex)
                lb_order.Items.RemoveAt(lb_order.SelectedIndex)

                ' disable buttons upon removing
                btn_remove_one.Enabled = False
                btn_removeAll.Enabled = False
                btn_remove_one.Text = "Remove One"
                btn_removeAll.Text = "Remove All"

            Else
                ' get the price
                mlb_quantity.Items.Item(lb_order.SelectedIndex) = curr_amount - 1
                mlb_price.Items.RemoveAt(lb_order.SelectedIndex)
                mlb_price.Items.Insert(lb_order.SelectedIndex, (CDbl(mlb_quantity.Items.Item((lb_order.SelectedIndex))) * sizePrice))
            End If
        End If
        ' change labels
        change_pizzanum_lab()
        getprice()
    End Sub

    '=====================================
    ' REMOVE ALL
    '=====================================
    Private Sub btn_removeAll_Click(sender As Object, e As EventArgs) Handles btn_removeAll.Click
        If Not lb_order.SelectedItem = Nothing Then
            ' remove pizza from listbox
            mlb_quantity.Items.RemoveAt(lb_order.SelectedIndex)
            mlb_price.Items.RemoveAt(lb_order.SelectedIndex)
            lb_order.Items.RemoveAt(lb_order.SelectedIndex)
            ' disable buttons upon removing
            btn_remove_one.Enabled = False
            btn_removeAll.Enabled = False
            btn_remove_one.Text = "Remove One"
            btn_removeAll.Text = "Remove All"
        End If
        ' change labels
        change_pizzanum_lab()
        getprice()
    End Sub

    '=====================================
    ' BACK BUTTON PRESSED
    '=====================================
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_back.Click

        ' define local checking variable
        Dim yn = vbYes

        ' check for cancelling
        If lb_order.Items.Count > 0 Then
            yn = MsgBox("This will cancel the current order. Continue?", MsgBoxStyle.YesNo)
        End If

        ' exit form
        If yn = vbYes Then

            'hide current form
            Me.Hide()

            ' close the "print reciept" box if it is open
            Dim formCollection As New FormCollection
            formCollection = Application.OpenForms()
            Dim formPrint As Form = frm_orderPrint
            If formPrint.Visible Then
                formPrint.Hide()
            End If

            ' clear the boxes in the order form
            dgv_pizzas.ClearSelection()
            dgv_pizzas.Rows(0).Selected = True
            mlb_price.Items.Clear()
            mlb_quantity.Items.Clear()
            lb_order.Items.Clear()
            frm_order_search.Show()
        End If
    End Sub

    '=====================================
    ' CLEAR ORDER
    '=====================================
    Private Sub btn_clear_Click(sender As Object, e As EventArgs) Handles btn_clear.Click
        Dim a = MsgBox("Are you sure to want to clear the current order?", MsgBoxStyle.YesNo)
        If a = vbYes Then
            lb_order.Items.Clear()
            mlb_price.Items.Clear()
            mlb_quantity.Items.Clear()

            'disable buttons
            btn_remove_one.Enabled = False
            btn_removeAll.Enabled = False
            btn_remove_one.Text = "Remove One"
            btn_removeAll.Text = "Remove All"

            'changelabels
            getprice()
            change_pizzanum_lab()
        End If
    End Sub

    '=====================================
    ' LISTBOX TIMER TICK
    '=====================================
    Private Sub tim_updateListbox_Tick(sender As Object, e As EventArgs) Handles tim_updateListbox.Tick
        ' synchronize listbox indexes
        mlb_quantity.TopIndex = lb_order.TopIndex
        mlb_price.TopIndex = lb_order.TopIndex

        ' control btn_confirmed
        If lb_order.Items.Count > 0 Then
            btn_confirm.Enabled = True
        Else
            btn_confirm.Enabled = False
        End If
    End Sub

    Private Sub frm_pizzaList_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    '=====================================
    ' CHECKBOX STATE CHANGED
    '=====================================
    Private Sub chk_takeaway_CheckedChanged(sender As Object, e As EventArgs) Handles chk_takeaway.CheckedChanged
        If chk_takeaway.Checked Then
            lab_orderTitle.Text = "Order for Member: " & selected_member & ", to take away."
            getprice()
        Else
            lab_orderTitle.Text = "Order for Member: " & selected_member & ", to deliver."
            getprice()
        End If
    End Sub

    '=====================================
    ' CONFIRM BUTTON CLICKED
    '=====================================
    Private Sub btn_confirm_Click(sender As Object, e As EventArgs) Handles btn_confirm.Click

        ' check for existing orders form, then create new
        Dim formCollection As New FormCollection
        formCollection = Application.OpenForms()
        Dim formPrint As Form = frm_orderPrint
        If formPrint.Visible Then
            formPrint.Hide()
        End If

        'populate orderArray
        orderArray(0) = selected_member
        orderArray(1) = getRawPrice()
        orderArray(3) = chk_takeaway.Checked
        orderArray(2) = getDiscount()

        'populate pizza_array
        ReDim pizzaArray(lb_order.Items.Count - 1)
        ReDim quantArray(lb_order.Items.Count - 1)
        ReDim priceArray(lb_order.Items.Count - 1)
        For n = 0 To lb_order.Items.Count - 1
            pizzaArray(n) = lb_order.Items.Item(n)
            quantArray(n) = mlb_quantity.Items.Item(n)
            priceArray(n) = mlb_price.Items.Item(n)
        Next

        ' show printing form
        frm_orderPrint.Show()
    End Sub
End Class