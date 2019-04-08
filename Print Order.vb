Public Class frm_orderPrint
    ' create globalvars
    Dim nl = Environment.NewLine ' used for quickly putting in new lines
    Dim selectedMember
    Dim raw_price
    Dim discount_value
    Dim takeaway_true
    Dim memberTitle(6) ' used for defining the strings to label the member categories.
    Dim memberValue(6) ' used for defining the strings to label the values of each category. 

    '=====================================
    ' VISIBLE CHANGED
    '=====================================
    Private Sub frm_orderPrint_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        'visible = true
        If Me.Visible Then

            'define vars
            Dim selectedMember = orderArray(0)
            Dim raw_price As Double = orderArray(1)
            Dim discount_value As Double = orderArray(2)
            Dim takeaway_true = orderArray(3)
            memberTitle(0) = "Title: "
            memberTitle(1) = "Forename: "
            memberTitle(2) = "Surname: "
            memberTitle(3) = "Address: "
            memberTitle(4) = "Phone Number: "
            memberTitle(5) = "Membership: "
            memberTitle(6) = "Delivery Method: "

            ' clear list boxes
            mlb_name.Items.Clear()
            mlb_quantity.Items.Clear()
            mlb_price.Items.Clear()

            ' reset global arrays
            For k = 0 To orderArray.Length - 1
                orderArray(k) = 0
            Next

            'set offical label &date
            lab_official.Text = "Dream Pizza " & Date.Today & nl & "© PizzaSoft, 2017"
            lab_date.Text = "Date Ordered: " & Date.Today

            'open database for adding user data
            'establish database vars:
            Dim con = New OleDb.OleDbConnection
            Dim da As OleDb.OleDbDataAdapter
            Dim ds As DataSet = New DataSet
            Dim sql As String
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=|DataDirectory|dreamPizza.accdb; Jet OLEDB:Database Password = ;"
            con.Open()

            sql = "SELECT [title], [firstName], [lastName], [address], [phoneNo], [membership], [delivery] FROM [tbl_members] WHERE memberNo = " & selectedMember

            da = New OleDb.OleDbDataAdapter(sql, con)
            da.Fill(ds, "Anything")

            'set array of member values
            For k = 0 To 6
                memberValue(k) = ds.Tables(0).Rows(0).ItemArray(k)
            Next

            ' add data to form ==================================================

            'member state
            lab_memberType.Text = "Member Status: " & memberValue(5)

            'take away/deliver
            If takeaway_true Then
                lab_takeAway.Text = "TO TAKE AWAY"

            Else
                lab_takeAway.Text = "TO DELIVER"
            End If

            'name
            lab_userData.Text = "Name: "
            For k = 0 To 2
                lab_userData.Text = lab_userData.Text & ds.Tables(0).Rows(0).ItemArray(k) & " "
            Next

            ' set phone no
            lab_userData.Text = lab_userData.Text & nl & memberTitle(4) & ds.Tables(0).Rows(0).ItemArray(4) & nl

            'set address label
            lab_address.Text = ds.Tables(0).Rows(0).ItemArray(3)

            'set delivery label
            lab_delivery.Text = ds.Tables(0).Rows(0).ItemArray(6)

            'get raw and discount prices
            lab_price.Text = "Price: $" & raw_price

            'add discount
            Dim totalPrice = raw_price
            If discount_value = -1 Then
                totalPrice = raw_price + 5
                lab_discount.Text = "Extra $5 Delivery Fee"
            ElseIf Not discount_value = 0 Then
                totalPrice *= discount_value
                lab_discount.Text = "Discount: " & ((1 - discount_value) * 100) & "% "
            Else
                lab_discount.Text = "Discount: N/A"
            End If

            'add gst
            totalPrice *= 1.15

            ' set total price value
            lab_totalPrice.Text = "Total Price (+GST): $" & totalPrice

            'set pizza data     
            Dim len = pizzaArray.Length - 1
            For k = 0 To len
                mlb_name.Items.Add(pizzaArray(k))
                mlb_quantity.Items.Add(quantArray(k))
                mlb_price.Items.Add("$" & priceArray(k))
            Next

            'reset pizza arrays
            ReDim pizzaArray(0)
            ReDim quantArray(0)
            ReDim priceArray(0)
            pizzaArray(0) = 0
            quantArray(0) = 0
            priceArray(0) = 0

            ' close conection
            con.Close()
            con.Dispose()
            da.Dispose()
            ds.Dispose()
        End If
    End Sub

    '=====================================
    ' PRINT FORM
    '=====================================
    Private Sub btn_print_Click(sender As Object, e As EventArgs) Handles btn_print.Click
        btn_print.Visible = False
        pf_order.Print()
        btn_print.Visible = True
    End Sub
End Class