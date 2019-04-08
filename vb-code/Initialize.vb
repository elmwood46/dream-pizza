Imports System.Data.SqlClient
Imports System.Text

Module Initialize
    '=====================================
    ' ESTABLISH GLOBAL VARIABLES
    '=====================================

    ' establish GST multiplier constant for future-proofing
    Public gst_value As Integer = 1.15

    ' global variable: current selected member
    Public currentSelected As String

    ' variables for efficiently manipulating the order information across forms
    Public orderArray(3)
    Public pizzaArray(0)
    Public quantArray(0)
    Public priceArray(0)

    '=========================================================
    ' FUNCTION: CHECK FOR QUOTATION/SPEECH MARKS IN USER INPUT
    ' This is necessary to prevent SQL injection errors,
    ' in certain user input scenarios. 
    ' It returns false if there are speech marks and true 
    ' if there are no speech marks. 
    '=========================================================
    Public Function checkSpeechMarks()
        ' check all input spaces (combo/text boxes) in the form 
        For Each cntrl As Control In Form.ActiveForm.Controls
            ' If there is a panel, check all input controls in the panel. 
            If TypeOf cntrl Is Panel Then
                For Each panCntrl As Control In cntrl.Controls
                    If TypeOf panCntrl Is TextBox Or TypeOf panCntrl Is ComboBox Then
                        If panCntrl.Text.Contains("'") Or panCntrl.Text.Contains("""") Then
                            MsgBox("User input may not contain inverted commas (') or quotation marks (""). ")
                            Return False
                            Exit Function
                        End If
                    End If
                Next
                ' if not a panel, check for textbox or combobox
            ElseIf TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
                If cntrl.Text.Contains("'") Or cntrl.Text.Contains("""") Then
                    MsgBox("User input may not contain inverted commas (') or quotation marks (""). ")
                    Return False
                    Exit Function
                End If
            End If
        Next
        ' if there are no speech marks, return true
        Return True
    End Function

End Module
