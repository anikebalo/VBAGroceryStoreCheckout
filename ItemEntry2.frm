VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ItemEntry2 
   ClientHeight    =   9800.001
   ClientLeft      =   260
   ClientTop       =   1010
   ClientWidth     =   12220
   OleObjectBlob   =   "ItemEntry2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ItemEntry2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ready As Boolean 'set variable to determine whether entry is completely error free

Private Sub UserForm_Initialize() 'reset form values whenever opened
Subtotal = 0
province = ""
optioncnt = 0
quantity1.Value = 0
price1.Value = 0

End Sub

Private Sub EnterTransaction_Click()

'set variables
Dim ctl As Control
Dim optioncnt As Integer
Dim ProvinceError As Boolean
Dim NameError As Boolean
Dim QuantError As Boolean
Dim Price98Error As Boolean
Dim PriceError As Boolean
Dim findecimal As Integer
Dim afterdecimal As String
Dim Lastprovince As String


'Set "error checkers" as False that will turn true if valid
ProvinceError = False
NameError = False
QuantError = False
Price98Error = False
PriceError = False
ready = False

'remember the province selected in the previous entry so comparisons can be done to see if it was changed between entries
Lastprovince = province

'look through every control on the userform
For Each ctl In Me.Controls
    'find the option button and if one is selected, count it and compare if it is the same province as in the previous transaction
    If TypeName(ctl) = "OptionButton" Then
        If ctl.Value = True Then
            'save the province
            province = ctl.Caption
            'warn the user if the province has changed between transactions
            If province <> Lastprovince And Lastprovince <> "" Then
                MsgBox "Please be aware that you have changed your province between transactions." & Chr(10) & "Only the last selected province before the transaction is confirmed will be used." & Chr(10) & "You have currently selected: " & province
            End If
            optioncnt = optioncnt + 1
        End If
    End If
    
    'find the name textbox
    If Left(ctl.Name, 4) = "name" Then
        If ctl.Value <> "" Then
            'if the value is not blank, turn textbox green
            ctl.BackColor = RGB(216, 245, 176)
        ElseIf ctl.Value = "" Then
            'if it is blank, turn textbox red and turn on error
            ctl.BackColor = RGB(235, 33, 33)
            NameError = True
        End If
    End If

    'find the quantity textbox
    If Left(ctl.Name, 5) = "quant" Then
        'if the value is not numeric change to textbox colour red and turn on error
        If Not (IsNumeric(ctl.Value)) Then
            ctl.BackColor = RGB(235, 33, 33)
            QuantError = True
        Else
        'if the value is greater than 0 then turn the value to the nearest integer and turn textbox green
            If ctl.Value > 0 Then
                If Int(ctl.Value) = 0 Then
                    ctl.Value = 1
                Else
                    ctl.Value = Int(ctl.Value)
                    ctl.BackColor = RGB(216, 245, 176)
                End If
            Else
                'if the value is less than zero, turn on the error and change textbox red
                ctl.BackColor = RGB(235, 33, 33)
                QuantError = True
            End If
        End If
    End If
    
    'find the price typebocx
    If Left(ctl.Name, 5) = "price" Then
            'Recognize that a value has changed and its numeric
            
            If ctl.Value <> 0 And IsNumeric(ctl.Value) Then
                'Proper input would have The decimal place at be the last 3 values. Take the length of the value inputted and subtract 2 to find where the decimal should be
                findecimal = Len(ctl.Value) - 2
                
                    'if the number of digits inputted is too small, change textbox colour red and turn on error
                    If findecimal <= 0 Then
                    ctl.BackColor = RGB(235, 33, 33)
                    Price98Error = True
                    Else
                        'Extract the decimal place and the numbers after it
                        afterdecimal = Mid(ctl.Value, findecimal, 3)
                        
                        'ensure extracted value looks like this, if not turn textbox red and turn on error
                        If afterdecimal <> ".98" Then
                            ctl.BackColor = RGB(235, 33, 33)
                            Price98Error = True
                        'if extracted value is .98 then turn textbox green
                        Else
                        ctl.BackColor = RGB(216, 245, 176)
                        End If
                    End If
            Else:
                'if textbox is empty or is not numeric, turn textbox red an turn on error
                ctl.BackColor = RGB(235, 33, 33)
                PriceError = True
            End If
        End If
Next
    
    
'if no province was selected then turn on province error
If optioncnt = 0 Then
    ProvinceError = True
End If

'if no province is selected then inform the user
If ProvinceError = True Then
    MsgBox "Please select a province!", vbCritical, "No Province Error"
End If

'if no name is entered then inform the user
If NameError = True Then
    MsgBox "Please enter an item name!", vbCritical, "No Name Error"
End If

'if no quantity is entered then inform the user
If QuantError = True Then
    MsgBox "Please enter an integer item quantity!", vbCritical, "No Name Error"
End If

'if price is entered but does not end in .98, then inform the user
If Price98Error = True Then
    MsgBox "Prices must end in 0.98!", vbCritical, "Price Error"
End If

'if no price is entered then inform the user
If PriceError = True Then
   MsgBox "Please ensure you are entering a price! The value must end in 0.98", vbCritical, "No Name Error"
End If

'ensures that all errors are turned off before entering transaction
If ProvinceError = False And NameError = False And QuantError = False And Price98Error = False And PriceError = False Then
    ready = True
End If

'perform calculations and move values
If ready = True Then
    'update subtotal value
    Subtotal = Subtotal + (price1 * quantity1)
    
    'remember last price of item purchased in case need to undo
    lastprice = price1 * quantity1
    
    'insert added name, quantity, unit and final price to final receipt
    Receipt.ListBox1.AddItem quantity1 & " " & named1 & " @ " & Format(price1, "Currency")
    Receipt.ListBox2.AddItem Format(quantity1 * price1, "Currency")
    
    'insert added name, quantity, unit and final price to temporary receipt
    Me.ListBox1.AddItem quantity1 & " " & named1 & " @ " & Format(price1, "Currency")
    
    'update running subtotal value
    runningsub.Value = Format(Subtotal, "Currency")
    
    'reset input boxes back to initialized values and revert back to white textboxes
    named1.Value = ""
    named1.BackColor = RGB(255, 255, 255)
    quantity1.Value = 0
    quantity1.BackColor = RGB(255, 255, 255)
    price1.Value = 0
    price1.BackColor = RGB(255, 255, 255)
    
    'identify proper taxes to use based on province inputted
    Select Case province
            Case Is = "AB"
                Receipt.GST.Caption = Format(0.05, "0.00%")
                Receipt.PST.Caption = Format(0, "0.00%")
                Receipt.HST.Caption = Format(0, "0.00%")
                totaltaxes = 0.05
             Case Is = "ON"
                Receipt.GST.Caption = Format(0, "0.00%")
                Receipt.PST.Caption = Format(0, "0.00%")
                Receipt.HST.Caption = Format(0.15, "0.00%")
                totaltaxes = 0.15
             Case Is = "SK"
                Receipt.GST.Caption = Format(0.05, "0.00%")
                Receipt.PST.Caption = Format(0.06, "0.00%")
                Receipt.HST.Caption = Format(0, "0.00%")
                totaltaxes = 0.11
             Case Is = "BC"
                Receipt.GST.Caption = Format(0.05, "0.00%")
                Receipt.PST.Caption = Format(0.07, "0.00%")
                Receipt.HST.Caption = Format(0, "0.00%")
                totaltaxes = 0.12
        End Select
    
    'determine if discount needed
    If Subtotal > 2000 Then
        DiscountVal = Subtotal * 0.15
    Else: DiscountVal = 0
    End If
    
End If

End Sub

Private Sub CompletePurchase_Click()

'set variables
Dim confirm As String
Dim ctl As Control

'ask user to confirm actions
confirm = MsgBox("Are you ready to finalize purchases?", vbYesNo, "Confirm Purchases")
    
    'if user does want to confirm transaction
    If confirm = vbYes Then
        
        'check if there have been no previously entered items or if at least one of the input boxes have been used. If so, allow users to add an item to their transaction this way
        If Receipt.ListBox1.ListCount = 0 Or named1.Value <> "" Or quantity1.Value <> 0 Or price1.Value <> 0 Then
            'enter the last transaction
            Call EnterTransaction_Click
                'if the transaction is error free then add it to the receipt and reset input values
                If ready = True Then
                    Receipt.Show
                    runningsub.Value = ""
                    Me.ListBox1.Clear
                    For Each ctl In Me.Controls 'for each control in the userform
                        If TypeName(ctl) = "OptionButton" Then 'look for the option buttons and ensure one is selected
                            ctl.Value = False
                        End If
                    Next
                End If
        Else
            'if the complete transaction button is just being used to complete the transaction (not add a new one), reset the input form and show receipt
            runningsub.Value = ""
            Me.ListBox1.Clear
            
            For Each ctl In Me.Controls 'for each control in the userform
                        If TypeName(ctl) = "OptionButton" Then 'look for the option buttons and ensure one is selected
                            ctl.Value = False
                        End If
            Next
            
            Receipt.Show
            Subtotal = 0
            province = ""
        End If
       
    End If
    
End Sub

Private Sub UndoLast_Click() 'undo the last entered transaction

'set variables
Dim ItemTarget1 As Long
Dim ItemTarget2 As Long
Dim ItemTarget3 As Long
Dim confirmundo As String

'find how long each listbox is (on the final reciept and the temp receipt)
ItemTarget1 = Receipt.ListBox1.ListCount
ItemTarget2 = Receipt.ListBox2.ListCount
ItemTarget3 = Me.ListBox1.ListCount

'confirm that user wants to undo the last entry
confirmundo = MsgBox("Are you sure you want to undo the last entry?", vbYesNo, "Confirm Purchases")

If confirmundo = vbYes Then
    'ensure there is at least one item on the final receipt
    If ItemTarget1 > 0 Then
        'remove the last item on each list box
        Receipt.ListBox1.RemoveItem ItemTarget1 - 1
        Receipt.ListBox2.RemoveItem ItemTarget2 - 1
        Me.ListBox1.RemoveItem ItemTarget3 - 1
        
        'update the subtotal
        Subtotal = Subtotal - lastprice
        
        'if the subtotal is now zero, reset the temporary subtotal value, otherwise, show the updated subtotal
        If Subtotal = 0 Then
            runningsub.Value = ""
        Else
            runningsub.Value = Format(Subtotal, "Currency")
        End If
        
        'reset input values
        price1.Value = 0
        quantity1.Value = 0
        named1.Value = ""
        
    'inform users if no transactions have been entered
    Else
        MsgBox "No transactions have been entered", vbCritical, "Undo Error"
    End If
End If

End Sub

Private Sub ClearEntry_Click() 'Remove transaction inputs that have not yet been entered

'if in the middle of a string of transactions, remove all input data but keep province
If Subtotal <> 0 Then

    For Each ctl In Me.Controls 'for each control in the userform & reset all colours to white
            If Left(ctl.Name, 4) = "name" Then
                ctl.Value = ""
                ctl.BackColor = RGB(255, 255, 255)
            ElseIf Left(ctl.Name, 5) = "quant" Then
                ctl.Value = 0
                ctl.BackColor = RGB(255, 255, 255)
            ElseIf Left(ctl.Name, 5) = "price" Then
                ctl.Value = 0
                ctl.BackColor = RGB(255, 255, 255)
            End If
    Next
    
Else

'if starting a string of transactions, remove all input data including province & reset all colours to white
    For Each ctl In Me.Controls 'for each control in the userform
            If TypeName(ctl) = "OptionButton" Then
                ctl.Value = False
            ElseIf Left(ctl.Name, 4) = "name" Then
                ctl.Value = ""
                ctl.BackColor = RGB(255, 255, 255)
            ElseIf Left(ctl.Name, 5) = "quant" Then
                ctl.Value = 0
                ctl.BackColor = RGB(255, 255, 255)
            ElseIf Left(ctl.Name, 5) = "price" Then
                ctl.Value = 0
                ctl.BackColor = RGB(255, 255, 255)
            End If
    Next

End If

End Sub

Private Sub RestartTransaction_Click() 'Remove all existing transactions and start again

    '
    Call ClearEntry_Click 'Call ClearEntry to reset most of the values
    
    
    'Rest province input incase not coveredby above step
    For Each ctl In Me.Controls 'for each control in the userform
            If TypeName(ctl) = "OptionButton" Then 'look for the option buttons and ensure one is selected
                ctl.Value = False
            End If
    Next
    
    'Reset values and clear listboxes
    Subtotal = 0
    province = ""
    runningsub.Value = ""
    Me.ListBox1.Clear
    Receipt.ListBox1.Clear
    Receipt.ListBox2.Clear

End Sub

Private Sub CancelTransaction_Click() 'Stop transaction process and close input form

'set variables
Dim canceltrans As String

'confirm user wants to close
canceltrans = MsgBox("Are you sure you want to cancel your transaction?", vbYesNo, "Cancel Transaction")
        'if yes, unload both input form and final receipt
        If canceltrans = vbYes Then
            Unload Me
            Unload Receipt
        End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'if input form is closed, unload input form and receipt
Unload Me
Unload Receipt

End Sub

Private Sub quantity1_AfterUpdate()
'prevent textbox from being blank, always reset to 0
If quantity1.Value = "" Then
    quantity1.Value = 0
End If
End Sub
Private Sub price1_AfterUpdate()
'prevent textbox from being blank, always reset to 0
If price1.Value = "" Then
    price1.Value = 0
End If
End Sub
Private Sub named1_Change()
'if textbox is changed, reset colour to white
named1.BackColor = RGB(255, 255, 255)
End Sub
Private Sub quantity1_Change()
'if textbox is changed, reset colour to white
quantity1.BackColor = RGB(255, 255, 255)
End Sub
Private Sub price1_Change()
'if textbox is changed, reset colour to white
price1.BackColor = RGB(255, 255, 255)
End Sub

