VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Receipt 
   ClientHeight    =   8500.001
   ClientLeft      =   50
   ClientTop       =   180
   ClientWidth     =   8010
   OleObjectBlob   =   "Receipt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Receipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate() 'every time Receipt is opened do the following

'insert subtotal amount
receiptsub.Caption = Format(Subtotal, "Currency")

'insert discount amount
Discount.Caption = Format(DiscountVal, "Currency")

'calculate and insert total amount
Total.Caption = Format((Subtotal - DiscountVal) * (1 + totaltaxes), "Currency")

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'if receipt is closed,reset subtotal, province and unload receipt

    Subtotal = 0
    province = ""
    Unload Me
    
End Sub

Private Sub EndTransaction_Click()

'Unload receipt and item entry form

Unload Me
Unload ItemEntry2

End Sub
