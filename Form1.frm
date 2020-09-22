VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 
   Caption         =   "Event Viewer"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   OleObjectBlob   =   "Form1.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub


Private Sub Form1_Load()

End Sub


Private Sub Command2_Click()

End Sub

Private Sub ComboBox1_Change()
List1.AddItem "Typing/Change"
End Sub

Private Sub ComboBox1_Click()
List1.AddItem "Click"
End Sub

Private Sub ComboBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
List1.AddItem "DoubleClick"
End Sub

Private Sub ComboBox1_DropButtonClick()
List1.AddItem "DropClick"
End Sub

Private Sub ComboBox1_Enter()
List1.AddItem "Enter"
End Sub

Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
List1.AddItem "Keydown"
End Sub

Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
List1.AddItem "KeyPress"
End Sub

Private Sub ComboBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
List1.AddItem "KeyUp"
End Sub

Private Sub ComboBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
List1.AddItem "Mousedown"
End Sub

Private Sub ComboBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
List1.AddItem "MouseMove"
End Sub

Private Sub ComboBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
List1.AddItem "MouseUp"
End Sub

Private Sub CommandButton1_Click()
Form1.Width = "184.5"
End Sub

Private Sub CommandButton2_Click()
If Combobox1.Text = "" Then
MsgBox "nothing in box", vbExclamation, "EventView"
Else
Combobox1.RemoveItem text2.Text
End If
End Sub

Private Sub CommandButton3_Click()
Combobox1.AddItem Text1.Text
End Sub

Private Sub CommandButton4_Click()
Combobox1.Clear
Combobox1.Text = "EventView"
End Sub

Private Sub CommandButton5_Click()
End
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub
