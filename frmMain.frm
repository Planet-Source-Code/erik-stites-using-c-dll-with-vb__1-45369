VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "C++ DLL in VB"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShift 
      Caption         =   "Shift"
      Height          =   375
      Left            =   495
      TabIndex        =   1
      Top             =   855
      Width           =   1320
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   495
      TabIndex        =   0
      Text            =   "12"
      Top             =   450
      Width           =   1320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Here is the declaration that calls the function in the dll
'Notice that the data type is Long. In Visual Basic, a long is 4 bytes, and an Integer is 2 bytes while in C++, both an int and a long are 4 bytes.
Private Declare Function BitShift Lib "shiftdll.dll" (ByVal Data As Long) As Long

'When the button is pressed, the value in the textbox will be divided by 2 and the result will be placed back into the textbox.
Private Sub cmdShift_Click()
    txtInput.Text = BitShift(Val(txtInput.Text))
End Sub
