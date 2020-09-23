VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   2265
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close Child 2"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1620
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
