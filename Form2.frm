VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   1230
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3075
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1230
   ScaleWidth      =   3075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close Child 1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
