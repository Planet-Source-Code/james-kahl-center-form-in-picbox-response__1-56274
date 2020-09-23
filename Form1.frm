VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Test 2"
      Height          =   495
      Left            =   10380
      TabIndex        =   3
      Top             =   7980
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   3975
      Left            =   660
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   581
      TabIndex        =   2
      ToolTipText     =   "Drag me arround the form (before executing) and press 'Test 2'"
      Top             =   3060
      Width           =   8775
   End
   Begin VB.PictureBox Picture1 
      Height          =   1635
      Left            =   2940
      ScaleHeight     =   1575
      ScaleWidth      =   3435
      TabIndex        =   1
      ToolTipText     =   "Drag me arround the form (before executing) and press 'Test 1'"
      Top             =   660
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test 1"
      Height          =   495
      Left            =   10380
      TabIndex        =   0
      Top             =   7380
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THE OBJECT OF THIS LITTLE MODULE IS TO
'CENTER A CHILD FORM INTO ANOTHER CONTROL [PICTUREBOX]
'ON THE MAIN FORM.

'THE PURPOSE BEING TO HOLD EXECUTION OF CODE UNTIL CRITERIA
'IS ENTERED INTO THE CHILD FORM AND ACCEPTED.

'USE YOUR IMAGINATION WITH THIS ONE...

'REMEMBER!

'THE CHILD FORM(S) MUST BE 60 TWIPS LESS THAN THE CONTAINER [PICTUREBOX] IN HEIGHT & WIDTH
'TO ACCOMODATE FOR THE BORDER (I.E.  THE CHILD1 FORM IS 3435 x 1575, THE CONTAINER IS 3495 x 1635)

Private Sub Command1_Click()
    centerObject Me, Form2, Picture1, False
    Form2.Show vbModal, Me
End Sub

Private Sub Command2_Click()
    centerObject Me, Form3, Picture2
    Form3.Show vbModal, Me
End Sub
