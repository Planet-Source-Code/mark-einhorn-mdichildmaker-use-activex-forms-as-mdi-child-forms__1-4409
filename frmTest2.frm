VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_oMDIContainerForm As Form

Public Property Set MDIContainerForm(oForm As Form)
    Set m_oMDIContainerForm = oForm
End Property


Private Sub Form_Activate()
    On Error Resume Next
    'TODO - Subclass children forms within the cMDIChildMaker to encapsulate this
    m_oMDIContainerForm.SetFocus
End Sub

