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
      Left            =   1440
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   960
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
Option Explicit

Private m_oMDIContainerForm As Form

Public Property Set MDIContainerForm(oForm As Form)
    Set m_oMDIContainerForm = oForm
End Property

Public Sub Menu_Build()
    m_oMDIContainerForm.mnuTop1.Visible = False
    m_oMDIContainerForm.mnuTop2.Visible = False
End Sub


