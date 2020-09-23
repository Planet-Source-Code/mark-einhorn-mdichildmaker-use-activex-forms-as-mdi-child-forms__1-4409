VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form2"
   ClientHeight    =   3825
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   ScaleHeight     =   3825
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   2835
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oMDIContainerForm As Form

Public Property Set MDIContainerForm(oForm As Form)
    Set m_oMDIContainerForm = oForm
End Property

Public Sub ResizeCtls()
    If m_oMDIContainerForm Is Nothing Then
        Exit Sub
    End If

    'Put resize code here
    '------------------------------------------------
    Me.Combo1.Left = 0
    Me.Combo1.Top = 0
    Me.Combo1.Width = m_oMDIContainerForm.ScaleWidth

    Me.Picture1.Left = 0
    Me.Picture1.Height = 500
    Me.Picture1.Top = m_oMDIContainerForm.ScaleHeight - Form1.Picture1.Height
    Me.Picture1.Width = m_oMDIContainerForm.ScaleWidth
    '------------------------------------------------
End Sub

Public Sub Menu_Build()
    m_oMDIContainerForm.mnuTop1.Caption = "&File"
    m_oMDIContainerForm.mnuTop2.Caption = "&Edit"
    m_oMDIContainerForm.mnuTop1_Menus(0).Caption = "&Open"
    Load m_oMDIContainerForm.mnuTop1_Menus(1)
    m_oMDIContainerForm.mnuTop1_Menus(1).Caption = "&Save"
End Sub

Public Sub Menu_Selection(mnu As Menu)
    Select Case mnu.Name
        Case "mnuTop1_Menus"
            Select Case mnu.Index
                Case 0
                    MsgBox "Open Menu!"
                Case 1
                    MsgBox "Save Menu!"
            End Select
        Case "mnuTop2"
            MsgBox "Edit Menu!"
    End Select
End Sub

Private Sub Command1_Click()
    MsgBox "test"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oMDIContainerForm = Nothing
End Sub
