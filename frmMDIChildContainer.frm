VERSION 5.00
Begin VB.Form frmMDIChildContainer 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   8055
   Begin VB.Menu mnuTop1 
      Caption         =   "mnuTop1"
      Begin VB.Menu mnuTop1_Menus 
         Caption         =   "mnuTop1_Menus"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTop2 
      Caption         =   "mnuTop2"
   End
End
Attribute VB_Name = "frmMDIChildContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' www.cyberbiz.com - mje@cyberbiz.com - PH: 773-338-3755
'
'LICENSE:
'--------
' You have a royalty-free right to use, modify, reproduce and distribute all
' MDIChildMaker files.
' By using any portion of MDIChildMaker you agree that CyberBiz, as well as all
' developers and distributors of MDIChildMaker have no warranty, obligations
' or liability for any MDIChildMaker related file (code, executable, etc.).
'-------------------------------------------------------------------------------

'Check the README1st.txt file in related documents for information on this application.

Private m_oMDIChildMaker As MDIChildMaker

Public Property Set MDIChildMaker(oObject As MDIChildMaker)
    Set m_oMDIChildMaker = oObject
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'BUGFIX This is a quick fix - we should really subclass this form
    'and test for mousedown in the non-client area, and make sure the child
    'get's activated...
    m_oMDIChildMaker.ChildForm.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    m_oMDIChildMaker.ChildForm.Menu_Build
End Sub


Private Sub Form_Resize()
    'm_oMDIChildMaker might not be set yet
    If Not m_oMDIChildMaker Is Nothing Then
        m_oMDIChildMaker.Resize
    End If
End Sub

Private Sub mnuTop1_Click()
    On Error Resume Next
    m_oMDIChildMaker.ChildForm.Menu_Selection mnuTop1
End Sub

Private Sub mnuTop1_Menus_Click(Index As Integer)
    m_oMDIChildMaker.ChildForm.Menu_Selection mnuTop1_Menus(Index)
End Sub

Private Sub mnuTop2_Click()
    On Error Resume Next
    m_oMDIChildMaker.ChildForm.Menu_Selection mnuTop2
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_oMDIChildMaker = Nothing
End Sub

