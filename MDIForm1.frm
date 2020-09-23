VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MDIForm1"
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

Option Explicit


Private Sub MDIForm_Load()
    Dim oForms As prjMDIForms.cForms
    Set oForms = New prjMDIForms.cForms
    CreateMDIChild oForms.CreateForm1()
    CreateMDIChild oForms.CreateForm2()
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload Me
End Sub


