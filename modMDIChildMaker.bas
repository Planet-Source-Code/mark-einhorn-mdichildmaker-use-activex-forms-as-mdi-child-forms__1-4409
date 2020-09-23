Attribute VB_Name = "modMDIChildMaker"
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

Public Sub CreateMDIChild(frm As Form)
    Dim oMDIChildMaker As MDIChildMaker
    Dim frmNewContainerForm As Form
    Dim frmNewChildForm As Form
    
    '------------------------------------------
    'We MUST instantiate the container form as new to reuse this form as the container
    Set frmNewContainerForm = New frmMDIChildContainer
    'NOTE: here we're referencing a newly instantiated form
    Set frmNewChildForm = frm
    Set frmNewChildForm.MDIContainerForm = frmNewContainerForm

    Set oMDIChildMaker = New MDIChildMaker
    Set oMDIChildMaker.ParentForm = frmNewContainerForm
    Set oMDIChildMaker.ChildForm = frmNewChildForm
    
    Set frmNewContainerForm.MDIChildMaker = oMDIChildMaker
    
    oMDIChildMaker.MakeMDIChild
    '------------------------------------------
End Sub

