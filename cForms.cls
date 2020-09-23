VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get CreateForm1() As Object
    Set CreateForm1 = New Form1
End Property

Public Property Get CreateForm2() As Object
    Set CreateForm2 = New Form2
End Property

Private Sub Class_Initialize()
    Unload Form1
    Unload Form2
End Sub
