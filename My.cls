VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "My"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Computer"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarComputer As Computer
Private Sub Class_Initialize()
   'create the mComputer object when the My class is created
   Set mvarComputer = New Computer
End Sub
Public Property Get Computer() As Computer
   Set Computer = mvarComputer
End Property
Public Property Set Computer(vData As Computer)
   Set mvarComputer = vData
End Property
Private Sub Class_Terminate()
   Set mvarComputer = Nothing
End Sub
