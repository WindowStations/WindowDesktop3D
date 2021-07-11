VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Screen"
Option Explicit
Private mvarScreen As Screen
Private Sub Class_Initialize()
   'create the mScreen object when the Computer class is created
   Set mvarScreen = New Screen
End Sub
Public Property Get Screen() As Screen
   Set Screen = mvarScreen
End Property
Public Property Set Screen(vData As Screen)
   Set mvarScreen = vData
End Property
Private Sub Class_Terminate()
   Set mvarScreen = Nothing
End Sub
