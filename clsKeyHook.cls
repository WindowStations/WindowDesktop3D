VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event OnKeyWDown()
Public Event OnKeyWUp()
Public Event OnKeyADown()
Public Event OnKeyAUp()
Public Event OnKeySDown()
Public Event OnKeySUp()
Public Event OnKeyDDown()
Public Event OnKeyDUp()
Public Event OnKeyEscDown()
Public Event OnKeyEscUp()

Public Sub KeyWDown()
    RaiseEvent OnKeyWDown
End Sub
Public Sub KeyWUp()
    RaiseEvent OnKeyWUp
End Sub

Public Sub KeyADown()
    RaiseEvent OnKeyADown
End Sub
Public Sub KeyAUp()
    RaiseEvent OnKeyAUp
End Sub

Public Sub KeySDown()
    RaiseEvent OnKeySDown
End Sub
Public Sub KeySUp()
    RaiseEvent OnKeySUp
End Sub

Public Sub KeyDDown()
    RaiseEvent OnKeyDDown
End Sub
Public Sub KeyDUp()
    RaiseEvent OnKeyDUp
End Sub

Public Sub KeyEscDown()
    RaiseEvent OnKeyEscDown
End Sub
Public Sub KeyEscUp()
    RaiseEvent OnKeyEscUp
End Sub
