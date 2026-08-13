VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDefaultMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mValue As Long

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
    Value = mValue
End Property

Public Property Let Value(ByVal RHS As Long)
    mValue = RHS
End Property
