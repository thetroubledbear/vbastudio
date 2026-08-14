VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ASSERT_ERROR_BASE As Long = vbObjectError + 1000

Public Sub AreEqual(expected As Variant, actual As Variant, Optional message As String = "")
    If expected <> actual Then
        Fail "AreEqual failed: expected <" & expected & "> but got <" & actual & ">. " & message
    End If
End Sub

Public Sub IsTrue(condition As Boolean, Optional message As String = "")
    If Not condition Then Fail "IsTrue failed. " & message
End Sub

Public Sub IsFalse(condition As Boolean, Optional message As String = "")
    If condition Then Fail "IsFalse failed. " & message
End Sub

Public Sub Fail(Optional message As String = "")
    Err.Raise ASSERT_ERROR_BASE, "clsAssert", message
End Sub
