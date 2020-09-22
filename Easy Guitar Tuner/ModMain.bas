Attribute VB_Name = "ModMain"
Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_USEREX_CLASSES = &H200

Public Sub Main()
On Error Resume Next

Dim objRet As tagInitCommonControlsEx

With objRet
    .lngSize = LenB(objRet)
    .lngICC = ICC_USEREX_CLASSES
End With

InitCommonControlsEx objRet
frmMain.Show
End Sub

Public Sub HCenter(ByRef CenterObject As Object, ByRef CenterOn As Object)
On Error Resume Next

CenterObject.Left = (CenterOn.Width * 0.5) - (CenterObject.Width * 0.5)
End Sub

Public Sub VCenter(ByRef CenterObject As Object, ByRef CenterOn As Object)
On Error Resume Next

CenterObject.Top = (CenterOn.Height * 0.5) - (CenterObject.Height * 0.5)
End Sub
