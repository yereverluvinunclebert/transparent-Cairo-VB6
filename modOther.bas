Attribute VB_Name = "modOther"
Option Explicit

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
      ByVal hObject As Long, _
      ByVal nCount As Long, _
      lpObject As Any _
   ) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" ( _
      ByVal lpDriverName As String, _
      lpDeviceName As Any, _
      lpOutput As Any, _
      lpInitData As Any _
    ) As Long
    
Private Type tBitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function BitBlt Lib "gdi32" ( _
       ByVal hDestDC As Long, _
       ByVal x As Long, ByVal y As Long, _
       ByVal nWidth As Long, ByVal nHeight As Long, _
       ByVal hSrcDC As Long, _
       ByVal xSrc As Long, ByVal ySrc As Long, _
       ByVal dwRop As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : HBitmapFromPicture
' Author    : beededea
' Date      : 02/04/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function HBitmapFromPicture(picThis As IPicture) As Long

   ' Create a copy of the bitmap:
   Dim lhDC As Long
   Dim lhDCCopy As Long
   Dim lhBmpCopy As Long
   Dim lhBmpCopyOld As Long
   Dim lhBmpOld As Long
   Dim lhDCC As Long
   Dim tBM As tBitmap

    On Error GoTo HBitmapFromPicture_Error

   GetObjectAPI picThis.Handle, Len(tBM), tBM
   lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lhDC = CreateCompatibleDC(lhDCC)
   lhBmpOld = SelectObject(lhDC, picThis.Handle)

   lhDCCopy = CreateCompatibleDC(lhDCC)
   lhBmpCopy = CreateCompatibleBitmap(lhDCC, tBM.bmWidth, tBM.bmHeight)
   lhBmpCopyOld = SelectObject(lhDCCopy, lhBmpCopy)

   BitBlt lhDCCopy, 0, 0, tBM.bmWidth, tBM.bmHeight, lhDC, 0, 0, vbSrcCopy

   If Not (lhDCC = 0) Then
      DeleteDC lhDCC
   End If
   If Not (lhBmpOld = 0) Then
      SelectObject lhDC, lhBmpOld
   End If
   If Not (lhDC = 0) Then
      DeleteDC lhDC
   End If
   If Not (lhBmpCopyOld = 0) Then
      SelectObject lhDCCopy, lhBmpCopyOld
   End If
   If Not (lhDCCopy = 0) Then
      DeleteDC lhDCCopy
   End If

   HBitmapFromPicture = lhBmpCopy

    On Error GoTo 0
    Exit Function

HBitmapFromPicture_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HBitmapFromPicture of Module mGDIPImageList"

End Function

