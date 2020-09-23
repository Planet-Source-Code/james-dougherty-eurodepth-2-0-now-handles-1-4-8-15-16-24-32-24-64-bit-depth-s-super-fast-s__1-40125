VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EuroDepth"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶               © 2002 Ariel Productions              ¶¶|'
  '|¶¶                 All Rights Reserved                 ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶             Programmer - James Dougherty            ¶¶|'
  '|¶¶             Source - EuroDepth.cls                  ¶¶|'
  '|¶¶             Object - N/A                            ¶¶|'
  '|¶¶             Version - 2.0                           ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'
  '|¶¶                                                     ¶¶|'
  '|¶¶ EuroDepth is free software; you can redistribute    ¶¶|'
  '|¶¶ it and/or modify it under the terms of the GNU      ¶¶|'
  '|¶¶ General Public License as published by the Free     ¶¶|'
  '|¶¶ Software Foundation; either version 2 of the        ¶¶|'
  '|¶¶ License, or (at your option) any later version.     ¶¶|'
  '|¶¶                                                     ¶¶|'
  '|¶¶ This program is distributed in the hope that it     ¶¶|'
  '|¶¶ will be useful, but WITHOUT ANY WARRANTY; without   ¶¶|'
  '|¶¶ even the implied warranty of MERCHANTABILITY or     ¶¶|'
  '|¶¶ FITNESS FOR A PARTICULAR PURPOSE.  See the GNU      ¶¶|'
  '|¶¶ General Public License for more details.            ¶¶|'
  '|¶¶                                                     ¶¶|'
  '|¶¶ You should have received a copy of the GNU General  ¶¶|'
  '|¶¶ Public License along with Foobar; if not, write to  ¶¶|'
  '|¶¶ the Free Software Foundation, Inc., 59 Temple       ¶¶|'
  '|¶¶ Place, Suite 330, Boston, MA 02111-1307  USA        ¶¶|'
  '|¶¶                                                     ¶¶|'
  '|¶¶¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¶¶|'
  '|¶¶        http://www.gnu.org/copyleft/gpl.html         ¶¶|'
  '|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|'

Option Explicit

Public Enum ED_BitsPerPixelDepth
 Colors_1Bit = 1
 Colors_4Bit = 4
 Colors_8Bit = 8
 Colors_15Bit = 15
 Colors_16Bit = 16
 Colors_24Bit_32k = 24
 Colors_24Bit_64k = 32
End Enum

Private Type RGB32
 Blue As Byte
 Green As Byte
 Red As Byte
 Reserved As Byte
End Type

Private Type BITMAPINFOHEADER
 biSize As Long
 biWidth As Long
 biHeight As Long
 biPlanes As Integer
 biBitCount As Integer
 biCompression As Long
 biSizeImage As Long
 biXPelsPerMeter As Long
 biYPelsPerMeter As Long
 biClrUsed As Long
 biClrImportant As Long
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
 bmiColors(0 To 255) As RGB32
End Type

Private m_i As Long
Private m_iX As Long
Private m_iY As Long
Private m_iBPR As Long
Private m_iWidth As Long
Private m_iHeight As Long
Private m_Pixels() As RGB32
Private m_btBytes() As Byte
Private m_btCIndex() As Byte
Private m_CurrentBPP As Long
Private m_tBitmapInfo As BITMAPINFO

Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'||'
'| FUNCTION - PrepareVariables()                                                   |'
'|                                                                                 |'
'| NOTES:                                                                          |'
'| This function prepares the bitmap to be worked with. It creates bitmap in       |'
'| memory and fills in the pixel arrays.                                           |'
'||'

Private Sub PrepareVariables(ByVal EBPP As ED_BitsPerPixelDepth, ByRef PictureBox As Object, Optional ByRef DestinationPictureBox As Object)
 Dim iBmpHandle As Long
 Dim iPallete As Long
 Dim iBitmap As Long
 Dim iBPP As Long
 Dim iDC As Long
  
 ReDim m_Pixels(0, 0)
 ReDim m_btBytes(0, 0)
 ReDim m_btCIndex(0, 0)
 Erase m_Pixels
 Erase m_btBytes
 Erase m_btCIndex
 
 iBPP = EBPP
 m_CurrentBPP = iBPP
 If (iBPP <> 1) And (iBPP <> 4) And (iBPP <> 8) And (iBPP <> 15) And (iBPP <> 16) And (iBPP <> 24) And (iBPP <> 32) Then Exit Sub
 m_iWidth = PictureBox.ScaleX(PictureBox.Image.Width, vbHimetric, vbPixels)
 m_iHeight = PictureBox.ScaleY(PictureBox.Image.Height, vbHimetric, vbPixels)
 
 iDC = CreateCompatibleDC(PictureBox.hdc)
 If iDC = 0 Then Exit Sub
 
 iBmpHandle = CreateCompatibleBitmap(PictureBox.hdc, m_iWidth, m_iHeight)
 If iBmpHandle = 0 Then Exit Sub
 
 iBitmap = SelectObject(iDC, iBmpHandle)
 PictureBox.AutoRedraw = True

 If PictureBox.Image.hPal <> 0 Then
  iPallete = SelectPalette(iDC, PictureBox.Image.hPal, False)
  If iPallete = 0 Then Exit Sub
  RealizePalette iDC
 End If

 If BitBlt(iDC, 0, 0, m_iWidth, m_iHeight, PictureBox.hdc, 0, 0, &HCC0020) = 0 Then Exit Sub
 iBmpHandle = SelectObject(iDC, iBitmap)
 If iBmpHandle = 0 Then Exit Sub

 With m_tBitmapInfo.bmiHeader
  .biPlanes = 1
  .biBitCount = iBPP
  .biWidth = m_iWidth
  .biCompression = 0&
  .biHeight = m_iHeight
  .biSize = Len(m_tBitmapInfo.bmiHeader)
 End With

 m_iBPR = ((m_iWidth * iBPP + 31) \ 32) * 4
 ReDim m_btBytes(0 To m_iBPR - 1, 0 To m_iHeight - 1)
 ReDim m_Pixels(0 To m_iWidth - 1, 0 To m_iHeight - 1)
 If GetDIBits(iDC, iBmpHandle, 0, m_iHeight, m_btBytes(0, 0), m_tBitmapInfo, 0) = 0 Then Exit Sub
 DeleteObject iBmpHandle
 DeleteObject iDC

End Sub

'||'
'| FUNCTION - SetDepth_01Bit()                                                     |'
'||'

Public Sub SetDepth_01Bit(ByVal PictureBox As Object, Optional ByVal DestinationPictureBox As Object)
 Dim iSV As Long
 
 PrepareVariables Colors_1Bit, PictureBox, DestinationPictureBox
 ReDim m_btCIndex(0 To (8 * m_iBPR) - 1, 0 To m_iHeight - 1)
 For m_iY = 0 To m_iHeight - 1
  For m_iX = 0 To m_iBPR - 1
   iSV = 128
   For m_i = 0 To 7
    If m_btBytes(m_iX, m_iY) And iSV Then
     m_btCIndex(8 * m_iX + m_i, m_iHeight - 1 - m_iY) = 1
    Else
     m_btCIndex(8 * m_iX + m_i, m_iHeight - 1 - m_iY) = 0
    End If
    iSV = iSV \ 2
   Next
  Next
 Next
 Finalize PictureBox, DestinationPictureBox
 
End Sub

'||'
'| FUNCTION - SetDepth_04Bit()                                                     |'
'||'

Public Sub SetDepth_04Bit(ByVal PictureBox As Object, Optional ByVal DestinationPictureBox As Object)
 
 PrepareVariables Colors_4Bit, PictureBox, DestinationPictureBox
 ReDim m_btCIndex(0 To (2 * m_iBPR) - 1, 0 To m_iHeight - 1)
 For m_iY = 0 To m_iHeight - 1
  For m_iX = 0 To m_iBPR - 1
   m_btCIndex(2 * m_iX, m_iHeight - 1 - m_iY) = m_btBytes(m_iX, m_iY) \ 16
   m_btCIndex(2 * m_iX + 1, m_iHeight - 1 - m_iY) = m_btBytes(m_iX, m_iY) Mod 16
  Next
 Next
 Finalize PictureBox, DestinationPictureBox
 
End Sub

'||'
'| FUNCTION - SetDepth_08Bit()                                                     |'
'||'

Public Sub SetDepth_08Bit(ByVal PictureBox As Object, Optional ByVal DestinationPictureBox As Object)
 
 PrepareVariables Colors_8Bit, PictureBox, DestinationPictureBox
 ReDim m_btCIndex(0 To m_iWidth - 1, 0 To m_iHeight - 1)
 For m_iY = 0 To m_iHeight - 1
  For m_iX = 0 To m_iWidth - 1
   m_btCIndex(m_iX, m_iHeight - 1 - m_iY) = m_btBytes(m_iX, m_iY)
  Next
 Next
 Finalize PictureBox, DestinationPictureBox
 
End Sub

'||'
'| FUNCTION - SetDepth_15Bit()                                                     |'
'||'

Public Sub SetDepth_15Bit(ByVal PictureBox As Object, Optional ByVal DestinationPictureBox As Object)
 Dim iTBit As Long
 
 PrepareVariables Colors_16Bit, PictureBox, DestinationPictureBox
 For m_iY = 0 To m_iHeight - 1
  For m_iX = 0 To m_iWidth - 1
   With m_Pixels(m_iX, m_iHeight - 1 - m_iY)
    iTBit = m_btBytes(m_iX * 2, m_iY) + m_btBytes(m_iX * 2 + 1, m_iY) * 256&
    .Blue = iTBit Mod 32
    iTBit = iTBit / 32
    .Green = iTBit Mod 32
    iTBit = iTBit / 32
    .Red = iTBit Mod 32
   End With
  Next
 Next
 Finalize PictureBox, DestinationPictureBox
 
End Sub

'||'
'| FUNCTION - SetDepth_16Bit()                                                     |'
'||'

Public Sub SetDepth_16Bit(ByVal PictureBox As Object, Optional ByVal DestinationPictureBox As Object)
 Dim iTBit As Long
 
 PrepareVariables Colors_16Bit, PictureBox, DestinationPictureBox
 For m_iY = 0 To m_iHeight - 1
  For m_iX = 0 To m_iWidth - 1
   With m_Pixels(m_iX, m_iHeight - 1 - m_iY)
    If .Red > &H1F Then .Red = &H1F
    If .Green > &H3F Then .Green = &H3F
    If .Blue > &H1F Then .Blue = &H1F
    iTBit = .Blue + 32 * (.Green + CLng(.Red) * 64)
    .Red = m_btBytes(m_iX * 1 + 0, m_iY) = (iTBit Mod 256) And &HFF
    .Green = m_btBytes(m_iX * 1 + 1, m_iY) = (iTBit \ 256) And &HFF
    .Blue = m_btBytes(m_iX * 1 + 2, m_iY) = (iTBit \ 256) And &HFF
   End With
  Next
 Next
 Finalize PictureBox, DestinationPictureBox
 
End Sub

'||'
'| FUNCTION - SetDepth_24Bit_32K()                                                 |'
'||'

Public Sub SetDepth_24Bit_32K(ByVal PictureBox As Object, Optional ByVal DestinationPictureBox As Object)
  
 PrepareVariables Colors_24Bit_32k, PictureBox, DestinationPictureBox
 For m_iY = 0 To m_iHeight - 1
  For m_iX = 0 To m_iWidth - 1
   m_Pixels(m_iX, m_iHeight - 1 - m_iY).Blue = m_btBytes(m_iX * 3 + 0, m_iY)
   m_Pixels(m_iX, m_iHeight - 1 - m_iY).Green = m_btBytes(m_iX * 3 + 1, m_iY)
   m_Pixels(m_iX, m_iHeight - 1 - m_iY).Red = m_btBytes(m_iX * 3 + 2, m_iY)
  Next
 Next
 Finalize PictureBox, DestinationPictureBox
  
End Sub

'||'
'| FUNCTION - SetDepth_24Bit_64k()                                                 |'
'||'

Public Sub SetDepth_24Bit_64k(ByVal PictureBox As Object, Optional ByVal DestinationPictureBox As Object)
 
 PrepareVariables Colors_24Bit_64k, PictureBox, DestinationPictureBox
 For m_iY = 0 To m_iHeight - 1
  For m_iX = 0 To m_iWidth - 1
   m_Pixels(m_iX, m_iHeight - 1 - m_iY).Blue = m_btBytes(m_iX * 3 + 0, m_iY)
   m_Pixels(m_iX, m_iHeight - 1 - m_iY).Green = m_btBytes(m_iX * 3 + 1, m_iY)
   m_Pixels(m_iX, m_iHeight - 1 - m_iY).Red = m_btBytes(m_iX * 3 + 2, m_iY)
  Next
 Next
 Finalize PictureBox, DestinationPictureBox
 
End Sub

'||'
'| FUNCTION - Finalize()                                                           |'
'|                                                                                 |'
'| NOTES:                                                                          |'
'| After the image is created and transformed in memory, the image gets blitted    |'
'| to the destination picture box. Then all the arrays get cleared and erased.     |'
'||'

Private Sub Finalize(ByRef PictureBox As Object, Optional ByRef DestinationPictureBox As Object)
 
 'Blit the image
 If Not DestinationPictureBox Is Nothing Then
  If StretchDIBits(DestinationPictureBox.hdc, 0, 0, m_iWidth, m_iHeight, 0, 0, m_iWidth, m_iHeight, m_btBytes(0, 0), m_tBitmapInfo, 0, &HCC0020) = &HFFFF Then Exit Sub
  DestinationPictureBox.Picture = DestinationPictureBox.Image
 Else
  If StretchDIBits(PictureBox.hdc, 0, 0, m_iWidth, m_iHeight, 0, 0, m_iWidth, m_iHeight, m_btBytes(0, 0), m_tBitmapInfo, 0, &HCC0020) = &HFFFF Then Exit Sub
  PictureBox.Picture = PictureBox.Image
 End If
 
End Sub

'||'
'| FUNCTION - Class_Terminate                                                      |'
'||'

Private Sub Class_Terminate()
 'Clear all memory
 ReDim m_Pixels(0, 0)
 ReDim m_btBytes(0, 0)
 ReDim m_btCIndex(0, 0)
 Erase m_Pixels
 Erase m_btBytes
 Erase m_btCIndex
End Sub
