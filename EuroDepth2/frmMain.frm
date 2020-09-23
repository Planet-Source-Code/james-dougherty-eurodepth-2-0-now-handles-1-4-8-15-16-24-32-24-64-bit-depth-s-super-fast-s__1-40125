VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EuroDepth Demo"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   610
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDepth 
      Height          =   315
      Left            =   1973
      TabIndex        =   2
      Top             =   4680
      Width           =   2415
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4110
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   1
      Top             =   5040
      Width           =   6360
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4110
      Left            =   0
      Picture         =   "frmMain.frx":511F
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   0
      Width           =   6360
   End
   Begin VB.Label lblCompile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Please Compile)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4440
      TabIndex        =   4
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time - 0.0000 Milliseconds"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   428
      TabIndex        =   3
      Top             =   4080
      Width           =   5505
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'If you had seen version 1.0 you can see this is a major inprovement.
'Better coded and optimized. All the code has been redone to avoid unnecessary repetition.


'The EuroDepth object
Private EuroDepth As New EuroDepth

'|----------------------------------------------------------------------------------------------------------------------------------------------------
'|The Timing Code Is By:
'|KPD-Team 2001
'|URL: http://www.allapi.net/
'|E-Mail: KPDTeam@Allapi.net
'|----------------------------------------------------------------------------------------------------------------------------------------------------
  Private Type LARGE_INTEGER
   LowPart As Long
   HighPart As Long
  End Type

  Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
  Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'|----------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub cmbDepth_Click()
 Dim liFrequency As LARGE_INTEGER
 Dim liStart As LARGE_INTEGER
 Dim liStop As LARGE_INTEGER
 Dim cuFrequency As Currency
 Dim cuStart As Currency
 Dim cuStop As Currency
 Dim sResult As String
 
 'Start the timing process
 QueryPerformanceFrequency liFrequency
 cuFrequency = LargeIntToCurrency(liFrequency)
 QueryPerformanceCounter liStart
 
 'Change the color depth depending on the index of the combo box
 'NOTE: A destination picture is optional but I suggest you use one.
 If cmbDepth.ListIndex = 0 Then
  EuroDepth.SetDepth_01Bit picSource, picDestination
 ElseIf cmbDepth.ListIndex = 1 Then
  EuroDepth.SetDepth_04Bit picSource, picDestination
 ElseIf cmbDepth.ListIndex = 2 Then
  EuroDepth.SetDepth_08Bit picSource, picDestination
 ElseIf cmbDepth.ListIndex = 3 Then
  EuroDepth.SetDepth_15Bit picSource, picDestination
 ElseIf cmbDepth.ListIndex = 4 Then
  EuroDepth.SetDepth_16Bit picSource, picDestination
 ElseIf cmbDepth.ListIndex = 5 Then
  EuroDepth.SetDepth_24Bit_32K picSource, picDestination
 ElseIf cmbDepth.ListIndex = 6 Then
  EuroDepth.SetDepth_24Bit_64k picSource, picDestination
 End If
 
 'End the timing process
 QueryPerformanceCounter liStop
 cuStart = LargeIntToCurrency(liStart)
 cuStop = LargeIntToCurrency(liStop)
 
 'Compute the result
 sResult = Format$(CStr((cuStop - cuStart) / cuFrequency), "0.0000")
 
 'Output the resule
 lblTime.Caption = "Time - " & sResult & " Milliseconds"
 
End Sub

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
 CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
 LargeIntToCurrency = LargeIntToCurrency * 10000
End Function

Private Sub Form_Load()
 
 'Fill in the combo box
 cmbDepth.AddItem "2 Colors (1 Bit)", 0
 cmbDepth.AddItem "16 Colors (4 Bit)", 1
 cmbDepth.AddItem "256 Colors (8 Bit)", 2
 cmbDepth.AddItem "(15 Bit)", 3
 cmbDepth.AddItem "(16 Bit)", 4
 cmbDepth.AddItem "32,000 Colors (24 Bit)", 5
 cmbDepth.AddItem "64,000 Colors (24 Bit)", 6
 'Set the default index to 5
 cmbDepth.ListIndex = 5
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Clear EuroDepth out of memeory
 Set EuroDepth = Nothing
End Sub

