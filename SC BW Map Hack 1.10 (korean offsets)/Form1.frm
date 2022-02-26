VERSION 5.00
Begin VB.Form frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
Caption         =   "StarCraft BroodWar Map Hack 1.10 - 2003"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3000
      Top             =   480
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.Label lblMapOn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
Caption         =   "Map On (del)"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblMapOff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
Caption         =   "Map Off (del)"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label lblState 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "State:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Dim Offset(1 To 13) As Long
Dim OldData(1 To 13) As String
Dim NewData(1 To 13) As String

Private Sub Form_Load()
If App.PrevInstance = True Then
    Unload Me
    End
End If

Call Timer_Timer

MsgBox "StarCraft & StarCraft Brood War Map Hack 1.10 - 2003" & vbCrLf & vbCrLf & _
"It work's only with StarCraft & StarCraft Brood War 1.10" & vbCrLf & _
"Press DELETE to turn map hack on/off" & vbCrLf & vbCrLf & _
"Coded by kostas86" & vbCrLf & "e-mail: kakavoulisk@gmail.com" & , vbInformation
End Sub

Private Sub Form_Activate()
LoadOffsets
LoadOldData
LoadNewData

Timer.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End
End Sub

Private Sub lblMapOn_Click()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
Dim x As Integer

hwnd = FindWindow(vbNullString, "Brood War")
If (hwnd = 0) Then hwnd = FindWindow(vbNullString, "StarCraft")

If (hwnd = 0) Then
    lblState.Caption = "State: Game not running."
    Exit Sub
End If

GetWindowThreadProcessId hwnd, pid

pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)

If (pHandle = 0) Then
    lblState.Caption = "State: Couldn't get a process handle."
    Exit Sub
End If

For x = 1 To 13
    WriteProcessMemory pHandle, Offset(x), NewData(x), Len(NewData(x)), 0&
Next x

CloseHandle pHandle
End Sub

Private Sub lblMapOff_Click()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
Dim x As Integer

hwnd = FindWindow(vbNullString, "Brood War")
If (hwnd = 0) Then hwnd = FindWindow(vbNullString, "StarCraft")

If (hwnd = 0) Then
    lblState.Caption = "State: Game not running."
    Exit Sub
End If

GetWindowThreadProcessId hwnd, pid

pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)

If (pHandle = 0) Then
    lblState.Caption = "State: Couldn't get a process handle."
    Exit Sub
End If

For x = 1 To 13
    WriteProcessMemory pHandle, Offset(x), OldData(x), Len(OldData(x)), 0&
Next x

CloseHandle pHandle
End Sub

Private Sub Timer_Timer()
Dim hwnd As Long
Dim pid As Long
Dim pHandle As Long
Dim data(1 To 13) As String * 37
Dim x As Integer
Dim dlen As Integer

hwnd = FindWindow(vbNullString, "Brood War")
If (hwnd = 0) Then hwnd = FindWindow(vbNullString, "StarCraft")

If (hwnd = 0) Then
    lblState.Caption = "State: Game not running."
    lblMapOn.Enabled = False
    lblMapOff.Enabled = False
    lblMapOn.ForeColor = &HC0C0C0
    lblMapOff.ForeColor = &HC0C0C0
    Exit Sub
End If

GetWindowThreadProcessId hwnd, pid

pHandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)

If (pHandle = 0) Then
    lblState.Caption = "State: Couldn't get a process handle."
    lblMapOn.Enabled = False
    lblMapOff.Enabled = False
    lblMapOn.ForeColor = &HC0C0C0
    lblMapOff.ForeColor = &HC0C0C0
    Exit Sub
End If

For x = 1 To 13
    ReadProcessMemory pHandle, Offset(x), data(x), 37, 0&
Next x

CloseHandle pHandle

For x = 1 To 13
    dlen = Len(OldData(x))
    If Mid(data(x), 1, dlen) <> OldData(x) Then GoTo 1
Next x

lblState.Caption = "State: Map Hack Deactivated."
lblMapOn.Enabled = True
lblMapOff.Enabled = False
lblMapOn.ForeColor = &H0&
lblMapOff.ForeColor = &HC0C0C0

If GetAsyncKeyState(vbKeyDelete) <> 0 Then
    Call lblMapOn_Click
    lblMapOn.Enabled = False
    lblMapOff.Enabled = True
    lblMapOn.ForeColor = &HC0C0C0
    lblMapOff.ForeColor = &H0&
    lblState.Caption = "State: Map Hack Activated."
End If

Exit Sub

1:

For x = 1 To 13
    dlen = Len(NewData(x))
    If Mid(data(x), 1, dlen) <> NewData(x) Then GoTo 2
Next x

lblState.Caption = "State: Map Hack Activated."
lblMapOn.Enabled = False
lblMapOff.Enabled = True
lblMapOn.ForeColor = &HC0C0C0
lblMapOff.ForeColor = &H0&

If GetAsyncKeyState(vbKeyDelete) <> 0 Then
    Call lblMapOff_Click
    lblMapOn.Enabled = True
    lblMapOff.Enabled = False
    lblMapOn.ForeColor = &H0&
    lblMapOff.ForeColor = &HC0C0C0
    lblState.Caption = "State: Map Hack Deactivated."
End If

Exit Sub

2:

lblState.Caption = "State: Error."
lblMapOn.Enabled = False
lblMapOff.Enabled = False
lblMapOn.ForeColor = &HC0C0C0
lblMapOff.ForeColor = &HC0C0C0
End Sub

Private Function Hex2ASCII(sText As String) As String
On Error Resume Next
Dim sBuff() As String, A As Long
sBuff() = Split(sText, Space$(1))
For A = 0 To UBound(sBuff)
Hex2ASCII = Hex2ASCII & Chr$("&h" & sBuff(A))
DoEvents
Next A
End Function

Private Function Hex2Dec(sText As String) As Long
On Error GoTo err
Dim H As String
H = sText
Dim Tmp$
Dim lo1 As Integer, lo2 As Integer
Dim hi1 As Long, hi2 As Long
Const Hx = "&H"
Const BigShift = 65536
Const LilShift = 256, Two = 2
Tmp = H
If UCase(Left$(H, 2)) = "&H" Then Tmp = Mid$(H, 3)
Tmp = Right$("0000000" & Tmp, 8)
If IsNumeric(Hx & Tmp) Then
lo1 = CInt(Hx & Right$(Tmp, Two))
hi1 = CLng(Hx & Mid$(Tmp, 5, Two))
lo2 = CInt(Hx & Mid$(Tmp, 3, Two))
hi2 = CLng(Hx & Left$(Tmp, Two))
Hex2Dec = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
End If
Exit Function
err:
End Function

Private Sub LoadOffsets()
Offset(1) = Hex2Dec("46D2F4")
Offset(2) = Hex2Dec("46D0A3")
Offset(3) = Hex2Dec("46E743")
Offset(4) = Hex2Dec("404C9B")
Offset(5) = Hex2Dec("4C4589")
Offset(6) = Hex2Dec("46C517")
Offset(7) = Hex2Dec("46C527")
Offset(8) = Hex2Dec("4105D2")
Offset(9) = Hex2Dec("480866")
Offset(10) = Hex2Dec("480898")
Offset(11) = Hex2Dec("4808AB")
Offset(12) = Hex2Dec("4808DC")
Offset(13) = Hex2Dec("480951")
End Sub

Private Sub LoadOldData()
OldData(1) = Hex2ASCII("33 D0 85 C9 89 15 F0 29 65 00 74 11 84 05 6C E3 65 00 75 11 A1 70 E3 65 00 85 C0 EB 06 85 05 40 30 65 00 74 32")
OldData(2) = Hex2ASCII("33 D0 85 C9 89 15 F0 29 65 00 74 11 84 05 6C E3 65 00 75 11 A1 70 E3 65 00 85 C0 EB 06 85 05 40 30 65 00 74 59")
OldData(3) = Hex2ASCII("88 4D FF")
OldData(4) = Hex2ASCII("0F 85 8B 00 00 00")
OldData(5) = Hex2ASCII("0F 84 F0 00 00 00")
OldData(6) = Hex2ASCII("74 08")
OldData(7) = Hex2ASCII("0F 84 08 FF FF FF")
OldData(8) = Hex2ASCII("84 C8")
OldData(9) = Hex2ASCII("90 90 90 90 90 90 90 90 90 90")
OldData(10) = Hex2ASCII("90 90 90 90 90 90 90")
OldData(11) = Hex2ASCII("90 90 90 90")
OldData(12) = Hex2ASCII("90 90 90 90")
OldData(13) = Hex2ASCII("33 C0")
End Sub

Private Sub LoadNewData()
NewData(1) = Hex2ASCII("85 C9 74 13 84 05 6C E3 65 00 75 19 A1 70 E3 65 00 85 C0 74 42 EB 0E 85 05 40 30 65 00 75 06 FF 0D 08 2F 65 00")
NewData(2) = Hex2ASCII("85 C9 74 13 84 05 6C E3 65 00 75 19 A1 70 E3 65 00 85 C0 74 69 EB 0E 85 05 40 30 65 00 75 06 FF 0D 08 2F 65 00")
NewData(3) = Hex2ASCII("88 45 FF")
NewData(4) = Hex2ASCII("0F 85 00 00 00 00")
NewData(5) = Hex2ASCII("0F 84 00 00 00 00")
NewData(6) = Hex2ASCII("EB 08")
NewData(7) = Hex2ASCII("E9 09 FF FF FF 90")
NewData(8) = Hex2ASCII("84 C9")
NewData(9) = Hex2ASCII("A0 40 30 65 00 8A 4F 0C EB 28")
NewData(10) = Hex2ASCII("84 C8 75 0F 5F 5E C3")
NewData(11) = Hex2ASCII("33 C0 EB 2D")
NewData(12) = Hex2ASCII("EB 75 EB 86")
NewData(13) = Hex2ASCII("EB 8B")
End Sub
