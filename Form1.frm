VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Inet checker"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5340
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C000&
      Caption         =   "Clear log"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "Hide"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "Stop log"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Start log"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00800000&
      Height          =   2790
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   0
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet log times"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu M 
      Caption         =   "m"
      Visible         =   0   'False
      Begin VB.Menu Ex 
         Caption         =   "Show Inet checker"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Private Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, _
  ByVal wParam As Long, lParam As Any) As Long
Const LB_SETTABSTOPS = &H192


Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "Kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "Kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "Kernel32" (ByVal hPass As Long)
 
'declarations
Private Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type

'declare constants (still in declarations)
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input for
'the icon in the taskbar

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203 'Double-click
Private Const WM_LBUTTONDOWN = &H201 'Button down
Private Const WM_LBUTTONUP = &H202 'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206 'Double-click
Private Const WM_RBUTTONDOWN = &H204 'Button down
Private Const WM_RBUTTONUP = &H205 'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA


Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type

Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure) for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX structures is passed to GDI
'along with a list of array indexes that describe separate triangles. GDI performs linear interpolation between triangle vertices
'and fills the interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Function LongToUShort(Unsigned As Long) As Integer
    'A small function to convert from long to unsigned short
    LongToUShort = CInt(Unsigned - &H10000)
End Function



Private Sub Command1_Click()
List1.AddItem "Started checking" & vbTab & Date & "  " & Time

Open "c:\internet log.txt" For Append As #1
      Print #1, "Internet checker started  " & Date & "  " & Time & vbCrLf
      Close #1

Timer1.Enabled = True
Command1.Caption = "Checking"

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

List1.AddItem "Finished checking" & vbTab & Date & "  " & Time

Timer1.Enabled = False
Timer2.Enabled = False
Timer2.Interval = 1000

Open "c:\internet log.txt" For Append As #1
      Print #1, vbCrLf & "Internet checker stopped  " & Date & "  " & Time & vbCrLf
      Close #1

Command1.Caption = "Start log"
counter = 0
End Sub

Private Sub Command4_Click()
Me.Hide
nid.cbSize = Len(nid)
nid.hwnd = Me.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Me.Icon
nid.szTip = "Inet checker" & vbNullChar
'Call Shell_NotifyIcon function to add to tray

Me.ScaleMode = vbTwips

Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Command5_Click()
List1.Clear
End Sub

Private Sub Ex_Click()
Shell_NotifyIcon NIM_DELETE, nid
Me.ScaleMode = vbPixels
Form1.Show
End Sub

Private Sub Form_Load()
Open "c:\internet log.txt" For Append As #1
      Print #1, "Program opened  " & Date & "  " & Time & vbCrLf
      Close #1
      
Dim LBTab(1) As Long
LBTab(0) = 75
LBTab(1) = 120
SendMessageArray List1.hwnd, LB_SETTABSTOPS, 2, LBTab(0)

Me.ScaleMode = vbPixels

MsgBox "take this msgbox out of Form_Load" & vbCrLf & "its just to let you know this" _
& vbCrLf & "program leaves a log .txt file and it" & vbCrLf & "would defeat the object to show" _
& vbCrLf & "where it is in the program" & vbCrLf & "C:\internet log.txt"

MsgBox "oh!, and this one too." & vbCrLf & "Dont forget to vote LOL"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim msg As Long
Dim sFilter As String
msg = x / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONDOWN
' Me.PopupMenu M 'if u want u can call the popup menu
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
Case WM_RBUTTONDOWN
Case WM_RBUTTONUP
Me.PopupMenu M 'call the popup menu
Case WM_RBUTTONDBLCLK
End Select

End Sub


Private Sub Form_Paint()
Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT

    'from black
    With vert(0)
        .x = 0
        .y = 0
        .Red = 0&
        .Green = 0&
        .Blue = 0&
        .Alpha = 0&
    End With

    'to blue
    With vert(1)
        .x = Me.ScaleWidth
        .y = Me.ScaleHeight
        .Red = 0&
        .Green = LongToUShort(&HFF00&)
        .Blue = LongToUShort(&HFF00&)
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    GradientFillRect Me.hdc, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H

End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid

Open "c:\internet log.txt" For Append As #1
      Print #1, vbCrLf & "Program closed  " & Date & "  " & Time
      Close #1


Set Form1 = Nothing

End Sub


Private Sub List1_DblClick()
If Command1.Caption = "Checking" Then
MsgBox "Stop the log or your actions will be recorded", vbInformation, "Before execution"
ElseIf Left$(List1.Text, 4) <> "http" Then
MsgBox "Not a valid URL", vbInformation, "Not valid"
Else
ShellExecute List1.hwnd, vbNullString, List1.Text, vbNullString, "C:\", SW_SHOWNORMAL
End If
End Sub


Private Sub Timer1_Timer()

Dim r As String, list As String
Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    r = Process32First(hSnapShot, uProcess)
    
Do While r
    list = list & Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
    r = Process32Next(hSnapShot, uProcess)
Loop
  CloseHandle hSnapShot

If InStr(1, list, "IEXPLORE.EXE") > 0 And counter = 0 Then
List1.AddItem "Internet opened" & vbTab & Date & "  " & Time
Open "c:\internet log.txt" For Append As #1
      Print #1, "Internet opened    " & Date & "  " & Time
      Close #1
Timer2.Enabled = True
Timer2.Interval = 10000
counter = 1

ElseIf InStr(1, list, "IEXPLORE.EXE") = 0 And counter = 1 Then
List1.AddItem "Internet closed" & vbTab & Date & "  " & Time
Open "c:\internet log.txt" For Append As #1
      Print #1, "Internet closed    " & Date & "  " & Time
      Close #1
counter = 0
Timer2.Enabled = False
Timer2.Interval = 1000
End If


End Sub


Private Sub Timer2_Timer()

  EnumWindows AddressOf EnumProc, 0

   
 If Mid$(Label2.Caption, 1, InStr(10, Label2.Caption, "/") - 1) <> List1.list(List1.ListCount - 1) Then
   List1.AddItem Mid$(Label2.Caption, 1, InStr(10, Label2.Caption, "/") - 1)
   Open "c:\internet log.txt" For Append As #1
      Print #1, Mid$(Label2.Caption, 1, InStr(10, Label2.Caption, "/") - 1)
      Close #1
 End If
   
End Sub


