VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Faisal's Menu"
   ClientHeight    =   4065
   ClientLeft      =   1395
   ClientTop       =   1710
   ClientWidth     =   7380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   7380
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   6720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   120
      Picture         =   "frmMain.frx":0442
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label UsrOption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label UsrOption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   2040
      Width           =   990
   End
   Begin VB.Label UsrOption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   0
      Left            =   6600
      TabIndex        =   2
      Top             =   3600
      Width           =   660
   End
   Begin VB.Label UsrOption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   2
      Left            =   2400
      TabIndex        =   1
      Top             =   1440
      Width           =   1410
   End
   Begin VB.Label UsrOption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   1
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lblMaxWidth As Long      'This variable is for the widest label
Private Sub ButtonChanger(X, Y)

For CN = 0 To UsrOption1.UBound
    With UsrOption1(CN)
    If X > .Left And X < .Left + .Width And Y > .Top And Y < .Top + .Height Then
        .ForeColor = RGB(255, 0, 0)
    Else
        .ForeColor = RGB(0, 0, 0)
    End If
    End With
Next CN

End Sub



Private Sub Form_Load()
'*****************************************************************


Rem **** Set Constants
UserScreenX = SysInfo1.WorkAreaWidth
UserScreenY = SysInfo1.WorkAreaHeight

'!!! If the SysInfo control doesn't come across properly
'!!! Replace SysInfo1.WorkAreaWidth with your screen's Width (640,800,1024,1280...)
'!!! Replace SysInfo1.WorkAreaHeight with your screen's Height (480,600,768,1024...)


'!!! Don't forget to look at the MouseMove events in UsrOption1 and in the form

Rem **** Set Display Parameters
With frmMain
    .Width = 7500
    .Height = 4500
    .Left = (UserScreenX / 2) - (.Width / 2)
    .Top = (UserScreenY / 2) - (.Height / 2)
    UsrOption1(0).Left = .Width - UsrOption1(0).Width - 200
    UsrOption1(0).Top = .Height - UsrOption1(0).Height - 400
End With

'Look for widest label excluding the EXIT label (index 0)
'Hint..... Look for lblMaxWidth in the form resize
'This works great when the label property AutoSize is true
lblMaxWidth = 0
For X = 1 To UsrOption1.UBound
    With UsrOption1(X)
        If lblMaxWidth < .Width Then lblMaxWidth = .Width
        .Left = (frmMain.Width / 2) - (.Width / 2)
        .Top = (frmMain.Height / (UsrOption1.UBound + 2)) * X
    End With
Next X

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonChanger X, Y

End Sub


Private Sub Form_Resize()

'Make sure form doesn't get too small
If frmMain.Width < lblMaxWidth * 1.5 Then
    frmMain.Width = lblMaxWidth * 1.5
    Exit Sub
End If
If frmMain.Height < UsrOption1(1).Height * (UsrOption1.UBound + 4) Then
    frmMain.Height = UsrOption1(1).Height * (UsrOption1.UBound + 4)
    Exit Sub
End If

'Position EXIT label in bottom right corner
With UsrOption1(0)
    .Left = frmMain.Width - .Width - 200
    .Top = frmMain.Height - .Height - 400
End With

'Locate the rest of the labels on the form
For X = 1 To UsrOption1.UBound
    With UsrOption1(X)
        .Left = (frmMain.Width / 2) - (.Width / 2)
        .Top = (frmMain.Height / (UsrOption1.UBound + 2)) * X
    End With
Next X

End Sub

Private Sub UsrOption1_Click(Index As Integer)

If Index = 0 Then End

If Index = 1 Then
    MsgBox "Label Index " & Index & " pressed", vbOKOnly
    'Unload frmMain
End If
If Index = 2 Then
    MsgBox "Label Index " & Index & " pressed", vbOKOnly
    'Unload frmMain
End If
If Index = 3 Then
    MsgBox "Label Index " & Index & " pressed", vbOKOnly
    'Unload frmMain
End If
If Index = 4 Then
    MsgBox "Label Index " & Index & " pressed", vbOKOnly
    'Unload frmMain
End If

End Sub

Private Sub UsrOption1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonChanger X + UsrOption1(Index).Left, Y + UsrOption1(Index).Top
End Sub


