VERSION 5.00
Begin VB.Form frmProcess 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Controller"
   ClientHeight    =   5505
   ClientLeft      =   2985
   ClientTop       =   3180
   ClientWidth     =   11010
   Icon            =   "Process.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00400040&
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00400040&
      Caption         =   "&Start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00400040&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   150
      Left            =   6120
      TabIndex        =   4
      Top             =   4320
      Width           =   4815
   End
   Begin VB.ListBox lstBlocked 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3900
      Left            =   6120
      TabIndex        =   3
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00400040&
      Caption         =   "&Get"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ListBox lstList 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4620
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   4080
      Top             =   5040
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "KEYS :  F1   Help  F9   Hide  F10 Show"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1185
      Index           =   4
      Left            =   5040
      TabIndex        =   12
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "List of Applications to be Block :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   3570
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Active Window handles and Names :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3570
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Name of Application to be Blocked ( e.g. Nescape Navigator )"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   1
      Left            =   6120
      TabIndex        =   5
      Top             =   4440
      Width           =   3450
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Click here to get all windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   5160
      Width           =   3090
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdRemove_Click()
    If lstBlocked.ListCount = 0 Then
        MsgBox "Nothing to remove", vbInformation
    Else
        If lstBlocked.ListIndex = -1 Then
            MsgBox "Proper selection required.", vbInformation
            Exit Sub
        End If
        lstBlocked.RemoveItem lstBlocked.ListIndex
    End If
End Sub

Private Sub cmdStart_Click()
    bStart = IIf(bStart = True, False, True)
    If bStart = True Then
        cmdStart.Caption = "&Stop"
    Else
        cmdStart.Caption = "&Start"
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo err

    lstList.Clear
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
err:
Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo cmdAdd_Click_err
    If Trim$(txtName.Text) = "" Then
    MsgBox "Enter the name of application", vbInformation
    Exit Sub
    End If

    lstBlocked.AddItem Trim(txtName.Text)
    txtName.Text = ""

Exit Sub
cmdAdd_Click_err:
    Debug.Print err.Description
End Sub




Private Sub Form_Unload(Cancel As Integer)
Unload frmHelp
Unload Me
End
End Sub

Private Sub lstList_DblClick()
On Error GoTo err
    Dim sWindowText  As String
    Dim l            As Long
    
   
    l = InStr(1, Trim$(lstList.List(lstList.ListIndex)), Chr$(134))
    If l = Len(Trim$(lstList.List(lstList.ListIndex))) Then
    Else
        sWindowText = Mid$(lstList.List(lstList.ListIndex), l + 1, Len(Trim$(lstList.List(lstList.ListIndex))))
        txtName.Text = Trim$(sWindowText)
    End If
    
err:
Debug.Print err.Description
Exit Sub
End Sub

Private Sub lstList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lstList.ListCount > 0 Then
        lstList.ToolTipText = "Double click to block the Active Window handler or Name"
    Else
        lstList.ToolTipText = "Click Get to get the Active Windows handler and Names"
    End If
End Sub

Private Sub Timer_Timer()
    If GetAsyncKeyState(VK_F9) Then Me.Hide
    If GetAsyncKeyState(VK_F10) Then Me.Show
    If GetAsyncKeyState(VK_F1) Then frmHelp.Show
    If lstBlocked.ListCount > 0 Then
        cmdStart.Enabled = True
        cmdRemove.Enabled = True
    Else
        cmdStart.Enabled = False
        cmdRemove.Enabled = False
        bStart = False
    End If
    
    If bStart = True Then
        EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    End If
End Sub
