VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Application"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3600
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send Data to Database File"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtAge 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Female"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton optGender 
      Caption         =   "Male"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtFullName 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
End
End Sub

Private Sub Command1_Click()
cmdExit.Enabled = False
Inet1.OpenURL "http://localhost/Put.asp?FullName=" + txtFullName.Text + "=Gender=" + optGender(Index).Caption + "=Age=" + txtAge.Text
DoEvents
cmdExit.Enabled = True
End Sub
