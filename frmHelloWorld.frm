VERSION 5.00
Begin VB.Form frmHelloWorld 
   Caption         =   "Hello World Software"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12435
   Icon            =   "frmHelloWorld.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Browse Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      Picture         =   "frmHelloWorld.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   5535
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6510
      Picture         =   "frmHelloWorld.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   5535
   End
   Begin VB.CommandButton cmdClickMe 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Hi, Click Me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   390
      Picture         =   "frmHelloWorld.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hello World"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   99.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   30
      TabIndex        =   0
      Top             =   840
      Width           =   13095
   End
End
Attribute VB_Name = "frmHelloWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClickMe_Click()
    MsgBox "Hello, This is my software !", vbInformation
End Sub

Private Sub cmdExit_Click()
    Dim intSure As Integer
    intSure = MsgBox("Do you want to close this software ?", vbYesNo)
    If intSure = 6 Then
        End
    End If
End Sub
