VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Racing Game"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Left            =   6480
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   6120
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   0
   End
   Begin VB.Label SpeedC 
      BackStyle       =   0  'Transparent
      Caption         =   "0 MPH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label SpeedV 
      BackStyle       =   0  'Transparent
      Caption         =   "0 MPH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Viper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GTR - 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Viper Won"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ready..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Line FL 
      BorderWidth     =   4
      X1              =   13560
      X2              =   13560
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2400
      X2              =   13560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image yellow 
      Height          =   585
      Left            =   120
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image green 
      Height          =   495
      Left            =   240
      Picture         =   "Form1.frx":C14A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        
'Author: Jonathan Santos
'Please vote for me
'Feel free to edit to edit the code just put my name somewhere
'Questions or comentaries: JJonathann@msn.com

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'The keypress functions:
If KeyCode = vbKeySpace And go = True Then 'Check if the cars can takeoff
Timer2.Interval = 10
End If
If KeyCode = vbKeyShift And go = True Then 'Check if the cars can takeoff
Timer3.Interval = 10
End If
   
End Sub

Private Sub Form_Load()
FormSize 'Enlarge the form
Set_Car_Properties 'Sets the cars spped, acceleration and maxspeed
End Sub

Private Sub Timer1_Timer()
Start ' Activates the Keypressing functions
End Sub

Private Sub Timer2_Timer()
ForwardV 'Moves the viper forward
End Sub

Private Sub Timer3_Timer()
ForwardC 'Moves the corvette '65 forward
End Sub


