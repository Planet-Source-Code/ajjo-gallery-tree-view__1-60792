VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About CDBOL"
   ClientHeight    =   1980
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   4620
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   1366.631
   ScaleMode       =   0  'User
   ScaleWidth      =   4338.419
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   """ The best power of a human being is his ability to program himself to be effective and challenging in life """
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1035
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gallery"
      Height          =   195
      Index           =   0
      Left            =   1950
      TabIndex        =   0
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()

    Unload Me

End Sub

