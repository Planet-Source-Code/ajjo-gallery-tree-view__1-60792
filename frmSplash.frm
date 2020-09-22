VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3255
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         Height          =   255
         Left            =   6600
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSplash.frx":000C
         Height          =   495
         Index           =   2
         Left            =   720
         TabIndex        =   4
         Top             =   2400
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSplash.frx":00B1
         Height          =   735
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Top             =   1560
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSplash.frx":01AC
         Height          =   735
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "GALLERY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub FormStayOnTop(FormToSet As Form, OnTop As Boolean)

    Dim lHwnd As Long
    Dim lFlags As Long
    Dim lPosFlag As Long
    
    lHwnd = FormToSet.hwnd
    lFlags = &H2 Or &H1 Or &H40 Or &H10
    
    Select Case OnTop
    
        Case True
            lPosFlag = -1
        Case False
             lPosFlag = -2
             
    End Select
    
    SetWindowPos lHwnd, lPosFlag, 0, 0, 0, 0, lFlags
    
End Sub

Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Call FormStayOnTop(Me, True)
    
    Me.Show
    DoEvents
    
    Form1.Show
    
    
End Sub



