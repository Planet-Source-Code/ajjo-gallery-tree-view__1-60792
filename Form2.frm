VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hulk View"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0442
   ScaleHeight     =   7605
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Height          =   7215
      Left            =   10560
      Picture         =   "Form2.frx":6A8F
      ScaleHeight     =   7155
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Start Slide"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Info"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Exit"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Form2.frx":930C
         Height          =   4635
         Left            =   0
         TabIndex        =   10
         Top             =   2520
         Width           =   1080
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11280
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   11880
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   0
      Width           =   150
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   11520
      TabIndex        =   4
      Top             =   1800
      Width           =   135
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5775
      Left            =   9840
      Max             =   100
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   100
      TabIndex        =   2
      Top             =   6240
      Width           =   9975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      Height          =   7155
      Left            =   120
      ScaleHeight     =   7095
      ScaleWidth      =   10335
      TabIndex        =   0
      Top             =   120
      Width           =   10395
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000007&
         Height          =   7695
         Left            =   0
         ScaleHeight     =   7635
         ScaleWidth      =   10635
         TabIndex        =   1
         Top             =   0
         Width           =   10695
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim calculatecounter, holdleft As Integer
Dim fso As New FileSystemObject

Private Sub Command1_Click()

    Picture1.Move 10, 10
    Picture2.Move 0, 0
    
    If Command1.Caption = "Start Slide" Then
    Command1.Caption = "Stop Slide"
    Timer1.Enabled = True
    Else
    Command1.Caption = "Start Slide"
    Timer1.Enabled = False
    Picture1.SetFocus
    End If

End Sub

Private Sub Command2_Click()

    Dim txtfile As File
    Dim str As String
    
    On Error GoTo a
    Set txtfile = fso.GetFile(List1.Text)
    
    
    str = MsgBox("File Path                ---    " + CStr(txtfile.Path) + vbCrLf _
                + "File Type            ---    " + CStr(txtfile.Type) + vbCrLf _
                + "File Size           ---   " + CStr(txtfile.Size) + vbCrLf _
                + "Created           ---   " + CStr(txtfile.DateCreated) + vbCrLf _
                + "Last Modified   ---   " + CStr(txtfile.DateLastModified) + vbCrLf _
                + "Last Accessed       ---   " + CStr(txtfile.DateLastAccessed), , "File Info")
    
a:
    
    Picture1.SetFocus

End Sub

Private Sub Command3_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    List1.Visible = False
    Text1.Visible = False
    
    Timer1.Enabled = False
    
    
    Form2.ScaleMode = vbPixels
    Picture1.ScaleMode = vbPixels
    
    Picture2.AutoSize = True
    
    
    Picture1.BorderStyle = 0
    Picture2.BorderStyle = 0
     
    
    Picture1.Move 10, 10
    Picture2.Move 0, 0
    
    
    Call scrolldefined
    
    
    VScroll1.Visible = False
    HScroll1.Visible = False

End Sub
Private Sub scrolldefined()

   HScroll1.Top = Picture1.Height
   HScroll1.Left = 0
  HScroll1.Width = Picture1.Width

   VScroll1.Top = 0
   VScroll1.Left = Picture1.Width
   VScroll1.Height = Picture1.Height

   HScroll1.Max = Picture2.Width - Picture1.Width
   VScroll1.Max = Picture2.Height - Picture1.Height

End Sub


Private Sub HScroll1_Change()

      Picture2.Left = -HScroll1.Value
           
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)

    calculatecounter = Val(Text1.Text)
    
    Select Case KeyCode
    
    Case 13
                Unload Me
                Form1.Show
                
    Case 39
    
                If Picture2.Width > Picture1.Width Then
    
                    If HScroll1.Value + 20 <= HScroll1.Max Then
                    Picture2.Left = Picture2.Left - 20
                    HScroll1.Value = HScroll1.Value + 20
                    Else
                    Picture2.Left = Picture1.Left - HScroll1.Max
                    End If
    
              End If
              
    Case 37
    
            If Picture2.Width > Picture1.Width Then
    
                    If HScroll1.Value - 20 > HScroll1.Min Then
                    Picture2.Left = Picture2.Left + 20
                    HScroll1.Value = HScroll1.Value - 20
                    Else
                    Picture2.Left = 0
                    End If
    
              End If
    
    Case 40
    
                 If Picture2.Height > Picture1.Height Then
    
                    If VScroll1.Value + 20 <= VScroll1.Max Then
                    Picture2.Top = Picture2.Top - 20
                    VScroll1.Value = VScroll1.Value + 20
                    Else
                    Picture2.Top = Picture1.Top - VScroll1.Max
                    End If
    
              End If
              
    Case 38
    
              If Picture2.Height > Picture1.Height Then
    
                    If VScroll1.Value - 20 > VScroll1.Min Then
                    Picture2.Top = Picture2.Top + 20
                    VScroll1.Value = VScroll1.Value - 20
                    Else
                    Picture2.Top = 0
                    End If
    
              End If
              
    Case 34
    
            If calculatecounter < List1.ListCount - 1 Then
            calculatecounter = calculatecounter + 1
            List1.ListIndex = calculatecounter
            Picture2.Picture = LoadPicture(List1.Text)
            Text1.Text = Val(Text1.Text) + 1
            refreshall
            End If
            
    Case 33
    
    
            If calculatecounter > 0 Then
            calculatecounter = calculatecounter - 1
            List1.ListIndex = calculatecounter
            Picture2.Picture = LoadPicture(List1.Text)
              Text1.Text = Val(Text1.Text) - 1
            refreshall
            End If
            
    Case 35
    
            If List1.ListCount > 0 Then
            List1.ListIndex = List1.ListCount - 1
            Picture2.Picture = LoadPicture(List1.Text)
              Text1.Text = List1.ListCount - 1
            End If
            
    Case 36
    
            If List1.ListCount > 0 Then
            List1.ListIndex = 0
            Picture2.Picture = LoadPicture(List1.Text)
              Text1.Text = 0
            End If
    End Select

    HScroll1.Max = Picture2.Width - Picture1.Width
    VScroll1.Max = Picture2.Height - Picture1.Height

End Sub

Private Sub refreshall()
    
    VScroll1.Value = 0
    Picture2.Left = 0
    Picture2.Top = 0

End Sub

Private Sub Timer1_Timer()

    If List1.ListIndex + 1 < List1.ListCount Then
    
        List1.ListIndex = List1.ListIndex + 1
        Text1.Text = List1.ListIndex
        Picture2.Picture = LoadPicture(List1.Text)
    
    Else
    
        List1.ListIndex = 0
        Picture2.Picture = LoadPicture(List1.Text)
    
    End If

End Sub
