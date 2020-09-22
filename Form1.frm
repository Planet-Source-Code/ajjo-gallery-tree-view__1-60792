VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gallery"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   7875
   ScaleWidth      =   11865
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   12000
      TabIndex        =   9
      Top             =   240
      Width           =   270
   End
   Begin MSComctlLib.ProgressBar pgbar 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   7440
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   6855
      Left            =   3240
      ScaleHeight     =   6795
      ScaleWidth      =   7875
      TabIndex        =   5
      Top             =   480
      Width           =   7935
      Begin VB.Image Image1 
         Height          =   1095
         Index           =   0
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   3480
      TabIndex        =   4
      Top             =   8040
      Width           =   495
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000006&
      Height          =   6885
      Left            =   11160
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   4560
      TabIndex        =   2
      Top             =   8040
      Width           =   465
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   5040
      TabIndex        =   1
      Top             =   8040
      Width           =   495
   End
   Begin MSComctlLib.ImageList ilsBook 
      Left            =   3960
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8E62
            Key             =   "FolderClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":92B6
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":970A
            Key             =   "Card"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9B5E
            Key             =   "Main"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treNames 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   13573
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ilsBook"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Path:"
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nod As Node
Dim i, j As Integer
Dim calculate As Variant
Dim aaa(0 To 100) As String

Dim glob, counter    As Integer
Dim a(0 To 25) As String
Dim findit As Integer

'The list of all the drives available in the local disk gets pumped to the tree view
Private Sub shallbecalled()

    Dim str As String
    
    For i = 0 To Drive1.ListCount - 1
    
        str = Trim(Left(Drive1.List(i), 3)) + "\"
        
          Set nod = treNames.Nodes.Add(, , str, str, "FolderClosed", "FolderOpen")
          Call resolve(str)
    
    Next i

End Sub
Private Sub resolve(what As String)

    Dim str As String
    
        On Error GoTo aa
        Dir1.Path = what
        
            For j = 0 To Dir1.ListCount - 1
    
                    calculate = Split(Dir1.List(j), "\")
                    str = Dir1.List(j) + "\"
                    On Error Resume Next
                    Set nod = treNames.Nodes.Add(what, tvwChild, str, calculate(UBound(calculate)), "FolderClosed", "FolderOpen")
                    
            Next j
            
            Exit Sub
            
aa:

    On Error Resume Next
    Set nod = treNames.Nodes.Add(what, tvwChild, what + "...", "None", "Card", "Card")

End Sub

Private Sub Command1_Click()

    frmAbout.Show vbModal

End Sub

Private Sub Form_Load()

    'Initialize
    pgbar.Value = 0
    
    Drive1.Visible = False
    Dir1.Visible = False
    File1.Visible = False
    List2.Visible = False

    glob = 5
    counter = 0
    
    Call shallbecalled

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim frm As Form
    
    For Each frm In Forms
    Unload frm
    Next

End Sub

Private Sub treNames_Collapse(ByVal Node As MSComctlLib.Node)

    Node.Image = "FolderClosed"

End Sub

Private Sub Image1_Click(Index As Integer)

    Dim i As Integer

    Form2.List1.Clear
    
    Form2.Text1.Text = Index
    
    For i = 1 To counter
    
        Form2.List1.AddItem a(i)
    
    Next i

     If Form2.List1.ListCount > 0 Then
     
        Form2.List1.ListIndex = Val(Form2.Text1.Text)
        Form2.Picture2.Picture = LoadPicture(Form2.List1.Text)
     
     End If
     
     Label2.Caption = Form2.List1.Text
     
     Form2.Show vbModal
  
End Sub

'Each folder details gets pumped to the tree view
Private Sub treNames_Expand(ByVal Node As MSComctlLib.Node)
    
    Node.Image = "FolderOpen"
    
    Dim letsresolve As Integer
    
    letsresolve = 0

        Dim str As String
        Dim str1 As String
        
        str = ""
        calculate = Split(Node.FullPath, "\")
    
        For i = 0 To UBound(calculate)
    
            If i <> 1 Then
            
                str = str + calculate(i) + "\"
                
            End If
    
        Next i
    
    
    For i = 0 To Drive1.ListCount - 1
    
        str1 = Trim(Left(Drive1.List(i), 3)) + "\"
        
        If (str <> str1) Then
        
            letsresolve = letsresolve + 1
        
        End If
    
    Next i
    
        If letsresolve = Drive1.ListCount Then
            
            Call resolve(str)
        
        End If
    
       If Node.Child.Text = "None" Then
          
             On Error GoTo aa
            Dir1.Path = str
            Call resolve(str)
       
       End If
       
    On Error GoTo aa:
    
        Dir1.Path = str
    
        j = Dir1.ListCount
    
        For i = 0 To j - 1
    
            aaa(i) = Dir1.List(i) + "\"
        
        Next i
        
    
        For i = 0 To j - 1
    
             Dir1.Path = aaa(i)
             
    
            If Dir1.ListCount > 0 Then
        
                    calculate = Split(Dir1.List(0), "\")
                    str = Dir1.List(0) + "\"
                    On Error Resume Next
                    Set nod = treNames.Nodes.Add(aaa(i), tvwChild, str, calculate(UBound(calculate)), "FolderClosed", "FolderOpen")
            
       
            End If
        
        Next i
        
    Exit Sub
aa:
        
        If Err.Number = 68 Then
        
            str = MsgBox("Device Not Ready", vbExclamation, "Device Status")
        
        End If

End Sub

Private Sub treNames_NodeClick(ByVal Node As MSComctlLib.Node)

    Image1(0).Picture = LoadPicture("")
    Dim i, letsresolve As Integer
    Dim str1, str As String
    
    letsresolve = 0
    
    
    For i = 0 To Drive1.ListCount - 1
    
    str1 = Trim(Left(Drive1.List(i), 3)) + "\"
    
    If (Node.FullPath <> str1) Then
    letsresolve = letsresolve + 1
    End If
    
    Next i
    
        If letsresolve = Drive1.ListCount Then
        
           
        On Error GoTo aa
        File1.Path = Left(Node.FullPath, 3) + Right(Node.FullPath, Len(Node.FullPath) - 4)
        Label2.Caption = File1.Path
        
        Else
         
        On Error GoTo aa:
        File1.Path = Node.FullPath
        Label2.Caption = File1.Path
        
        End If
    Call countall
    Call unloadall
    Call alterall(0, 5)
    
aa:
End Sub

'======================================================================================
Private Sub countall()

    Dim i As Integer
    Dim j As Integer
    List2.Clear
    List1.Clear
    
    For i = 0 To File1.ListCount - 1
    
        File1.ListIndex = i
        
        If (Right(File1.FileName, 3) = "jpg" Or Right(File1.FileName, 4) = "jpeg" Or Right(File1.FileName, 3) = "bmp") Then
        
            List2.AddItem File1.Path + "\" + File1.FileName
            
        End If
    
    Next i
    
    j = List2.ListCount \ 25
    
    
    For i = 1 To j
    
        List1.AddItem i
    
    Next i
    
    j = List2.ListCount - (j * 25)
    List1.AddItem j

End Sub

Private Sub unloadall()

    Dim i As Integer
    
    For i = 1 To Image1.UBound
    
        Unload Image1(i)
        
    Next i

End Sub

Private Sub alterall(intgetvalue As Integer, globrefreshed As Integer)

    counter = 0

    Dim i As Integer

    For i = intgetvalue To intgetvalue + 24
    
        If i <= List2.ListCount - 1 Then
        
        List2.ListIndex = i
        
                If i = globrefreshed Then
        
                    globrefreshed = globrefreshed + 5
        
                    Call resetall(counter)
                    Image1(counter).Picture = LoadPicture(List2.Text)
        
                Else
                    
                    If i = 0 Then
        
                            Image1(0).Picture = LoadPicture(List2.Text)
        
        
                    Else
        
        
                            Call position(counter)
                            Image1(counter).Picture = LoadPicture(List2.Text)
        
        
                    End If
        
        
                 End If
        
            counter = counter + 1
            
            a(counter) = List2.Text
        
        End If
        
        
        pgbar.Value = pgbar.Value + 4
        
        If pgbar.Value >= 100 Then
        
            pgbar.Value = 0
        
        End If
    
    Next i


End Sub

Private Sub position(Index As Integer)

    If Index <> 0 Then
    
        Load Image1(Index)
        Image1(Index).Left = Image1(Index - 1).Left + 1560
        Image1(Index).Height = Image1(Index - 1).Height
        Image1(Index).Top = Image1(Index - 1).Top
        Image1(Index).Width = Image1(Index - 1).Width
        Image1(Index).Visible = True
    
    End If

End Sub

Private Sub resetall(Index As Integer)

    Load Image1(Index)
    Image1(Index).Left = Image1(Index - 1).Left - 6240
    Image1(Index).Height = Image1(Index - 1).Height
    Image1(Index).Top = Image1(Index - 1).Top + 1350
    Image1(Index).Width = Image1(Index - 1).Width
    Image1(Index).Visible = True

End Sub

Private Sub List1_Click()
    
    Dim j, i, k As Integer
    
    Call unloadall
    
    If List1.ListIndex = 0 Then
    
        Call alterall(0, 5)
    
    ElseIf List1.ListIndex = List1.ListCount - 1 Then
    
        k = Val(List1.Text)
        
        i = 25 * (k - (k - Val(List1.List(List1.ListIndex - 1))))
        glob = i + 5
        
        Call alterall(Val(i), Val(glob))
        
        
    Else
        
        i = 25 * (Val(List1.Text) - 1)
        glob = i + 5
        
        Call alterall(Val(i), Val(glob))
    
    End If


End Sub

