VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get All Files :"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   720
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   10335
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   10
      Scrolling       =   1
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   10335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START >>"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder To Get Files :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1740
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   10095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working ....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Index           =   0
      Left            =   3840
      TabIndex        =   4
      Top             =   2040
      Width           =   2790
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author  : Abdalla Mahmoud
'Age     : 14
'Country : Egypt
'Cisty   : Mansoura
'E-Mails : la3toot@hotmail.com ; la3toot@yahoo.com
'*******************************************************
'* Please Read This Code and Send Me Mail about
'Your Comments
'* Please Excuse My English If Is Bad , Because I Still
'Young And English Is Not My Language .
'*******************************************************
'This Is The Main Function You Nead
'*******************************************************
Const vbAllAttr As Long = vbArchive Or vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem
Function GetAllFiles(ByVal Folder As String, ByRef Folders As Collection, ByRef Files As Collection, Optional Attributes As VbFileAttribute = vbAllAttr) As Long
'This Function Get All Files That Is Located In
'Given Folder Including SubFolders , Also Get
'All Folders Located In It and Return Results To
'Two Collections , 1.For Folders , 2.For Files
'This Function Return a Long Value
'0 = Success ; 1 = Folder Not Exists ; -1 = Other Error
On Error GoTo FunErr

Dim CollFolders As New Collection
Dim CollFiles   As New Collection
Dim NowDir      As String
Dim TmpDir      As String
Dim Pos         As Long
Dim Tmp         As String
Dim FileSysObj  As Object
Dim Max         As Long

'We Use This Function To Create FileSystemObject
'You Will Not Have To Include Your Project a FileSystemObject
Set FileSysObj = CreateObject("Scripting.FileSystemObject")
'Check Folder String
Folder = Folder & IIf(Right(Folder, 1) = "\", vbNullString, "\")
'Determine If Folder Exists With FileSystemObject
If FileSysObj.FolderExists(Folder) = False Then
    If FileSysObj.DriveExists(Folder) = False Then
        GoTo NotFound
    End If
End If
'Add Given Folder To Folders' Collection
CollFolders.Add Folder
'Load First Values To Variables
Pos = 1: Max = 1
'Try To Rear The Code and Understand It
Do
    If Max < Pos Then GoTo Finish
    NowDir = CollFolders.Item(Pos)
    TmpDir = Dir(NowDir, Attributes)
    Do While TmpDir <> ""
        If TmpDir <> "." And TmpDir <> ".." Then
            Tmp = NowDir & TmpDir
            If FileSysObj.FolderExists(Tmp) = False Then
                CollFiles.Add Tmp
            Else
                CollFolders.Add Tmp & "\"
                Max = Max + 1
            End If
        End If
        TmpDir = Dir
        DoEvents
    Loop
    Pos = Pos + 1
    DoEvents
Loop

Finish:
    Set Folders = CollFolders
    Set Files = CollFiles
    GetAllFiles = 0
    GoTo SetNothing
Exit Function

FunErr:
    GetAllFiles = -1
    GoTo SetNothing
Exit Function

NotFound:
    GetAllFiles = 1
    GoTo SetNothing
Exit Function

SetNothing:
    Set CollFiles = Nothing
    Set CollFolders = Nothing
    Set FileSysObj = Nothing
End Function
'*******************************************************
'The Rest Of Codes You Will Not Need It
'It's For The Other Work For Programme
'*******************************************************

Private Sub Command1_Click()
On Error Resume Next

Dim X1 As Collection
Dim X2 As Collection
Dim I  As Long
Dim Cnt As Long
Dim Ret As Long

List1.Visible = False
Timer1.Enabled = True
Ret = GetAllFiles(Text1.Text, X1, X2)
If Ret Then
    Select Case Ret
    Case 1
        MsgBox "Folder Not Exists  .", vbCritical, "Abdalla For Programming"
    Case -1
        MsgBox "Error While Getting Files .", vbCritical, "Abdalla For Programming"
    End Select
    List1.Visible = True
    Exit Sub
End If
List1.Clear
PB.Max = X2.Count
lbl(1).Caption = PB.Max & "  Files Found , Now Adding These Files To List ."
List1.Visible = False
For I = 1 To PB.Max
    List1.AddItem X2(I)
    PB.Value = I
    DoEvents
Next
List1.Visible = True
MsgBox X2.Count & " Files Found , In " & X1.Count & " Folders .", , "Abdalla For Programming ."
PB.Value = 0
lbl(1).Caption = ""
Timer1.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
KeyAscii = 0
Command1_Click
End Sub

Private Sub Timer1_Timer()
If lbl(0).Caption = "Working ." Then
    lbl(0).Caption = "Working .."
ElseIf lbl(0).Caption = "Working .." Then
    lbl(0).Caption = "Working ..."
ElseIf lbl(0).Caption = "Working ..." Then
    lbl(0).Caption = "Working ...."
ElseIf lbl(0).Caption = "Working ...." Then
    lbl(0).Caption = "Working ....."
ElseIf lbl(0).Caption = "Working ....." Then
    lbl(0).Caption = "Working "
ElseIf lbl(0).Caption = "Working " Then
    lbl(0).Caption = "Working ."
End If
End Sub
