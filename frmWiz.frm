VERSION 5.00
Begin VB.Form frmWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wizard Sample"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "frmWiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSteps 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   0
      Left            =   2160
      TabIndex        =   11
      Top             =   0
      Width           =   6375
      Begin VB.Shape Shape5 
         BorderColor     =   &H00800000&
         BorderWidth     =   4
         Height          =   255
         Left            =   5040
         Top             =   2640
         Width           =   255
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00008000&
         BorderWidth     =   4
         Height          =   255
         Left            =   5040
         Top             =   2280
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         BorderWidth     =   4
         FillColor       =   &H000000C0&
         Height          =   255
         Left            =   5040
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This wizard will create a document suitable for Email merging using the parameters that you specify."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblHead 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Wizard Sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   4
      Left            =   2160
      TabIndex        =   15
      Top             =   0
      Width           =   6375
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Step 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   3
      Left            =   2160
      TabIndex        =   14
      Top             =   0
      Width           =   6375
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Step 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   2
      Left            =   2160
      TabIndex        =   13
      Top             =   0
      Width           =   6375
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Step 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   1
      Left            =   2160
      TabIndex        =   12
      Top             =   0
      Width           =   6375
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Step 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5010
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox picWiz 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   2175
      TabIndex        =   4
      Top             =   0
      Width           =   2175
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finish"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   10
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   9
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   7
         Top             =   960
         Width           =   465
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   465
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   120
         Width           =   420
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   1  'Square
         Top             =   2400
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   360
         Shape           =   1  'Square
         Top             =   1920
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   360
         Shape           =   1  'Square
         Top             =   1440
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   360
         Shape           =   1  'Square
         Top             =   960
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   360
         Shape           =   1  'Square
         Top             =   480
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   120
         Shape           =   1  'Square
         Top             =   120
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   2295
         Left            =   240
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   5
      Left            =   2160
      TabIndex        =   16
      Top             =   0
      Width           =   6375
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Wizard Sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2040
         TabIndex        =   24
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Those are all the answers that the wizard needs to create mail merge document. Click finish to create the document"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2760
         TabIndex        =   23
         Top             =   720
         Width           =   3495
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00000080&
         BorderWidth     =   4
         FillColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         Top             =   120
         Width           =   255
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00008000&
         BorderWidth     =   4
         Height          =   255
         Left            =   120
         Top             =   480
         Width           =   255
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00800000&
         BorderWidth     =   4
         Height          =   255
         Left            =   480
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2160
      Top             =   3840
      Width           =   6375
   End
End
Attribute VB_Name = "frmWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Integer

Private Sub MoveIndex(PrevIndex As Integer, NextIndex As Integer)
    'Are we at the last step?
    If PrevIndex = shpSteps.Count - 1 Then
        'Yes
        shpSteps(PrevIndex).FillColor = vbRed
    Else
        'No
        shpSteps(PrevIndex).FillColor = &H808080
    End If
    'Set bold off
    lblSteps(PrevIndex).FontBold = False
    
    'Set next step color and font bold
    shpSteps(NextIndex).FillColor = vbGreen
    lblSteps(NextIndex).FontBold = True
End Sub

Private Sub cmdBack_Click()
    Index = Index - 1
    If Index <= 0 Then
        Index = 0
        cmdBack.Enabled = False
    End If
    
    MoveIndex Index + 1, Index

    'Set the frames
    fraSteps(Index).ZOrder 0
    
    'Set command buttons
    cmdNext.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    'Don't do anything
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    'Put some code here to check if all data captured on the previous screens
    'was correct and then proceed
    PerformSomething
    Unload Me
End Sub

Private Sub PerformSomething()
    MsgBox "Finished creating document", vbOKOnly, "Finished"
End Sub

Private Sub cmdNext_Click()
    Index = Index + 1
    If Index >= shpSteps.Count - 1 Then
        Index = shpSteps.Count - 1
        cmdNext.Enabled = False
    End If
    
    MoveIndex Index - 1, Index
    
    'Set the frames
    fraSteps(Index).ZOrder 0
    
    'Set command buttons
    cmdBack.Enabled = True
End Sub

Private Sub Form_Load()
    Index = 0
End Sub
