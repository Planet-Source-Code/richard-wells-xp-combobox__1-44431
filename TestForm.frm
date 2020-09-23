VERSION 5.00
Begin VB.Form TestForm 
   BackColor       =   &H00E0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XP Combo Class"
   ClientHeight    =   2775
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Enable"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox Combo4 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "Combo4"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   765
      Picture         =   "TestForm.frx":0000
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   135
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   525
      Picture         =   "TestForm.frx":032E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   135
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   300
      Picture         =   "TestForm.frx":065C
      ScaleHeight     =   270
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   135
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   60
      Picture         =   "TestForm.frx":098A
      ScaleHeight     =   270
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   135
      Width           =   240
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Read the file "Read me b4 trying to run the project.txt"
'Before running this project.
Private Sub Command1_Click()
    If Combo4.Enabled = True Then
        Combo4.Enabled = False
        Command1.Caption = "Enable"
    Else
        Combo4.Enabled = True
        Command1.Caption = "Disable"
    End If
End Sub

Private Sub Form_Load()

'If all goes well this is all you have to do.
    DrawTheXPCombos Me
'//
'This is my ugly custom scheme.
'Un-rem to reveal its dark side.
    'DrawTheXPCombos Me, vbRed, vbGreen, vbBlue, pic(0).Picture, pic(1).Picture, pic(2).Picture, pic(3).Picture
'//

'Nasty bug.
'Cant seem to show the combo as being disabled on Form_Load.
'Need help here :-(
'Also dosent support the Simple Combo style either
'as i dont use it.
Combo4.Enabled = False
'//
End Sub


