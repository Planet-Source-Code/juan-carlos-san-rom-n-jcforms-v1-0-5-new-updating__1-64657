VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Another form"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin jcFormsDemo.jcForms jcForms1 
      Align           =   1  'Align Top
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   4868
      ThemeColor      =   0
      BorderStyle     =   0
      MaxButton       =   0   'False
      MinButton       =   0   'False
      CustomBackColor =   16777215
      Style           =   2
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   270
         TabIndex        =   2
         Top             =   930
         Width           =   4545
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Form changes its color tonality when it gets or loses focus"
            Height          =   345
            Left            =   270
            TabIndex        =   3
            Top             =   360
            Width           =   4065
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   315
         Left            =   3780
         TabIndex        =   1
         Top             =   2130
         Width           =   1065
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
