VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About jcForms 1.05"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin jcFormsDemo.jcForms jcForms1 
      Align           =   1  'Align Top
      Height          =   3090
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   5450
      BackColorStyle  =   1
      BorderStyle     =   0
      MaxButton       =   0   'False
      MinButton       =   0   'False
      CustomBackColor =   16776948
      Style           =   5
      ChangeAllBackgrounds=   -1  'True
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FEECDD&
         Caption         =   "Close"
         Height          =   345
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2490
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   1425
         Left            =   210
         TabIndex        =   2
         Top             =   750
         Width           =   4875
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Show some information
    Label1.Caption = " jcForms 1.05 - A nice control to skin your form" + vbCrLf + vbCrLf + _
            "by Juan Carlos San Román Arias" + vbCrLf + _
            "2006" + vbCrLf + vbCrLf + vbCrLf + _
            "See code in Demo form to learn how to handle added menu items"
End Sub
