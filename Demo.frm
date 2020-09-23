VERSION 5.00
Begin VB.Form Demo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "jcForms v 1.0.5"
   ClientHeight    =   7380
   ClientLeft      =   3945
   ClientTop       =   2340
   ClientWidth     =   7560
   ForeColor       =   &H8000000E&
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   7560
   Begin jcFormsDemo.jcForms jcForms1 
      Align           =   1  'Align Top
      Height          =   7380
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   13018
      ColorFrom       =   16512
      WindowState     =   2
      CustomBackColor =   16777215
      Style           =   2
      Begin VB.CommandButton Command4 
         Caption         =   "Change jcForms Menu Captions"
         Height          =   375
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "You can change caption in jcForms menu items and title bar buttons"
         Top             =   6750
         Width           =   2565
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Another Form"
         Height          =   375
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Load another form"
         Top             =   6750
         Width           =   1485
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         Caption         =   "Caption buttons:"
         ForeColor       =   &H80000008&
         Height          =   1515
         Left            =   240
         TabIndex        =   31
         Top             =   5010
         Width           =   1935
         Begin VB.CheckBox Chkclose 
            Appearance      =   0  'Flat
            Caption         =   "Close"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   240
            TabIndex        =   34
            ToolTipText     =   "Show or hide close button"
            Top             =   1050
            Value           =   1  'Checked
            Width           =   1305
         End
         Begin VB.CheckBox ChkMax 
            Appearance      =   0  'Flat
            Caption         =   "Maximize"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   240
            TabIndex        =   33
            ToolTipText     =   "Show or hide maximize button"
            Top             =   660
            Value           =   1  'Checked
            Width           =   1305
         End
         Begin VB.CheckBox ChkMin 
            Appearance      =   0  'Flat
            Caption         =   "Minimize "
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   240
            TabIndex        =   32
            ToolTipText     =   "Show or hide minimize button"
            Top             =   270
            Value           =   1  'Checked
            Width           =   1305
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         Caption         =   "Borderstyle:"
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   2340
         TabIndex        =   28
         Top             =   5190
         Width           =   2355
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            Caption         =   "Sizable"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   210
            TabIndex        =   30
            Top             =   810
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            Caption         =   "Fixed"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   29
            Top             =   390
            Width           =   1185
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "General:"
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   4860
         TabIndex        =   23
         Top             =   4575
         Width           =   2325
         Begin VB.CheckBox chkTitleBarShadow 
            Appearance      =   0  'Flat
            Caption         =   "TitleBarShadow"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   390
            TabIndex        =   38
            ToolTipText     =   "Show or hide titlebar shadow"
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Add menu item"
            Height          =   375
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Add or remove menu item to jcforms"
            Top             =   1065
            Width           =   1605
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Change Caption"
            Height          =   375
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Change caption of the form"
            Top             =   630
            Width           =   1605
         End
         Begin VB.ComboBox CboIconSize 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "Demo.frx":08CA
            Left            =   1290
            List            =   "Demo.frx":08D7
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Change icon titlebar size "
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "IconSize:"
            Height          =   195
            Left            =   420
            TabIndex        =   26
            Top             =   315
            Width           =   660
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "Custom colors:"
         ForeColor       =   &H80000008&
         Height          =   3695
         Left            =   4860
         TabIndex        =   18
         Top             =   765
         Width           =   2325
         Begin VB.OptionButton OptnColors 
            Appearance      =   0  'Flat
            Caption         =   "ColorTo"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   1020
            Width           =   1785
         End
         Begin VB.OptionButton OptnColors 
            Appearance      =   0  'Flat
            Caption         =   "ColorFrom"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   660
            Width           =   1785
         End
         Begin VB.OptionButton OptnColors 
            Appearance      =   0  'Flat
            Caption         =   "CustomBackColor"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   330
            Value           =   -1  'True
            Width           =   1785
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00F8F3F1&
            ForeColor       =   &H80000008&
            Height          =   2115
            Left            =   110
            Picture         =   "Demo.frx":08E7
            ScaleHeight     =   139
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   139
            TabIndex        =   19
            Top             =   1460
            Width           =   2115
            Begin VB.Shape ShpColor 
               BackColor       =   &H00F8F3F1&
               BorderColor     =   &H00FFFFFF&
               FillColor       =   &H00404040&
               Height          =   270
               Left            =   1080
               Top             =   60
               Width           =   270
            End
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   2340
         TabIndex        =   13
         Top             =   3540
         Width           =   2355
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "ChangeAllBackgrounds"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   150
            TabIndex        =   27
            ToolTipText     =   "Check it if you want to change backcolor of all controls"
            Top             =   0
            Width           =   2025
         End
         Begin VB.OptionButton OptnBackcolorStyle 
            Appearance      =   0  'Flat
            Caption         =   "Custom"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   "Use Custom BackColor"
            Top             =   1110
            Width           =   1065
         End
         Begin VB.OptionButton OptnBackcolorStyle 
            Appearance      =   0  'Flat
            Caption         =   "Auto"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   15
            ToolTipText     =   "It automatically changes with theme color selection"
            Top             =   750
            Width           =   1065
         End
         Begin VB.OptionButton OptnBackcolorStyle 
            Appearance      =   0  'Flat
            Caption         =   "Default"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   14
            ToolTipText     =   "It sets button face color as backcolor"
            Top             =   390
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Style:"
         ForeColor       =   &H80000008&
         Height          =   2625
         Left            =   2340
         TabIndex        =   8
         Top             =   765
         Width           =   2355
         Begin VB.OptionButton OptnStyle 
            Appearance      =   0  'Flat
            Caption         =   "Style6"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   37
            Top             =   2220
            Width           =   915
         End
         Begin VB.OptionButton OptnStyle 
            Appearance      =   0  'Flat
            Caption         =   "Style5"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   36
            Top             =   1842
            Width           =   915
         End
         Begin VB.OptionButton OptnStyle 
            Appearance      =   0  'Flat
            Caption         =   "Style4"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   12
            Top             =   1464
            Width           =   915
         End
         Begin VB.OptionButton OptnStyle 
            Appearance      =   0  'Flat
            Caption         =   "Style3"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   11
            Top             =   1086
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton OptnStyle 
            Appearance      =   0  'Flat
            Caption         =   "Style2"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   10
            Top             =   708
            Width           =   915
         End
         Begin VB.OptionButton OptnStyle 
            Appearance      =   0  'Flat
            Caption         =   "Style1"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   9
            Top             =   330
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Theme colors:"
         ForeColor       =   &H80000008&
         Height          =   4125
         Left            =   240
         TabIndex        =   1
         Top             =   765
         Width           =   1935
         Begin VB.OptionButton OptnTheme 
            Appearance      =   0  'Flat
            Caption         =   "AutoDetect"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   17
            Top             =   3540
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton OptnTheme 
            Appearance      =   0  'Flat
            Caption         =   "CustomTheme"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   7
            Top             =   3030
            Width           =   1395
         End
         Begin VB.OptionButton OptnTheme 
            Appearance      =   0  'Flat
            Caption         =   "Norton"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   6
            Top             =   2520
            Width           =   1395
         End
         Begin VB.OptionButton OptnTheme 
            Appearance      =   0  'Flat
            Caption         =   "Visual Studio"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   5
            Top             =   2010
            Width           =   1395
         End
         Begin VB.OptionButton OptnTheme 
            Appearance      =   0  'Flat
            Caption         =   "Olive"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   4
            Top             =   1500
            Width           =   1395
         End
         Begin VB.OptionButton OptnTheme 
            Appearance      =   0  'Flat
            Caption         =   "Silver"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   3
            Top             =   990
            Width           =   1395
         End
         Begin VB.OptionButton OptnTheme 
            Appearance      =   0  'Flat
            Caption         =   "Blue"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   480
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PicIniColor(0 To 2) As String

Private Sub CboIconSize_Click()
    
    jcForms1.IconSize = Val(CboIconSize.Text)

End Sub

Private Sub Check1_Click()
    
    Dim intI As Integer
    
    jcForms1.ChangeAllBackgrounds = Check1.Value
    
    For intI = 0 To 2
        Me.OptnBackcolorStyle(intI).Enabled = Not Me.OptnBackcolorStyle(intI).Enabled
    Next intI

End Sub

Private Sub Chkclose_Click()
    
    jcForms1.CloseButton = Chkclose.Value

End Sub

Private Sub ChkMax_Click()
    
    jcForms1.MaxButton = ChkMax.Value

End Sub

Private Sub ChkMin_Click()
    
    jcForms1.MinButton = ChkMin.Value

End Sub

Private Sub chkTitleBarShadow_Click()

    jcForms1.TitleBarShadow = chkTitleBarShadow.Value

End Sub

Private Sub Command1_Click()
    
    If Me.Caption = "jcForms v 1.0.5" Then
        Me.Caption = "Caption changed at runtime"
    Else
        Me.Caption = "jcForms v 1.0.5"
    End If
    
    jcForms1.Refresh

End Sub

Private Sub Command2_Click()
    
    If Command2.Caption = "Add menu item" Then
        Me.jcForms1.FormMenuAdd "MenuItem Added at runtime", True
        Command2.Caption = "Remove menu item"
    Else
        Command2.Caption = "Add menu item"
        Me.jcForms1.FormMenuRemove 13
    End If

End Sub

Private Sub Command3_Click()
    Form1.Show
End Sub

Private Sub Command4_Click()
    jcForms1.SetjcFormsMenuCaption jcRestore, "Restore changed"
    jcForms1.SetjcFormsMenuCaption jcMaximize, "Maximize changed"
    jcForms1.SetjcFormsMenuCaption jcMinimize, "Minimize changed"
    jcForms1.SetjcFormsMenuCaption jcClose, "Close changed"
    jcForms1.SetjcFormsMenuCaption jcAlwaysOnTop, "Always on Top changed"
End Sub

Private Sub Form_Load()
    
    PicIniColor(0) = "47"
    PicIniColor(1) = "25"
    PicIniColor(2) = "20"
    CboIconSize.Text = 16

    OptnTheme(6).Value = True
    OptnColors(0).Value = True
    jcForms1.CustomBackColor = RGB(240, 239, 176)

    ShpColor.Move 1 + Val(Left(PicIniColor(0), 1)) * 17, 1 + Val(Right(PicIniColor(0), 1)) * 17

    'you can add your own menu items
    jcForms1.FormMenuAdd "About ...", True
    jcForms1.FormMenuAdd "MenuItem Added 2", , , True
    jcForms1.FormMenuAdd "MenuItem Added 3"
    jcForms1.FormMenuAdd "MenuItem Added 4"
End Sub

Private Sub Option1_Click(Index As Integer)
    
    jcForms1.BorderStyle = Index

End Sub

Private Sub OptnBackcolorStyle_Click(Index As Integer)
    
    jcForms1.BackColorStyle = Index
    
    If Index = 2 Then
        OptnColors(0).Value = True
    End If
    
End Sub

Private Sub OptnColors_Click(Index As Integer)
    
    ShpColor.Move 1 + Val(Left(PicIniColor(Index), 1)) * 17, 1 + Val(Right(PicIniColor(Index), 1)) * 17
    
    If Index = 0 Then
        If OptnBackcolorStyle(2).Value = False Then
            OptnBackcolorStyle(2).Value = True
        End If
    End If

End Sub

Private Sub OptnStyle_Click(Index As Integer)
    
    jcForms1.Style = Index

End Sub

Private Sub OptnTheme_Click(Index As Integer)
    
    jcForms1.ThemeColor = Index
    
    If Index = 5 Then
        OptnColors(1).Value = True
    Else
        'OptnColors(0).Value = True
    End If

End Sub

Private Sub picColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim MyRow       As Single
    Dim MyCol       As Single
    
    MyRow = Fix((X - 2) / 17)
    MyCol = Fix((Y - 2) / 17)
    
    ShpColor.Move 1 + MyRow * 17, 1 + MyCol * 17
    
    If OptnColors(0).Value = True Then
        jcForms1.CustomBackColor = picColor.POINT(3 + MyRow * 17, 3 + MyCol * 17)
        PicIniColor(0) = MyRow & MyCol
    ElseIf OptnColors(1).Value = True Then
        jcForms1.ColorFrom = picColor.POINT(3 + MyRow * 17, 3 + MyCol * 17)
        PicIniColor(1) = MyRow & MyCol
    ElseIf OptnColors(2).Value = True Then
        jcForms1.ColorTo = picColor.POINT(3 + MyRow * 17, 3 + MyCol * 17)
        PicIniColor(2) = MyRow & MyCol
    End If
    
End Sub

Private Sub jcForms1_MenuItemSelected(MenuItem As Integer, MenuCaption As String)
    
    'You can put here your code to handle added menu items
    
    If MenuItem = 8 Then
        frmAbout.Show
    Else
        MsgBox "MenuItem = " & MenuItem & " clicked, MenuCaption = " & MenuCaption
    End If
End Sub

