VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Color Code Generator"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Color Combination"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2580
      Left            =   2700
      TabIndex        =   5
      Top             =   135
      Width           =   3120
      Begin VB.HScrollBar hscBlue 
         Height          =   240
         Left            =   750
         Max             =   255
         TabIndex        =   17
         Top             =   2115
         Width           =   2175
      End
      Begin VB.HScrollBar hscGreen 
         Height          =   240
         Left            =   750
         Max             =   255
         TabIndex        =   16
         Top             =   1800
         Width           =   2175
      End
      Begin VB.HScrollBar hscRed 
         Height          =   240
         Left            =   750
         Max             =   255
         TabIndex        =   15
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Label lblPBlue 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   210
         Left            =   2520
         TabIndex        =   29
         Top             =   1095
         Width           =   420
      End
      Begin VB.Label lblPGreen 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   210
         Left            =   2520
         TabIndex        =   28
         Top             =   840
         Width           =   420
      End
      Begin VB.Label lblPRed 
         AutoSize        =   -1  'True
         Caption         =   "100%"
         Height          =   210
         Left            =   2520
         TabIndex        =   27
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lblBlue 
         BackColor       =   &H8000000D&
         Height          =   210
         Left            =   2130
         TabIndex        =   26
         Top             =   1095
         Width           =   270
      End
      Begin VB.Label lblGreen 
         BackColor       =   &H8000000D&
         Height          =   210
         Left            =   2130
         TabIndex        =   25
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblRed 
         BackColor       =   &H8000000D&
         Height          =   210
         Left            =   2130
         TabIndex        =   24
         Top             =   585
         Width           =   270
      End
      Begin VB.Label lblHRed 
         AutoSize        =   -1  'True
         Caption         =   "255"
         Height          =   210
         Left            =   1575
         TabIndex        =   23
         Top             =   585
         Width           =   270
      End
      Begin VB.Label lblHGreen 
         AutoSize        =   -1  'True
         Caption         =   "255"
         Height          =   210
         Left            =   1575
         TabIndex        =   22
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lblHBlue 
         AutoSize        =   -1  'True
         Caption         =   "255"
         Height          =   210
         Left            =   1575
         TabIndex        =   21
         Top             =   1095
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   1575
         TabIndex        =   20
         Top             =   315
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ASCII"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   765
         TabIndex        =   19
         Top             =   315
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   14
         Top             =   2115
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   13
         Top             =   1815
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   12
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label lblABlue 
         AutoSize        =   -1  'True
         Caption         =   "255"
         Height          =   210
         Left            =   825
         TabIndex        =   11
         Top             =   1110
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   10
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label lblAGreen 
         AutoSize        =   -1  'True
         Caption         =   "255"
         Height          =   210
         Left            =   825
         TabIndex        =   9
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   8
         Top             =   840
         Width           =   510
      End
      Begin VB.Label lblARed 
         AutoSize        =   -1  'True
         Caption         =   "255"
         Height          =   210
         Left            =   825
         TabIndex        =   7
         Top             =   600
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   6
         Top             =   600
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1905
      Left            =   60
      TabIndex        =   3
      Top             =   810
      Width           =   2565
      Begin VB.Label lblPreview 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   105
         TabIndex        =   4
         Top             =   270
         Width           =   2340
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   2565
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   285
         Left            =   1770
         TabIndex        =   18
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lblHexCode 
         AutoSize        =   -1  'True
         Caption         =   "#000000"
         Height          =   210
         Left            =   1020
         TabIndex        =   2
         Top             =   210
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hex Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   135
         TabIndex        =   1
         Top             =   210
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
' Project   :   ColorCode
' Procedure :   frmMain
' Author    :   Mihir Solanki
' DateTime  :   28/03/2003
' Version   :   1.1
' Purpose   :   Create Hex Color Code for HTML Pages
' Copyright :   Copyright © 2000-2005 Mihir Solanki,UK
' Contact   :   mihir_solanki@lycos.co.uk
'**************************************************************************************
'
Option Explicit

Private Sub cmdCopy_Click()
    'Copy to clipboard
    Clipboard.SetText lblHexCode
End Sub

Private Sub Form_Load()
    'Center Me
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    'Default Value
    hscBlue.Value = 100
    hscGreen.Value = 100
    hscRed.Value = 100
    
    'Display Changed Values
    Value_Changed hscRed.Value, hscGreen.Value, hscBlue.Value
    
End Sub

Private Sub hscBlue_Change()

    'Display Changed Values
    Value_Changed hscRed.Value, hscGreen.Value, hscBlue.Value
End Sub

Private Sub hscGreen_Change()
        
    'Display Changed Values
    Value_Changed hscRed.Value, hscGreen.Value, hscBlue.Value
End Sub

Private Sub hscRed_Change()
    
    'Display Changed Values
    Value_Changed hscRed.Value, hscGreen.Value, hscBlue.Value
End Sub
Private Sub Value_Changed(ired As Integer, igreen As Integer, iblue As Integer)
    
    'Preview in Label Control
    lblPreview.BackColor = RGB(ired, igreen, iblue)
    
    'ASCII Values
    lblARed = ired
    lblAGreen = igreen
    lblABlue = iblue
    
    'HEX Values
    'Bug Fixed : For Value 10 to 15.
    
    lblHRed = Format(Hex(ired), "00")
    lblHRed = IIf(Len(lblHRed) = 1, "0" & CStr(lblHRed), lblHRed)
    lblHGreen = Format(Hex(igreen), "00")
    lblHGreen = IIf(Len(lblHGreen) = 1, "0" & lblHGreen, lblHGreen)
    lblHBlue = Format(Hex(iblue), "00")
    lblHBlue = IIf(Len(lblHBlue) = 1, "0" & lblHBlue, lblHBlue)
    
    'Complete HEX Code
    lblHexCode = "#" & lblHRed & lblHGreen & lblHBlue

    'Display Individual Color
    lblRed.BackColor = RGB(ired, 0, 0)
    lblGreen.BackColor = RGB(0, igreen, 0)
    lblBlue.BackColor = RGB(0, 0, iblue)
    
    'Percentage of Individual Color
    lblPRed = Format(ired / 256, "00%")     ' i.e 255=100%
    lblPGreen = Format(igreen / 256, "00%")
    lblPBlue = Format(iblue / 256, "00%")
    
End Sub
