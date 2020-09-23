VERSION 5.00
Begin VB.Form frmUnitsConverter 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Units converter version 1.1 (© S.J.G. Strijk)"
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmToMetric.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   725
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFF00&
      Caption         =   "Torque"
      Height          =   1050
      Left            =   120
      TabIndex        =   80
      Top             =   7320
      Width           =   5265
      Begin VB.ComboBox cboOPTorque 
         Height          =   315
         Left            =   3720
         TabIndex        =   84
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtOPTorque 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2760
         TabIndex        =   83
         Top             =   600
         Width           =   1005
      End
      Begin VB.ComboBox cboIPTorque 
         Height          =   315
         Left            =   1080
         TabIndex        =   82
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtIPTorque 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   81
         ToolTipText     =   "Input data 10000max."
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   3960
         TabIndex        =   89
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2760
         TabIndex        =   88
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1320
         TabIndex        =   87
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   85
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFF00&
      Caption         =   "Outgassing rates"
      Height          =   1050
      Left            =   120
      TabIndex        =   55
      Top             =   9720
      Width           =   5535
      Begin VB.ComboBox cboOGRunitOut 
         Height          =   315
         Left            =   3840
         TabIndex        =   65
         Text            =   "Combo1"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtOGRout 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2880
         MaxLength       =   9
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cboOGRunitIn 
         Height          =   315
         Left            =   1080
         TabIndex        =   62
         Text            =   "Combo1"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtOGRin 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   9
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   4080
         TabIndex        =   73
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1440
         TabIndex        =   72
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2880
         TabIndex        =   67
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2660
         TabIndex        =   64
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFF00&
      Caption         =   "Pressure"
      Height          =   1050
      Left            =   120
      TabIndex        =   54
      Top             =   8520
      Width           =   5295
      Begin VB.ComboBox cboOutPressure 
         Height          =   315
         Left            =   3720
         TabIndex        =   60
         Text            =   "Combo1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtOutPressure 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2760
         MaxLength       =   9
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cboInPressure 
         Height          =   315
         Left            =   1080
         TabIndex        =   57
         Text            =   "Combo1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtInPressure 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   56
         Text            =   "Text1"
         ToolTipText     =   "Enter value <100000"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   3960
         TabIndex        =   71
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1320
         TabIndex        =   70
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2760
         TabIndex        =   66
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2540
         TabIndex        =   58
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Magnetic Fieldstrength"
      Height          =   1050
      Left            =   120
      TabIndex        =   45
      Top             =   6120
      Width           =   4545
      Begin VB.ComboBox cboOPMagneticfield 
         Height          =   315
         Left            =   3360
         TabIndex        =   49
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtOPMagneticfield 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cboIPMagneticfield 
         Height          =   315
         Left            =   1080
         TabIndex        =   47
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtIPMagneticfield 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   46
         ToolTipText     =   "Input data 10000max."
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   74
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   3480
         TabIndex        =   53
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2400
         TabIndex        =   52
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1080
         TabIndex        =   51
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Temperature"
      Height          =   1050
      Left            =   120
      TabIndex        =   36
      Top             =   4920
      Width           =   4545
      Begin VB.TextBox txtIPtemperature 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   40
         ToolTipText     =   "Input data 10000max."
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cboIPtemperature 
         Height          =   315
         Left            =   850
         TabIndex        =   39
         Top             =   600
         Width           =   1250
      End
      Begin VB.TextBox txtOPtemperature 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   38
         Top             =   600
         Width           =   765
      End
      Begin VB.ComboBox cboOPtemperature 
         Height          =   315
         Left            =   3160
         TabIndex        =   37
         Top             =   600
         Width           =   1250
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   79
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2400
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   3480
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Mass"
      Height          =   1050
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   4545
      Begin VB.ComboBox cboOPMass 
         Height          =   315
         Left            =   3360
         TabIndex        =   33
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtOPMass 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   32
         Top             =   600
         Width           =   1000
      End
      Begin VB.ComboBox cboIPMass 
         Height          =   315
         Left            =   1080
         TabIndex        =   30
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtIPMass 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   28
         ToolTipText     =   "Input data 10000000 max."
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   78
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   3480
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2400
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Volume"
      Height          =   1050
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4545
      Begin VB.ComboBox cboOPVolume 
         Height          =   315
         Left            =   3360
         TabIndex        =   26
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtOPVolume 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   24
         Top             =   600
         Width           =   1000
      End
      Begin VB.ComboBox cboIPVolume 
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtIPVolume 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   21
         ToolTipText     =   "Input data 10000000 max."
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   77
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   3480
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Area"
      Height          =   1050
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4545
      Begin VB.ComboBox cboOPArea 
         Height          =   315
         Left            =   3360
         TabIndex        =   15
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtOPArea 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   14
         Top             =   600
         Width           =   1000
      End
      Begin VB.ComboBox cboIPArea 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtIPArea 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   12
         ToolTipText     =   "Input data 10000000 max."
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   76
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Length"
      Height          =   1050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4545
      Begin VB.ComboBox cboOPLength 
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtOPLength 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Top             =   600
         Width           =   1000
      End
      Begin VB.ComboBox cboIPLength 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   1020
      End
      Begin VB.TextBox txtIPLength 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   8
         TabIndex        =   4
         ToolTipText     =   "Input data 10000000 max."
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFF00&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   75
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "OP Value"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "IP Value"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "Unit"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmUnitsConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This units calculator was written by S.J.G. Strijk, March2002 - Nov2009
'issue 1.0  Aug.2003
'issue 1.1  Nov.2009 Torque, Pressure & Outgassing conversion added
'
'--below the list of subs used in the form---------
'--highlighting the sub name and using the find button to go there quickly  :) ---------------
'
'Private Sub Close_Click()
'Public Function GetInfo(ByVal lInfo As Long) As String
'Private Sub Form_Load()
'Private Sub Form_Keypress(KeyAscii As Integer)
'-----Length conversions-------
'Private Sub cboIPLength_Click()
'Private Sub cboIPLength_Change()
'Private Sub cboOPLength_Click()
'Private Sub cboOPLength_Change()
'Private Sub txtIPLength_Change()
'Private Sub OPLength(temp As Double)
'Private Sub txtIPLength_KeyPress(KeyAscii As Integer)
'Private Sub txtOPLength_KeyPress(KeyAscii As Integer)
'Private Sub cboIPLength_KeyPress(KeyAscii As Integer)
'Private Sub cboOPLength_KeyPress(KeyAscii As Integer)
'----Area conversions-------------
'Private Sub cboIPArea_Click()
'Private Sub cboIPArea_Change()
'Private Sub cboOPArea_Click()
'Private Sub cboOPArea_Change()
'Private Sub txtIPArea_Change()
'Private Sub OPArea(temp As Double)
'Private Sub txtIPArea_KeyPress(KeyAscii As Integer)
'Private Sub txtOPArea_KeyPress(KeyAscii As Integer)
'Private Sub cboIPArea_KeyPress(KeyAscii As Integer)
'Private Sub cboOPArea_KeyPress(KeyAscii As Integer)
'----Volume conversions ------
'Private Sub cboIPVolume_Click()
'Private Sub cboIPVolume_Change()
'Private Sub cboOPVolume_Click()
'Private Sub cboOPVolume_Change()
'Private Sub txtIPVolume_Change()
'Private Sub OPVolume(temp As Double)
'Private Sub txtIPVolume_KeyPress(KeyAscii As Integer)
'Private Sub txtOPVolume_KeyPress(KeyAscii As Integer)
'Private Sub cboIPVolume_KeyPress(KeyAscii As Integer)
'Private Sub cboOPVolume_KeyPress(KeyAscii As Integer)
'---Mass conversions--------
'Private Sub cboIPMass_Click()
'Private Sub cboIPMass_Change()
'Private Sub cboOPMass_Click()
'Private Sub cboOPMass_Change()
'Private Sub txtIPMass_Change()
'Private Sub OPMass(temp As Double)
'Private Sub txtIPMass_KeyPress(KeyAscii As Integer)
'Private Sub txtOPMass_KeyPress(KeyAscii As Integer)
'Private Sub cboIPMass_KeyPress(KeyAscii As Integer)
'Private Sub cboOPMass_KeyPress(KeyAscii As Integer)
'---Torque conversions-----
'Private Sub cboIPTorque_Click()
'Private Sub cboIPTorque_Change()
'Private Sub cboOPTorque_Click()
'Private Sub cboOPTorque_Change()
'Private Sub txtIPTorque_Change()
'Private Sub OPTorque(temp As Double)
'Private Sub txtIPTorque_KeyPress(KeyAscii As Integer)
'Private Sub txtOPTorque_KeyPress(KeyAscii As Integer)
'Private Sub cboIPTorque_KeyPress(KeyAscii As Integer)
'Private Sub cboOPTorque_KeyPress(KeyAscii As Integer)
'----Temperature conversions --------
'Private Sub cboIPtemperature_Click()
'Private Sub cboIPtemperature_Change()
'Private Sub cboOPtemperature_Click()
'Private Sub cboOPtemperature_Change()
'Private Sub txtIPtemperature_Change()
'Private Sub OPtemperature(temp As Double)
'Private Sub txtOPtemperature_KeyPress(KeyAscii As Integer)
'Private Sub cboIPtemperature_KeyPress(KeyAscii As Integer)
'Private Sub cboOPtemperature_KeyPress(KeyAscii As Integer)
'Private Sub txtIPtemperature_KeyPress(KeyAscii As Integer)
'----Magnetic field conversions-----------
'Private Sub cboIPMagneticfield_Click()
'Private Sub cboIPMagneticfield_Change()
'Private Sub cboOPMagneticfield_Click()
'Private Sub cboOPMagneticfield_Change()
'Private Sub txtIPMagneticfield_Change()
'Private Sub OPMagneticfield(temp As Double) 'passing the reference fieldstrength
'Private Sub txtOPMagneticfield_KeyPress(KeyAscii As Integer)
'Private Sub cboIPMagneticfield_KeyPress(KeyAscii As Integer)
'Private Sub cboOPMagneticfield_KeyPress(KeyAscii As Integer)
'Private Sub txtIPMagneticfield_KeyPress(KeyAscii As Integer)
'
'Private Sub DataEntry(KeyAscii As Integer, IPtxt As String, signflg As Boolean)
'
' for info 79.58A/m = 1 Oersted

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'----Variables-------
Public StandaardArea As Double, StandaardVolume As Double, StandaardMass As Double
Public RefTemperature As Double, RefMagneticField As Double, StandaardLengte As Double
Public StandaardPressure As Double, RefTorque As Double
'----Strings---------
Public strLocal As String, display As String
'-----Flags---------
Public flgLocal As Boolean

Const LOCALE_USER_DEFAULT = &H400
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

Dim IPtxt As String, temp As Double, temp1 As Double

Private Sub Close_Click()
  Unload Me
End Sub

Public Function GetInfo(ByVal lInfo As Long) As String
Dim Buffer As String, Ret As String
'--sub to find the local decimal to be a comma (Europe) or period (USA,UK a.o.)-----
Buffer = String$(256, 0)
  Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    If Ret > 0 Then
        GetInfo = Left$(Buffer, Ret - 1)
    Else
        GetInfo = ""
    End If
End Function

Private Sub Form_Load()
strLocal = GetInfo(&H16)   ' ask for the local decimal separator i.e. "," or "."
If strLocal = "," Then flgLocal = False
If strLocal = "." Or strLocal = "" Then flgLocal = True

cboIPLength.Clear
cboIPLength.AddItem "micron"
cboIPLength.AddItem "mm"
cboIPLength.AddItem "cm"
cboIPLength.AddItem "dm"
cboIPLength.AddItem "m"
cboIPLength.AddItem "Km"
cboIPLength.AddItem "Mil"
cboIPLength.AddItem "Inch"
cboIPLength.AddItem "uInch"
cboIPLength.AddItem "Foot"
cboIPLength.AddItem "Yard"
cboIPLength.AddItem "Mile"
cboIPLength.Text = cboIPLength.List(7)
cboIPLength.Visible = True

cboOPLength.Clear
cboOPLength.AddItem "micron"
cboOPLength.AddItem "mm"
cboOPLength.AddItem "cm"
cboOPLength.AddItem "dm"
cboOPLength.AddItem "m"
cboOPLength.AddItem "Km"
cboOPLength.AddItem "Mil"
cboOPLength.AddItem "Inch"
cboOPLength.AddItem "uInch"
cboOPLength.AddItem "Foot"
cboOPLength.AddItem "Yard"
cboOPLength.AddItem "Mile"
cboOPLength.Text = cboOPLength.List(1)
cboOPLength.Visible = True

cboIPArea.Clear
cboIPArea.AddItem "cm^2"
cboIPArea.AddItem "dm^2"
cboIPArea.AddItem "m^2"
cboIPArea.AddItem "ares"
cboIPArea.AddItem "hectare"
cboIPArea.AddItem "Km^2"
cboIPArea.AddItem "acre"
cboIPArea.AddItem "inch^2"
cboIPArea.AddItem "feet^2"
cboIPArea.AddItem "yard^2"
cboIPArea.AddItem "Mile^2"
cboIPArea.Text = cboIPArea.List(7)
cboIPArea.Visible = True

cboOPArea.Clear
cboOPArea.AddItem "cm^2"
cboOPArea.AddItem "dm^2"
cboOPArea.AddItem "m^2"
cboOPArea.AddItem "ares"
cboOPArea.AddItem "hectare"
cboOPArea.AddItem "Km^2"
cboOPArea.AddItem "acre"
cboOPArea.AddItem "inch^2"
cboOPArea.AddItem "feet^2"
cboOPArea.AddItem "yard^2"
cboOPArea.AddItem "Mile^2"
cboOPArea.Text = cboOPArea.List(1)
cboOPArea.Visible = True

cboIPVolume.Clear
cboIPVolume.AddItem "cm^3"
cboIPVolume.AddItem "cc = ml"
cboIPVolume.AddItem "cl"
cboIPVolume.AddItem "Litre"
cboIPVolume.AddItem "m^3"
cboIPVolume.AddItem "inch^3"
cboIPVolume.AddItem "feet^3"
cboIPVolume.AddItem "yard^3"
cboIPVolume.AddItem "UK gal."
cboIPVolume.AddItem "US gal."
cboIPVolume.Text = cboIPVolume.List(5)
cboIPVolume.Visible = True

cboOPVolume.Clear
cboOPVolume.AddItem "cm^3"
cboOPVolume.AddItem "cc = ml"
cboOPVolume.AddItem "cl"
cboOPVolume.AddItem "Litre"
cboOPVolume.AddItem "m^3"
cboOPVolume.AddItem "inch^3"
cboOPVolume.AddItem "feet^3"
cboOPVolume.AddItem "yard^3"
cboOPVolume.AddItem "UK gal."
cboOPVolume.AddItem "US gal."
cboOPVolume.Text = cboOPVolume.List(8)
cboOPVolume.Visible = True

cboIPMass.Clear
cboIPMass.AddItem "mg"
cboIPMass.AddItem "g"
cboIPMass.AddItem "Kg"
cboIPMass.AddItem "m Tonne"
cboIPMass.AddItem "grain"
cboIPMass.AddItem "oz"
cboIPMass.AddItem "lb"
cboIPMass.Text = cboIPMass.List(6)
cboIPMass.Visible = True

cboOPMass.Clear
cboOPMass.AddItem "mg"
cboOPMass.AddItem "g"
cboOPMass.AddItem "Kg"
cboOPMass.AddItem "m Tonne"
cboOPMass.AddItem "grain"
cboOPMass.AddItem "oz"
cboOPMass.AddItem "lb"
cboOPMass.Text = cboOPMass.List(2)
cboOPMass.Visible = True

cboIPtemperature.Clear
cboIPtemperature.AddItem "Fahrenheit"
cboIPtemperature.AddItem "Celsius"
cboIPtemperature.AddItem "Kelvin"
cboIPtemperature.Text = cboIPtemperature.List(0)

cboOPtemperature.Clear
cboOPtemperature.AddItem "Fahrenheit"
cboOPtemperature.AddItem "Celsius"
cboOPtemperature.AddItem "Kelvin"
cboOPtemperature.Text = cboOPtemperature.List(1)

cboIPTorque.Clear
cboIPTorque.AddItem "gcm"
cboIPTorque.AddItem "kgcm"
cboIPTorque.AddItem "kgm"
cboIPTorque.AddItem "Ncm"
cboIPTorque.AddItem "Nm"
cboIPTorque.AddItem "OunceInches"
cboIPTorque.AddItem "PoundInches"
cboIPTorque.AddItem "Poundfeet"
cboIPTorque.Text = cboIPTorque.List(5)

cboOPTorque.Clear
cboOPTorque.AddItem "gcm"
cboOPTorque.AddItem "kgcm"
cboOPTorque.AddItem "kgm"
cboOPTorque.AddItem "Ncm"
cboOPTorque.AddItem "Nm"
cboOPTorque.AddItem "OunceInches"
cboOPTorque.AddItem "PoundInches"
cboOPTorque.AddItem "Poundfeet"
cboOPTorque.Text = cboOPTorque.List(3)

cboIPMagneticfield.Clear
cboIPMagneticfield.AddItem "Gauss"
cboIPMagneticfield.AddItem "Tesla"
cboIPMagneticfield.AddItem "Oersted"
cboIPMagneticfield.AddItem "A/m"
cboIPMagneticfield.Text = cboIPMagneticfield.List(0)

cboOPMagneticfield.Clear
cboOPMagneticfield.AddItem "Gauss"
cboOPMagneticfield.AddItem "Tesla"
cboOPMagneticfield.AddItem "Oersted"
cboOPMagneticfield.AddItem "A/m"
cboOPMagneticfield.Text = cboOPMagneticfield.List(1)

cboInPressure.Clear
cboInPressure.AddItem "Atmosphere"
cboInPressure.AddItem "Torr"
cboInPressure.AddItem "Bar"
cboInPressure.AddItem "mBar"
cboInPressure.AddItem "Pascal"
cboInPressure.AddItem "hPa"
cboInPressure.AddItem "kPa"
cboInPressure.Text = cboInPressure.List(0)
cboInPressure.Visible = True

cboOutPressure.Clear
cboOutPressure.AddItem "Atmosphere"
cboOutPressure.AddItem "Torr"
cboOutPressure.AddItem "Bar"
cboOutPressure.AddItem "mBar"
cboOutPressure.AddItem "Pascal"
cboOutPressure.AddItem "hPa"
cboOutPressure.AddItem "kPa"
cboOutPressure.Text = cboOutPressure.List(3)
cboOutPressure.Visible = True

cboOGRunitIn.Clear    'Outgassing rate
cboOGRunitIn.AddItem "Pa m/s"
cboOGRunitIn.AddItem "mBar-l/s-cm^2"
cboOGRunitIn.AddItem "Torr-l/s-cm^2"
cboOGRunitIn.Text = cboOGRunitIn.List(0)
cboOGRunitIn.Visible = True

cboOGRunitOut.Clear
cboOGRunitOut.AddItem "Pa m/s"
cboOGRunitOut.AddItem "mBar-l/s-cm^2"
cboOGRunitOut.AddItem "Torr-l/s-cm^2"
cboOGRunitOut.Text = cboOGRunitOut.List(1)
cboOGRunitOut.Visible = True

txtInPressure.Text = 1
txtOGRin.Text = Format(0.000015, "scientific")
txtIPLength.Text = Format(1, "##0.0#")
txtIPArea.Text = Format(1, "##0.0#")
txtIPVolume.Text = Format(1, "##0.0#")
txtIPTorque.Text = Format(1, "##0.0#")
txtIPMass.Text = Format(1, "##0.0#")
txtIPtemperature.Text = Format(32, "##0.0#")
txtIPMagneticfield.Text = Format(1, "##0.0##")

End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc("Q")
    Unload Me
  Case Asc("q")
    Unload Me
  End Select
End Sub

Private Sub DefineDisplayFormat(OPvalue As Double, display As String)
  If OPvalue <= 0.001 Then display = "scientific"
  If OPvalue > 0.000999 And OPvalue <= 0.999999 Then display = "0.######"
  If OPvalue > 0.999999 And OPvalue <= 9.999999 Then display = "0.######"
  If OPvalue > 9.999999 And OPvalue <= 99.99999 Then display = "#0.#####"
  If OPvalue > 99.99999 And OPvalue <= 999.9999 Then display = "##0.####"
  If OPvalue > 999.9999 And OPvalue <= 9999.999 Then display = "###0.###"
  If OPvalue > 9999.999 And OPvalue <= 99999.99 Then display = "####0.##"
  If OPvalue > 99999.99 And OPvalue <= 999999.9 Then display = "#####0.#"
  If OPvalue > 999999.9 And OPvalue <= 9999999 Then display = "########"
  If OPvalue > 9999999 Then display = "scientific"
End Sub

'--------------------------------------Length conversions --------------

Private Sub txtIPLength_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtIPLength, False
End Sub

Private Sub txtOPLength_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPLength_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOPLength_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

' below 4 subs ensures an up-to-date calculation being performed
Private Sub cboIPLength_Click()
  txtIPLength_Change
End Sub

Private Sub cboIPLength_Change()
  txtIPLength_Change
End Sub

Private Sub cboOPLength_Click()
  txtIPLength_Change
End Sub

Private Sub cboOPLength_Change()
  txtIPLength_Change
End Sub

Private Sub txtIPLength_Change()
temp = Val(Replace(txtIPLength, ",", "."))
If temp > 10000000 Then
  temp = 10000000
  txtIPLength.Text = temp
End If
Select Case cboIPLength.Text
  Case "micron"
    StandaardLengte = temp / 1000000  '--normalise all inputs to meter
  Case "mm"
    StandaardLengte = temp / 1000
  Case "cm"
    StandaardLengte = temp / 100
  Case "dm"
    StandaardLengte = temp / 10
  Case "m"
    StandaardLengte = temp
  Case "Km"
    StandaardLengte = temp * 1000
  Case "Mil"
    StandaardLengte = temp * 0.0254 * 0.001
  Case "Inch"
    StandaardLengte = temp * 0.0254
  Case "uInch"
    StandaardLengte = temp * 0.0254 * 0.000001
  Case "Foot"
    StandaardLengte = temp * 0.3048
  Case "Yard"
    StandaardLengte = temp * 0.9144
  Case "Mile"
    StandaardLengte = temp * 1609.34
End Select
Call OPLength(StandaardLengte)  'convert to m
End Sub

Private Sub OPLength(temp As Double)
Dim OPvalue As Double
Select Case cboOPLength.Text
  Case "micron"
    OPvalue = temp * 1000000  '--calc the meter to the output value
  Case "mm"
    OPvalue = temp * 1000
  Case "cm"
    OPvalue = temp * 100
  Case "dm"
    OPvalue = temp * 10
  Case "m"
    OPvalue = temp
  Case "Km"
    OPvalue = temp / 1000
  Case "Mil"
    OPvalue = temp / (0.0254 * 0.001)
  Case "Inch"
    OPvalue = temp / 0.0254
  Case "uInch"
    OPvalue = temp / (0.0254 * 0.000001)
  Case "Foot"
    OPvalue = temp / 0.3048
  Case "Yard"
    OPvalue = temp / 0.9144
  Case "Mile"
    OPvalue = temp / 1609.34
End Select
DefineDisplayFormat OPvalue, display
txtOPLength.Text = Format(OPvalue, display)
End Sub
  
'---------------------------------Area calculations---------------------
Private Sub txtIPArea_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtIPArea, False
End Sub

Private Sub txtOPArea_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPArea_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOPArea_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPArea_Click()
  txtIPArea_Change
End Sub

Private Sub cboIPArea_Change()
  txtIPArea_Change
End Sub

Private Sub cboOPArea_Click()
  txtIPArea_Change
End Sub

Private Sub cboOPArea_Change()
  txtIPArea_Change
End Sub

Private Sub txtIPArea_Change()
Dim temp As Double
temp = Val(Replace(txtIPArea, ",", "."))
If temp > 10000000 Then
  temp = 10000000
  txtIPArea.Text = temp
End If
Select Case cboIPArea.Text
  Case "cm^2"
    StandaardArea = temp * 0.0001 '--normalise all inputs to sq.meter
  Case "dm^2"
    StandaardArea = temp * 0.01
  Case "m^2"
    StandaardArea = temp
  Case "ares"
    StandaardArea = temp * 100
  Case "acre"
    StandaardArea = temp * 4046.8564
  Case "hectare"
    StandaardArea = temp * 10000
  Case "Km^2"
    StandaardArea = temp * 1000000#
  Case "inch^2"
    StandaardArea = temp * 0.00064516
  Case "feet^2"
    StandaardArea = temp * 0.09290304
  Case "yard^2"
    StandaardArea = temp * 0.83612736
  Case "Mile^2"
    StandaardArea = temp * 2589988.110336
End Select
Call OPArea(StandaardArea)
End Sub

Private Sub OPArea(temp As Double)
Dim OPvalue As Double
Select Case cboOPArea.Text
  Case "cm^2"
    OPvalue = temp / 0.0001 '--normalise all inputs to sq.meter
  Case "dm^2"
    OPvalue = temp / 0.01
  Case "m^2"
    OPvalue = temp
  Case "ares"
    OPvalue = temp / 100
  Case "acre"
    OPvalue = temp / 4046.8564
  Case "hectare"
    OPvalue = temp / 10000
  Case "Km^2"
    OPvalue = temp / 1000000#
  Case "inch^2"
    OPvalue = temp / 0.00064516
  Case "feet^2"
    OPvalue = temp / 0.09290304
  Case "yard^2"
    OPvalue = temp / 0.83612736
  Case "Mile^2"
    OPvalue = temp / 2589988.110336
End Select
DefineDisplayFormat OPvalue, display
txtOPArea.Text = Format(OPvalue, display)
End Sub

'--------------------------------Volume calculations---------------------

Private Sub txtIPVolume_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtIPVolume, False
End Sub

Private Sub txtOPVolume_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPVolume_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOPVolume_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPVolume_Click()
  txtIPVolume_Change
End Sub

Private Sub cboIPVolume_Change()
  txtIPVolume_Change
End Sub

Private Sub cboOPVolume_Click()
  txtIPVolume_Change
End Sub

Private Sub cboOPVolume_Change()
  txtIPVolume_Change
End Sub

Private Sub txtIPVolume_Change()
Dim temp As Double
temp = Val(Replace(txtIPVolume, ",", "."))
If temp > 10000000 Then
  temp = 10000000
  txtIPVolume.Text = temp
  Beep
End If
Select Case cboIPVolume.Text
  Case "cm^3"
    StandaardVolume = temp * 0.001    '--normalise all inputs to litre
  Case "cc = ml"
    StandaardVolume = temp * 0.001
  Case "cl"
    StandaardVolume = temp * 0.01
  Case "Litre"
    StandaardVolume = temp
  Case "m^3"
    StandaardVolume = temp * 1000
  Case "inch^3"
    StandaardVolume = temp * 0.016387064
  Case "feet^3"
    StandaardVolume = temp * 28.316846592
  Case "yard^3"
    StandaardVolume = temp * 764.554857984
  Case "UK gal."
    StandaardVolume = temp * 4.5459631
  Case "US gal."
    StandaardVolume = temp * 3.7853060057
End Select
Call OPVolume(StandaardVolume)
End Sub

Private Sub OPVolume(temp As Double)
Dim OPvalue As Double
Select Case cboOPVolume.Text
  Case "cm^3"
    OPvalue = temp / 0.001    '--normalise all inputs to litre
  Case "cc = ml"
    OPvalue = temp / 0.001
  Case "cl"
    OPvalue = temp / 0.01
  Case "Litre"
    OPvalue = temp
  Case "m^3"
    OPvalue = temp / 1000
  Case "inch^3"
   OPvalue = temp / 0.016387064
  Case "feet^3"
    OPvalue = temp / 28.316846592
  Case "yard^3"
    OPvalue = temp / 764.554857984
  Case "UK gal."
    OPvalue = temp / 4.5459631
  Case "US gal."
   OPvalue = temp / 3.7853060057
End Select
DefineDisplayFormat OPvalue, display
txtOPVolume.Text = Format(OPvalue, display)
End Sub


'--------------------------------Mass calculations---------------------

Private Sub txtIPMass_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtIPMass, False
End Sub

Private Sub txtOPMass_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPMass_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOPMass_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPMass_Click()
  txtIPMass_Change
End Sub

Private Sub cboIPMass_Change()
  txtIPMass_Change
End Sub

Private Sub cboOPMass_Click()
  txtIPMass_Change
End Sub

Private Sub cboOPMass_Change()
  txtIPMass_Change
End Sub

Private Sub txtIPMass_Change()
Dim temp As Double
temp = Val(Replace(txtIPMass, ",", "."))
If temp > 10000000 Then
  temp = 10000000
  txtIPMass.Text = temp
  Beep
End If
Select Case cboIPMass.Text
  Case "mg"
    StandaardMass = temp * 0.000001      '--normalise all inputs to Kg
  Case "g"
    StandaardMass = temp * 0.001
  Case "Kg"
    StandaardMass = temp
  Case "m Tonne"
    StandaardMass = temp * 1000
  Case "grain"
    StandaardMass = temp * 0.0000647989
  Case "oz"
    StandaardMass = temp * 0.0283495231
  Case "lb"
    StandaardMass = temp * 0.45359237
End Select
Call OPMass(StandaardMass)
End Sub

Private Sub OPMass(temp As Double)
Dim OPvalue As Double
Select Case cboOPMass.Text
  Case "mg"
    OPvalue = temp / 0.000001      '--normalise all inputs to Kg
  Case "g"
    OPvalue = temp / 0.001
  Case "Kg"
    OPvalue = temp
  Case "m Tonne"
    OPvalue = temp / 1000
  Case "grain"
    OPvalue = temp / 0.0000647989
  Case "oz"
   OPvalue = temp / 0.0283495231
  Case "lb"
    OPvalue = temp / 0.45359237
End Select
DefineDisplayFormat OPvalue, display
txtOPMass.Text = Format(OPvalue, display)
End Sub

'----------------------------Temperature calculations---------------------
Private Sub txtIPtemperature_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtIPtemperature, True    'accept neg temperatures
End Sub

Private Sub txtOPtemperature_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPtemperature_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOPtemperature_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPtemperature_Click()
  txtIPtemperature_Change
End Sub

Private Sub cboIPtemperature_Change()
  txtIPtemperature_Change
End Sub

Private Sub cboOPtemperature_Click()
  txtIPtemperature_Change
End Sub

Private Sub cboOPtemperature_Change()
  txtIPtemperature_Change
End Sub

Private Sub txtIPtemperature_Change()
Dim temp As Double
temp = Val(Replace(txtIPtemperature, ",", "."))
If temp > 10000 Then
  temp = 10000
  txtIPtemperature.Text = 10000
End If
Select Case cboIPtemperature.Text
  Case "Fahrenheit"
    RefTemperature = 5 * (temp - 32) / 9 + 273.15 '--normalise all inputs to Kelvin
    If RefTemperature < 0 Then
      RefTemperature = 0
      txtIPtemperature.Text = -459.67
    End If
  Case "Celsius"
    RefTemperature = temp + 273.15
    If RefTemperature < 0 Then
      RefTemperature = 0
      txtIPtemperature.Text = -273.15
    End If
  Case "Kelvin"
    RefTemperature = temp
    If RefTemperature < 0 Then
      RefTemperature = 0
      txtIPtemperature.Text = 0
    End If
End Select
Call OPtemperature(RefTemperature)
End Sub

Private Sub OPtemperature(temp As Double)
Dim OPvalue As Double
Select Case cboOPtemperature.Text
  Case "Fahrenheit"
    OPvalue = (9 * (temp - 273.15) / 5) + 32
  Case "Celsius"
    OPvalue = temp - 273.15
  Case "Kelvin"
    OPvalue = temp
End Select
If OPvalue > -1000 And OPvalue < 1000 Then txtOPtemperature.Text = Format(OPvalue, "##0.0##")
If OPvalue > 999 And OPvalue < 10000 Then txtOPtemperature.Text = Format(OPvalue, "###0.0#")
End Sub

'----------------------------Torque calculations---------------------

Private Sub txtIPTorque_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtIPTorque, False
End Sub

Private Sub txtOPTorque_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPTorque_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOPTorque_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPTorque_Click()
  txtIPTorque_Change
End Sub

Private Sub cboIPTorque_Change()
  txtIPTorque_Change
End Sub

Private Sub cboOPTorque_Click()
  txtIPTorque_Change
End Sub

Private Sub cboOPTorque_Change()
  txtIPTorque_Change
End Sub

Private Sub txtIPTorque_Change()
Dim temp As Double
temp = Val(Replace(txtIPTorque, ",", "."))
If temp > 10000 Then
  temp = 10000
  txtIPTorque.Text = 10000
End If
Select Case cboIPTorque.Text
  Case "gcm"
    RefTorque = temp * 0.00980665 '--normalise all inputs to Newtoncentimeter
  Case "kgcm"
    RefTorque = temp * 9.80665
  Case "kgm"
    RefTorque = temp * 980.665
  Case "Ncm"
    RefTorque = temp
  Case "Nm"
    RefTorque = temp * 100
  Case "OunceInches"
    RefTorque = temp * 0.706155181422
  Case "PoundInches"
    RefTorque = temp * 11.2984829027617
  Case "Poundfeet"
    RefTorque = temp * 135.58179483314
    If temp < 0 Then
      temp = 0
      txtIPTorque.Text = 0
    End If
End Select
Call OPTorque(RefTorque)
End Sub

Private Sub OPTorque(temp As Double)
Dim OPvalue As Double
Select Case cboOPTorque.Text
  Case "gcm"
    OPvalue = temp * 101.971621297793
  Case "kgcm"
    OPvalue = temp * 0.1019716213
  Case "kgm"
    OPvalue = temp * 0.001019716213
  Case "Ncm"
    OPvalue = temp
  Case "Nm"
    OPvalue = temp * 0.01
  Case "OunceInches"
    OPvalue = temp * 1.416119326612
  Case "PoundInches"
    OPvalue = temp * 0.08850745791
  Case "Poundfeet"
    OPvalue = temp * 0.007375621492
End Select
DefineDisplayFormat OPvalue, display
txtOPTorque.Text = Format(OPvalue, display)
End Sub


'-----------------------Magnetic fieldstrenght calculations---------------------

Private Sub txtIPMagneticfield_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtIPMagneticfield, False
End Sub

Private Sub txtOPMagneticfield_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPMagneticfield_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOPMagneticfield_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboIPMagneticfield_Click()
  txtIPMagneticfield_Change
End Sub

Private Sub cboIPMagneticfield_Change()
  txtIPMagneticfield_Change
End Sub

Private Sub cboOPMagneticfield_Click()
  txtIPMagneticfield_Change
End Sub

Private Sub cboOPMagneticfield_Change()
  txtIPMagneticfield_Change
End Sub

Private Sub txtIPMagneticfield_Change()
Dim temp As Double
'-- 1Gauss = 1Oersted = 0.0001Tesla = 79.57747A/m ------reference data-------
temp = Val(Replace(txtIPMagneticfield, ",", "."))
If temp > 10000 Then
  temp = 10000
  txtIPMagneticfield.Text = 10000
End If
Select Case cboIPMagneticfield.Text
  Case "A/m"
    RefMagneticField = temp / 79.57747 'convert A/m to Gauss
  Case "Gauss"
    RefMagneticField = temp           'normalise all inputs to Gauss
  Case "Tesla"
    RefMagneticField = temp * 10000   'convert Tesla to Gauss
  Case "Oersted"
    RefMagneticField = temp           'convert Oersted to Gauss (1 G = 1 Oe)
    If temp < 0 Then
      temp = 0
      txtIPMagneticfield.Text = 0
    End If
End Select
Call OPMagneticfield(RefMagneticField)
End Sub

Private Sub OPMagneticfield(temp As Double) 'passing the reference fieldstrength
Dim OPvalue As Double
Select Case cboOPMagneticfield.Text
  Case "A/m"
    OPvalue = temp * 79.57747 'convert A/m to Gauss
  Case "Gauss"
    OPvalue = temp            '(Gauss is reference)
  Case "Tesla"
    OPvalue = temp * 0.0001   'convert Gauss to Tesla
  Case "Oersted"
    OPvalue = temp            'convert Gauss to Oersted (1 G= 1 Oe)
End Select
DefineDisplayFormat OPvalue, display
txtOPMagneticfield.Text = Format(OPvalue, display)
End Sub


'--------------------------------Pressure conversion calculations-----------------------

Private Sub txtInpressure_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtInPressure, True 'allow neg inputs
End Sub

Private Sub txtOutPressure_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOutPressure_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboInPressure_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboInPressure_Click()
  txtInPressure_Change
End Sub

Private Sub cboInPressure_Change()
  txtInPressure_Change
End Sub

Private Sub cboOutPressure_Click()
  txtInPressure_Change
End Sub

Private Sub cboOutPressure_Change()
  txtInPressure_Change
End Sub

Private Sub txtInPressure_Change()
Dim temp
temp = Val(Replace(txtInPressure, ",", "."))
If temp = 0 Then Exit Sub
If temp > 1000000 Then
  temp = 1000000
  txtInPressure.Text = temp
End If
If temp > 1E+99 Then
  temp = 1E+99
  txtInPressure.Text = Format(temp, "scientific")
End If
If temp < 1E-60 Then
  temp = 1E-60
  txtInPressure.Text = Format(temp, "scientific")
End If
Select Case cboInPressure.Text
  Case "Atmosphere"
    StandaardPressure = temp * 101325   'normalise all inputs to pascal
  Case "Torr"
    StandaardPressure = temp * 133.322
  Case "Bar"
    StandaardPressure = temp * 100000
  Case "mBar"
    StandaardPressure = temp * 100
  Case "Pascal"
    StandaardPressure = temp * 1
  Case "kPa"
    StandaardPressure = temp * 1000
  Case "hPa"
    StandaardPressure = temp * 100
End Select
Call OPPressure(StandaardPressure)    'Converted to Pascal
End Sub

Private Sub OPPressure(temp As Double)
Dim OPvalue As Double
Select Case cboOutPressure.Text
  Case "Atmosphere"
    OPvalue = temp / 101325
  Case "Torr"
    OPvalue = temp * 0.0075006
  Case "Bar"
    OPvalue = temp / 100000
  Case "mBar"
    OPvalue = temp / 100
  Case "Pascal"
    OPvalue = temp
  Case "kPa"
    OPvalue = temp / 1000
  Case "hPa"
    OPvalue = temp / 100
End Select
DefineDisplayFormat OPvalue, display
txtOutPressure.Text = Format(OPvalue, display)
End Sub

'-------------------------------Outgassing rates calculations---------------------
Private Sub txtOGRin_KeyPress(KeyAscii As Integer)
  DataEntry KeyAscii, txtOGRin, False 'don't allow neg inputs
End Sub

Private Sub txtOGRout_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOGRunitIn_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOGRunitOut_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub cboOGRunitIn_Click()
  txtOGRin_Change
End Sub

Private Sub cboOGRunitIn_Change()
  txtOGRin_Change
End Sub

Private Sub cboOGRunitOut_Click()
  txtOGRin_Change
End Sub

Private Sub cboOGRunitOut_Change()
  txtOGRin_Change
End Sub

Private Sub txtOGRin_Change()
Dim OGR As Double
temp = Val(Replace(txtOGRin, ",", "."))
If temp = 0 Then Exit Sub
If temp > 1000 Then
  temp = 1000
  txtOGRin.Text = Format(temp, "scientific")
End If
If temp > 1E+60 Then
  temp = 1E+60
  txtOGRin = Format(temp, "scientific")
End If
If temp < 1E-60 Then
  temp = 1E-60
  txtOGRin = Format(temp, "scientific")
End If
Select Case cboOGRunitIn.Text
  Case "Pa m/s"
    OGR = temp    'normalise all to Pa m/s
  Case "mBar-l/s-cm^2"
    OGR = temp * 1000
  Case "Torr-l/s-cm^2"
    OGR = temp * 1333
End Select
  Call OutgassingRate(OGR)  '---converted to Pa m/s
End Sub

Private Sub OutgassingRate(temp As Double)
Dim OPvalue As Double
Select Case cboOGRunitOut.Text
  Case "Pa m/s"
    OPvalue = temp
  Case "mBar-l/s-cm^2"
    OPvalue = temp / 1000
  Case "Torr-l/s-cm^2"
    OPvalue = temp / 1333
  End Select
DefineDisplayFormat OPvalue, display
txtOGRout.Text = Format(OPvalue, display)
End Sub

Private Sub DataEntry(KeyAscii As Integer, IPtxt As String, signflg As Boolean)
Dim a As Integer
'if signflg is true ->pos & neg. numbers requested,if signflg is false ->only pos. numbers requested
If signflg = False And KeyAscii = Asc("-") Then 'no neg numbers requested, ignore input
  If Right(Left(IPtxt, a), 1) = "-" Or Val(IPtxt) = 0 Then
    KeyAscii = 0
    Exit Sub
  End If
End If
If flgLocal = True Then '---if decimal separator is a period (US, UK, a.o.) --
'KeyAscii values --> 46 is "del", 8 is backspace
  If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = Asc("-") _
      Or KeyAscii = Asc("e") Or KeyAscii = Asc("E") Or KeyAscii = Asc(".")) Then
      KeyAscii = 0
    Exit Sub
  End If
Else '-------- if decimal separator is a comma (European continent) --
  If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = Asc("-") _
      Or KeyAscii = Asc("e") Or KeyAscii = Asc("E") Or KeyAscii = Asc(",")) Then
      KeyAscii = 0
    Exit Sub
  End If
End If
If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
IPtxt = Replace(IPtxt, ",", ".")  'make sure separator is a period (VB default)
  For a = 1 To Len(IPtxt)
    If Right(Left(IPtxt, a), 1) = "." Then  'avoid multiple separators
      KeyAscii = 0
      Exit Sub
    End If
  Next a
End If
If KeyAscii = Asc("e") Or KeyAscii = Asc("E") Then  'avoid multiple e or E
  For a = 1 To Len(IPtxt)
    If Right(Left(IPtxt, a), 1) = "e" Or Right(Left(IPtxt, a), 1) = "E" Then
      KeyAscii = 0
      Exit Sub
    End If
  Next a
  For a = 1 To Len(IPtxt)   'avoid an "e" after a minus sign
    If Right(Left(IPtxt, a), 1) = "-" Then
      KeyAscii = 0
      Exit Sub
    End If
  Next a
End If
If KeyAscii = Asc("-") Then
  For a = 1 To Len(IPtxt)
    If Right(Left(IPtxt, a), 1) = "-" Then 'avoid multiple minuses
      KeyAscii = 0
      Exit Sub
    End If
  Next a
End If
If IPtxt = "" And KeyAscii = Asc("e") Or KeyAscii = Asc("E") Then KeyAscii = 0
If IPtxt = "0" And KeyAscii = Asc("0") Then KeyAscii = 0  'avoid multiple zeros
If IPtxt = "-0" And KeyAscii = Asc("0") Then KeyAscii = 0  'avoid multiple zeros with neg input
If IPtxt = "0" And KeyAscii >= Asc("0") Then IPtxt = Replace(IPtxt, "0", "") 'remove leading zeros
IPtxt = IPtxt
End Sub


