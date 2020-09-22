VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Sunrise Calculator"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtpDay 
      Height          =   315
      Left            =   1380
      TabIndex        =   18
      Top             =   2400
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      _Version        =   393216
      Format          =   22806529
      CurrentDate     =   38850
   End
   Begin VB.CheckBox chkDaySavings 
      Height          =   255
      Left            =   1380
      TabIndex        =   16
      Top             =   2100
      Width           =   255
   End
   Begin VB.TextBox txtTimeZone 
      Height          =   315
      Left            =   1380
      TabIndex        =   14
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   315
      Left            =   1740
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sun"
      Height          =   1695
      Left            =   60
      TabIndex        =   5
      Top             =   3120
      Width           =   4635
      Begin VB.Label lblSunset 
         BackColor       =   &H80000018&
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1320
         Width           =   3555
      End
      Begin VB.Label lblSunNoon 
         BackColor       =   &H80000018&
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   840
         Width           =   3555
      End
      Begin VB.Label lblSunrise 
         BackColor       =   &H80000018&
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   3555
      End
      Begin VB.Label Label6 
         Caption         =   "Sunset"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1380
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Noon:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Sunrise:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.TextBox txtLongitude 
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   1260
      Width           =   1455
   End
   Begin VB.TextBox txtLatitude 
      Height          =   315
      Left            =   1380
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Day:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2460
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Day Savings:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Time Zone:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Longitude:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1260
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Latitude:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Sunrise Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4395
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================
'           Copyright Information
'==========================================================
'Program Name     : Mewsoft Qibla Direction Compass
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Home Page        : http://www.islamware.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================
Option Explicit

Private cSun As clsSunrise

Private Sub Form_Load()
    
    Set cSun = New clsSunrise
    
    'Latitude: North = +, South = -
    'Longitude: West = +, East = -
    
    txtLatitude.Text = 30.05
    txtLongitude.Text = -31.25
    txtTimeZone.Text = -2
    chkDaySavings.Value = vbChecked
    
    cmdCalculate_Click
   
End Sub

Private Sub cmdCalculate_Click()
    
    If Not (IsNumeric(txtLatitude.Text)) Then Exit Sub
    If Not (IsNumeric(txtLongitude.Text)) Then Exit Sub
    If Not (IsNumeric(txtTimeZone.Text)) Then txtTimeZone.Text = 0
    
    cSun.Latitude = CDbl(txtLatitude.Text)
    cSun.Longitude = CDbl(txtLongitude.Text)
    cSun.TimeZone = CDbl(txtTimeZone.Text)
    
    cSun.DateDay = dtpDay.Value
    cSun.TimeZone = -2
    cSun.DaySavings = chkDaySavings.Value

    cSun.CalculateSun
    
    lblSunrise.Caption = cSun.Sunrise
    lblSunNoon.Caption = cSun.SolarNoon  'SunTransit
    lblSunset.Caption = cSun.Sunset

End Sub

Private Sub chkDaySavings_Click()
    cmdCalculate_Click
End Sub

Private Sub dtpDay_Change()
    cmdCalculate_Click
End Sub

Private Sub txtLatitude_Change()
    cmdCalculate_Click
End Sub

Private Sub txtLongitude_Change()
    cmdCalculate_Click
End Sub

Private Sub txtTimeZone_Change()
    cmdCalculate_Click
End Sub
