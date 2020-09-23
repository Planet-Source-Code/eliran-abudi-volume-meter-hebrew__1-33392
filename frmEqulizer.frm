VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "òåöîä"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1515
   Icon            =   "frmEqulizer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   1515
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "òåöîä"
      Height          =   4695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   840
         Top             =   3240
      End
      Begin VB.CheckBox chkVolMute 
         Alignment       =   1  'Right Justify
         Caption         =   "äùú÷"
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   4080
         Width           =   855
      End
      Begin Project1.Progress Progress2 
         Height          =   4020
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   7091
         Picture         =   "frmEqulizer.frx":000C
         Orientation     =   1
      End
      Begin ComctlLib.Slider sldMain 
         Height          =   3495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   6165
         _Version        =   327682
         Orientation     =   1
         TickStyle       =   2
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "éöéàú ÷åì"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   4440
         Width           =   825
      End
      Begin VB.Label lblMain 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblMain"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   3840
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vol As New cVolume
Private Sub chkVolMute_Click()
     vol.VolumeMute = IIf((chkVolMute.Value = 1), True, False)
End Sub
Private Sub Form_Load()
    Set vol = New cVolume
    
    With sldMain
        .Min = vol.VolumeMin
        .Max = vol.VolumeMax
        .TickFrequency = (.Max - .Min) \ 10
        .LargeChange = .TickFrequency
    End With
    ' Main-------------------------------------------------------------------
    Progress2.Max = vol.MaxVolumeMeterOutput
End Sub
Private Sub Form_Paint()
    sldMain.Value = sldMain.Max - vol.VolumeLevel
    lblMain.Caption = format$((vol.VolumeLevel / vol.VolumeMax), "##0 %")
    chkVolMute.Value = IIf(vol.VolumeMute, 1, 0)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Set vol = Nothing
End Sub
Private Sub sldMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Beep
End Sub
Private Sub sldMain_Scroll()
    vol.VolumeLevel = sldMain.Max - sldMain.Value
    lblMain.Caption = format$((vol.VolumeLevel / vol.VolumeMax), "##0 %")
End Sub
Private Sub Timer1_Timer()
    Progress2.Position = vol.CurrentVolumeMeterOutput
End Sub
