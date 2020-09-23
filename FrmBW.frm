VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmBW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smart Black and White - Loc2K"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicDither 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4680
      Picture         =   "FrmBW.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   2640
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   2640
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   120
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Load Image..."
      Filter          =   "Windows Bitmap (*.bmp)|*.bmp|All Files (*.*)|*.*"
      Flags           =   4
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton CmdDither 
      Caption         =   "Dither"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   7560
      Width           =   1335
   End
   Begin VB.PictureBox PcWS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7500
      Left            =   0
      ScaleHeight     =   7440
      ScaleWidth      =   7440
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      Begin VB.PictureBox PicBW 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   2895
         TabIndex        =   5
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Label LblProg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   8160
      Width           =   1335
   End
End
Attribute VB_Name = "FrmBW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Const SRCCOPY = &HCC0020

Private Sub CmdLoad_Click()
    On Error GoTo ErrorHandler
    CD1.ShowOpen
    PicBW.Picture = LoadPicture(CD1.FileName)
ErrorHandler:
End Sub

Private Sub CmdDither_Click()
    PicBW.Cls
    GradInt = (PicDither.Width / 15) / 11
    For v = 0 To PicBW.Height / 15 - 1
        For h = 0 To PicBW.Width / 15 - 1
            CurRGB = (RGBCon(GetPixel(PicBW.hdc, h, v), 1) + RGBCon(GetPixel(PicBW.hdc, h, v), 2) + RGBCon(GetPixel(PicBW.hdc, h, v), 3)) / 3
            PalLoc = 0
            Do Until GradInt * PalLoc > (CurRGB / 255) * (PicDither.Width / 15)
                PalLoc = PalLoc + 1
            Loop
            PalLoc = PalLoc - 2
            Sclh = h
            Sclv = v
            If h > 15 Then Sclh = h Mod 16
            If v > 15 Then Sclv = v Mod 16
            SetPixel PicBW.hdc, h, v, GetPixel(PicDither.hdc, GradInt * PalLoc + Sclh, Sclv)
        Next h
        PicBW.Refresh
        LblProg.Caption = Format(100 * (v + h / (PicBW.Width / 15 - 1)) / (PicBW.Height / 15 - 1), "0.00") & "%"
        LblProg.Refresh
    Next v
End Sub

Private Function RGBCon(RGBColor, CType As Integer)
    'Convert RGB integer to R (CType = 1), G (CType = 2), or B (CType = 3) integer
    If CType = 1 Then
        CType = 3
    ElseIf CType = 3 Then
        CType = 1
    End If
    HRGB = Left("000000", 6 - Len(Hex(RGBColor))) & Hex(RGBColor)
    RGBCon = Val("&H" & Mid(HRGB, 1 + 2 * (CType - 1), 2))
End Function

Private Sub CmdSave_Click()
    SavePicture PicBW.Image, App.Path & "\Temp.bmp"
    MsgBox "Saved to: " & App.Path & "\Temp.bmp", , "Saved"
End Sub
