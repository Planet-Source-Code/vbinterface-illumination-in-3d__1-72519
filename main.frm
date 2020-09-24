VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    Illumination in 3D World"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11100
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   740
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picColorBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   2610
      Picture         =   "main.frx":0442
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   15
      Top             =   5790
      Width           =   3330
   End
   Begin VB.HScrollBar hsbIntensity 
      Height          =   195
      Left            =   2610
      Max             =   100
      TabIndex        =   14
      Top             =   6870
      Value           =   100
      Width           =   3330
   End
   Begin VB.HScrollBar hsbLight 
      Height          =   165
      Index           =   1
      Left            =   2610
      Max             =   180
      Min             =   -180
      TabIndex        =   9
      Top             =   5610
      Width           =   3330
   End
   Begin VB.HScrollBar hsbLight 
      Height          =   165
      Index           =   0
      Left            =   2610
      Max             =   180
      Min             =   -180
      TabIndex        =   8
      Top             =   5430
      Width           =   3330
   End
   Begin VB.PictureBox picLeaves 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000C0C0&
      Height          =   1065
      Left            =   7500
      Picture         =   "main.frx":13CE4
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   226
      TabIndex        =   4
      Top             =   6030
      Width           =   3390
      Begin VB.Label lblWish 
         BackStyle       =   0  'Transparent
         Caption         =   "I would like to work on 2D/3D graphics simulation programs as a freelancer."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   645
         Left            =   1140
         TabIndex        =   6
         Top             =   45
         Width           =   2130
      End
      Begin VB.Label lblContact 
         BackStyle       =   0  'Transparent
         Caption         =   "bytelogik@gmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1140
         TabIndex        =   5
         Top             =   765
         Width           =   2130
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4965
      Left            =   330
      ScaleHeight     =   331
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   0
      Top             =   300
      Width           =   5595
   End
   Begin VB.PictureBox picNormalMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00505050&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4965
      Left            =   330
      Picture         =   "main.frx":18E22
      ScaleHeight     =   331
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   20
      Top             =   300
      Width           =   5595
   End
   Begin VB.Label Label9 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous :       Vertex Mesh Deformation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F9F9F&
      Height          =   495
      Left            =   6210
      TabIndex        =   22
      Top             =   5970
      Width           =   1305
   End
   Begin VB.Label Label10 
      BackColor       =   &H00505050&
      Caption         =   "Read supporting article :   http://bytelogik.wordpress.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   6210
      TabIndex        =   21
      Top             =   5670
      Width           =   4665
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0027766F&
      X1              =   414
      X2              =   725
      Y1              =   156
      Y2              =   156
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0027766F&
      X1              =   414
      X2              =   725
      Y1              =   82
      Y2              =   82
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0027766F&
      X1              =   414
      X2              =   725
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0027766F&
      X1              =   414
      X2              =   725
      Y1              =   34
      Y2              =   34
   End
   Begin VB.Label Label8 
      BackColor       =   &H00505050&
      BackStyle       =   0  'Transparent
      Caption         =   "Coming Next :       Loading models in the 3D world."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0048D0BC&
      Height          =   495
      Left            =   6210
      TabIndex        =   19
      Top             =   6600
      Width           =   1305
   End
   Begin VB.Label Label7 
      BackColor       =   &H00505050&
      Caption         =   " Click on the colors. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0048D0BC&
      Height          =   825
      Left            =   330
      TabIndex        =   18
      Top             =   6030
      Width           =   2265
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":61014
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1740
      Left            =   6210
      TabIndex        =   17
      Top             =   3810
      Width           =   4695
   End
   Begin VB.Label lblTopA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":61237
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   990
      Left            =   6210
      TabIndex        =   16
      Top             =   1290
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00505050&
      Caption         =   " Light Intensity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0048D0BC&
      Height          =   195
      Left            =   330
      TabIndex        =   13
      Top             =   6870
      Width           =   2265
   End
   Begin VB.Label Label4 
      BackColor       =   &H00505050&
      Caption         =   " Light Color  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0048D0BC&
      Height          =   255
      Left            =   330
      TabIndex        =   12
      Top             =   5790
      Width           =   2265
   End
   Begin VB.Label Label3 
      BackColor       =   &H00505050&
      Caption         =   " Vertical"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0048D0BC&
      Height          =   165
      Left            =   1560
      TabIndex        =   11
      Top             =   5610
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackColor       =   &H00505050&
      Caption         =   " Horizontal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0048D0BC&
      Height          =   165
      Left            =   1560
      TabIndex        =   10
      Top             =   5430
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackColor       =   &H00505050&
      Caption         =   " Light Rotation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0048D0BC&
      Height          =   345
      Left            =   330
      TabIndex        =   7
      Top             =   5430
      Width           =   1215
   End
   Begin VB.Label lblMiddle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":61355
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1200
      Left            =   6210
      TabIndex        =   3
      Top             =   2430
      Width           =   4695
   End
   Begin VB.Label lblTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":614B1
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   630
      Left            =   6210
      TabIndex        =   2
      Top             =   570
      Width           =   4695
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Illumination in 3-Dimensions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   6210
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Code logic only for personal learning
    'Contact bytelogik@gmail.com for commercial use
    'Read supporting article :   http://bytelogik.wordpress.com
    
    '-----------------------------------------------------------------
    'You can use GDI-DIB functions instead of GetPixel and SetPixel to
    'speed up rendering time
    
    '-----------------------------------------------------------------
    Dim sW As Long, sH As Long, pX As Long, pY As Long
    Dim SurfaceNormal() As D3DVECTOR4
    Dim UseSurfaceNormal() As Byte
    Dim PixelColor As Long
    Dim PixelRed As Integer, PixelGreen As Integer, PixelBlue As Integer
    Dim NormColor As Long, MonoColor As Long
    Dim NormRed As Integer, NormGreen As Integer, NormBlue As Integer
    Dim LightVect As D3DVECTOR4, tLightVect As D3DVECTOR4, NormLight As D3DVECTOR4
    Dim LightIntensity As Single
    Dim LightColor As Long
    Dim LightRed As Integer, LightGreen As Integer, LightBlue As Integer
    Dim LightMatrix As D3DMATRIX
    Dim DotP As Single, ShadeValue As Single
    Dim AvoidColor As Long
    Private Const PIBY180 = 3.14 / 180
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Sub Form_Load()
    sW = pic.ScaleWidth: sH = pic.ScaleHeight
    ReDim SurfaceNormal(sW, sH)
    ReDim UseSurfaceNormal(sW, sH)
    AvoidColor = RGB(80, 80, 80)
    '----------------------------------------
    LightVect.X = 0: LightVect.Y = 0: LightVect.z = 500
    D3DXVec4Normalize NormLight, LightVect
    LightRed = 255: LightGreen = 255: LightBlue = 255
    LightIntensity = 1
    '----------------------------------------
    LoadSurfaceNormals      'eats loading time due to GetPixel
    IlluminateModel
End Sub
Sub LoadSurfaceNormals()
    MousePointer = 11
    For pX = 0 To sW
        For pY = 0 To sH
            NormColor = GetPixel(picNormalMap.hdc, pX, pY)
            If NormColor <> AvoidColor Then         'do not consider background color
                ColorLongToRGB NormColor, NormRed, NormGreen, NormBlue
                SurfaceNormal(pX, pY).X = ((NormRed / 255) * 2) - 1
                SurfaceNormal(pX, pY).Y = ((NormGreen / 255) * 2) - 1
                SurfaceNormal(pX, pY).z = ((NormBlue / 255) * 2) - 1
                UseSurfaceNormal(pX, pY) = 1
            Else
                UseSurfaceNormal(pX, pY) = 0
            End If
        Next pY
    Next pX
    MousePointer = 0
End Sub
Private Sub hsbIntensity_Change()
    LightIntensity = hsbIntensity.Value / 100
    IlluminateModel
End Sub
Private Sub hsbLight_Change(Index As Integer)
    'transform light ray
    D3DXMatrixRotationYawPitchRoll LightMatrix, -hsbLight(0).Value * PIBY180, hsbLight(1).Value * PIBY180, 0
    D3DXVec4Transform tLightVect, LightVect, LightMatrix
    D3DXVec4Normalize NormLight, tLightVect
    IlluminateModel
End Sub
Private Sub picColorBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LightColor = picColorBox.Point(X, Y)
    ColorLongToRGB LightColor, LightRed, LightGreen, LightBlue
    IlluminateModel
End Sub
Sub IlluminateModel()
    MousePointer = 11
    For pX = 0 To sW
        For pY = 0 To sH
            If UseSurfaceNormal(pX, pY) = 1 Then
                DotP = D3DXVec4Dot(NormLight, SurfaceNormal(pX, pY))
                If DotP < 0 Then DotP = 0
                If DotP > 1 Then DotP = 1
                ShadeValue = DotP * LightIntensity        'Apply the light intensity
                PixelRed = ShadeValue * LightRed
                PixelGreen = ShadeValue * LightGreen
                PixelBlue = ShadeValue * LightBlue
                PixelColor = RGB(PixelRed, PixelGreen, PixelBlue)
                SetPixel pic.hdc, pX, pY, PixelColor      'A more time consuming function
            End If
        Next pY
    Next pX
    pic.Refresh
    MousePointer = 0
End Sub
Sub ColorLongToRGB(ByVal SplitColor As Long, ByRef RedValue As Integer, ByRef GreenValue As Integer, ByRef BlueValue As Integer)
    'There is another simple logic for splitting the long color into R,G,B
    'I will mention it in my next submission
    
    RedValue = Abs(SplitColor Mod &H100)
    SplitColor = Abs(SplitColor \ &H100)
    GreenValue = Abs(SplitColor Mod &H100)
    SplitColor = Abs(SplitColor \ &H100)
    BlueValue = Abs(SplitColor Mod &H100)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Erase SurfaceNormal
    Erase UseSurfaceNormal
End Sub
