VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "GrayScale"
      Height          =   375
      Left            =   3360
      TabIndex        =   22
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Do"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   7200
      Width           =   1335
   End
   Begin MSComctlLib.Slider PbR 
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   5400
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
   End
   Begin VB.PictureBox pc 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   4680
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Negative"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mix"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   4080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Max             =   100
      SelStart        =   50
      Value           =   50
   End
   Begin VB.PictureBox PtO 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   120
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   2
      ToolTipText     =   "Click And Move!"
      Top             =   120
      Width           =   4095
   End
   Begin VB.PictureBox P2 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   5160
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   1935
      Left            =   4680
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComctlLib.Slider PbG 
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   5760
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
   End
   Begin MSComctlLib.Slider PbB 
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   6120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
   End
   Begin MSComctlLib.Slider PbI 
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   6480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
   End
   Begin MSComctlLib.Slider PbC 
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   6840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
   End
   Begin VB.Label Label7 
      Caption         =   "Set :-"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Mix Percentage"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Cont :"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Bright :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Blue :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Green :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Red :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   288
      X2              =   568
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Image S2 
      Height          =   3870
      Left            =   4320
      Top             =   3840
      Width           =   4200
   End
   Begin VB.Image S1 
      Height          =   3870
      Left            =   4320
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'////             Advanced  Color Controls                  ////
'////             Code By: Marco Samy Nasif                 ////
'////             mail:marco_s2@hotmail.com                 ////
'////             Call:  (+20) 12 72 42 974                 ////
'////             /////////////////////////                 ////
'////             /////////////////////////                 ////
'////             Arabic Republic Of  EGYPT                 ////
'////             Copyright (c)2002,   FREE                 ////
'////             To  Use  Or Include  into                 ////
'////             Your    Own    Programs .                 ////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////

Private Sub Command1_Click()
'Setting Values to Min
PB.Min = 0
PB.Value = 0
PB.Max = PtO.Width
'Begin
Dim Cl1, Cl2, Cl3
'For Each Point(Pixel) in the photo
For x = 1 To PtO.Width
PB.Value = x
For y = 1 To PtO.Height
'how we make two photo transpacy?
'we reads the color of the first photo and the color on the second photo
'the we get the average of the two values
'we do that for every pixel in the two photos
'we need to make the two photos in similar dimensions first
Cl1 = P1.Point(x, y) 'the color of the first photo
Cl2 = P2.Point(x, y) 'the color of the second photo
Cl3 = ColorPercentage(Cl1, Cl2, Slider1.Value) 'get the midcolor
PtO.PSet (x, y), Cl3 'setting new color
Next y
Next x
'setting dimensions
pc.Width = PtO.Width
pc.Height = PtO.Height
'painting the final picture
pc.PaintPicture PtO.Image, 0, 0
End Sub
Private Sub Command2_Click()
'Setting Values to Min
PB.Min = 0
PB.Value = 0
PB.Max = PtO.Width
'Begin
Dim Cl1, Cl3
'For Each Point(Pixel) in the photo
For x = 1 To PtO.Width
PB.Value = x
For y = 1 To PtO.Height
'how to negative a color?
'to negative color we get the Invert of the Values of the Reg and Green and Blue Values
'because the maimum value of any one (Red or Green or Blue) is 255 so we get th negative as the following
Cl1 = pc.Point(x, y) 'getting the color of current pixel in the photo
Cl3 = NegativeColor(Cl1) 'getting the negative of it
'setting new color value to current pixel
PtO.PSet (x, y), Cl3
Next y
Next x
'setting dimensions
pc.Width = PtO.Width
pc.Height = PtO.Height
'writing the final photo on the picture control
pc.PaintPicture PtO.Image, 0, 0
End Sub
Private Sub Command3_Click()
'Modify Photo Colors
''''''''''
'Setting Values to Min
PB.Min = 0
PB.Value = 0
PB.Max = PtO.Width
'Begin
Dim Cl1, Cl2, Cl3
'For Each Point(Pixel) in the photo
For x = 1 To PtO.Width
PB.Value = x 'moving progress
For y = 1 To PtO.Height
Cl1 = pc.Point(x, y)
'adding Red Modification
Cl2 = vbRed
Cl3 = ColorPercentage(Cl1, Cl2, 100 - PbR.Value)
Cl1 = Cl3
'adding Green Modification
Cl2 = vbGreen
Cl3 = ColorPercentage(Cl1, Cl2, 100 - PbG.Value)
Cl1 = Cl3
'adding Blue Modification
Cl2 = vbBlue
Cl3 = ColorPercentage(Cl1, Cl2, 100 - PbB.Value)
Cl1 = Cl3
'adding Brightness Modification
Cl2 = vbWhite
Cl3 = ColorPercentage(Cl1, Cl2, 100 - PbI.Value)
Cl1 = Cl3
'adding Blackness Modification
Cl2 = vbBlack
Cl3 = ColorPercentage(Cl1, Cl2, 100 - PbC.Value)
'Setting the new color
PtO.PSet (x, y), Cl3
Next y
Next x
End Sub
Private Sub Command4_Click()
'about
MsgBox "Copyright (c) Marco Samy Nasif" & vbCrLf & "send to : marco_s2@hotmail.com" & vbCrLf & "El-Minia , Egypt.", vbInformation
End Sub
Private Sub Command5_Click()
'setting values to min
PB.Min = 0
PB.Value = 0
PB.Max = PtO.Width
'begin
Dim Cl1, Cl3
'for every point in the two photos
For x = 1 To PtO.Width
PB.Value = x 'moving progress
For y = 1 To PtO.Height
'how to gray sacle?
'gray scale is to get the value of the gray from a fixed color
'of course that will make (Red = Green = Blue)
'See GrayColor in the Color Control Module
Cl1 = pc.Point(x, y) 'Color of the Picture
Cl3 = GrayColor(Cl1) ' the gray of this color
PtO.PSet (x, y), Cl3 'setting new color
Next y
Next x
'modifying dimensions
pc.Width = PtO.Width
pc.Height = PtO.Height
'Painting the final photo
pc.PaintPicture PtO.Image, 0, 0
End Sub
Private Sub Form_Load()
'Some Load Commands
'Setting Sizes
P1.Width = S1.Width
P1.Height = S1.Height
PtO.Width = S1.Width
PtO.Height = S1.Height
P2.Width = S1.Width
P2.Height = S1.Height
'Loading Pictures from files
S1.Picture = LoadPicture(App.Path & "\..\Images\1.jpg")
S2.Picture = LoadPicture(App.Path & "\..\Images\2.jpg")
'Painting the two Photos
P1.PaintPicture S1.Picture, 0, 0, S1.Width, S1.Height
P2.PaintPicture S2.Picture, 0, 0, S2.Width, S2.Height
End Sub
Private Sub PtO_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Extracting Values using the commands in the MouseMove
PtO_MouseMove Button, Shift, x, y
End Sub
Private Sub PtO_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'here are extracting rgb valus
If Button = 1 Then
Dim Rr, Rg, Rb, PointC
'the color of the current point under the mouse cursor
PointC = PtO.Point(x, y)
'getting valus of RGB
ColorRGB PointC, Rr, Rg, Rb
'Setting Sliders Valus
PbR.Value = Rr / 255 * 100
PbG.Value = Rg / 255 * 100
PbB.Value = Rb / 255 * 100
End If
End Sub
'I Hope You've Learned Something
'You Can Vote If You Want
