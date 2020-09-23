VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   518
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   7920
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Click Me Please"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   3615
   End
   Begin VB.PictureBox P3 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   2
      Top             =   3840
      Width           =   3855
   End
   Begin VB.PictureBox P2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   3840
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Using WinDIBits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      TabIndex        =   6
      Top             =   6480
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0116
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3840
      TabIndex        =   5
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Transpacy Control"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   3840
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rember these color depths
'0 Blue
'1 Green
'2 Red
'3 B all -- not used
Private Sub Command2_Click()
Dim DI1() As Byte, DI2() As Byte
Dim nWid As Long, nHei As Long, PCNT As Long, vMid As Long
PCNT = 50
GetDIs P1, DI1(), nWid, nHei
GetDIs P2, DI2()
Dim NewByte() As Byte
'Loading into memory
ReDim NewByte(3, nWid, nHei)
'we need th dimensions that the getdis used to modify byte
'getdis exported these dimensions in the two varibles we created(nWid , nHei)
'//////////Begin Here
'begin with the transpacy effect
'how we make two photo transpacy?
'we reads the color of the first photo and the color on the second photo
'the we get the average of the two values
'we do that for every pixel in the two photos
'we need to make the two photos in similar dimensions first
For PCNT = 0 To 100 Step 5
For NX = 0 To nWid 'the maximum number of width bytes
For NY = 0 To nHei 'the maximum number of height bytes
'setting bytes as the PCNT varible tells us
'PCNT is where we are working now
NewByte(0, NX, NY) = (DI1(0, NX, NY) * ((100 - PCNT) / 100)) + (DI2(0, NX, NY) * ((PCNT) / 100))
NewByte(1, NX, NY) = (DI1(1, NX, NY) * ((100 - PCNT) / 100)) + (DI2(1, NX, NY) * ((PCNT) / 100))
NewByte(2, NX, NY) = (DI1(2, NX, NY) * ((100 - PCNT) / 100)) + (DI2(2, NX, NY) * ((PCNT) / 100))
Next: Next
SetDIs P3, NewByte()
P3.Refresh
PB.Value = PCNT
Next
Label1.Caption = "Negative Control"
DoEvents
'negative
For PCNT = 0 To 100 Step 5
'how to negative a color?
'to negative color we get the Invert of the Values of the Reg and Green and Blue Values
'because the maimum value of any one (Red or Green or Blue) is 255 so we get th negative as the following
'Red = 255 - Red
'Green = 255 - Green
'Blue = 255 - Blue
For NX = 0 To nWid 'the maximum number of width bytes
For NY = 0 To nHei 'the maximum number of height bytes
'setting bytes as the PCNT varible tells us
'PCNT is where we are working now
NewByte(0, NX, NY) = ((255 - (DI1(0, NX, NY))) * ((100 - PCNT) / 100)) + (DI1(0, NX, NY) * ((PCNT) / 100))
NewByte(1, NX, NY) = ((255 - (DI1(1, NX, NY))) * ((100 - PCNT) / 100)) + (DI1(1, NX, NY) * ((PCNT) / 100))
NewByte(2, NX, NY) = ((255 - (DI1(2, NX, NY))) * ((100 - PCNT) / 100)) + (DI1(2, NX, NY) * ((PCNT) / 100))
Next: Next
SetDIs P3, NewByte()
P3.Refresh
PB.Value = PCNT
Next
Label1.Caption = "GrayScale Control"
DoEvents
'grayscale
For PCNT = 0 To 100 Step 5
For NX = 0 To nWid 'the maximum number of width bytes
For NY = 0 To nHei 'the maximum number of height bytes
'how to gray sacle?
'gray scale is to get the value of the gray from a fixed color
'of course that will make (Red = Green = Blue)
vMid = (0.5 + (0.299 * (DI2(2, NX, NY))) + (0.587 * (DI2(1, NX, NY))) + (0.114 * (DI2(0, NX, NY))))
'setting bytes as the PCNT varible tells us
'PCNT is where we are working now
NewByte(0, NX, NY) = (vMid * ((100 - PCNT) / 100)) + (DI2(0, NX, NY) * ((PCNT) / 100))
NewByte(1, NX, NY) = (vMid * ((100 - PCNT) / 100)) + (DI2(1, NX, NY) * ((PCNT) / 100))
NewByte(2, NX, NY) = (vMid * ((100 - PCNT) / 100)) + (DI2(2, NX, NY) * ((PCNT) / 100))
Next: Next
SetDIs P3, NewByte()
P3.Refresh
PB.Value = PCNT
Next
'now we are done
MsgBox "Done!.", vbInformation
End Sub
Private Sub Form_Load()
'Loading Pictures from files
P1.Picture = LoadPicture(App.Path & "\..\Images\1.jpg")
P2.Picture = LoadPicture(App.Path & "\..\Images\2.jpg")
End Sub
'Hope you've learned something ....
'You Can Vote If You Want
