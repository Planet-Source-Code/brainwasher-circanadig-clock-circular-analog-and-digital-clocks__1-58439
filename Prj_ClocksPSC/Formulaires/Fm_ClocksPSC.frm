VERSION 5.00
Begin VB.Form Fm_ClocksPSC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Analog and Digital clocks..."
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   398
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_Quit 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Timer Timer_Digital 
      Enabled         =   0   'False
      Interval        =   950
      Left            =   4440
      Top             =   1200
   End
   Begin VB.Timer Timer_Analog 
      Enabled         =   0   'False
      Interval        =   950
      Left            =   3240
      Top             =   1200
   End
   Begin VB.PictureBox Pic_Final 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   120
      Picture         =   "Fm_ClocksPSC.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   1200
      Width           =   3030
   End
   Begin VB.PictureBox Pic_Initial 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   120
      Picture         =   "Fm_ClocksPSC.frx":27144
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label Lbl_Digital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   7
      Left            =   7320
      TabIndex        =   14
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label Lbl_Digital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   6
      Left            =   6960
      TabIndex        =   13
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label Lbl_Digital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   5
      Left            =   6630
      TabIndex        =   12
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label Lbl_Digital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   4
      Left            =   6450
      TabIndex        =   11
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label Lbl_Digital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   3
      Left            =   6090
      TabIndex        =   10
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label Lbl_Digital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   2
      Left            =   5745
      TabIndex        =   9
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label Lbl_Digital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   1
      Left            =   5580
      TabIndex        =   8
      Top             =   1650
      Width           =   360
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   528
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Label Lbl_Info 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Fm_ClocksPSC.frx":4E288
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   4920
      TabIndex        =   6
      Top             =   2730
      Width           =   3015
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   240
      Picture         =   "Fm_ClocksPSC.frx":4E33E
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "Create a graphical circular progress bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "This tutorial will show you how to create such a circular progress bar. The image can be changed."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label Lbl_Digital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   0
      Left            =   5220
      TabIndex        =   3
      Top             =   1650
      Width           =   360
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1530
      Left            =   4920
      Picture         =   "Fm_ClocksPSC.frx":4F02E
      Top             =   1200
      Width           =   3030
   End
   Begin VB.Label Lbl_Time 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--:--:--"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "Fm_ClocksPSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'___________________________________________________________________________
' Program name      : Analog and Digital clocks.
' Description       : A nice way to create clocks. From code to fake.
' Company           : MELANTECH
' Authors           : Weitten Pascal
'___________________________________________________________________________
'
' Date              : (c) 2005.01.22
' Version N°        : V0.1
' Customer          : Internal stuff.
'
' Last Modification : 2005.01.22
'___________________________________________________________________________
' TODO :
'       -
'       -
'___________________________________________________________________________
'
' Don't set the timer's interval at 1000. It may cause some lags
' due to processing time and some times the needles are jumping one
' position. To avoid it use an interval >500 and <1000.
'___________________________________________________________________________
'
' Both clocks were drawn by myself, based on tutorials found on web sites
' dedicated to PhotopShop. I'm not a graphician, so don't be too harsh.
'___________________________________________________________________________
'

Const Pi = 3.141592654

'We'll use tables for the cosine and sine tables. As the
'seconds and minutes use 60 positions we'll create one
'single table for both of them. For the hours we'll use
'a separate table. As there only 12 hours displayed on
'an analog clock.
Dim CosTable(61) As Double
Dim SinTable(61) As Double
Dim CosTableH(25) As Double
Dim SinTableH(25) As Double

'Digital Clock
'This one is a bit of a fake.
'Anyway it looks nice.
'We'll use the original digit position when the value is not 1
'if the value is 1, then we have to move the label (digit) a bit to the
'right so it's position matches the graphical clock.
Dim HourLeftIs1 As Integer, HourLeftIsNot1 As Integer
Dim HourRightIs1 As Integer, HourRightIsNot1 As Integer
Dim MinuteLeftIs1 As Integer, MinuteLeftIsNot1 As Integer
Dim MinuteRightIs1 As Integer, MinuteRightIsNot1 As Integer
Dim SecondLeftIs1 As Integer, SecondLeftIsNot1 As Integer
Dim SecondRightIs1 As Integer, SecondRightIsNot1 As Integer

Private Sub Cmd_Quit_Click()
    End
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    'Analog Clock
    'First we'll generate the sine and cosine tables.
    'This method was prefered to real time calculations
    'as there might be some lag. It also eases up the use of
    'only one function.
    Call GenerateTables
    
    'Now we can enable the timers.
    Timer_Analog.Enabled = True
    Timer_Digital.Enabled = True
    
    'Digital Clock.
    'We setup these positions once for all.
    'Not a piece of optimization, but functionnal.
    'Hour
    HourLeftIs1 = Lbl_Digital(0).Left + 3
    HourLeftIsNot1 = Lbl_Digital(0).Left
    HourRightIs1 = Lbl_Digital(1).Left + 3
    HourRightIsNot1 = Lbl_Digital(1).Left
    
    'Minutes
    MinuteLeftIs1 = Lbl_Digital(3).Left + 3
    MinuteLeftIsNot1 = Lbl_Digital(3).Left
    MinuteRightIs1 = Lbl_Digital(4).Left + 3
    MinuteRightIsNot1 = Lbl_Digital(4).Left
   
    'Seconds
    SecondLeftIs1 = Lbl_Digital(6).Left + 3
    SecondLeftIsNot1 = Lbl_Digital(6).Left
    SecondRightIs1 = Lbl_Digital(7).Left + 3
    SecondRightIsNot1 = Lbl_Digital(7).Left
End Sub

Private Sub Timer_Analog_Timer()
    Dim strTime As String
    Dim angSeconds As Integer
    Dim angMinutes As Integer
    Dim angHour As Integer

    On Error Resume Next
    'There's nothing really special there.
    'We just split the current time in hour, minutes and seconds.
    strTime = CStr(Format(Time, "hh:mm:ss"))
    angSeconds = CInt(Right(strTime, 2))
    angMinutes = CInt(Mid$(strTime, 4, 2))
    angHour = CInt(Left(strTime, 2))

    'We the draw each needle for the hour, minutes and seconds.
    'Draw_Needle(intNeedleRadius, intNeedleAngle , intNeedleColour, boolMS)
    'intNeedleRadius: Lenght of the needle.
    'intNeedleAngle : Angle of the needle (in fact the hour minute or second)
    'intNeedleColour: Colour of the needle.
    'boolMS         : Are we using Hours' table or Minutes and Seconds' table?
    Call Draw_Needle(80, angSeconds, 0, True)
    Call Draw_Needle(70, angMinutes, 1, True)
    Call Draw_Needle(50, angHour, 1, False)
    
    Lbl_Time.Caption = strTime     'Just in case we had doubts.
End Sub

'___________________________________________________________________________
'
'Function description:
'
'Draw_Needle(intNeedleRadius, intNeedleAngle , intNeedleColour, boolMS)
'intNeedleRadius: Lenght of the needle.
'intNeedleAngle : Angle of the needle (in fact the hour minute or second)
'intNeedleColour: Colour of the needle.
'boolMS         : Are we using Hours' table or Minutes and Seconds' table?
'
'___________________________________________________________________________
'
Function Draw_Needle(intNeedleRadius As Integer, intNeedleAngle As Integer, intNeedleColour As Integer, boolMS As Boolean)
    Dim X As Integer, Y As Integer
    Dim X2 As Double, Y2 As Double
    Dim i As Long
    Dim CosX As Double, CosY As Double, SinX As Double, SinY As Double
    Dim lngInitialPixelColour As Long
    
    On Error Resume Next
    'We define the center of our graphical clock.
    'Normaly you should use Pic_Final.Width/2 and Pic_Final.Height/2 to
    'get the center. But you might have noticed, mys clock ain't centered
    'on my image (I was maybe tired while creating the clock image).
    'So I fixed the positions.
    X = 96 '(Pic_Final.Width / 2)
    Y = 94 '(Pic_Final.Height / 2)
    
    'We are going to draw the Needle.
    'Are we using a minutes or seconds value?
    If boolMS Then
        CosX = CosTable(intNeedleAngle + 1)  'We started our cos and sin table at -1.
        SinX = SinTable(intNeedleAngle + 1)  'So to get the right angle we add 1.
    Else
        CosX = CosTableH(intNeedleAngle + 1) 'In case we are drawing the hours needle.
        SinX = SinTableH(intNeedleAngle + 1)
    End If
    
    For i = 0 To intNeedleRadius
        X2 = X + (i * CosX)
        Y2 = Y + (i * SinX)
        'The following is just here to make things look nice.
        If intNeedleColour = 0 Then
            Pic_Final.PSet (X2, Y2), RGB(255, 50, 50)
        ElseIf intNeedleColour = 1 Then
            Pic_Final.PSet (X2, Y2), RGB(0, 0, 0)
        ElseIf intNeedleColour = 2 Then
            Pic_Final.PSet (X2, Y2), RGB(155, 155, 155)
        End If
    Next i
            
    'To make it realistic, we'll have to hide the previous needle's position.
    'Here we go
    If boolMS Then
        CosX = CosTable(intNeedleAngle)     'The previous angle value, remember we started
        SinX = SinTable(intNeedleAngle)     'Our tables at -1.
    Else
        CosX = CosTableH(intNeedleAngle)
        SinX = SinTableH(intNeedleAngle)
    End If
    
    For i = 0 To intNeedleRadius
        X2 = X + (i * CosX)
        Y2 = Y + (i * SinX)
        lngInitialPixelColour = Pic_Initial.Point(X2, Y2) 'Get teh pixels colour.
        Pic_Final.PSet (X2, Y2), lngInitialPixelColour    'Put this colour to the final picture.
    Next i
End Function


Sub GenerateTables()
    Dim i As Integer
    
    'Generate the tables for the Minutes and Seconds.
    'This part needs some comments.
    'We are using an analog clock, this means we start at 12
    'To do this on a circle your starting angle must be
    '270° clockwise or -90° unclockwise. I chose -90.
    'In fact variable i we'll be the minutes or seconds.
    'We also mutltiply the value by 6 cause there are 60 minutes
    'in one hour and 60 seconds in one minute. On a 360° circle
    'this means 60 divisions of 6° (60*6=360).
    'Example: If we want the angle for 15 seconds. We have 15*6=90°
    'but as we want to start at 12, we substract our -90 angle (it could also be 90+270)
    'We then convert this value to radians by doing this: * Pi/180.
    'We got our angle and just have to calculate teh sine and cosine values.
    'We also start our tables at -1 and 0, because we'll use these
    'tables to erase the previous angle needle. And as we can get a 0 value
    'we could have faced a crash when trying to erase -1 angle. Tricky again.
    For i = -1 To 60
        Let CosTable(i + 1) = Cos(((i * 6) - 90) * (Pi / 180))
        Let SinTable(i + 1) = Sin(((i * 6) - 90) * (Pi / 180))
    Next i
    
    'Generate the tables for the Hours.
    'The only change here, is that an analog clock displays 12 hours per day.
    'So 360/12=30° divisions. I guess you got it now.
    For i = -1 To 24
        Let CosTableH(i + 1) = Cos(((i * 30) - 90) * (Pi / 180))
        Let SinTableH(i + 1) = Sin(((i * 30) - 90) * (Pi / 180))
    Next i
End Sub

Private Sub Timer_Digital_Timer()
    Dim strTime As String
    Dim strHour As String, strMinutes As String, strSeconds As String
    Dim intValLeft As Integer, intValRight As Integer
    
    On Error Resume Next
    strTime = CStr(Format(Time, "hh:mm:ss"))
    strSeconds = CStr(Right(strTime, 2))
    strMinutes = CStr(Mid$(strTime, 4, 2))
    strHour = CStr(Left(strTime, 2))

    'The following code might look tricky, and it is. We are
    'facking the digital clock. And want it to look nice.
    'It can be done in a different way, but it'll be up to you
    'to discover it.

    'Hour.
    intValLeft = CInt(Left(strHour, 1))
    intValRight = CInt(Right(strHour, 1))
    If intValLeft = 1 Then Lbl_Digital(0).Left = HourLeftIs1 Else Lbl_Digital(0).Left = HourLeftIsNot1
    If intValRight = 1 Then Lbl_Digital(1).Left = HourRightIs1 Else Lbl_Digital(1).Left = HourRightIsNot1
    Lbl_Digital(0).Caption = intValLeft
    Lbl_Digital(1).Caption = intValRight

    'Minutes.
    intValLeft = CInt(Left(strMinutes, 1))
    intValRight = CInt(Right(strMinutes, 1))
    If intValLeft = 1 Then Lbl_Digital(3).Left = MinuteLeftIs1 Else Lbl_Digital(3).Left = MinuteLeftIsNot1
    If intValRight = 1 Then Lbl_Digital(4).Left = MinuteRightIs1 Else Lbl_Digital(4).Left = MinuteRightIsNot1
    Lbl_Digital(3).Caption = intValLeft
    Lbl_Digital(4).Caption = intValRight

    'Seconds.
    intValLeft = CInt(Left(strSeconds, 1))
    intValRight = CInt(Right(strSeconds, 1))
    If intValLeft = 1 Then Lbl_Digital(6).Left = SecondLeftIs1 Else Lbl_Digital(6).Left = SecondLeftIsNot1
    If intValRight = 1 Then Lbl_Digital(7).Left = SecondRightIs1 Else Lbl_Digital(7).Left = SecondRightIsNot1
    Lbl_Digital(6).Caption = intValLeft
    Lbl_Digital(7).Caption = intValRight
End Sub
