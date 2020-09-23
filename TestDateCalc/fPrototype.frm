VERSION 5.00
Begin VB.Form fTestDateCalc 
   Caption         =   "Test Date Routines"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   330
      Index           =   6
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3255
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   330
      Index           =   5
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2835
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   1470
      TabIndex        =   3
      Text            =   "45"
      Top             =   630
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   330
      Index           =   4
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2415
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   330
      Index           =   3
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1995
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&GO"
      Default         =   -1  'True
      Height          =   330
      Left            =   3780
      TabIndex        =   7
      Top             =   1260
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   2
      Left            =   1470
      TabIndex        =   6
      Text            =   "1"
      Top             =   1050
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   1470
      TabIndex        =   1
      Text            =   "2001"
      Top             =   210
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "WeekDay Name:"
      Height          =   225
      Index           =   6
      Left            =   105
      TabIndex        =   14
      Top             =   3255
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Month Name:"
      Height          =   225
      Index           =   5
      Left            =   105
      TabIndex        =   12
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Week Number:"
      Height          =   225
      Index           =   4
      Left            =   105
      TabIndex        =   2
      Top             =   630
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Week Number:"
      Height          =   225
      Index           =   3
      Left            =   105
      TabIndex        =   10
      Top             =   2415
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Result Date:"
      Height          =   225
      Index           =   2
      Left            =   105
      TabIndex        =   8
      Top             =   1995
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   105
      X2              =   4410
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Label Label2 
      Caption         =   "(1=Sunday, 2=Monday, 3=Tuesday, ...)"
      Height          =   225
      Left            =   105
      TabIndex        =   5
      Top             =   1470
      Width           =   3060
   End
   Begin VB.Label Label1 
      Caption         =   "Start Day:"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Year:"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "fTestDateCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fTestDateCalc
' Author:       Slider
' Date:         11/11/2001
' Version:      01.00.00
' Description:  Tests date routines.
' Edit History: 01.00.00 11/11/01 Initial Release
'
'===========================================================================

Option Explicit

Public Enum DayConstants
    ocalSunday = 1
    ocalMonday = 2
    ocalTuesday = 3
    ocalWednesday = 4
    ocalThursday = 5
    ocalFriday = 6
    ocalSaturday = 7
End Enum

'===========================================================================
'
Private Sub Command1_Click()

    Dim sMonth() As String
    Dim sDOW()   As String
    Dim dDate    As Date
    Dim eWeekDay As DayConstants

    eWeekDay = CByte(Text1(2).Text)
    dDate = WeekNum2Date(CLng(Text1(0).Text), CLng(Text1(1).Text), eWeekDay)
    GetMonthNames False, sMonth()
    GetWeekDayNames False, sDOW(), eWeekDay

    Text1(3).Text = FormatDateTime(dDate, vbShortDate)
    Text1(4).Text = WeekNumber(CDate(Text1(3).Text))
    Text1(5).Text = sMonth(VBA.Month(dDate))
    Text1(6).Text = sDOW(VBA.Weekday(dDate, eWeekDay))

End Sub
