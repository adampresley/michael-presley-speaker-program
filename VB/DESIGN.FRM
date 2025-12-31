VERSION 2.00
Begin Form Design 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculated Box Parameters"
   ClientHeight    =   5940
   ClientLeft      =   1140
   ClientTop       =   1485
   ClientWidth     =   7710
   ControlBox      =   0   'False
   Height          =   6345
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7710
   Top             =   1140
   Width           =   7830
   WindowState     =   2  'Maximized
   Begin CommandButton cmdVent 
      Caption         =   "&Vent"
      Height          =   350
      Left            =   4320
      TabIndex        =   25
      Top             =   5550
      Width           =   1095
   End
   Begin OptionButton GraphType 
      Caption         =   "Ported Box"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   21
      Top             =   4680
      Width           =   1335
   End
   Begin OptionButton GraphType 
      Caption         =   "Closed Box"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   20
      Top             =   4680
      Value           =   -1  'True
      Width           =   1335
   End
   Begin CommandButton cmdParamEnt 
      Caption         =   "&Parameters"
      Height          =   350
      Left            =   4320
      TabIndex        =   14
      Top             =   6360
      Width           =   1095
   End
   Begin CommandButton cmdDataBase 
      Caption         =   "&DataBase"
      Height          =   350
      Left            =   4320
      TabIndex        =   13
      Top             =   5955
      Width           =   1095
   End
   Begin CommandButton cmdGraph 
      Caption         =   "&Graph"
      Height          =   350
      Left            =   4320
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin Frame Frame2 
      Caption         =   "Ported Box"
      Height          =   2055
      Left            =   5880
      TabIndex        =   2
      Top             =   4680
      Width           =   3255
      Begin TextBox Fb_Text_PB 
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   720
         Width           =   615
      End
      Begin TextBox Vb_Text_PB 
         Height          =   285
         Left            =   2520
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin Label Hump_Text_PB 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Deviation From Flat Responce:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   2295
      End
      Begin Label F3_Text_PB 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Enclosure Volume (cu. ft.):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Enclosure Tuning:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "-3db Point (F3):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   16
         Top             =   1080
         Width           =   1695
      End
   End
   Begin Frame Frame1 
      Caption         =   "Closed Box"
      Height          =   1695
      Left            =   480
      TabIndex        =   1
      Top             =   5040
      Width           =   3495
      Begin TextBox Vb_Text_CB 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin Label Qtc_Text_CB 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin Label Sr_Text_CB 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin Label F3_Text_CB 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "-3db Point (F3):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "System Resonance:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "System Q:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Enclosure Volume (cu. ft.):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin GRAPH Graph1 
      AsciiFSize      =   "150~100~100~80"
      AsciiFStyle     =   "0"
      AsciiLegend     =   "~~Q =1"
      AsciiXPos       =   "0~0~0~10~100^"
      AutoInc         =   0  'Off
      BorderStyle     =   0  'None
      BottomTitle     =   "Frequency"
      DrawMode        =   3  'Blit
      GraphTitle      =   "Predicted Responce"
      GraphType       =   6  'Line
      GridStyle       =   3  'Horizontal and Vertical
      Height          =   4455
      Left            =   0
      LeftTitle       =   "db"
      LegendStyle     =   1  'Color
      NumPoints       =   200
      NumSets         =   2
      RandomData      =   0  'Off
      TabIndex        =   0
      ThickLines      =   0  'Off
      TickEvery       =   10
      Top             =   0
      Width           =   9615
      YAxisMax        =   10
      YAxisMin        =   -10
      YAxisPos        =   1  'Left
      YAxisStyle      =   2  'User-defined
      YAxisTicks      =   5
   End
   Begin Label Label4 
      Caption         =   "Hz."
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      Top             =   4725
      Visible         =   0   'False
      Width           =   375
   End
   Begin Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Calculating:"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   4725
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   4725
      Visible         =   0   'False
      Width           =   615
   End
End
Dim Qtc
Dim VB
Dim Sr
Dim F3
Dim Fb
Dim FirstTime

Sub cmdDataBase_Click ()
   SpkData.Show
End Sub

Sub cmdGraph_Click ()

If GraphType(0).Value = True Then
   'set parameters form closed box data
   VB = Vb_Text_CB.Text
   Qtc = Qtc_Text_CB.Caption
   Sr = Sr_Text_CB.Caption
   F3 = F3_Text_CB.Caption
   'graph the data
   GraphPoints 2, "Closed; Vb=" & Format(VB, "##.#") & " Q=" & Format(Qtc, "#.##")
   'change the Vb text back to black
   Vb_Text_CB.ForeColor = &H0
ElseIf GraphType(1).Value = True Then
   'set parameters form ported box data
   VB = Vb_Text_PB.Text
   Fb = Fb_Text_PB.Text
   F3 = F3_Text_PB.Caption
   'graph the data
   GraphPoints 4, "Ported; Vb=" & Format(VB, "#.##") & " Fb=" & Format(Fb, "##.#")
   'change the Vb text back to black
   Vb_Text_PB.ForeColor = &H0
End If

'Redraw the graph
Graph1.DrawMode = 2

End Sub

Sub cmdParamEnt_Click ()
   ParamEnt.Show
End Sub

Sub cmdVent_Click ()
   MsgBox ("This Option Not Available")
End Sub

Sub Form_Activate ()

On Error GoTo BadMath
If BeenHere Then Exit Sub
FirstTime = True

'___________________________________________
'Ported Box Design
VB = 15 * QTS ^ 2.87 * VAS
Fb = .42 * QTS ^ -.9 * Fs
F3 = .26 * QTS ^ -1.4 * Fs
Hump = 20 * Log(2.6 * QTS * (VAS / VB) ^ .35) / Log(10)

'Fill in the ported box text boxes
Vb_Text_PB.Text = Format(VB, "##.#")
Fb_Text_PB.Text = Format(Fb, "#.##")
F3_Text_PB.Caption = Format(F3, "##.#")
Hump_Text_PB.Caption = Format(Hump, "##.#")

'Graph the data
GraphPoints 3, "Ported; Vb=" & Format(VB, "#.##") & " Fb=" & Format(Fb, "##.#")

'___________________________________________
'CLOSED BOX DESIGN Q of 1
Q1 = 1 / QTS
Q2 = (Q1 ^ 2) - 1
VB = VAS / Q2
Sr = Fs * Q1
F3 = Sr * .79
Qtc = (Sr * QTS) / Fs

'Fill in the closed box text boxes
Vb_Text_CB.Text = Format(VB, "##.#")
Qtc_Text_CB.Caption = Format(Qtc, "#.##")
Sr_Text_CB.Caption = Format(Sr, "##.#")
F3_Text_CB.Caption = Format(F3, "##.#")
'________________________________________________

'Graph the data
GraphPoints 1, "Closed; Vb=" & Format(VB, "##.#") & " Q=" & Format(Qtc, "#.##")

'Redraw the graph
Graph1.DrawMode = 2


Gone:
BeenHere = True
FirstTime = False
Exit Sub

BadMath:
MsgBox ("Math Error")
Resume Gone

End Sub

Sub Form_Load ()

'Test Data ******************
'   Vas = 2.17
'   Fb = 0
'   Qts = .6
'   Vb = .6
'   Fs = 31
'****************************
   
   CenterForm Design
   Show
End Sub

Sub GraphPoints (SetNum%, Legend$)

On Error GoTo ErrorMsg

'Test Data ******************
'VAS = 2.17
'Fb = 0
'Qts = .6
'VB = .6
'Fs = 31
'****************************

Graph1.AutoInc = 0
Graph1.NumPoints = 200
Graph1.NumSets = 4
Graph1.ThisSet = SetNum%
Graph1.ThisPoint = SetNum%
Graph1.LegendText = Legend$
'Graph1.LabelEvery = 100

'Show the Freq counter
Label1.Visible = True
Label2.Visible = True
Label4.Visible = True

'Disable the buttons
cmdGraph.Enabled = False
cmdDatabase.Enabled = False
cmdParamEnt.Enabled = False
cmdVent.Enabled = False

A = (Fb ^ 2) / (Fs ^ 2)
B = A / QTS + (Fb / (7 * Fs))
C = 1 + A + (Fb / (7 * Fs * QTS)) + (VAS / VB)
D = 1 / QTS + (Fb / (7 * Fs))

For Freq = 1 To 200 Step 1
   F9 = Freq / Fs
   F5 = F9 ^ 2
   F3 = F9 ^ 3
   F4 = F9 ^ 4
   F6 = (F4 - C * F5 + A) ^ 2
   F7 = (B * F9 - D * F3) ^ 2
   Graph1.GraphData = 20 * (Log(F4 / (F6 + F7) ^ .5) / Log(10))
   
   If Graph1.GraphData < -10 Then
      Graph1.GraphData = -10
   End If
   
   Graph1.ThisPoint = Freq
   Graph1.XPosData = Freq
   
   DoEvents
   Label2.Caption = Freq

Next Freq

GoodBye:
'Hide the freq counter
Label1.Visible = False
Label2.Visible = False
Label4.Visible = False

'Enable the buttons
cmdGraph.Enabled = True
cmdDatabase.Enabled = True
cmdParamEnt.Enabled = True
cmdVent.Enabled = True
Exit Sub

ErrorMsg:
MsgBox ("Math Error")
Resume GoodBye

End Sub

Sub GraphType_Click (Index As Integer)
'
End Sub

Sub Vb_Text_CB_Change ()

If FirstTime Then Exit Sub
On Error GoTo MathError

VB = Vb_Text_CB.Text
A = VAS / VB
Sr = Fs * Sqr(A + 1)
Qtc = (Sr * QTS) / Fs
F3 = Sr * Sqr(((1 / Qtc ^ 2 - 2) + Sqr((1 / Qtc ^ 2 - 2) ^ 2 + 4)) / 2)

'Fill in the closed box text boxes
Qtc_Text_CB.Caption = Format(Qtc, "#.##")
Sr_Text_CB.Caption = Format(Sr, "##.#")
F3_Text_CB.Caption = Format(F3, "##.#")

'Make the text red until the graph is recalculated
Vb_Text_CB.ForeColor = &HFF

Bye:
FirstTime = False
Exit Sub

MathError:
Resume Bye

End Sub

Sub Vb_Text_PB_Change ()

If FirstTime Then Exit Sub
On Error GoTo MathError2

VB = Vb_Text_PB.Text
Fb = (Fs * (VAS / VB) ^ .32)
F3 = Fs * Sqr(VAS / VB)
Hump = 20 * Log(2.6 * QTS * (VAS / VB) ^ .35) / Log(10)

'Fill in the closed box text boxes
Fb_Text_PB.Text = Format(Fb, "#.##")
F3_Text_PB.Caption = Format(F3, "##.#")
Hump_Text_PB.Caption = Format(Hump, "##.#")

'Make the text red until the graph is recalculated
Vb_Text_PB.ForeColor = &HFF

Bye2:
FirstTime = False
Exit Sub

MathError2:
Resume Bye2

End Sub

