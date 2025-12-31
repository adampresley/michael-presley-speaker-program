VERSION 2.00
Begin Form ParamEnt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Speaker Parameters"
   ClientHeight    =   5175
   ClientLeft      =   1605
   ClientTop       =   2565
   ClientWidth     =   6450
   ControlBox      =   0   'False
   Height          =   5580
   Left            =   1545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6450
   Top             =   2220
   Width           =   6570
   Begin CommandButton cmdDatabase 
      Caption         =   "Data&base"
      Height          =   375
      Left            =   5040
      TabIndex        =   26
      Top             =   240
      Width           =   1095
   End
   Begin CommandButton cmdDesign 
      Caption         =   "&Design"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin Frame Frame3 
      Caption         =   "Calculated Parameters"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   5775
      Begin Label Label14 
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin Label Label13 
         Height          =   255
         Left            =   4320
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin Label Label12 
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   720
         Width           =   735
      End
      Begin Label Label11 
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Vas ="
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Qts ="
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
      Begin Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Qes ="
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Qms ="
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
   End
   Begin Frame Frame2 
      Height          =   1695
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Width           =   5775
      Begin TextBox F1_Text 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin TextBox F2_Text 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin TextBox TFS_Text 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin Label Label4 
         Caption         =   "Frequency Below Resonance Where Impedance ="
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4455
      End
      Begin Label Label4 
         Caption         =   "Frequency Above Resonance Where Impedance ="
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   4455
      End
      Begin Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Speaker Resonance In Test Box:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2040
         LinkItem        =   "e"
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
   End
   Begin Frame Frame1 
      Height          =   1935
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   4455
      Begin TextBox RE_Text 
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin TextBox FS_Text 
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin TextBox ZMAX_Text 
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin TextBox TVB_Text 
         Height          =   285
         Left            =   3360
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Voice Coil dc Resistance:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Free Air Resonance:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Voice Coil Impedance at Resonance:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   2775
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Volume Of Test Box:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
   End
End
Dim VasTemp
Dim QtsTemp
Dim FsTemp

Sub cmdDataBase_Click ()
   SpkData.Show
End Sub

Sub cmdDesign_Click ()
   Design.Show
End Sub

Sub EmptyTest1 (OK1)
   If Val(RE_Text.Text) > 0 And Val(FS_Text.Text) > 0 And Val(ZMAX_Text.Text) > 0 And Val(TVB_Text.Text) > 0 Then
      OK1 = True
   End If
End Sub

Sub EmptyTest2 (OK2)
   If Val(TFS_Text.Text) > 0 And Val(F1_Text.Text) > 0 And Val(F2_Text.Text) > 0 Then
      OK2 = True
   End If
End Sub

Sub F1_Text_Change ()
   F1 = Val(F1_Text.Text)
   EmptyTest2 OK2
   If OK2 Then ShowAnswers
End Sub

Sub F1_Text_GotFocus ()
   HighLightText F1_Text
End Sub

Sub F2_Text_Change ()
   F2 = Val(F2_Text.Text)
   EmptyTest2 OK2
   If OK2 Then ShowAnswers
End Sub

Sub F2_Text_GotFocus ()
   HighLightText F2_Text
End Sub

Sub Form_Activate ()
   VasTemp = Vas
   QtsTemp = Qts
   FsTemp = Fs
End Sub

Sub Form_Deactivate ()
      Vas = VasTemp
      Qts = QtsTemp
      Fs = FsTemp
End Sub

Sub Form_Load ()
   CenterForm ParamEnt
   'Show
   'RE_Text.Text = ""
   'FS_Text.Text = ""
   'ZMAX_Text.Text = ""
   'TVB_Text.Text = ""
   'TFS_Text.Text = ""
   'F1_Text.Text = ""
   'F2_Text.Text = ""
End Sub

Sub FS_Text_Change ()
   Fs = Val(FS_Text.Text)
   EmptyTest1 OK1
   If OK1 Then
      MakeAvailable
      EmptyTest2 OK2
      If OK2 Then ShowAnswers
   Else
      MakeUnAvailable
   End If
End Sub

Sub FS_Text_GotFocus ()
   HighLightText FS_Text
End Sub

Sub HighLightText (ObjectName As Control)
    ObjectName.SelStart = 0
    ObjectName.SelLength = Len(ObjectName.Text)
End Sub

Sub MakeAvailable ()
   TFS_Text.Enabled = True
   F1_Text.Enabled = True
   F2_Text.Enabled = True
End Sub

Sub MakeUnAvailable ()
   TFS_Text.Enabled = False
   F1_Text.Enabled = False
   F2_Text.Enabled = False
End Sub

Sub RE_Text_Change ()
   RE = Val(RE_Text.Text)
   EmptyTest1 OK1
   If OK1 Then
      MakeAvailable
      RO = ZMAX / RE
      RF = Sqr(RO) * RE
      Label4(0).Caption = "Frequency Below Resonance Where Impedance = " & Format(RF, "##.#") & " ohms"
      Label4(1).Caption = "Frequency Above Resonance Where Impedance = " & Format(RF, "##.#") & " ohms"
   Else
      MakeUnAvailable
   End If
   EmptyTest2 OK2
   If OK2 Then ShowAnswers
End Sub

Sub RE_Text_GotFocus ()
   HighLightText RE_Text
End Sub

Sub ShowAnswers ()
   On Error GoTo ErrHandle

   QMS = Fs * Sqr(RO) / (F2 - F1)
   Label11.Caption = Format(QMS, "##.##")
   
   QES = QMS / (RO - 1)
   Label12.Caption = Format(QES, "##.##")
   
   Qts = (QMS * QES) / (QMS + QES)
   Label13.Caption = Format(Qts, "##.##")
   
   Vas = TVB * (1.149 * ((TFS / Fs) ^ 2 - 1))
   Label14.Caption = Format(Vas, "##.##")

ByeBye:
Exit Sub

ErrHandle:
MsgBox ("Math Error")
Resume ByeBye

End Sub

Sub TFS_Text_Change ()
   TFS = Val(TFS_Text.Text)
   EmptyTest2 OK2
   If OK2 Then ShowAnswers
End Sub

Sub TFS_Text_GotFocus ()
   HighLightText TFS_Text
End Sub

Sub TVB_Text_Change ()
   TVB = Val(TVB_Text.Text)
   EmptyTest1 OK1
   If OK1 Then
      MakeAvailable
      RO = ZMAX / RE
      RF = Sqr(RO) * RE
      Label4(0).Caption = "Frequency Below Resonance Where Impedance = " & Format(RF, "##.#") & " ohms"
      Label4(1).Caption = "Frequency Above Resonance Where Impedance = " & Format(RF, "##.#") & " ohms"
   Else
      MakeUnAvailable
   End If
   EmptyTest2 OK2
   If OK2 Then ShowAnswers
End Sub

Sub TVB_Text_GotFocus ()
   HighLightText TVB_Text
End Sub

Sub ZMAX_Text_Change ()
   ZMAX = Val(ZMAX_Text.Text)
   EmptyTest1 OK1
   If OK1 Then
      MakeAvailable
      RO = ZMAX / RE
      RF = Sqr(RO) * RE
      Label4(0).Caption = "Frequency Below Resonance Where Impedance = " & Format(RF, "##.#") & " ohms"
      Label4(1).Caption = "Frequency Above Resonance Where Impedance = " & Format(RF, "##.#") & " ohms"
   Else
      MakeUnAvailable
   End If
   EmptyTest2 OK2
   If OK2 Then ShowAnswers
End Sub

Sub ZMAX_Text_GotFocus ()
   HighLightText ZMAX_Text
End Sub

