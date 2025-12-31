VERSION 2.00
Begin Form SpkData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Speaker Data"
   ClientHeight    =   4020
   ClientLeft      =   945
   ClientTop       =   2025
   ClientWidth     =   8070
   ControlBox      =   0   'False
   Height          =   4425
   Icon            =   SPKDATA.FRX:0000
   Left            =   885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8070
   Top             =   1680
   Width           =   8190
   Begin TrueGrid Table1 
      AllowArrows     =   -1  'True
      AllowTabs       =   -1  'True
      BackColor       =   &H00C0C0C0&
      DataSource      =   "Data1"
      Editable        =   -1  'True
      EditDropDown    =   -1  'True
      ExposeCellMode  =   0  'Expose upon selection
      FetchMode       =   0  'By cell
      HeadingHeight   =   1
      Height          =   495
      HorzLines       =   2  '3D
      Layout          =   SPKDATA.FRX:0302
      LayoutIndex     =   1
      Left            =   2040
      LinesPerRow     =   1
      MarqueeUnique   =   -1  'True
      SplitPropsGlobal=   -1  'True
      SplitTabMode    =   0  'Don't tab across splits
      TabCapture      =   0   'False
      TabIndex        =   39
      Top             =   0
      UseBookmarks    =   -1  'True
      Visible         =   0   'False
      Width           =   1575
      WrapCellPointer =   0   'False
   End
   Begin CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   3480
      Width           =   975
   End
   Begin Frame Frame1 
      Height          =   3135
      Left            =   240
      TabIndex        =   38
      Top             =   120
      Width           =   6495
      Begin TextBox Text1 
         DataField       =   "Notes"
         DataSource      =   "Data1"
         Height          =   735
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   2280
         Width           =   2655
      End
      Begin Data Data1 
         Caption         =   "Data1"
         Connect         =   ""
         DatabaseName    =   "SPEAKER.MDB"
         Exclusive       =   0   'False
         Height          =   270
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordSource    =   "Speaker Parameters"
         Top             =   0
         Visible         =   0   'False
         Width           =   1140
      End
      Begin TextBox Vas_Text 
         DataField       =   "Vas"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   600
      End
      Begin TextBox Fs_Text 
         DataField       =   "Fs"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   600
      End
      Begin TextBox Qts_Text 
         DataField       =   "Qts"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   1560
         Width           =   600
      End
      Begin TextBox Manufacturer 
         DataField       =   "Manufacturer"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin TextBox ModelNumber 
         DataField       =   "Model Number"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4320
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin TextBox Re_Text 
         DataField       =   "Re"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5640
         TabIndex        =   7
         Top             =   840
         Width           =   600
      End
      Begin TextBox Zmax_Text 
         DataField       =   "Zmax"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5640
         TabIndex        =   8
         Top             =   1200
         Width           =   600
      End
      Begin TextBox Tvb_Text 
         DataField       =   "Tvb"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5640
         TabIndex        =   9
         Top             =   1560
         Width           =   600
      End
      Begin TextBox F1_Text 
         DataField       =   "F1"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   1920
         Width           =   600
      End
      Begin TextBox F2_Text 
         DataField       =   "F2"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   2280
         Width           =   600
      End
      Begin TextBox Tfs_Text 
         DataField       =   "Tfs"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   5640
         TabIndex        =   10
         Top             =   1920
         Width           =   615
      End
      Begin CommandButton Command4 
         Caption         =   "&Parameters"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin CommandButton Command2 
         Caption         =   "&Design"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   1095
      End
      Begin Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   40
         Top             =   2280
         Width           =   615
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Free Air Resonance (Fs):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   1815
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Q (Qts):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   28
         Top             =   1560
         Width           =   1335
      End
      Begin Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Equivalant Air Mass (Vas):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1935
      End
      Begin Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Manufacturer"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Model #"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "V.C. dc Resistance (Re):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   3600
         TabIndex        =   32
         Top             =   840
         Width           =   1935
      End
      Begin Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "V.C. Impedance @ Fs (Zmax):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   33
         Top             =   1200
         Width           =   2175
      End
      Begin Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Test Box Volume (Tvb, cu. ft.):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   1560
         Width           =   2295
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "F1 (< Fs):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   35
         Top             =   1920
         Width           =   1575
      End
      Begin Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "F2 (> Fs):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   36
         Top             =   2280
         Width           =   1575
      End
      Begin Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Spk Resonance In Test Box (Tfs):"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   1920
         Width           =   2655
      End
   End
   Begin CommandButton cmdMark 
      Caption         =   "&Mark"
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   3000
      Width           =   975
   End
   Begin CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   25
      Top             =   3480
      Width           =   975
   End
   Begin CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3480
      TabIndex        =   23
      Top             =   3480
      Width           =   975
   End
   Begin CommandButton cmdAddNew 
      Caption         =   "&Add New"
      Height          =   375
      Left            =   2400
      TabIndex        =   24
      Top             =   3480
      Width           =   975
   End
   Begin CommandButton cmdReturn 
      Caption         =   "&Return"
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   3480
      Width           =   975
   End
   Begin CommandButton cmdFindNext 
      Caption         =   "Find Ne&xt"
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin CommandButton cmdFindFirst 
      Caption         =   "Find Fi&rst"
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   6960
      TabIndex        =   20
      Top             =   3480
      Width           =   975
   End
   Begin CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin CommandButton cmdPrior 
      Caption         =   "&Prior"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
End
Option Explicit

'Form-level variables
Dim strCriterion As String
Dim strBookmark As String

Sub cmdAddNew_Click ()
   'Add a new record after refreshing the Dyanset
   Data1.Refresh
   'Dim lngPubID As Long
   'Find the last value of PubID and add 1
   Data1.Recordset.MoveLast
   'lngPubID = Data1.Recordset!PubID
   'If Text5.Text <> "" Then
      'Increment PubID if the last record is valid
   '  lngPubID = lngPubID + 1
   'End If
   'Position the record pointer on the tentative append record
   Data1.Recordset.AddNew
   'Pass the PubID value to the PubID text box
   'Text4.Text = lngPubID
End Sub

Sub cmdBrowse_Click ()
   Table1.Width = SpkData.Width
   Table1.Height = SpkData.Height
   Table1.Top = 0
   Table1.Left = 0
   Table1.Visible = True
   Table1.SetFocus
   'Mark the current record
   strBookmark = Data1.Recordset.Bookmark
End Sub

Sub cmdCancel_Click ()
   'Cancel an AddNew or Update operation
   'Restore the value of text boxes to the original field values
   Data1.Refresh
   'If IsNull(Text4.Text) Or Text4.Text = "" Or Text5.Text = "" Then
      'Don't attempt to update controls from an aborted append
      If Len(strBookmark) > 0 Then
	 'Return to the bookmarked record
	 Data1.Recordset.Bookmark = strBookmark
      Else
	 'Go to the last record
	 Data1.Recordset.MoveLast
      End If
   'End If
End Sub

Sub cmdDelete_Click ()
   'Delete the current record with confirmation
   If MsgBox("Are you sure you want to delete this record?", 36, "Delete Record") = 6 Then
      'Test if the to-be-deleted record is the bookmarked record
      If Len(strBookmark) > 0 And strBookmark = Data1.Recordset.Bookmark Then
	 strBookmark = ""
      End If
      'Delete the record
      Data1.Recordset.Delete
      'Refresh the Recordset
      Data1.Refresh
      If Len(strBookmark) > 0 Then
	 Data1.Recordset.Bookmark = strBookmark
      Else
	 Data1.Recordset.MoveLast
      End If
   End If
End Sub

Sub cmdFindFirst_Click ()
   'Declare local variables
   Dim strPrompt As String
   Dim strTitle As String
   Dim strEntry As String
   
   'Set up the prompt and the title, then open the input box
   strPrompt = "Enter the name (or partial name) of the manufacturer."
   strTitle = "Find First Occurance of Manufacturer"
   strEntry = InputBox$(strPrompt, strTitle)

   If Len(Trim$(strEntry)) = 2 Then
      'The user made an entry, set a bookmark
      strBookmark = Data1.Recordset.Bookmark
      'Create the criterion string
      strCriterion = "State = '" & UCase$(strEntry) & "'"
      'Invoke the FindFirst method
      Data1.Recordset.FindFirst strCriterion
      If Data1.Recordset.NoMatch Then
	 'No matching record was found
	 MsgBox "No records match '" & UCase$(strEntry) & "'.", , "Record Not Found"
	 Data1.Recordset.Bookmark = strBookmark
      Else
	 'A match was found, place the caret in the State field
	 'Text9.SetFocus
      End If
   End If
End Sub

Sub cmdFindNext_Click ()
   'Check to see if the criterion has been entered
   If Len(strCriterion) > 0 Then
      'Invoke the FindNext method
      Data1.Recordset.FindNext strCriterion
      If Data1.Recordset.NoMatch Then
	 'No matching record was found
	 MsgBox "No more records match " & strCriterion, , "Record Not Found"
	 Data1.Recordset.Bookmark = strBookmark
      Else
	 'A match was found, place the caret in the State field
	 'Text9.SetFocus
      End If
   Else
      'There was no criterion
      MsgBox "You need to enter a criterion with the Find First button first", , "No Criteria Entered"
   End If
End Sub

Sub cmdFirst_Click ()
   Data1.Recordset.MoveFirst
End Sub

Sub cmdLast_Click ()
   Data1.Recordset.MoveLast
End Sub

Sub cmdMark_Click ()
   'Mark the current record
   strBookmark = Data1.Recordset.Bookmark
End Sub

Sub cmdNext_Click ()
   Data1.Recordset.MoveNext
   If Data1.Recordset.EOF Then
      Beep
      Data1.Recordset.MoveFirst
   End If
End Sub

Sub cmdPrior_Click ()
   Data1.Recordset.MovePrevious
   If Data1.Recordset.BOF Then
      Beep
      Data1.Recordset.MoveLast
   End If
End Sub

Sub cmdQuit_Click ()
   End
End Sub

Sub cmdReturn_Click ()
   If Len(strBookmark) > 0 Then
      'Move the record pointer to the Bookmarked position
      Data1.Recordset.Bookmark = strBookmark
   Else
      Beep
   End If
End Sub

Sub cmdUpdate_Click ()
   'Transfer the data from changed text box values to the current record
   'You need to move the record pointer to cause the update to "take"
   Data1.Recordset.MoveNext
   Data1.Recordset.MovePrevious
End Sub

Sub Command2_Click ()
   SpkData.Hide
   Design.Show
End Sub

Sub Command4_Click ()
   SpkData.Hide
   ParamEnt.Show
End Sub

Sub Data1_Validate (Action As Integer, Save As Integer)
   Select Case Action
      Case 1

      Case 2
      'Move First
      
      Case 3
      'Move Prior
      
      Case 4
      'Move Next
      Case 5
      'Move Last
      
      Case 6
      'Data_ActionUpdate
      Case 7
      Case 8
      Case 9
      Case 10
      'Data_ActionClose
   End Select
End Sub

Sub F1_Text_Change ()
   F1 = Val(F1_Text.Text)
   BeenHere = False
End Sub

Sub F2_Text_Change ()
   F2 = Val(F2_Text.Text)
   BeenHere = False
End Sub

Sub Form_Load ()
   Data1.DatabaseName = app.Path + "\SPEAKER.MDB"
   CenterForm SpkData
   SpkData.Hide
   Load ParamEnt
   ParamEnt.Hide
   Load Design
End Sub

Sub FS_Text_Change ()
   Fs = Val(Fs_Text.Text)
   BeenHere = False
End Sub

Sub Qts_Text_Change ()
   Qts = Val(Qts_Text.Text)
   BeenHere = False
End Sub

Sub RE_Text_Change ()
   Re = Val(Re_Text.Text)
   BeenHere = False
End Sub

Sub Table1_Append ()
   If MsgBox("Add new record?", 1) = 2 Then Exit Sub
   Data1.Recordset.AddNew
End Sub

Sub Table1_KeyDown (KeyCode As Integer, Shift As Integer)

'User pressed ENTER
If KeyCode = 13 Then
   Table1.Visible = False
'User pressed ESC
ElseIf KeyCode = 27 Then
   If Len(strBookmark) > 0 Then
      'Move the record pointer to the Bookmarked position
      Data1.Recordset.Bookmark = strBookmark
      Table1.Visible = False
   Else
      Beep
   End If
End If

End Sub

Sub TFS_Text_Change ()
   Tfs = Val(Tfs_Text.Text)
   BeenHere = False
End Sub

Sub TVB_Text_Change ()
   Tvb = Val(Tvb_Text.Text)
   BeenHere = False
End Sub

Sub Vas_Text_Change ()
  Vas = Val(Vas_Text.Text)
  BeenHere = False
End Sub

Sub ZMAX_Text_Change ()
   Zmax = Val(Zmax_Text.Text)
   BeenHere = False
End Sub

