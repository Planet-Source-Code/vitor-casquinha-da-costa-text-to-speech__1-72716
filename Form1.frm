VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text to Speak"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   9255
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   4200
      Top             =   2880
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   5880
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2880
      Width           =   9255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type here"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last Phrases"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'object to text to speech
Dim voice As SpVoice

'variable to store the last string readed
Dim LastString As String

Private Sub Form_Load()
'initialize the speech object
Set voice = New SpVoice

'be sure, the progress bar has value zero
PB1 = 0
End Sub

Private Sub List1_DblClick()
'when clicking on list item, computer will speak everything except the index
voice.Speak Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 6), SVSFlagsAsync
End Sub

Private Sub Text1_Change()
'every time we type anything, the progress bar value is set to zero
PB1 = 0
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

'verify if [Enter] key or [Return] key is pressed
If KeyCode = vbEnter Or KeyCode = vbKeyReturn Then
    'then add the text from textbox to the listbox
    List1.AddItem Format(List1.ListCount + 1, "000") & " - " & Text1
    List1.ListIndex = List1.ListCount - 1
    voice.Speak Text1, SVSFlagsAsync
    Text1 = ""
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

'if nothing in the textbox exit from event
If Len(Text1) = 0 Then Exit Sub

'if textbox content different from the text in variable "LastString"
'add 1 to progress bar
If Text1 <> LastString Then PB1 = PB1 + 1

'if progress bar reach 100, it's assumed you stop typing and computer
'will speak everything in the text box
If PB1 >= 100 Then
    LastString = Text1
    voice.Speak Text1, SVSFlagsAsync
    PB1 = 0
End If
End Sub
