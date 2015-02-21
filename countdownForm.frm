VERSION 5.00
Begin VB.Form countdownForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoShutdown"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   788
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label cdLabel 
      Caption         =   "Shutdown in"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "countdownForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' todo
' properly format time as 00:00
' beautify and center text displacement
' standard button size, center button on form
' if easy: make this app one instance only...
'          double clicking it activates main form...

Public Event Cancelled()
Public Interval As Long
Dim WithEvents Timer As SelfTimer
Attribute Timer.VB_VarHelpID = -1

Private Sub Form_Load()
    MakeTopMost Me.hWnd
    Set Timer = New SelfTimer
    Timer.Interval = 1000
    Timer.Enabled = True
    UpdateLabel
End Sub

Private Sub timer_Timer(ByVal Seconds As Currency)
    If Interval > 0 Then
        Interval = Interval - 1
        Timer.Interval = 1000
        UpdateLabel
    Else
        Timer.Enabled = False
        Unload Me
    End If
End Sub

Private Sub UpdateLabel()
    Dim s As Long: s = Interval Mod 60
    Dim m As Long: m = Interval \ 60
    cdLabel.Caption = "Shutdown in " & Format(m, "00") & ":" & Format(s, "00")
End Sub

Private Sub cancelButton_Click()
    RaiseEvent Cancelled
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        RaiseEvent Cancelled
    End If
End Sub
