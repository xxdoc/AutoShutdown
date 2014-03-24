VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "AutoShutdown"
   ClientHeight    =   2160
   ClientLeft      =   8235
   ClientTop       =   4485
   ClientWidth     =   4395
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox warnIntervalTB 
      Height          =   405
      Left            =   2280
      TabIndex        =   7
      Text            =   "1"
      Top             =   1560
      Width           =   360
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   405
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   714
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "warnIntervalTB"
      BuddyDispid     =   196619
      OrigLeft        =   3240
      OrigTop         =   1560
      OrigRight       =   3495
      OrigBottom      =   2055
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox warnCB 
      Caption         =   "&Warn before shutdown by"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox silentStartCB 
      Caption         =   "Start timer &on startup"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox autoStartCB 
      Caption         =   "&Run on Windows logon"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton stopButton 
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtPicker 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "h:mm tt"
      Format          =   95027203
      UpDown          =   -1  'True
      CurrentDate     =   0.628472222222222
   End
   Begin VB.CommandButton startButton 
      Caption         =   "&Start"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "minute(s)"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Menu trayMenu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu openMenuItem 
         Caption         =   "Open"
      End
      Begin VB.Menu exitMenuItem 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' todo
' update tip with correct interval on hover?
' catch after minimize event: then min button minimizes to only if running

Dim cfgFile As String
Dim WithEvents myTimer As SelfTimer
Attribute myTimer.VB_VarHelpID = -1
Private WithEvents tray As frmSysTray
Attribute tray.VB_VarHelpID = -1
Private easter As EasterEggForm
Attribute easter.VB_VarHelpID = -1
Dim fired As Boolean
Dim testMode As Boolean
Dim warnSeconds As Long

Private Sub Form_Load()
    MakeTopMost Me.hWnd
    Set myTimer = New SelfTimer
    myTimer.Enabled = False
    cfgFile = App.Path & "\" & App.EXEName & ".cfg"
    LoadSettings
    
    If silentStartCB.Value Then
        startButton_Click
    End If
End Sub

Private Sub LoadSettings()
    ' add error handling
    Dim st As String
    st = ReadIni(cfgFile, "Main", "alarmTime")
    If st <> "" Then
        DTPicker.Value = TimeValue(CDate(st))
    End If

    st = ReadIni(cfgFile, "Main", "autoStart")
    If st <> "" Then
        autoStartCB.Value = CInt(st)
    Else
        autoStartCB.Value = 0
    End If
    st = ReadIni(cfgFile, "Main", "silentStart")
    If st <> "" Then
        silentStartCB.Value = CInt(st)
    Else
        silentStartCB.Value = 0
    End If
    st = ReadIni(cfgFile, "Main", "warnInterval")
    If st <> "" Then
        warnIntervalTB.Text = st
    End If
    st = ReadIni(cfgFile, "Main", "warn")
    If st <> "" Then
        warnCB.Value = CInt(st)
    End If
End Sub

Private Sub CreateShortcut()
    Dim WshObj As Object
    Set WshObj = CreateObject("WScript.Shell")
    Dim shortcutPath As String
    Dim startupDir As String
    startupDir = WshObj.SpecialFolders("Startup")
    shortcutPath = startupDir & "\" & App.EXEName & ".lnk"
    If Dir(shortcutPath) <> "" Then
        Exit Sub
    End If

    Dim shlObj As Object
    Dim fso As Object
    Dim shortcut As Object
    Dim targetPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    targetPath = App.Path & "\" & App.EXEName & ".exe"
    Set shortcut = WshObj.CreateShortcut(shortcutPath)
    shortcut.targetPath = targetPath
    shortcut.WorkingDirectory = App.Path
    shortcut.Save
End Sub

Private Sub DeleteShortcut()
    Dim WshObj As Object
    Set WshObj = CreateObject("WScript.Shell")
    Dim shortcutPath As String
    Dim startupDir As String
    startupDir = WshObj.SpecialFolders("Startup")
    shortcutPath = startupDir & "\" & App.EXEName & ".lnk"
    If Dir(shortcutPath) = "" Then
        Exit Sub
    End If
    Kill shortcutPath
End Sub

Private Sub myTimer_Timer(ByVal Seconds As Currency)
    If Not fired Then
        fired = True
        myTimer.Interval = warnSeconds * 1000!
        myTimer.Enabled = True
        
        If warnCB.Value Then
            If warnIntervalTB.Text = "1" Then
                tray.ShowBalloonTip "Computer will shutdown in 1 minute", Me.Caption, NIIF_INFO
            Else
                tray.ShowBalloonTip "Computer will shutdown in " & warnIntervalTB.Text & " minutes", Me.Caption, NIIF_INFO
            End If
        End If
    Else
        ToggleTimer
        If Not testMode Then
            shutdown
        Else
            MakeNormal Me.hWnd
            MsgBox "shutdown"
            MakeTopMost Me.hWnd
        End If
    End If
End Sub

Private Sub ToggleTimer()
    Dim b As Boolean
    b = Not myTimer.Enabled
    myTimer.Enabled = b
    DTPicker.Enabled = Not b
    startButton.Enabled = Not b
    stopButton.Enabled = b
    If warnCB.Value Then
        ' essential, so that warnSeconds doesn't get changed after we start Timer
        warnIntervalTB.Enabled = Not b
        UpDown1.Enabled = Not b
    End If
End Sub

Private Sub shutdown()
    If ShutdownPC Then
        Unload tray
        Unload Me
    Else
        MakeNormal Me.hWnd
        MsgBox "Shutdown error"
        MakeTopMost Me.hWnd
    End If
End Sub

Private Sub startButton_Click()
    ' if time has passed, fire tomorrow!
    'MsgBox dtPicker.Day
    'Debug.Print Format(dtPicker.Value, "dd MM yyyy hh:mm am/pm")
    'MsgBox Format(dtPicker.Value, "h:mm AM/PM")
    Dim Seconds As Long
    ' won't use TimeValue() so that the subtraction takes into account if the
    ' time to fire is tomorrow
    Seconds = DateDiff("s", TimeValue(Now), TimeValue(DTPicker.Value))
    If Seconds <= 0 Then
        'Debug.Print Seconds
        ' ensure constants are long
        Seconds = Seconds + 24! * 3600!
    End If
    
    warnSeconds = CInt(warnIntervalTB.Text) * 60!
    If Seconds <= warnSeconds Then
        fired = True
    Else
        ' set it false to that when timer fires
        ' it displays the warn balloon tip, then sets itself
        ' to fire again after 60 seconds
        fired = False
        Seconds = Seconds - warnSeconds
    End If
    
    'Debug.Print DateDiff("h", Now, dtPicker.Value)
    'Debug.Print DateDiff("n", Now, dtPicker.Value)
    'Debug.Print "seconds: " & seconds
    ' set timer interval
    myTimer.Interval = Seconds * 1000
    ToggleTimer
    minToTray
End Sub

Sub minToTray()
    Me.Hide
    If tray Is Nothing Then
        Set tray = New frmSysTray
        With tray
            .AddMenuItem "&Open", "Open", True
            .AddMenuItem "&Exit", "Exit"
            .IconHandle = Me.Icon.Handle
        End With
    End If
    UpdateNotIconTip
End Sub

Sub UpdateNotIconTip()
    Dim Seconds As Long
    Seconds = DateDiff("s", TimeValue(Now), TimeValue(DTPicker.Value))
    If Seconds <= 0 Then
        Seconds = Seconds + 24! * 3600!
    End If
    
    Dim tip As String
    tip = "Shutdown at " & Format(DTPicker.Value, "h:nn AM/PM")
    tray.ToolTip = tip
    
    Dim hr, min, sec As Integer
    hr = Seconds \ 3600
    Seconds = Seconds Mod 3600
    min = Seconds \ 60
    sec = Seconds Mod 60
    
    Dim timeStr As String
    If hr > 0 Then
        If hr = 1 Then
            timeStr = hr & " hour"
        Else
            timeStr = hr & " hours"
        End If
        If min > 0 Or sec > 0 Then
            timeStr = timeStr & ", "
        End If
    End If
    
    If min > 0 Then
        If min = 1 Then
            timeStr = timeStr & min & " minute"
        Else
            timeStr = timeStr & min & " minutes"
        End If
        If sec > 0 Then
            timeStr = timeStr & ", "
        End If
    End If
    If sec > 0 Then
        If sec = 1 Then
            timeStr = timeStr & sec & " second"
        Else
            timeStr = timeStr & sec & " seconds"
        End If
    End If
    Dim bTip As String
    bTip = tip & vbNewLine & "Remaining: " & timeStr & vbNullChar
    tray.ShowBalloonTip bTip, Me.Caption, NIIF_INFO
End Sub

Private Sub stopButton_Click()
    fired = False
    ToggleTimer
    CloseTray
End Sub

Private Sub CloseTray()
    Unload tray
    Set tray = Nothing
End Sub

Private Sub RestoreMainForm()
    Me.Show
    Me.ZOrder
End Sub

Private Sub tray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
    Select Case sKey
        Case "Open"
            RestoreMainForm
            
        Case "Exit"
            Unload Me
    End Select
End Sub

Private Sub tray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    Select Case eButton
        Case vbLeftButton
            RestoreMainForm
            
        Case vbRightButton
            tray.ShowMenu
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyT
            If Shift = (vbCtrlMask Or vbShiftMask) Then
                MakeNormal Me.hWnd
                ToggleTestMode
                MakeTopMost Me.hWnd
            End If
        
        Case vbKeyA
            If Shift = vbCtrlMask Then
                MakeNormal Me.hWnd
                MsgBox "Author: El-Tabei M. El-Tabei"
                MakeTopMost Me.hWnd
            End If
        
        Case vbKeyH
            If Shift = (vbCtrlMask Or vbShiftMask Or vbAltMask) Then
                EasterEgg
            End If
    End Select
End Sub

Private Sub EasterEgg()
    If easter Is Nothing Then
        Set easter = New EasterEggForm
    End If
    easter.Left = Me.Left + Me.Width + 150
    easter.Top = Me.Top
    easter.Show
End Sub

Private Sub ToggleTestMode()
    If Not testMode Then
            'MsgBox "Test mode: On"
            Me.Caption = "AutoShutdown [Test Mode]"
        Else
            'MsgBox "Test mode: Off"
            Me.Caption = "AutoShutdown"
        End If
    testMode = Not testMode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.Visible And myTimer.Enabled Then
        Cancel = 1
        minToTray
        Exit Sub
    End If
        
    If Not tray Is Nothing Then
        CloseTray
    End If
    If Not easter Is Nothing Then
        Unload easter
        Set easter = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
End Sub

Private Sub SaveSettings()
    If autoStartCB.Value Then
        CreateShortcut
    Else
        DeleteShortcut
    End If
    WriteIni cfgFile, "Main", "alarmTime", Format(DTPicker.Value, "h:nn AM/PM")
    WriteIni cfgFile, "Main", "autoStart", autoStartCB.Value
    WriteIni cfgFile, "Main", "silentStart", silentStartCB.Value
    WriteIni cfgFile, "Main", "warn", warnCB.Value
    WriteIni cfgFile, "Main", "warnInterval", warnIntervalTB.Text
End Sub

Private Sub warnCB_Click()
    warnIntervalTB.Enabled = warnCB.Value
    UpDown1.Enabled = warnCB.Value
End Sub
