VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprodutor MP3 do JCFB"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Acerca..."
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtInfo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CDlg1 
      Left            =   5520
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   5415
      Begin VB.CheckBox chkLoop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Click to loop playback"
         Top             =   1480
         Width           =   190
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Parar"
         Height          =   495
         Left            =   4320
         Picture         =   "FrmPlayer.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Stop playing"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdFF 
         Caption         =   "Pra frente"
         Height          =   495
         Left            =   3480
         Picture         =   "FrmPlayer.frx":0497
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "5 Seconds Fast Forward"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdRew 
         Caption         =   "Pra Tras"
         Height          =   495
         Left            =   2640
         Picture         =   "FrmPlayer.frx":04F3
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "5 Seconds Back"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pausa"
         Height          =   495
         Left            =   1800
         Picture         =   "FrmPlayer.frx":054F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Pause playing"
         Top             =   1440
         Width           =   855
      End
      Begin VB.PictureBox picProgBack 
         Height          =   255
         Left            =   960
         ScaleHeight     =   195
         ScaleWidth      =   4155
         TabIndex        =   12
         Top             =   650
         Visible         =   0   'False
         Width           =   4215
         Begin VB.PictureBox picProg 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   255
            TabIndex        =   13
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Começar"
         Height          =   495
         Left            =   960
         Picture         =   "FrmPlayer.frx":05A9
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Start playing"
         Top             =   1440
         Width           =   855
      End
      Begin VB.Timer playTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4920
         Top             =   960
      End
      Begin MSComctlLib.Slider VolSlider 
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   2990
         _Version        =   393216
         Orientation     =   1
         Max             =   20
         TickStyle       =   2
         TextPosition    =   1
      End
      Begin VB.Label lblTimer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   960
         TabIndex        =   10
         Top             =   960
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Procurar"
      Height          =   285
      Left            =   4680
      TabIndex        =   1
      ToolTipText     =   "Browse for files to play"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter file path"
      Top             =   720
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   120
      Picture         =   "FrmPlayer.frx":05FE
      Top             =   3240
      Width           =   5025
   End
   Begin VB.Label Label1 
      Caption         =   "Tambem toca Wave e MIDI"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      BorderWidth     =   2
      X1              =   120
      X2              =   4200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Image ImgIco 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "FrmPlayer.frx":33D6
      Top             =   0
      Width           =   480
   End
   Begin VB.Image ImgIco 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "FrmPlayer.frx":3818
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblAbout 
      Caption         =   "JCFB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   4200
      MouseIcon       =   "FrmPlayer.frx":3C5A
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "About"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Toca toda a musica!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================
'Tocador de mp3 do JCFB
'==========================================
'Descripcao:
'Tocador de musica retardado, criado para evitar o windows media player
'Criado usando Visual Basic '98 - Sim o 1998!
'Este source code está em dominio publico, faz o que quiseres com ele

Dim PlayerIsPlaying As Boolean  'determina se o player esta a reproduzir
Dim Player As FilgraphManager   'Refere o reprodutor
Dim PlayerPos As IMediaPosition Referência ao tempo a que a musica já está a reproduzir
Dim PlayerAU As IBasicAudio     Referência para deteminar o volume do audio (usado no slider do lado esquerdo)
Dim i As Integer                'Icon index
Sub Play()

Dim CurState As Long
 'check player
 If Not Player Is Nothing Then
    'Get the state
    Player.GetState x, CurState
      
    If CurState = 1 Then
      PausePlay
      Exit Sub
    End If
 End If
 
 StartPlay 'Comeca a reproduzir o ficheiro

End Sub
Sub StartPlay()

On Error GoTo error                   'Manda o erro foder-se
   'Set objects
   Set Player = New FilgraphManager   Reprodutor
   Set PlayerPos = Player             Posição
   Set PlayerAU = Player              'Volume

   Player.RenderFile txtFile.Text     'carrega o ficheiro
   AdjustVolume                       'Seta o Volume
   Player.Run                         'Corre o Player
   PlayerIsPlaying = True             'ta a reproduzir
   playTimer.Enabled = True           Começa o temporizador
   lblStatus.Caption = "Playing..."   'Estado
   cmdBrowse.Enabled = False          'Sen ficheiro aberto para parar
   txtFile.Enabled = False            'Sem mudança
   VolSlider.SetFocus                 'Remove o foco do cmdPlay
   i = 0                              'Animar icon
   GetMP3Tags                         'Carrega tags mp3
   Me.Caption = Me.Caption & " [" & txtInfo.Text & "]"

Exit Sub
error:                                 'Anda com o erro
   StopPlay                            'Stop player
   cmdBrowse.Enabled = True            'Enable file browser
   txtFile.Enabled = True              'Enable file name text box
   lblStatus.Caption = Err.Description 'Write error description to status label
   
End Sub

Private Sub cmdBrowse_Click()

On Error GoTo error
 
 With CDlg1
   'On Cancel do nothing
   .CancelError = True
   
   .DialogTitle = "Browse For MP3 Files"
   .Filter = "MP3 Files (*.mp3) |*.mp3; |Wave Files (*.wav) |*.wav; |Midi Files (*.mid) |*.mid"
   .ShowOpen
   'Handle no filename
   If Len(.FileName) = 0 Then Exit Sub
  
   txtFile.Text = .FileName
   cmdPlay.SetFocus 'Press enter to play
 End With
 
error:  'Do nothing

End Sub

Private Sub cmdFF_Click()

'5 seconds fast forward
Dim CurPos  'Current position of the player
  
  If Player Is Nothing Then Exit Sub 'Not playing nothing to forward!

  CurPos = PlayerPos.CurrentPosition
  CurPos = CurPos + 5
  'If over the duration of the song set position to duration
  If CurPos > PlayerPos.Duration Then CurPos = PlayerPos.Duration
  'Go now
  PlayerPos.CurrentPosition = CurPos
  
End Sub

Private Sub cmdPause_Click()

  PausePlay
  
End Sub

Private Sub cmdPlay_Click()

  Play
  
End Sub

Private Sub cmdRew_Click()

'anda 5 s pra frente
Dim CurPos 'Posicao atual do player
  
  If Player Is Nothing Then Exit Sub 'Nao ta nada a dar, nao vai pra traz. lol

  CurPos = PlayerPos.CurrentPosition
  CurPos = CurPos - 5
  'Se chegarmos ao inico, reseta tudo
  If CurPos < 0 Then CurPos = 0
  'Agora vai
  PlayerPos.CurrentPosition = CurPos
  
End Sub

Private Sub cmdStop_Click()

  StopPlay
  
End Sub

Private Sub Command1_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()

    'Se nao ha ficheiro selecionado, desativa os botoes e elementos pra nao bugar o programa
    cmdPlay.Enabled = False
    cmdPause.Enabled = False
    cmdStop.Enabled = False
    cmdRew.Enabled = False
    cmdFF.Enabled = False
    chkLoop.Value = 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'para a reprodução
   StopPlay
   'configura os objetos para null
   Set Player = Nothing
   Set PlayerPos = Nothing
   Set PlayerAU = Nothing
   
End Sub


Private Sub lblAbout_Click()
'Mostra msg acerca do programa
Dim Msg As String

 Msg = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf
 Msg = Msg & "Por JCFB" & vbCrLf
 Msg = Msg & App.LegalCopyright & vbCrLf
 Msg = Msg & "Livre distribuição" & vbCrLf & vbCrLf
 Msg = Msg & " Faz a procaria que quiseres com esta aplicação."
 MsgBox Msg, vbInformation, "About " & App.Title 'info da aplicacao ao clicar no label com jcfb escrito

End Sub

Private Sub playTimer_Timer()

Contador só está disponivel quando a reproduzir um ficheiro
If PlayerIsPlaying = True Then
  'Para tudo quando a musica acaba
  If PlayerPos.CurrentPosition >= PlayerPos.Duration Then
     StopPlay
     PlayerIsPlaying = False
     playTimer.Enabled = False
     'Verifica se o botão para o loop está selecionado e realiza-o caso esteja
     If chkLoop.Value = 1 Then
          Play
     End If
     Exit Sub
  End If
  'Se ainda esta a reproduzir
  lblTimer.Caption = ToTimeValues(PlayerPos.CurrentPosition)
  'Animate icon
  If i > 1 Then i = 0
  Me.Icon = ImgIco(i).Picture
  i = i + 1
  'A barra de posição só funciona com mp3
  If LCase(Right(txtFile, 3)) = "mp3" Then
     ShowTime PlayerPos.CurrentPosition
  End If
 
End If

End Sub
Sub StopPlay()

  If Player Is Nothing Then Exit Sub 'Nada pra parar!
  playTimer.Enabled = False          'Sem contador apos parar
  
  'Para o barulho
  Player.Stop
  'configura o tempo e a barra de estado
  lblStatus.Caption = "Stopped"
  lblTimer.Caption = ""
  'Para selecionar outro ficheiro
  cmdBrowse.Enabled = True
  txtFile.Enabled = True
  'Esconde a barra de posiçao
  picProgBack.Visible = False
  'Icone
  Me.Icon = ImgIco(0).Picture
  'Limpa a info. do texto
  txtInfo.Text = ""
  Me.Caption = App.Title
  'foca ao reproduzir
  cmdPlay.SetFocus
    
End Sub

Sub PausePlay()
   
Static Paused As Boolean                'Se pausado
Dim CurState As Long                    'Estado atual do player

  If Player Is Nothing Then Exit Sub    'Nada pra pausar
     
     'Obtem o estado do player
     Player.GetState x, CurState
     
     If CurState = 2 Then
       'Se a reproduzir, para-o
       Paused = True
       Player.Pause
       lblStatus.Caption = "Paused"
     Else
       'Se já. pausado, volta a reproduzir
       Paused = False
       Player.Run
       lblStatus.Caption = "Playing..."
     End If
     
End Sub
Function ToTimeValues(ByVal Seconds As Long) As String

'Hora input do tipo "00:00:00"
Dim HH As Long                   'Horas
Dim MM As Long                   'Minutos
Dim SS As Long                   'Segundos
Dim tmp As String                'valor temporario

 'Valores velhos do tempo
 HH = Seconds \ 3600
 MM = Seconds \ 60 Mod 60
 SS = Seconds Mod 60
 
 'Se o ficheiro tem mais de uma hora
 If HH > 0 Then tmp = Format$(HH, "00:")
 ToTimeValues = tmp & Format$(MM, "00:") & Format$(SS, "00")
 
End Function

Sub ShowTime(ByVal iDur As Integer)
       
    'Set the position bar to player position
    picProgBack.Visible = True
    'Clear pictures
    picProgBack.Cls
    picProg.Cls
    'Customizable color of the position bar
    picProg.BackColor = vbBlue
    'Bar width depends on current position, when max is song duration
    picProg.Width = picProgBack.ScaleWidth * ((CInt(iDur)) / PlayerPos.Duration)
    picProgBack.CurrentX = picProgBack.ScaleWidth / 2
    picProgBack.CurrentY = 1
    picProg.CurrentX = picProg.ScaleWidth / 2
    picProg.CurrentY = 1
    
    DoEvents
    
End Sub

Private Sub txtFile_Change()

On Error Resume Next

Dim Ext As String
'Check if there is a string in txtFile
 If txtFile <> "" Then
    'File extension is the last three letters
    Ext = LCase(Right(txtFile, 4))
    'If extension is supported enable player buttons
    If Ext = ".mp3" Or Ext = ".wav" Or Ext = ".mid" Then
      cmdPlay.Enabled = True
      cmdPause.Enabled = True
      cmdStop.Enabled = True
      cmdRew.Enabled = True
      cmdFF.Enabled = True
      chkLoop.Value = 1
    End If
 Else
    'If no file name disable player buttons
    cmdPlay.Enabled = False
    cmdPause.Enabled = False
    cmdStop.Enabled = False
    cmdRew.Enabled = False
    cmdFF.Enabled = False
    chkLoop.Value = 2
 End If
 
End Sub

Private Sub VolSlider_Change()

  SetSliderText
  AdjustVolume
  
End Sub
Sub AdjustVolume()

    'Set the volume to slider value
    If Player Is Nothing Then Exit Sub  'No player
    Set PlayerAU = Player
    PlayerAU.Volume = ((20 - VolSlider.Value) * 5 * 40) - 4000
    
    SetSliderText
    
End Sub

Private Sub VolSlider_Click()

   SetSliderText
   
End Sub

Private Sub VolSlider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   SetSliderText
   
End Sub
Sub SetSliderText()

  'Slider volue to percent
  tmp = 100 - (VolSlider.Value * 5)
  VolSlider.Text = "Volume %" & CStr(tmp)
  
End Sub

Private Sub VolSlider_Scroll()

  SetSliderText
  AdjustVolume
  
End Sub
Sub GetMP3Tags()

On Error GoTo error

Dim PosInFile As Long        'Byte in mp3 file
Dim Mp3File As String        'Mp3 file
Dim Tags As String * 128     'String for tags
Dim vTitle As String         'MP3 Title tag
Dim vArtist As String        'MP3 artist tag

Mp3File = txtFile.Text

'Not a mp3 file
If LCase(Right(Mp3File, 4)) <> ".mp3" Then GoTo error2
    'function free file
    f = FreeFile
    'position of tags in mp3 file
    PosInFile = FileLen(Mp3File) - 127
    
    If PosInFile > 0 Then
       'get tags from file
       Open Mp3File For Binary As #f
              Get #f, PosInFile, Tags
       Close #f
       'unknown
       If UCase(Left(Tags, 3)) <> "TAG" Then GoTo error
       
       'tag string values
            vTitle = Replace(Trim(Mid(Tags, 4, 30)), Chr(0), "")
            vArtist = Replace(Trim(Mid(Tags, 34, 30)), Chr(0), "")

'I haven't used other tags but maybe you need them

'            vAlbum = Replace(Trim(Mid(Tags, 64, 30)), Chr(0), "")
'            vYear = Replace(Trim(Mid(Tags, 94, 4)), Chr(0), "")
'            vComment = Replace(Trim(Mid(Tags, 98, 30)), Chr(0), "")
'            vGenre = Asc(Mid(Tags, 128, 1))
            
            txtInfo.Text = vTitle & " by " & vArtist
            'We do not have enough place
            If Len(txtInfo) > 45 Then
              txtInfo.Text = Left(txtInfo, 42) & "..."
            End If
    Else
            GoTo error
    End If
    
    Close               'In case of unexpected error

Exit Sub
error:                  'Handle no tag
  
  vTitle = "Unknown"
  vArtist = "Unknown"
  txtInfo.Text = vTitle & " by " & vArtist
  Close                 'In case of unexpected error

Exit Sub
error2: 'Not an mp3 file

   txtInfo.Text = ""
   Close               'In case of unexpected error
   
End Sub