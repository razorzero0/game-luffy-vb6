VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16500
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   16500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P1result 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Enabled         =   0   'False
      Height          =   735
      Left            =   1920
      ScaleHeight     =   675
      ScaleWidth      =   2355
      TabIndex        =   55
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "P1 WIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   56
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.PictureBox P2death1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3105
      Left            =   2280
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3045
      ScaleWidth      =   12000
      TabIndex        =   53
      Top             =   2280
      Visible         =   0   'False
      Width           =   12060
      Begin VB.PictureBox P2death2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3105
         Left            =   1320
         Picture         =   "Form1.frx":76F62
         ScaleHeight     =   3045
         ScaleWidth      =   12000
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   12060
      End
   End
   Begin VB.PictureBox P1death1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3105
      Left            =   2880
      Picture         =   "Form1.frx":EDEC4
      ScaleHeight     =   3045
      ScaleWidth      =   12000
      TabIndex        =   51
      Top             =   2640
      Visible         =   0   'False
      Width           =   12060
      Begin VB.PictureBox P1death2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3105
         Left            =   1320
         Picture         =   "Form1.frx":164E26
         ScaleHeight     =   3045
         ScaleWidth      =   12000
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   12060
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   795
      Left            =   13560
      Picture         =   "Form1.frx":1DBD88
      ScaleHeight     =   735
      ScaleWidth      =   750
      TabIndex        =   49
      Top             =   480
      Width           =   810
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   795
      Left            =   240
      Picture         =   "Form1.frx":1DDAE2
      ScaleHeight     =   735
      ScaleWidth      =   750
      TabIndex        =   48
      Top             =   480
      Width           =   810
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   120
      Top             =   7320
   End
   Begin VB.PictureBox P2health1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   8760
      ScaleHeight     =   435
      ScaleWidth      =   4995
      TabIndex        =   45
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox P2health2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   6120
      ScaleHeight     =   435
      ScaleWidth      =   4995
      TabIndex        =   44
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox P1health1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   7560
      ScaleHeight     =   435
      ScaleWidth      =   4995
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox P1health2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   5160
      ScaleHeight     =   435
      ScaleWidth      =   4995
      TabIndex        =   42
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox P2gatling_gun1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3675
      Left            =   2520
      Picture         =   "Form1.frx":1DF83C
      ScaleHeight     =   3615
      ScaleWidth      =   48015
      TabIndex        =   40
      Top             =   1200
      Visible         =   0   'False
      Width           =   48075
      Begin VB.PictureBox P2gatling_gun2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3675
         Left            =   1320
         Picture         =   "Form1.frx":4149C2
         ScaleHeight     =   3615
         ScaleWidth      =   48015
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   48075
      End
   End
   Begin VB.PictureBox P2bazooka1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3375
      Left            =   4440
      Picture         =   "Form1.frx":649B48
      ScaleHeight     =   3315
      ScaleWidth      =   48030
      TabIndex        =   38
      Top             =   3840
      Visible         =   0   'False
      Width           =   48090
      Begin VB.PictureBox P2bazooka2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3375
         Left            =   0
         Picture         =   "Form1.frx":8501F2
         ScaleHeight     =   3315
         ScaleWidth      =   48030
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   48090
      End
   End
   Begin VB.PictureBox P2punch1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3390
      Left            =   1320
      Picture         =   "Form1.frx":A5689C
      ScaleHeight     =   3330
      ScaleWidth      =   24000
      TabIndex        =   36
      Top             =   2880
      Visible         =   0   'False
      Width           =   24060
      Begin VB.PictureBox P2punch2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3390
         Left            =   -360
         Picture         =   "Form1.frx":B5AB5E
         ScaleHeight     =   3330
         ScaleWidth      =   24000
         TabIndex        =   37
         Top             =   960
         Visible         =   0   'False
         Width           =   24060
      End
   End
   Begin VB.PictureBox P2defense1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3165
      Left            =   3240
      Picture         =   "Form1.frx":C5EE20
      ScaleHeight     =   3105
      ScaleWidth      =   12000
      TabIndex        =   32
      Top             =   4200
      Visible         =   0   'False
      Width           =   12060
      Begin VB.PictureBox P2defense2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3165
         Left            =   0
         Picture         =   "Form1.frx":CD8302
         ScaleHeight     =   3105
         ScaleWidth      =   12000
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   12060
      End
   End
   Begin VB.PictureBox P1defense1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3165
      Left            =   4560
      Picture         =   "Form1.frx":D517E4
      ScaleHeight     =   3105
      ScaleWidth      =   12000
      TabIndex        =   30
      Top             =   5160
      Visible         =   0   'False
      Width           =   12060
      Begin VB.PictureBox P1defense2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3165
         Left            =   0
         Picture         =   "Form1.frx":DCACC6
         ScaleHeight     =   3105
         ScaleWidth      =   12000
         TabIndex        =   31
         Top             =   1920
         Visible         =   0   'False
         Width           =   12060
      End
   End
   Begin VB.PictureBox P2runback1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3435
      Left            =   3360
      Picture         =   "Form1.frx":E441A8
      ScaleHeight     =   3375
      ScaleWidth      =   36015
      TabIndex        =   27
      Top             =   5160
      Visible         =   0   'False
      Width           =   36075
      Begin VB.PictureBox P2runback2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3435
         Left            =   0
         Picture         =   "Form1.frx":FCFD8E
         ScaleHeight     =   3375
         ScaleWidth      =   36015
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   36075
      End
   End
   Begin VB.PictureBox P2hurt1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3315
      Left            =   3600
      Picture         =   "Form1.frx":115B974
      ScaleHeight     =   3255
      ScaleWidth      =   60000
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   60060
      Begin VB.PictureBox P2hurt2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3315
         Left            =   0
         Picture         =   "Form1.frx":13D7596
         ScaleHeight     =   3255
         ScaleWidth      =   60000
         TabIndex        =   26
         Top             =   1920
         Visible         =   0   'False
         Width           =   60060
      End
   End
   Begin VB.PictureBox P2idle1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3375
      Left            =   120
      Picture         =   "Form1.frx":16531B8
      ScaleHeight     =   3315
      ScaleWidth      =   24030
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   24090
      Begin VB.PictureBox Picture1 
         Height          =   735
         Left            =   720
         ScaleHeight     =   675
         ScaleWidth      =   915
         TabIndex        =   46
         Top             =   2520
         Width           =   975
      End
      Begin VB.PictureBox P2idle2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3375
         Left            =   960
         Picture         =   "Form1.frx":17568A2
         ScaleHeight     =   3315
         ScaleWidth      =   24030
         TabIndex        =   24
         Top             =   1920
         Visible         =   0   'False
         Width           =   24090
      End
   End
   Begin VB.PictureBox P2run1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3435
      Left            =   2640
      Picture         =   "Form1.frx":1859F8C
      ScaleHeight     =   3375
      ScaleWidth      =   36015
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   36075
      Begin VB.PictureBox p2run2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3435
         Left            =   0
         Picture         =   "Form1.frx":19E5B72
         ScaleHeight     =   3375
         ScaleWidth      =   36015
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   36075
      End
   End
   Begin VB.PictureBox p1hurt 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3300
      Left            =   840
      Picture         =   "Form1.frx":1B71758
      ScaleHeight     =   3240
      ScaleWidth      =   60000
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   60060
      Begin VB.PictureBox p1hurt2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3300
         Left            =   840
         Picture         =   "Form1.frx":1DEA49A
         ScaleHeight     =   3240
         ScaleWidth      =   60000
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   60060
      End
   End
   Begin VB.PictureBox P1bazooka1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3375
      Left            =   3600
      Picture         =   "Form1.frx":20631DC
      ScaleHeight     =   3315
      ScaleWidth      =   48030
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   48090
      Begin VB.PictureBox P1bazooka2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3375
         Left            =   840
         Picture         =   "Form1.frx":2269886
         ScaleHeight     =   3315
         ScaleWidth      =   48030
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   48090
      End
   End
   Begin VB.PictureBox P1gutling_gun1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3675
      Left            =   1560
      Picture         =   "Form1.frx":246FF30
      ScaleHeight     =   3615
      ScaleWidth      =   48015
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   48075
      Begin VB.PictureBox P1gutling_gun2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3675
         Left            =   840
         Picture         =   "Form1.frx":26A50B6
         ScaleHeight     =   3615
         ScaleWidth      =   48015
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   48075
      End
   End
   Begin VB.PictureBox P1punch1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3390
      Left            =   2160
      Picture         =   "Form1.frx":28DA23C
      ScaleHeight     =   3330
      ScaleWidth      =   24000
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   24060
      Begin VB.PictureBox P1punch2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3390
         Left            =   2280
         Picture         =   "Form1.frx":29DE4FE
         ScaleHeight     =   3330
         ScaleWidth      =   24000
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   24060
      End
   End
   Begin VB.PictureBox P1runback1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3435
      Left            =   1560
      Picture         =   "Form1.frx":2AE27C0
      ScaleHeight     =   3375
      ScaleWidth      =   36015
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   36075
      Begin VB.PictureBox P1runback2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   3435
         Left            =   0
         Picture         =   "Form1.frx":2C6E3A6
         ScaleHeight     =   3375
         ScaleWidth      =   36015
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   36075
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   4440
      Top             =   3360
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7560
      Left            =   360
      Picture         =   "Form1.frx":2DF9F8C
      ScaleHeight     =   7500
      ScaleWidth      =   15000
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   15060
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   360
         Top             =   6960
      End
      Begin VB.PictureBox P2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   7560
         Left            =   1440
         Picture         =   "Form1.frx":2F6832E
         ScaleHeight     =   7500
         ScaleWidth      =   15000
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   15060
         Begin VB.PictureBox P1run1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   3330
            Left            =   2040
            Picture         =   "Form1.frx":30D66D0
            ScaleHeight     =   3270
            ScaleWidth      =   36030
            TabIndex        =   5
            Top             =   5880
            Visible         =   0   'False
            Width           =   36090
            Begin VB.PictureBox P1run2 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   3330
               Left            =   840
               Picture         =   "Form1.frx":3256122
               ScaleHeight     =   3270
               ScaleWidth      =   36030
               TabIndex        =   6
               Top             =   1200
               Visible         =   0   'False
               Width           =   36090
            End
         End
         Begin VB.PictureBox P1idle1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   3375
            Left            =   960
            Picture         =   "Form1.frx":33D5B74
            ScaleHeight     =   3315
            ScaleWidth      =   24030
            TabIndex        =   2
            Top             =   2520
            Visible         =   0   'False
            Width           =   24090
            Begin VB.PictureBox P1idle2 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               Height          =   3375
               Left            =   840
               Picture         =   "Form1.frx":34D925E
               ScaleHeight     =   3315
               ScaleWidth      =   24030
               TabIndex        =   3
               Top             =   1440
               Visible         =   0   'False
               Width           =   24090
            End
         End
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   14640
      TabIndex        =   50
      Top             =   1920
      Width           =   6645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   14640
      TabIndex        =   47
      Top             =   1080
      Width           =   6645
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   14160
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin WMPLibCtl.WindowsMediaPlayer defense2 
      Height          =   135
      Left            =   240
      TabIndex        =   35
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
      URL             =   "D:\SMT 7\MM Game\game\Song\defense2.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   238
   End
   Begin WMPLibCtl.WindowsMediaPlayer defense 
      Height          =   255
      Left            =   600
      TabIndex        =   34
      Top             =   8400
      Visible         =   0   'False
      Width           =   255
      URL             =   "D:\SMT 7\MM Game\game\Song\defense.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
   Begin WMPLibCtl.WindowsMediaPlayer hurt 
      Height          =   375
      Left            =   840
      TabIndex        =   29
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
      URL             =   "D:\SMT 7\MM Game\game\Song\hurt.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   661
   End
   Begin WMPLibCtl.WindowsMediaPlayer swoosh 
      Height          =   135
      Left            =   3480
      TabIndex        =   18
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
      URL             =   "D:\SMT 7\MM Game\game\Song\1swoosh.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   238
   End
   Begin WMPLibCtl.WindowsMediaPlayer gatling_gun 
      Height          =   135
      Left            =   2880
      TabIndex        =   17
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
      URL             =   "D:\SMT 7\MM Game\game\Song\0da5.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   238
   End
   Begin WMPLibCtl.WindowsMediaPlayer music 
      Height          =   135
      Left            =   2040
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
      URL             =   "D:\SMT 7\MM Game\game\music.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   238
   End
   Begin WMPLibCtl.WindowsMediaPlayer bazooka 
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
      URL             =   "D:\SMT 7\MM Game\game\Song\Songs_0-46.wav"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal _
hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal _
nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias _
"sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Dim lufi1kiri, lufi1atas, lufi1kol As Integer
Dim lufi2kiri, lufi2atas, lufi2kol As Integer
Dim health1kiri, health1atas, health1kol As Integer
Dim health2kiri, health2atas, health2kol As Integer
Dim pkiri, patas, pkol As Integer
Dim gerak, gerak2 As Integer
Dim detik, detik2 As Integer
Private Const x1 = 0
Private Const x2 = 1000
Private Const y1 = 0
Private Const y2 = 500
Dim allowMovements As Boolean
Dim range1, range2 As Integer





Private Sub Form_Activate()
allowMovements = True
On Error Resume Next
Do

P2.Cls
      
u& = BitBlt(P2.hDC, 0, 0, P2.ScaleWidth, P2.ScaleHeight, P1.hDC, 0, 0, vbSrcCopy)

      'u& = BitBlt(P2.hDC, 200, 300, 200, 200, P1result.hDC, 0, 0, vbSrcAnd)
      'u& = BitBlt(P2.hDC, 200, 300, 200, 200, P1result.hDC, 0, 0, vbSrcPaint)

      u& = BitBlt(P2.hDC, health1kiri, health1atas, 330, 200, P1health1.hDC, health1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, health1kiri, health1atas, 330, 200, P1health2.hDC, health1kol, 0, vbSrcPaint)
      
      u& = BitBlt(P2.hDC, health2kiri, health2atas, 330, 200, P2health1.hDC, health2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, health2kiri, health2atas, 330, 200, P2health2.hDC, health2kol, 0, vbSrcPaint)

If gerak = 0 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 200, P1idle1.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 200, P1idle2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 0 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2idle1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2idle2.hDC, lufi2kol, 0, vbSrcPaint)
End If

If gerak = 1 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 200, P1run1.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 200, P1run2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 1 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2run1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, p2run2.hDC, lufi2kol, 0, vbSrcPaint)
End If

If gerak = 2 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 200, P1runback1.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 200, P1runback2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 2 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2runback1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2runback2.hDC, lufi2kol, 0, vbSrcPaint)
End If

'punch
If gerak = 3 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 200, P1punch1.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 200, P1punch2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 3 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2punch1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2punch2.hDC, lufi2kol, 0, vbSrcPaint)
End If

'gatling gun punch
If gerak = 4 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 220, P1gutling_gun1.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 220, P1gutling_gun2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 4 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2gatling_gun1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2gatling_gun2.hDC, lufi2kol, 0, vbSrcPaint)
End If

'bazooka punch
If gerak = 5 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 350, 200, P1bazooka1.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 350, 200, P1bazooka2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 5 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2bazooka1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2bazooka2.hDC, lufi2kol, 0, vbSrcPaint)
End If

'defense
If gerak = 6 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 220, P1defense1.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 220, P1defense2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 6 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 220, P2defense1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 220, P2defense2.hDC, lufi2kol, 0, vbSrcPaint)
End If

'hurt
If gerak = 11 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 220, p1hurt.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 400, 220, p1hurt2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 11 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2hurt1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 400, 200, P2hurt2.hDC, lufi2kol, 0, vbSrcPaint)
End If

'death
If gerak = 20 Then
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 200, 200, P1death1.hDC, lufi1kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi1kiri, lufi1atas, 200, 200, P1death2.hDC, lufi1kol, 0, vbSrcPaint)
End If

If gerak2 = 20 Then
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 200, 200, P2death1.hDC, lufi2kol, 0, vbSrcAnd)
      u& = BitBlt(P2.hDC, lufi2kiri, lufi2atas, 200, 200, P2death2.hDC, lufi2kol, 0, vbSrcPaint)
End If




u& = BitBlt(hDC, x1, y1, x2, y2, P2.hDC, 24, 24, vbSrcCopy)
P2.Refresh
    
DoEvents
Sleep (1)

Loop


End Sub

Private Sub Form_KeyDown(Key As Integer, Shift As Integer)
If allowMovements = True Then
If Key = vbKeyRight Then gerak = 1
If Key = vbKeyLeft Then gerak = 2
If Key = vbKeyW Then gerak = 4
If Key = vbKeyA Then gerak = 3
If Key = vbKeyD Then gerak = 5
If Key = vbKeyS Then gerak = 6
If Key = vbKeyJ Then gerak2 = 1
If Key = vbKeyL Then gerak2 = 2
If Key = vbKeyU Then gerak2 = 3
If Key = vbKeyO Then gerak2 = 4
If Key = vbKeyP Then gerak2 = 5
If Key = vbKeyK Then gerak2 = 6
End If
If Key = vbKeyReturn Then Reset

End Sub

Private Sub Form_KeyUp(Key As Integer, Shift As Integer)
If Key = vbKeyRight Or Key = vbKeyLeft Or Key = vbKeyS Then
gerak = 0
lufi1kol = 0
End If

If Key = vbKeyJ Or Key = vbKeyL Or Key = vbKeyK Then
gerak2 = 0
lufi2kol = 0
End If

End Sub

Private Sub Form_Load()
lufi1atas = 230
lufi1kiri = 80

lufi2atas = 230
lufi2kiri = 400

health1kiri = 103
health1atas = 70

health2kiri = 590
health2atas = 70

detik = 120

gatling_gun.Controls.stop
bazooka.Controls.stop
swoosh.Controls.stop
hurt.Controls.stop
defense.Controls.stop
defense2.Controls.stop
detik = 0
detik2 = 0
End Sub

Private Sub Picture4_Click()

End Sub

Private Sub Timer1_Timer()

'idle
If gerak = 0 Then
lufi1kol = lufi1kol + 400
If lufi1kol >= (400 * 4) Then lufi1kol = 0: gerak = 0
End If

If gerak2 = 0 Then
lufi2kol = lufi2kol + 400
If lufi2kol >= (400 * 4) Then lufi2kol = 0: gerak2 = 0
End If

'run
If gerak = 1 Then
lufi1kol = lufi1kol + 400
If lufi1kol >= (400 * 6) Then lufi1kol = 0
lufi1kiri = lufi1kiri + 20
If lufi1kiri >= lufi2kiri + 150 Then lufi1kiri = lufi2kiri + 150: lufi2kiri = lufi2kiri + 5
End If

If gerak2 = 1 Then
lufi2kol = lufi2kol + 400
If lufi2kol >= (400 * 6) Then lufi2kol = 0
lufi2kiri = lufi2kiri - 20
If lufi2kiri + 150 < lufi1kiri Then lufi2kiri = lufi1kiri - 150: lufi1kiri = lufi1kiri - 5
End If



'run2
If gerak = 2 Then
lufi1kol = lufi1kol + 400
If lufi1kol >= (400 * 6) Then lufi1kol = 0
lufi1kiri = lufi1kiri - 20
End If

If gerak2 = 2 Then
lufi2kol = lufi2kol + 400
If lufi2kol >= (400 * 6) Then lufi2kol = 0
lufi2kiri = lufi2kiri + 20
End If



'punch
If gerak = 3 Then
lufi1kol = lufi1kol + 400
swoosh.Controls.play
If lufi1kol >= (400 * 4) Then lufi1kol = 0: gerak = 0
If lufi1kiri + 110 > lufi2kiri + 210 Then
detik = 0
If gerak2 <> 6 Then gerak2 = 11: health2kol = health2kol - 5 Else: defense.Controls.play
End If
End If

If gerak2 = 3 Then
lufi2kol = lufi2kol + 400
swoosh.Controls.play
If lufi2kol >= (400 * 4) Then lufi2kol = 0: gerak2 = 0
If lufi1kiri > lufi2kiri + 100 Then
detik = 0
If gerak <> 6 Then gerak = 11: health1kol = health1kol + 5 Else: defense.Controls.play

End If
End If


'punch gutling gun
If gerak = 4 Then
gatling_gun.Controls.play
lufi1kol = lufi1kol + 400
If lufi1kol >= (400 * 8) Then lufi1kol = 0: gerak = 0
If lufi1kiri > lufi2kiri And lufi1kol >= 600 Then
lufi2kiri = lufi1kiri + 5
detik = 0
If gerak2 <> 6 Then gerak2 = 11: health2kol = health2kol - 10 Else: defense2.Controls.play
End If
End If

If gerak2 = 4 Then
gatling_gun.Controls.play
lufi2kol = lufi2kol + 400
If lufi2kol >= (400 * 8) Then lufi2kol = 0: gerak2 = 0
If lufi1kiri > lufi2kiri And lufi2kol >= 600 Then
lufi1kiri = lufi2kiri - 5
detik = 0
If gerak <> 6 Then gerak = 11: health1kol = health1kol + 10 Else: defense2.Controls.play
End If
End If


'punch bazooka
If gerak = 5 Then
bazooka.Controls.play
lufi1kol = lufi1kol + 400
If lufi1kol >= (400 * 8) Then lufi1kol = 0: gerak = 0
If lufi1kiri > lufi2kiri And lufi1kol >= (400 * 6) Then
lufi2kiri = lufi1kiri + 10
detik = 0
If gerak2 <> 6 Then gerak2 = 11: health2kol = health2kol - 20 Else: defense.Controls.play
End If
End If

If gerak2 = 5 Then
bazooka.Controls.play
lufi2kol = lufi2kol + 400
If lufi2kol >= (400 * 8) Then lufi2kol = 0: gerak2 = 0
If lufi1kiri > lufi2kiri And lufi2kol >= (400 * 6) Then
lufi1kiri = lufi2kiri + 10
detik = 0
If gerak <> 6 Then gerak = 11: health1kol = health1kol + 20 Else: defense.Controls.play
End If
End If

'defense
If gerak = 6 Then
lufi1kol = lufi1kol + 400
If lufi1kol >= (400 * 2) Then lufi1kol = 400
End If

If gerak2 = 6 Then
lufi2kol = lufi2kol + 400
If lufi2kol >= (400 * 2) Then lufi2kol = 400
End If


'hurt
If gerak = 11 Then
hurt.Controls.play
lufi1kol = lufi1kol + 400
If lufi1kol >= (400 * 10) Then lufi1kol = 0: gerak = 0
End If

If gerak2 = 11 Then
hurt.Controls.play
lufi2kol = lufi2kol + 400
If lufi2kol >= (400 * 10) Then lufi2kol = 0: gerak2 = 0
End If

'death
If gerak = 20 Then
result
lufi1kiri = lufi2kiri - 5
    If lufi1kol < (200 * 4) Then
    lufi1kol = lufi1kol + 200
    Else
    lufi1kol = 600
    End If
End If

If gerak2 = 20 Then
result
lufi2kiri = lufi1kiri + 200
    If lufi2kol < (200 * 4) Then
    lufi2kol = lufi2kol + 200
    Else
    lufi2kol = 600
    End If
End If

'border
If lufi1kiri < -10 Then lufi1kiri = -10
If lufi2kiri > 620 Then lufi2kiri = 620

'health
If health2kol <= -330 Then Label4.Caption = "P1 WIN": gerak2 = 20
If health1kol >= 330 Then Label4.Caption = "P2 WIN": gerak = 20

detik = detik + 1
Label2.Caption = detik
If detik Mod 13 = 0 And allowMovements = True Then
'Aturan1
End If
End Sub
Private Sub result()
 allowMovements = False
 'hasil = 1
 ' Mengubah posisi PictureBox
P1result.Left = 6000 ' Mengubah posisi horizontal
P1result.Top = 1500  ' Mengubah posisi vertikal
' Mengubah menjadi terlihat (visible)
P1result.Visible = True
 
End Sub
Private Sub Reset()
P1result.Visible = False
hasil = 0
lufi1atas = 230
lufi1kiri = 80

lufi2atas = 230
lufi2kiri = 400

health1kiri = 103
health1atas = 70

health2kiri = 590
health2atas = 70

gerak = 0
gerak2 = 0

health1kol = 0
health2kol = 0

allowMovements = True
End Sub

Private Sub Timer2_Timer()
detik2 = detik2 + 1
Label3.Caption = detik2
End Sub

Private Sub Aturan1()
 'If lufi1kiri < lufi2kiri + 100 Then gerak2 = 1 Else gerak2 = angka2
 Select Case True
    Case lufi1kiri < lufi2kiri + 100
       Aturan2
    Case Else
       Aturan3
 End Select
End Sub

Private Function Aturan2()
gerak2 = 1
End Function

Private Function Aturan3() As Integer
random = Int((3 * Rnd) + 3)
gerak2 = random
End Function

Private Function Aturan4()
gerak2 = 6
End Function

Private Sub music_OpenStateChange(ByVal NewState As Long)
music.settings.volume = 25
  If NewState = wmppsMediaEnded Then
        music.Controls.play
    End If
End Sub


