VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Principal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Applied Genetics"
   ClientHeight    =   12990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   Icon            =   "genetica.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12990
   ScaleWidth      =   16650
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog d_alea_alea 
      Left            =   10200
      Top             =   11040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin AG.jcFrames Meniul 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Top             =   450
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   794
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   ""
      TextBoxHeight   =   25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin AG.isButton Deschidere_fisier 
         Height          =   300
         Left            =   1320
         TabIndex        =   27
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B012
         Style           =   7
         Caption         =   "Deschide"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.zulu bara_date 
         Height          =   255
         Left            =   11280
         TabIndex        =   26
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         Picture         =   "genetica.frx":2B02E
         BackColor       =   16761024
         ForeColor       =   16777215
         Appearance      =   0
         BorderStyle     =   0
         BarPicture      =   "genetica.frx":2B2A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Segments        =   -1  'True
      End
      Begin AG.isButton Salveza_fisier 
         Height          =   300
         Left            =   2520
         TabIndex        =   28
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B604
         Style           =   7
         Caption         =   "Salveaza"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Printare 
         Height          =   300
         Left            =   3720
         TabIndex        =   29
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B620
         Style           =   7
         Caption         =   "Compilare"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton isButton14 
         Height          =   300
         Left            =   7440
         TabIndex        =   30
         Top             =   90
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B63C
         Style           =   7
         Caption         =   "Ajutor"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Deschide_despre 
         Height          =   300
         Left            =   9960
         TabIndex        =   31
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B658
         Style           =   7
         Caption         =   "Despre"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Optiuni_program 
         Height          =   300
         Left            =   6120
         TabIndex        =   120
         Top             =   90
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B674
         Style           =   7
         Caption         =   "Optiuni"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton ONline_AG 
         Height          =   300
         Left            =   8760
         TabIndex        =   121
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B690
         Style           =   7
         Caption         =   "OnLine"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Run_Primar 
         Height          =   300
         Left            =   4920
         TabIndex        =   123
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B6AC
         Style           =   7
         Caption         =   "Start"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Importeaza_orice 
         Height          =   300
         Left            =   120
         TabIndex        =   161
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B6C8
         Style           =   7
         Caption         =   "Importeaza"
         iNonThemeStyle  =   0
         BackColor       =   12937777
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
   End
   Begin AG.jcFrames Top_miscare 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Top             =   0
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   794
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   "      Applied Genetics v1.2 beta"
      TextBoxHeight   =   25
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   14040
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin AG.isButton Iesire 
         Height          =   300
         Left            =   16320
         TabIndex        =   36
         Top             =   90
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B6E4
         Style           =   7
         Caption         =   " X"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton max_ecran 
         Height          =   300
         Left            =   15960
         TabIndex        =   37
         Top             =   90
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B700
         Style           =   7
         Caption         =   " 0"
         iNonThemeStyle  =   7
         Enabled         =   0   'False
         HighlightColor  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Fa_micuta 
         Height          =   300
         Left            =   15600
         TabIndex        =   38
         Top             =   90
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Icon            =   "genetica.frx":2B71C
         Style           =   7
         Caption         =   " _"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Label Printeaza_compileaza 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   78
         Top             =   120
         Width           =   6135
      End
      Begin AG.aicAlphaImage aicAlphaImage4 
         Height          =   345
         Left            =   0
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Image           =   "genetica.frx":2B738
         Scaler          =   1
         Props           =   5
      End
   End
   Begin VB.Timer Master_timer 
      Interval        =   10
      Left            =   10800
      Top             =   11040
   End
   Begin AG.jcFrames Proprietati_stanga 
      Height          =   5835
      Left            =   13280
      Top             =   6600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   10292
      BackColor       =   15783104
      FillColor       =   15783104
      MoverControle   =   -1  'True
      Caption         =   "Proprietati"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin AG.isButton Testeaza_CODUL 
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   5160
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         Icon            =   "genetica.frx":2B9F8
         Style           =   10
         Caption         =   "Start"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin TabDlg.SSTab Tab_proprietati 
         Height          =   4575
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   8070
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Enzime"
         TabPicture(0)   =   "genetica.frx":2BA14
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "de_la_distanta"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "d_x"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "d_n"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "dista"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "bara_date2"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "optiuni_proprietati_arata"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "ins_dreapta"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "ins_stanga"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "normal_dreapta"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "normal_stanga"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "del_stanga"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "del_dreapta"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Incepand_distanta_DEL"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "distanta_x"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "distanta_n"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "enzima"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "distanta_DEL"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Arata_fragmente"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "Gene"
         TabPicture(1)   =   "genetica.frx":2BA30
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame_Shift_plat"
         Tab(1).Control(1)=   "Delimitare_platforma"
         Tab(1).Control(2)=   "Notatii_platforma"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "S.I."
         TabPicture(2)   =   "genetica.frx":2BA4C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Bruta"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "status_combinatii"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Bruta_optiuni"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "S.R."
         TabPicture(3)   =   "genetica.frx":2BA68
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SR_platforma"
         Tab(3).Control(1)=   "Frame1"
         Tab(3).ControlCount=   2
         Begin VB.Frame SR_platforma 
            Caption         =   "Secvente repetitive"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   -74880
            TabIndex        =   135
            Top             =   480
            Width           =   2895
            Begin VB.HScrollBar r_v_la 
               Height          =   255
               Left            =   1320
               Max             =   255
               Min             =   1
               TabIndex        =   143
               Top             =   1800
               Value           =   12
               Width           =   1455
            End
            Begin VB.HScrollBar r_v_dela 
               Height          =   255
               Left            =   1320
               Max             =   255
               Min             =   1
               TabIndex        =   142
               Top             =   1440
               Value           =   2
               Width           =   1455
            End
            Begin VB.TextBox r_dela 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   141
               Top             =   1440
               Width           =   375
            End
            Begin VB.TextBox r_la 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   140
               Top             =   1800
               Width           =   375
            End
            Begin VB.TextBox rep_txt_max 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   139
               Top             =   720
               Width           =   615
            End
            Begin VB.TextBox rep_txt_min 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   138
               Top             =   360
               Width           =   615
            End
            Begin VB.HScrollBar rep_min 
               Height          =   255
               Left            =   1800
               Max             =   10000
               Min             =   1
               TabIndex        =   137
               Top             =   360
               Value           =   2
               Width           =   975
            End
            Begin VB.HScrollBar rep_max 
               Height          =   255
               Left            =   1800
               Max             =   10000
               Min             =   1
               TabIndex        =   136
               Top             =   720
               Value           =   12
               Width           =   975
            End
            Begin VB.Line Line6 
               BorderColor     =   &H00808080&
               X1              =   120
               X2              =   2760
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "b"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1080
               TabIndex        =   151
               Top             =   1800
               Width           =   210
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "b"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1080
               TabIndex        =   150
               Top             =   1440
               Width           =   210
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "De la:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   149
               Top             =   1440
               Width           =   420
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Pana la:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   148
               Top             =   1800
               Width           =   585
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Cel mult:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   147
               Top             =   720
               Width           =   630
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Cel putin:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   146
               Top             =   360
               Width           =   690
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "rep"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1440
               TabIndex        =   145
               Top             =   360
               Width           =   360
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "rep"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1440
               TabIndex        =   144
               Top             =   720
               Width           =   360
            End
         End
         Begin VB.Frame Frame_Shift_plat 
            Caption         =   "Frame Shift"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   -74880
            TabIndex        =   132
            Top             =   480
            Width           =   2895
            Begin VB.HScrollBar FrameShift_Val 
               Height          =   255
               Left            =   1200
               Max             =   3
               Min             =   1
               TabIndex        =   133
               Top             =   360
               Value           =   1
               Width           =   1575
            End
            Begin VB.Label FSV 
               Caption         =   "Frame Shift: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   134
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Delimitare_platforma 
            Caption         =   "Delimitare"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   -74880
            TabIndex        =   130
            Top             =   3600
            Width           =   2895
            Begin VB.CheckBox Delimitator_AUG 
               Caption         =   "Delimiteaza AUG"
               Height          =   375
               Left            =   120
               TabIndex        =   131
               ToolTipText     =   "Delimiteaza semnalul de start OutPut"
               Top             =   360
               Width           =   1575
            End
            Begin AG.isButton Mai_mult_codon 
               Height          =   300
               Left            =   1800
               TabIndex        =   159
               Top             =   360
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   529
               Icon            =   "genetica.frx":2BA84
               Style           =   7
               Caption         =   "CODONI"
               iNonThemeStyle  =   0
               Tooltiptitle    =   ""
               ToolTipIcon     =   0
               ToolTipType     =   1
               ttForeColor     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               RoundedBordersByTheme=   0   'False
            End
         End
         Begin VB.Frame Notatii_platforma 
            Caption         =   "Notatii"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   -74880
            TabIndex        =   124
            Top             =   1440
            Width           =   2895
            Begin VB.TextBox start_not 
               Height          =   285
               Left            =   1800
               TabIndex        =   126
               Text            =   "[+POZITIEb+]"
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox stop_not 
               Height          =   285
               Left            =   1800
               TabIndex        =   125
               Text            =   "[-POZITIEb-]"
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label Info_casuta 
               Caption         =   "Cuvantul ""POZITIE""  locatializeaza codonii de start sau de stop in secventa data."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Index           =   3
               Left            =   120
               TabIndex        =   129
               Top             =   360
               Width           =   2655
            End
            Begin VB.Label Info_casuta 
               Caption         =   "Notatie codoni STOP"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   128
               Top             =   1560
               Width           =   1575
            End
            Begin VB.Label Info_casuta 
               Caption         =   "Notatie codon START"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   127
               Top             =   1200
               Width           =   1695
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   -74880
            TabIndex        =   110
            Top             =   2880
            Width           =   2895
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Combinatii din:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   118
               Top             =   600
               Width           =   1050
            End
            Begin VB.Label c_len_r 
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   1200
               TabIndex        =   117
               Top             =   600
               Width           =   1575
            End
            Begin VB.Line Line5 
               X1              =   120
               X2              =   2760
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label c_t_r 
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   1440
               TabIndex        =   116
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label c_s_r 
               BackStyle       =   0  'Transparent
               Height          =   255
               Left            =   1080
               TabIndex        =   115
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label c_c_r 
               Height          =   255
               Left            =   1560
               TabIndex        =   114
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Combinatii curente:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   113
               Top             =   360
               Width           =   1395
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Combinatii totale:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   112
               Top             =   1200
               Width           =   1260
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Combinatii/s:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   111
               Top             =   960
               Width           =   930
            End
         End
         Begin VB.Frame Bruta_optiuni 
            Caption         =   "Secvente inverse"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   -74880
            TabIndex        =   88
            Top             =   360
            Width           =   2895
            Begin VB.CheckBox SI_grafic_step_by_step 
               Caption         =   "Arata S.I. in format grafic"
               Height          =   195
               Left            =   120
               TabIndex        =   164
               Top             =   2280
               Width           =   2415
            End
            Begin VB.CheckBox arata_SI 
               Caption         =   "Arata fragmentele in text clar"
               Height          =   195
               Left            =   120
               TabIndex        =   119
               Top             =   2040
               Value           =   1  'Checked
               Width           =   2415
            End
            Begin VB.HScrollBar Brut_la 
               Height          =   255
               Left            =   1320
               Max             =   255
               Min             =   1
               TabIndex        =   105
               Top             =   1680
               Value           =   12
               Width           =   1455
            End
            Begin VB.HScrollBar Brut_dela 
               Height          =   255
               Left            =   1320
               Max             =   255
               Min             =   1
               TabIndex        =   104
               Top             =   1320
               Value           =   10
               Width           =   1455
            End
            Begin VB.TextBox txtCharacterset 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               TabIndex        =   93
               Top             =   3480
               Width           =   1095
            End
            Begin VB.TextBox txtTarget 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   240
               TabIndex        =   92
               Text            =   $"genetica.frx":2BAA0
               Top             =   3120
               Width           =   2175
            End
            Begin VB.TextBox txtStartLen 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   91
               Top             =   1320
               Width           =   375
            End
            Begin VB.TextBox txtEndLen 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   90
               Top             =   1680
               Width           =   375
            End
            Begin VB.TextBox txtStartCombo 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   89
               Top             =   480
               Width           =   2655
            End
            Begin AG.isButton a_scriex 
               Height          =   300
               Left            =   120
               TabIndex        =   99
               Top             =   840
               Width           =   540
               _ExtentX        =   953
               _ExtentY        =   529
               Icon            =   "genetica.frx":2BBA3
               Style           =   7
               Caption         =   "A"
               iNonThemeStyle  =   0
               Tooltiptitle    =   ""
               ToolTipIcon     =   0
               ToolTipType     =   1
               ttForeColor     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               RoundedBordersByTheme=   0   'False
            End
            Begin AG.isButton c_scriex 
               Height          =   300
               Left            =   720
               TabIndex        =   100
               Top             =   840
               Width           =   540
               _ExtentX        =   953
               _ExtentY        =   529
               Icon            =   "genetica.frx":2BBBF
               Style           =   7
               Caption         =   "C"
               iNonThemeStyle  =   0
               Tooltiptitle    =   ""
               ToolTipIcon     =   0
               ToolTipType     =   1
               ttForeColor     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               RoundedBordersByTheme=   0   'False
            End
            Begin AG.isButton g_scriex 
               Height          =   300
               Left            =   1320
               TabIndex        =   101
               Top             =   840
               Width           =   540
               _ExtentX        =   953
               _ExtentY        =   529
               Icon            =   "genetica.frx":2BBDB
               Style           =   7
               Caption         =   "G"
               iNonThemeStyle  =   0
               Tooltiptitle    =   ""
               ToolTipIcon     =   0
               ToolTipType     =   1
               ttForeColor     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               RoundedBordersByTheme=   0   'False
            End
            Begin AG.isButton t_scriex 
               Height          =   300
               Left            =   1920
               TabIndex        =   102
               Top             =   840
               Width           =   540
               _ExtentX        =   953
               _ExtentY        =   529
               Icon            =   "genetica.frx":2BBF7
               Style           =   7
               Caption         =   "T"
               iNonThemeStyle  =   0
               Tooltiptitle    =   ""
               ToolTipIcon     =   0
               ToolTipType     =   1
               ttForeColor     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               RoundedBordersByTheme=   0   'False
            End
            Begin AG.isButton ACGT_sterge 
               Height          =   300
               Left            =   2520
               TabIndex        =   103
               Top             =   840
               Width           =   300
               _ExtentX        =   529
               _ExtentY        =   529
               Icon            =   "genetica.frx":2BC13
               Style           =   7
               Caption         =   "<"
               iNonThemeStyle  =   0
               Tooltiptitle    =   ""
               ToolTipIcon     =   0
               ToolTipType     =   1
               ttForeColor     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaskColor       =   0
               RoundedBordersByTheme=   0   'False
            End
            Begin VB.Label bzz2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "b"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1080
               TabIndex        =   107
               Top             =   1680
               Width           =   210
            End
            Begin VB.Label bzz1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "b"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1080
               TabIndex        =   106
               Top             =   1320
               Width           =   210
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00808080&
               X1              =   120
               X2              =   2760
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Set caractere:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   98
               Top             =   3480
               Width           =   1035
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Combinatie tinta 255 caractere:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   97
               Top             =   2880
               Width           =   2280
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "De la:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   96
               Top             =   1320
               Width           =   420
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Pana la:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   95
               Top             =   1680
               Width           =   585
            End
            Begin VB.Label C_brut_start 
               AutoSize        =   -1  'True
               Caption         =   "Combinatie de start: 0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   94
               Top             =   240
               Width           =   1605
            End
         End
         Begin VB.Frame status_combinatii 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   -74880
            TabIndex        =   79
            Top             =   2880
            Width           =   2895
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Combinatii/s:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   87
               Top             =   960
               Width           =   930
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Combinatii totale:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   86
               Top             =   1200
               Width           =   1260
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Combinatii curente:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   85
               Top             =   360
               Width           =   1395
            End
            Begin VB.Label lblCurrentCOmbo 
               Height          =   255
               Left            =   1560
               TabIndex        =   84
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblComboPS 
               Height          =   255
               Left            =   1080
               TabIndex        =   83
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label lblTotalCombo 
               Height          =   255
               Left            =   1440
               TabIndex        =   82
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00808080&
               X1              =   120
               X2              =   2760
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label lblComboLen 
               Height          =   255
               Left            =   1200
               TabIndex        =   81
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Combinatii din:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   600
               Width           =   1050
            End
         End
         Begin VB.Timer Bruta 
            Interval        =   10
            Left            =   -71280
            Top             =   1800
         End
         Begin VB.CheckBox Arata_fragmente 
            Caption         =   "Arata fragmentele de restrictie"
            Height          =   195
            Left            =   360
            TabIndex        =   61
            Top             =   4200
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.VScrollBar distanta_DEL 
            Height          =   2055
            Left            =   2520
            Max             =   100
            Min             =   1
            TabIndex        =   49
            Top             =   1500
            Value           =   1
            Width           =   255
         End
         Begin VB.TextBox enzima 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   420
            Width           =   2655
         End
         Begin VB.CheckBox distanta_n 
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   1200
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox distanta_x 
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   1440
            Width           =   255
         End
         Begin VB.VScrollBar Incepand_distanta_DEL 
            Height          =   2055
            Left            =   2280
            Max             =   100
            Min             =   1
            TabIndex        =   45
            Top             =   1500
            Value           =   1
            Width           =   255
         End
         Begin AG.isButton del_dreapta 
            Height          =   300
            Left            =   240
            TabIndex        =   50
            Top             =   2580
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Icon            =   "genetica.frx":2BC2F
            Style           =   7
            Caption         =   "Deletie 5'->3'"
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonType      =   1
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin AG.isButton del_stanga 
            Height          =   300
            Left            =   240
            TabIndex        =   51
            Top             =   2940
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Icon            =   "genetica.frx":2BC4B
            Style           =   7
            Caption         =   "Deletie 3'<-5'"
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonType      =   1
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin AG.isButton normal_stanga 
            Height          =   300
            Left            =   240
            TabIndex        =   52
            Top             =   2220
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Icon            =   "genetica.frx":2BC67
            Style           =   7
            Caption         =   "Normal 3'<-5'"
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonType      =   1
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin AG.isButton normal_dreapta 
            Height          =   300
            Left            =   240
            TabIndex        =   53
            Top             =   1860
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Icon            =   "genetica.frx":2BC83
            Style           =   7
            Caption         =   "Normal 5'->3'"
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonType      =   1
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin AG.isButton ins_stanga 
            Height          =   300
            Left            =   240
            TabIndex        =   54
            Top             =   3660
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Icon            =   "genetica.frx":2BC9F
            Style           =   7
            Caption         =   "Insertie 3'<-5'"
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonType      =   1
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin AG.isButton ins_dreapta 
            Height          =   300
            Left            =   240
            TabIndex        =   55
            Top             =   3300
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Icon            =   "genetica.frx":2BCBB
            Style           =   7
            Caption         =   "Insertie 5'->3'"
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonType      =   1
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin AG.isButton optiuni_proprietati_arata 
            Height          =   300
            Left            =   1920
            TabIndex        =   60
            Top             =   3660
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "genetica.frx":2BCD7
            Style           =   7
            Caption         =   "ENZIME"
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   1
            ttForeColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaskColor       =   0
            RoundedBordersByTheme=   0   'False
         End
         Begin AG.zulu bara_date2 
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            Picture         =   "genetica.frx":2BCF3
            BackColor       =   16761024
            ForeColor       =   16777215
            Appearance      =   0
            BorderStyle     =   0
            BarPicture      =   "genetica.frx":2BF65
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Segments        =   -1  'True
         End
         Begin VB.Label dista 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   255
            Left            =   2520
            TabIndex        =   59
            Top             =   1260
            Width           =   255
         End
         Begin VB.Label d_n 
            BackStyle       =   0  'Transparent
            Caption         =   "Distanta 1b ... nb"
            Height          =   255
            Left            =   480
            TabIndex        =   58
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label d_x 
            BackStyle       =   0  'Transparent
            Caption         =   "Distanta  nb"
            Height          =   255
            Left            =   480
            TabIndex        =   57
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label de_la_distanta 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   255
            Left            =   2280
            TabIndex        =   56
            Top             =   1260
            Width           =   255
         End
      End
      Begin AG.aicAlphaImage aicAlphaImage2 
         Height          =   1920
         Left            =   720
         Top             =   1440
         Width           =   1920
         _ExtentX        =   3413
         _ExtentY        =   3413
         Image           =   "genetica.frx":2C2C9
         Props           =   5
      End
   End
   Begin AG.jcFrames stanga_principal 
      Align           =   3  'Align Left
      Height          =   11685
      Left            =   0
      Top             =   900
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   20611
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   "Editare Manuala"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin AG.isButton coloreaza_cuvantul 
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Icon            =   "genetica.frx":2EB72
         Style           =   7
         Caption         =   "CORELATII CODONI"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Cauta_posibilitati 
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Icon            =   "genetica.frx":2EB8E
         Style           =   7
         Caption         =   "DETECTEAZA RESTRICTIE"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Transpozoni 
         Height          =   375
         Left            =   360
         TabIndex        =   42
         Top             =   1680
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Icon            =   "genetica.frx":2EBAA
         Style           =   7
         Caption         =   "SECVENTE INVERSE"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.jcFrames Info_GEL 
         Height          =   8415
         Left            =   10
         Top             =   2760
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   14843
         BackColor       =   15783104
         FillColor       =   15783104
         Caption         =   "Informatii GEL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderStyle     =   1
         Begin VB.HScrollBar timpul 
            Height          =   255
            Left            =   120
            Max             =   200
            Min             =   1
            TabIndex        =   65
            Top             =   4440
            Value           =   16
            Width           =   2775
         End
         Begin VB.PictureBox Gel_general 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H0000FF00&
            Height          =   3255
            Left            =   120
            ScaleHeight     =   213
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   198
            TabIndex        =   64
            Top             =   1200
            Width           =   3030
            Begin VB.Line multi_gel_inc2 
               BorderColor     =   &H0000FF00&
               BorderStyle     =   3  'Dot
               X1              =   200
               X2              =   184
               Y1              =   8
               Y2              =   8
            End
            Begin VB.Line multi_gel_inc1 
               BorderColor     =   &H0000FF00&
               BorderStyle     =   3  'Dot
               X1              =   8
               X2              =   0
               Y1              =   8
               Y2              =   8
            End
            Begin VB.Shape Select_multi_test 
               BorderColor     =   &H00C0C0C0&
               FillColor       =   &H00808080&
               Height          =   90
               Left            =   120
               Top             =   90
               Width           =   2655
            End
         End
         Begin VB.PictureBox Gelul 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H0000FF00&
            Height          =   3255
            Left            =   120
            ScaleHeight     =   213
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   198
            TabIndex        =   63
            Top             =   4680
            Width           =   3030
            Begin VB.Line Y_baze_linie 
               BorderColor     =   &H00008000&
               X1              =   200
               X2              =   0
               Y1              =   8
               Y2              =   8
            End
            Begin VB.Line Select_gel1 
               BorderColor     =   &H00008000&
               BorderStyle     =   3  'Dot
               X1              =   192
               X2              =   192
               Y1              =   0
               Y2              =   216
            End
            Begin VB.Line Select_gel2 
               BorderColor     =   &H00008000&
               BorderStyle     =   3  'Dot
               X1              =   192
               X2              =   192
               Y1              =   0
               Y2              =   216
            End
         End
         Begin AG.MorphDisplay ceasul 
            Height          =   300
            Left            =   120
            TabIndex        =   162
            Top             =   7920
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   529
            BorderColor     =   16777215
            BurnInColor     =   4210688
            BurnInColorNeg  =   4210688
            InterDigitGap   =   4
            InterDigitGapExp=   3
            NumDigits       =   4
            NumDigitsExp    =   2
            SegmentHeight   =   4
            SegmentHeightExp=   4
            SegmentLitColorNeg=   65535
            SegmentWidth    =   2
            SegmentWidthExp =   2
            Theme           =   6
            XOffsetExp      =   50
         End
         Begin VB.Label gel_enzima 
            BackStyle       =   0  'Transparent
            Caption         =   "Enzima : -"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Sunt_la_fragmentul 
            BackStyle       =   0  'Transparent
            Caption         =   "Fragment : 0"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label distanta_pe_GEL 
            BackStyle       =   0  'Transparent
            Caption         =   "Migrare : 0"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label timp_gelus 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   255
            Left            =   2880
            TabIndex        =   66
            Top             =   4440
            Width           =   255
         End
      End
      Begin AG.isButton ComparatorSec 
         Height          =   375
         Left            =   360
         TabIndex        =   69
         Top             =   2160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Icon            =   "genetica.frx":2EBC6
         Style           =   7
         Caption         =   "SECVENTE REPETATE"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton aduna_LA_OUTPUT 
         Height          =   300
         Left            =   120
         TabIndex        =   163
         Top             =   11280
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   529
         Icon            =   "genetica.frx":2EBE2
         Style           =   10
         Caption         =   "Aduna datele in OUTPUT - >"
         iNonThemeStyle  =   0
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.aicAlphaImage Treci_la_loc 
         Height          =   1920
         Index           =   1
         Left            =   720
         Top             =   6120
         Width           =   1920
         _ExtentX        =   3413
         _ExtentY        =   3413
         Image           =   "genetica.frx":2EBFE
         Props           =   5
      End
   End
   Begin AG.jcFrames Editare_mijloc 
      Align           =   3  'Align Left
      Height          =   11685
      Left            =   3300
      Top             =   900
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   20611
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   "Editare"
      TextBoxHeight   =   28
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin VB.TextBox out_put 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   72
         Top             =   8040
         Width           =   9735
      End
      Begin VB.PictureBox segment_total_gena 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   647
         TabIndex        =   62
         Top             =   600
         Width           =   9735
         Begin VB.Timer Arata_exp_semi_transparent 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   8400
            Top             =   120
         End
         Begin VB.PictureBox Stanga_af_gr 
            Appearance      =   0  'Flat
            BackColor       =   &H00F0D4C0&
            ForeColor       =   &H80000008&
            Height          =   2055
            Left            =   0
            ScaleHeight     =   135
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   143
            TabIndex        =   152
            Top             =   0
            Width           =   2175
            Begin VB.Image sageata_dir 
               Height          =   195
               Index           =   5
               Left            =   1920
               Picture         =   "genetica.frx":32503
               Top             =   1560
               Width           =   180
            End
            Begin VB.Image sageata_dir 
               Height          =   195
               Index           =   4
               Left            =   1920
               Picture         =   "genetica.frx":329E1
               Top             =   1320
               Width           =   180
            End
            Begin VB.Image sageata_dir 
               Height          =   195
               Index           =   3
               Left            =   1920
               Picture         =   "genetica.frx":32EBF
               Top             =   960
               Width           =   180
            End
            Begin VB.Image sageata_dir 
               Height          =   195
               Index           =   2
               Left            =   1920
               Picture         =   "genetica.frx":3339D
               Top             =   720
               Width           =   180
            End
            Begin VB.Image sageata_dir 
               Height          =   195
               Index           =   1
               Left            =   1920
               Picture         =   "genetica.frx":3387B
               Top             =   360
               Width           =   180
            End
            Begin VB.Image sageata_dir 
               Height          =   195
               Index           =   0
               Left            =   1920
               Picture         =   "genetica.frx":33D59
               Top             =   120
               Width           =   180
            End
            Begin VB.Label exp_sp_grafic 
               BackStyle       =   0  'Transparent
               Caption         =   "Corelatii codoni"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   120
               TabIndex        =   158
               Top             =   1530
               Width           =   1575
            End
            Begin VB.Label exp_sp_grafic 
               BackStyle       =   0  'Transparent
               Caption         =   "Secvente repetate"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   4
               Left            =   120
               TabIndex        =   157
               Top             =   1230
               Width           =   1815
            End
            Begin VB.Label exp_sp_grafic 
               BackStyle       =   0  'Transparent
               Caption         =   "Secvente inverse"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   3
               Left            =   120
               TabIndex        =   156
               Top             =   930
               Width           =   1695
            End
            Begin VB.Label exp_sp_grafic 
               BackStyle       =   0  'Transparent
               Caption         =   "Situsuri posibile"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   2
               Left            =   120
               TabIndex        =   155
               Top             =   630
               Width           =   1575
            End
            Begin VB.Label exp_sp_grafic 
               BackStyle       =   0  'Transparent
               Caption         =   "Selectare manuala"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   120
               TabIndex        =   154
               Top             =   330
               Width           =   1815
            End
            Begin VB.Label exp_sp_grafic 
               BackStyle       =   0  'Transparent
               Caption         =   "Restrictie normala"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   120
               TabIndex        =   153
               Top             =   30
               Width           =   1695
            End
         End
         Begin VB.Label anunt_schimbare_fundal 
            BackStyle       =   0  'Transparent
            Caption         =   "Va rog sa asteptati, modific fundalul ..."
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   2520
            TabIndex        =   160
            Top             =   480
            Visible         =   0   'False
            Width           =   6015
         End
         Begin VB.Shape select_segment_de_la_gel 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   2055
            Left            =   0
            Top             =   0
            Width           =   15
         End
         Begin VB.Line Segment_grafic 
            BorderColor     =   &H0000C0C0&
            BorderStyle     =   3  'Dot
            X1              =   328
            X2              =   328
            Y1              =   0
            Y2              =   136
         End
      End
      Begin VB.PictureBox Despartitor_in_out 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0D4C0&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   120
         MousePointer    =   7  'Size N S
         ScaleHeight     =   105
         ScaleWidth      =   9705
         TabIndex        =   41
         Top             =   7920
         Width           =   9735
      End
      Begin AG.isButton reset_text 
         Height          =   300
         Left            =   240
         TabIndex        =   32
         Top             =   120
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   529
         Icon            =   "genetica.frx":34237
         Style           =   7
         Caption         =   "R"
         iNonThemeStyle  =   7
         BackColor       =   15783104
         FontHighlightColor=   16777215
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin RichTextLib.RichTextBox COD_CODON 
         Height          =   5175
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   9128
         _Version        =   393217
         BackColor       =   4210752
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"genetica.frx":34253
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AG.isButton ADN_la_ARN 
         Height          =   300
         Left            =   6360
         TabIndex        =   33
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         Icon            =   "genetica.frx":48BB3
         Style           =   7
         Caption         =   "ADN-ARN"
         CaptionAlign    =   2
         iNonThemeStyle  =   4
         BackColor       =   15783104
         HighlightColor  =   16777215
         FontHighlightColor=   14737632
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Revers 
         Height          =   300
         Left            =   3720
         TabIndex        =   34
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Icon            =   "genetica.frx":48D0D
         Style           =   7
         Caption         =   "Revers"
         CaptionAlign    =   2
         iNonThemeStyle  =   4
         BackColor       =   15783104
         HighlightColor  =   16777215
         FontHighlightColor=   14737632
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Gaseste_Eruari 
         Height          =   300
         Left            =   720
         TabIndex        =   35
         Top             =   120
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   529
         Icon            =   "genetica.frx":48E67
         Style           =   7
         Caption         =   "F"
         BackColor       =   8421504
         FontColor       =   14737632
         FontHighlightColor=   16777215
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton cautare_secvente 
         Height          =   300
         Left            =   1200
         TabIndex        =   43
         Top             =   120
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   529
         Icon            =   "genetica.frx":48E83
         Style           =   7
         Caption         =   "S"
         BackColor       =   15783104
         FontColor       =   14737632
         FontHighlightColor=   16777215
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton ARN_la_ADN 
         Height          =   300
         Left            =   5040
         TabIndex        =   73
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Icon            =   "genetica.frx":48E9F
         Style           =   7
         Caption         =   "ARN-ADN"
         CaptionAlign    =   2
         iNonThemeStyle  =   4
         BackColor       =   14737632
         HighlightColor  =   16777215
         FontHighlightColor=   14737632
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Antisens 
         Height          =   300
         Left            =   2400
         TabIndex        =   74
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Icon            =   "genetica.frx":48FF9
         Style           =   7
         Caption         =   "Antisens"
         CaptionAlign    =   2
         iNonThemeStyle  =   4
         BackColor       =   8421504
         HighlightColor  =   16777215
         FontHighlightColor=   14737632
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton genereaza_nuc 
         Height          =   300
         Left            =   1680
         TabIndex        =   77
         Top             =   120
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   529
         Icon            =   "genetica.frx":49153
         Style           =   7
         Caption         =   "G"
         BackColor       =   8421504
         FontColor       =   14737632
         FontHighlightColor=   16777215
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   1
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin AG.isButton Adn_prot 
         Height          =   300
         Left            =   7800
         TabIndex        =   122
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         Icon            =   "genetica.frx":4916F
         Style           =   7
         Caption         =   "ADN-AA"
         CaptionAlign    =   2
         iNonThemeStyle  =   3
         BackColor       =   12937777
         HighlightColor  =   16777215
         FontHighlightColor=   14737632
         Tooltiptitle    =   "Explicatie"
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonType      =   1
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   2280
         X2              =   2280
         Y1              =   120
         Y2              =   360
      End
   End
   Begin AG.jcFrames Anunt_jos 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      Top             =   12585
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   714
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin VB.Label mesagerie_statica 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Selectat de la nucleotida"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label mesagerie_statica 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "pana la nucleotida"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Total_selectie 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6960
         TabIndex        =   18
         Top             =   120
         Width           =   855
      End
      Begin VB.Label mesagerie_statica 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total nucleotide selectate:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   17
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label selectare_stop 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.Label selectare_start 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Textul_selectat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7920
         TabIndex        =   14
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Textul_selectat_coada 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10560
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
   End
   Begin AG.jcFrames Dreapta_principal 
      Align           =   4  'Align Right
      Height          =   11685
      Left            =   13275
      Top             =   900
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   20611
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   "Mutari in timp real "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin VB.CheckBox In_OUTPUT 
         Height          =   200
         Left            =   960
         TabIndex        =   76
         Top             =   5400
         Value           =   1  'Checked
         Width           =   200
      End
      Begin VB.CheckBox fara_culoare 
         Height          =   200
         Left            =   960
         TabIndex        =   75
         Top             =   5160
         Value           =   1  'Checked
         Width           =   200
      End
      Begin VB.PictureBox Culoarea1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   40
         Top             =   5160
         Width           =   375
      End
      Begin VB.PictureBox Culoarea2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   39
         Top             =   5280
         Width           =   375
      End
      Begin VB.PictureBox Tabel_culori 
         AutoSize        =   -1  'True
         Height          =   1755
         Left            =   120
         MousePointer    =   2  'Cross
         Picture         =   "genetica.frx":492C9
         ScaleHeight     =   1695
         ScaleWidth      =   3000
         TabIndex        =   25
         Top             =   3360
         Width           =   3060
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Proceseaza  in OUTPUT"
         Height          =   195
         Left            =   1320
         TabIndex        =   109
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Coloreaza secventa"
         Height          =   255
         Left            =   1320
         TabIndex        =   108
         Top             =   5160
         Width           =   1815
      End
      Begin AG.aicAlphaImage Treci_la_loc 
         Height          =   1920
         Index           =   0
         Left            =   840
         Top             =   7800
         Width           =   1920
         _ExtentX        =   3413
         _ExtentY        =   3413
         Image           =   "genetica.frx":4E5A9
         Props           =   5
      End
      Begin AG.aicAlphaImage aicAlphaImage6 
         Height          =   1920
         Left            =   1200
         Top             =   960
         Width           =   1920
         _ExtentX        =   3413
         _ExtentY        =   3413
         Image           =   "genetica.frx":51EAE
         Props           =   5
      End
      Begin AG.aicAlphaImage aicAlphaImage3 
         Height          =   720
         Left            =   0
         Top             =   0
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1270
         Image           =   "genetica.frx":7B50C
         Props           =   5
      End
      Begin VB.Label nucleotide_la_suta 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label nucleotide_la_suta 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label nucleotide_la_suta 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label nucleotide_la_suta 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label nucleotide_la_suta 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label mesagerie_statica 
         BackStyle       =   0  'Transparent
         Caption         =   "A :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.Label mesagerie_statica 
         BackStyle       =   0  'Transparent
         Caption         =   "T :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
      Begin VB.Label mesagerie_statica 
         BackStyle       =   0  'Transparent
         Caption         =   "C :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label mesagerie_statica 
         BackStyle       =   0  'Transparent
         Caption         =   "G :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label mesagerie_statica 
         BackStyle       =   0  'Transparent
         Caption         =   "U :"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Posibil_tm 
         BackStyle       =   0  'Transparent
         Caption         =   "Tm = 0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label procent_GC 
         BackStyle       =   0  'Transparent
         Caption         =   "G+C = 0%"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Ce_E 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Primer = 0b"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   0
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   0
         X2              =   2400
         Y1              =   2400
         Y2              =   2400
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const WM_PASTE = &H302

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Public bara_de_delimitare_text As String

Dim nucleotida(1 To 5) As String
Dim aici_y_vechi As Variant
Dim old_Despartitor_in_out As Variant
Dim play_run As Variant

Dim sel_manuala As Long

Dim fragmentul_nr_x As Variant

Dim blocat As Boolean

Private sterge_selectie As Boolean

Public Nume_proiect As String
Public Obj_de_colorat_principal As Object

Public cmd_prm As String

'***************************************************************(*
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private WithEvents cBrute As clsBruteforce
Attribute cBrute.VB_VarHelpID = -1
Dim bStop As Boolean
Dim bLoopRunning As Boolean
'***************************************************************(*
Private afisat_pan As Boolean

Private Sub Catena_sora_Click()

S1 = Mid(COD_CODON.Text, 1, COD_CODON.SelStart)
p = COD_CODON.SelText
S2 = Mid(COD_CODON.Text, Len(S1) + Len(p) + 1, Len(COD_CODON.Text))
'***

If (p = "") Then Exit Sub
p_tmp = p
p = ""

For n = 1 To Len(p_tmp)
d = LCase(Mid(p_tmp, n, 1))

If (d = "u") Then MsgBox "Secventa selectata este ARN !": Exit Sub

If (d = "a") Then d = "t": GoTo 2
If (d = "t") Then d = "a": GoTo 2
If (d = "g") Then d = "c": GoTo 2
If (d = "c") Then d = "g": GoTo 2

2:
p = p & UCase(d)
Next n
'***
COD_CODON.Text = S1 & p & S2

End Sub

Private Sub Arata_exp_semi_transparent_Timer()

If afisat_pan = False Then
If Stanga_af_gr.Left >= 0 Then
Stanga_af_gr.Left = -5
afisat_pan = True
Arata_exp_semi_transparent.Enabled = False
Else
Stanga_af_gr.Left = Stanga_af_gr.Left + 30
End If

Else


If Stanga_af_gr.Left <= -(Stanga_af_gr.Width) Then
Stanga_af_gr.Left = -(Stanga_af_gr.Width)
afisat_pan = False
Arata_exp_semi_transparent.Enabled = False
Else
Stanga_af_gr.Left = Stanga_af_gr.Left - 30
End If
End If
'If Stanga_af_gr.Left <= Stanga_af_gr.Width Then
'Arata_exp_semi_transparent.Enabled = False
'Else
'Stanga_af_gr.Left = Stanga_af_gr.Left + 30
'End If


End Sub

Private Sub Brut_dela_Change()
If Brut_dela.Value > Brut_la.Value Then Brut_la.Value = Brut_dela.Value
End Sub

Private Sub Brut_la_Change()
If Brut_dela.Value > Brut_la.Value Then Brut_dela.Value = Brut_la.Value
If Brut_la.Value < Len(txtStartCombo.Text) Then
Brut_la.Value = Len(txtStartCombo.Text)
txtEndLen.Text = Brut_la.Value
End If
End Sub
'=============================================================================================================
Private Sub a_scriex_Click()
txtStartCombo.Enabled = True
If Len(txtStartCombo.Text) > 255 Then
MsgBox "Limita este de 255 baze !", vbApplicationModal + vbCritical, "Atentie !"
Exit Sub
End If
txtStartCombo.Text = txtStartCombo.Text & "A"
End Sub

Private Sub c_scriex_Click()
txtStartCombo.Enabled = True
If Len(txtStartCombo.Text) > 255 Then
MsgBox "Limita este de 255 baze !", vbApplicationModal + vbCritical, "Atentie !"
Exit Sub
End If
txtStartCombo.Text = txtStartCombo.Text & "C"
End Sub

Private Sub ComparatorSec_Click()
Tab_proprietati.Tab = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShutdownGDIPlus
End Sub

Private Sub g_scriex_Click()
txtStartCombo.Enabled = True
If Len(txtStartCombo.Text) > 255 Then
MsgBox "Limita este de 255 baze !", vbApplicationModal + vbCritical, "Atentie !"
Exit Sub
End If
txtStartCombo.Text = txtStartCombo.Text & "G"
End Sub

Private Sub Importeaza_orice_Click()
Dim CurEnvFile As String
'------------------------------------------------------------------------------------------------------
CommonDialog1.FileName = ""
CommonDialog1.Filter = "*.*"
CommonDialog1.ShowOpen
CurEnvFile = CommonDialog1.FileName
If CommonDialog1.FileName = "" Or CommonDialog1.FileTitle = "" Then Exit Sub


'------------------------------------------------------------------------------------------------------

  Filenr = FreeFile
  Dim BitArray As String
  Open CurEnvFile For Binary As #Filenr
  BitArray = Space(LOF(Filenr))
  Get #Filenr, , BitArray
  Close #Filenr

XBitArray = BitArray

po1 = InStr(BitArray, "ORIGIN")
If po1 <> 0 Then

de_la_gene_bank_tmp = Mid(BitArray, po1 + Len("ORIGIN"), Len(BitArray))
po2 = InStr(de_la_gene_bank_tmp, "//")
If po2 <> 0 Then

de_la_gene_bank_SEGMENT = Mid(de_la_gene_bank_tmp, 1, po2 - 1)
de_la_gene_bank_SEGMENT = Replace(UCase(de_la_gene_bank_SEGMENT), UCase("</A>"), "")
de_la_gene_bank_SEGMENT = Replace(UCase(de_la_gene_bank_SEGMENT), UCase("<A"), "")
de_la_gene_bank_SEGMENT = Replace(UCase(de_la_gene_bank_SEGMENT), UCase("name"), "")
de_la_gene_bank_SEGMENT = Replace(UCase(de_la_gene_bank_SEGMENT), UCase("sequence"), "")
'00000000000000000000000000000000000000000000
For R = 1 To Len(de_la_gene_bank_SEGMENT)
DoEvents
'--------
Principal.bara_date.Value = (Principal.bara_date.Max / Len(de_la_gene_bank_SEGMENT)) * R
'--------
nuc = Mid(de_la_gene_bank_SEGMENT, R, 1)
'00000000000000000000000000000000000000000000
If (filtru_nucleotide(nuc) <> 1) Then
b = b + 1
Else
n = n + 1
jr = jr & UCase(nuc)
End If
'00000000000000000000000000000000000000000000
Next R
'00000000000000000000000000000000000000000000

XBitArray = jr 'de_la_gene_bank_SEGMENT

Else
MsgBox "Structura fisierului este corupta !" '& vbCrLf & "Programul continua sa importeze continutul fisierului ca atare !"
Exit Sub
End If


Else
MsgBox "Fisierul nu contine formatul GeneBank !" '& vbCrLf & "Programul continua sa importeze continutul fisierului ca atare !"
Exit Sub
End If

COD_CODON.TextRTF = ""
COD_CODON.Text = LCase(XBitArray)

If GetSetting(App.ExeName, "Setari", "p1") <> "" Then
COD_CODON.SelStart = 0
COD_CODON.SelLength = Len(COD_CODON.Text)
COD_CODON.SelColor = GetSetting(App.ExeName, "Setari", "p7")
COD_CODON.SelStart = 0
Else
COD_CODON.SelStart = 0
COD_CODON.SelLength = Len(COD_CODON.Text)
COD_CODON.SelColor = vbWhite
COD_CODON.SelStart = 0
End If

XBitArray = Empty
BitArray = Empty
'--------
Principal.bara_date.Value = 0
'--------
End Sub

Private Sub Mai_mult_codon_Click()
deasupara_reprezentarii_grafice = False
Mai_mult.Show
Mai_mult.op_tab.Tab = 3
End Sub

Private Sub ONline_AG_Click()
R = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & adresa_site_general, vbMaximizedFocus)
End Sub

Private Sub Optiuni_program_Click()
deasupara_reprezentarii_grafice = False
Mai_mult.Show
'Mai_mult.op_tab.Tab = 2
End Sub

Private Sub r_v_dela_Change()
If r_v_dela.Value > r_v_la.Value Then r_v_la.Value = r_v_dela.Value
End Sub

Private Sub r_v_la_Change()
If r_v_dela.Value > r_v_la.Value Then r_v_dela.Value = r_v_la.Value
End Sub

Private Sub rep_max_Change()
If rep_min.Value > rep_max.Value Then rep_min.Value = rep_max.Value
End Sub

Private Sub rep_min_Change()
If rep_min.Value > rep_max.Value Then rep_max.Value = rep_min.Value
End Sub

Private Sub Run_Primar_Click()
Call Testeaza_CODUL_Click
End Sub

Private Sub sageata_dir_Click(Index As Integer)
sterge_linie = "2,20!22,40!42,62!64,84!86,104!106,126"

sterge_linie = Split(sterge_linie, "!")(Index)


sterge_linie1 = Split(sterge_linie, ",")(0)
sterge_linie2 = Split(sterge_linie, ",")(1)

segment_total_gena.Line (0, sterge_linie1)-(segment_total_gena.ScaleWidth, sterge_linie2), Mai_mult.R_cul_grafic.BackColor, BF
End Sub

Private Sub segment_total_gena_Click()
'segment_total_gena.BackColor = vbBlack
'segment_total_gena.ForeColor = vbWhite
     'segment_total_gena.CurrentY = 19
     'segment_total_gena.Print "Applied Genetics."
     Arata_exp_semi_transparent.Enabled = True
     If deasupara_reprezentarii_grafice = False Then Exit Sub
     
        COD_CODON.SetFocus
End Sub

Private Sub t_scriex_Click()
txtStartCombo.Enabled = True
If Len(txtStartCombo.Text) > 255 Then
MsgBox "Limita este de 255 baze !", vbApplicationModal + vbCritical, "Atentie !"
Exit Sub
End If
txtStartCombo.Text = txtStartCombo.Text & "T"
End Sub

Private Sub ACGT_sterge_Click()
If txtStartCombo.Text = "" Then
txtStartCombo.Enabled = False
Exit Sub
End If
txtStartCombo.Text = Mid(txtStartCombo.Text, 1, Len(txtStartCombo.Text) - 1)
End Sub

Private Sub txtStartCombo_Change()
C_brut_start.Caption = "Combinatie de start: " & Len(txtStartCombo.Text) & " b"
If Len(txtStartCombo.Text) > 0 Then
Brut_dela.Enabled = False
txtStartLen.Enabled = False

If Brut_la.Value < Len(txtStartCombo.Text) Then
Brut_la.Value = Len(txtStartCombo.Text)
txtEndLen.Text = Brut_la.Value
End If

Else
Brut_dela.Enabled = True
txtStartLen.Enabled = True
End If
End Sub
'=============================================================================================================


Private Sub COD_CODON_Click()

If Len(COD_CODON.Text) <= 0 Then Exit Sub
Segment_grafic.x1 = (Principal.segment_total_gena.ScaleWidth / Len(COD_CODON.Text)) * COD_CODON.SelStart
Segment_grafic.x2 = Segment_grafic.x1
deasupara_reprezentarii_grafice = True
End Sub


Private Sub COD_CODON_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Len(COD_CODON.SelText) <= 0) Then
deasupara_reprezentarii_grafice = True
Else
deasupara_reprezentarii_grafice = False
End If



If fara_culoare.Value = 1 Then
'-----------------------
If (Len(COD_CODON.SelText) <= 0) Then
For o = 0 To 4
nucleotide_la_suta(o).Caption = "0pb - [0%]"
Next o
Exit Sub
End If

For n = 1 To Len(COD_CODON.SelText)
DoEvents
'--------
If (Len(COD_CODON.SelText) <= 0) Then
Exit Sub
End If
Principal.bara_date.Value = (Principal.bara_date.Max / Len(COD_CODON.SelText)) * n
'--------
d = LCase(Mid(COD_CODON.SelText, n, 1))
If (d = "a") Then A = A + 1
If (d = "t") Then t = t + 1
If (d = "g") Then g = g + 1
If (d = "c") Then c = c + 1
If (d = "u") Then u = u + 1
Next n

tot_txt_cod = 0
tot_txt_cod = Len(COD_CODON.SelText)

nucleotide_la_suta(0).Caption = Int(A) & "b - [" & Int((100 / tot_txt_cod) * A) & "%]"
nucleotide_la_suta(1).Caption = Int(t) & "b - [" & Int((100 / tot_txt_cod) * t) & "%]"
nucleotide_la_suta(2).Caption = Int(c) & "b - [" & Int((100 / tot_txt_cod) * c) & "%]"
nucleotide_la_suta(3).Caption = Int(g) & "b - [" & Int((100 / tot_txt_cod) * g) & "%]"
nucleotide_la_suta(4).Caption = Int(u) & "b - [" & Int((100 / tot_txt_cod) * u) & "%]"
'-----------------------
Principal.bara_date.Value = 0
'-----------------------

'-----------------------
Segment_grafic.x1 = (Principal.segment_total_gena.ScaleWidth / Len(COD_CODON.Text)) * COD_CODON.SelStart
Segment_grafic.x2 = Segment_grafic.x1

If COD_CODON.SelLength = 0 Then
Exit Sub
End If

x1 = (Principal.segment_total_gena.ScaleWidth / Len(COD_CODON.Text)) * COD_CODON.SelStart
x2 = (Principal.segment_total_gena.ScaleWidth / Len(COD_CODON.Text)) * (COD_CODON.SelStart + COD_CODON.SelLength)

Principal.segment_total_gena.Line (x1, 22)-(x2, 40), Culoarea1.BackColor, BF

If In_OUTPUT.Value = 1 Then
If aduna_LA_OUTPUT.Value = False Then out_put.Text = Empty
If Antisens.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "Catena ANTISENS incepe de la pozitia [" & selectare_start.Caption & "] in secventa INPUT si este:" & vbCrLf & CATENA_PLUS(COD_CODON.SelText) & vbCrLf & "pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT. Lungimea catenei ANTISENS:[" & Total_selectie.Caption & "]" & vbCrLf & bara_de_delimitare_text
If Revers.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "Secventa INVERSATA provine din secventa de INPUT - de la nucleotida [" & selectare_start.Caption & "]:" & vbCrLf & REVERSUL(COD_CODON.SelText) & vbCrLf & "pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT. Lungimea catenei INVERSATE:[" & Total_selectie.Caption & "]" & vbCrLf & bara_de_delimitare_text
If ARN_la_ADN.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "[ARN in ADN] secventa de ADN incepe de la [" & selectare_start.Caption & "]:" & vbCrLf & ARN_ADN_f(COD_CODON.SelText) & vbCrLf & "pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT. Lungimea catenei de ADN:[" & Total_selectie.Caption & "]" & vbCrLf & bara_de_delimitare_text
If ADN_la_ARN.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "[ADN in ARN] secventa de ARN incepe de la [" & selectare_start.Caption & "]:" & vbCrLf & ADN_ARN_f(COD_CODON.SelText) & vbCrLf & "pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT. Lungimea catenei de ARN:[" & Total_selectie.Caption & "]" & vbCrLf & bara_de_delimitare_text

tmp_amino_str = X_codon_aminoacid(ADN_ARN_f(COD_CODON.SelText), amino_notatie_scurta)
nr_aa = Split(tmp_amino_str, "@")(0)
str_aa = Split(tmp_amino_str, "@")(1)
If Adn_prot.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "[ADN in PROTEINA] secventa de ADN incepe de la [" & selectare_start.Caption & "] pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT." & vbCrLf & "Lungimea lantului de aminoacizi:[" & nr_aa & " aa]" & vbCrLf & "Secventa este: " & vbCrLf & "[" & str_aa & "]" & vbCrLf

End If

Call HighLightSelection(Principal, COD_CODON, sel_manuala, False)
Else

If In_OUTPUT.Value = 1 Then
If aduna_LA_OUTPUT.Value = False Then out_put.Text = Empty
If Antisens.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "Catena ANTISENS incepe de la pozitia [" & selectare_start.Caption & "] in secventa INPUT si este:" & vbCrLf & CATENA_PLUS(COD_CODON.SelText) & vbCrLf & "pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT. Lungimea catenei ANTISENS:[" & Total_selectie.Caption & "]" & vbCrLf & bara_de_delimitare_text
If Revers.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "Secventa INVERSATA provine din secventa de INPUT - de la nucleotida [" & selectare_start.Caption & "]:" & vbCrLf & REVERSUL(COD_CODON.SelText) & vbCrLf & "pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT. Lungimea catenei INVERSATE:[" & Total_selectie.Caption & "]" & vbCrLf & bara_de_delimitare_text
If ARN_la_ADN.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "[ARN in ADN] secventa de ADN incepe de la [" & selectare_start.Caption & "]:" & vbCrLf & ARN_ADN_f(COD_CODON.SelText) & vbCrLf & "pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT. Lungimea catenei de ADN:[" & Total_selectie.Caption & "]" & vbCrLf & bara_de_delimitare_text
If ADN_la_ARN.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "[ADN in ARN] secventa de ARN incepe de la [" & selectare_start.Caption & "]:" & vbCrLf & ADN_ARN_f(COD_CODON.SelText) & vbCrLf & "pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT. Lungimea catenei de ARN:[" & Total_selectie.Caption & "]" & vbCrLf & bara_de_delimitare_text
tmp_amino_str = X_codon_aminoacid(ADN_ARN_f(COD_CODON.SelText), amino_notatie_scurta)
nr_aa = Split(tmp_amino_str, "@")(0)
str_aa = Split(tmp_amino_str, "@")(1)
If Adn_prot.Value = True Then out_put.Text = out_put.Text & vbCrLf & bara_de_delimitare_text & vbCrLf & "[ADN in PROTEINA] secventa de ADN incepe de la [" & selectare_start.Caption & "] pana la nucleotida [" & selectare_stop.Caption & "] din secventa introdusa in INPUT." & vbCrLf & "Lungimea lantului de aminoacizi:[" & nr_aa & " aa]" & vbCrLf & "Secventa este: " & vbCrLf & "[" & str_aa & "]" & vbCrLf


End If

End If
'-----------------------
deasupara_reprezentarii_grafice = True
End Sub

Private Sub Deschide_despre_Click()
deasupara_reprezentarii_grafice = False
Despre.Show
End Sub

Private Sub Deschidere_fisier_Click()
'On Error GoTo OpenCodeError:
Dim CurEnvFile As String

If cmd_prm <> "%1" Then
If cmd_prm = "" Then
'------------------------------------------------------------------------------------------------------
CommonDialog1.FileName = ""
CommonDialog1.Filter = "*.pro"
CommonDialog1.ShowOpen
CurEnvFile = CommonDialog1.FileName
If CommonDialog1.FileName = "" Or CommonDialog1.FileTitle = "" Then Exit Sub

Cale_proiect = Mid(CommonDialog1.FileName, 1, InStrRev(CommonDialog1.FileName, "\"))
proiectul = Mid(CommonDialog1.FileTitle, 1, InStrRev(CommonDialog1.FileTitle, ".") - 1)
'------------------------------------------------------------------------------------------------------
'Exit Sub
Else
CurEnvFile = cmd_prm
Cale_proiect = Mid(CurEnvFile, 1, InStrRev(CurEnvFile, "\"))
tmp = Mid(CurEnvFile, InStrRev(CurEnvFile, "\") + 1, Len(CurEnvFile))
proiectul = Mid(tmp, 1, InStrRev(tmp, ".") - 1)
cmd_prm = ""
'GoTo 1
End If

ElseIf cmd_prm = "" Then

End If

1:
Nume_proiect = proiectul

  Filenr = FreeFile
  Dim BitArray As String
  Open CurEnvFile For Binary As #Filenr
  BitArray = Space(LOF(Filenr))
  Get #Filenr, , BitArray
  Close #Filenr

ff1 = Mid(BitArray, 1, InStr(BitArray, "[#PROCESARE#]"))

tmp = Mid(BitArray, InStr(BitArray, "[#PROCESARE#]") + Len("[#PROCESARE#]"), Len(BitArray))
ff2 = Mid(tmp, 1, InStr(tmp, "[#PIC#]") - 1)
tmp = Empty

tmp = Mid(BitArray, InStr(BitArray, "[#PIC#]") + Len("[#PIC#]"), Len(BitArray))
ff3 = Mid(tmp, 1, InStr(tmp, "[#NR_GEL#]") - 1)
tmp = Empty

'segment_total_gena
'    Dim F As Integer
'    F = FreeFile
'    Open App.Path & "\gag.tmp" For Binary As #F
'    Put #F, , ff3
'    Close #F

segment_total_gena.Picture = LoadPicture(Cale_proiect & proiectul & "_ag.agx")
Gel_general.Picture = LoadPicture(Cale_proiect & proiectul & "_gl.agx")

'For u = 0 To Val(ff3) - 1
'Unload Lista_cu_geluri.Gel_nr_X(u)
'Next u

'For u = 0 To Val(ff3) - 1
'Load Lista_cu_geluri.Gel_nr_X(u + Lista_cu_geluri.Gel_nr_X.Count)
'Lista_cu_geluri.Gel_nr_X(u).Picture = LoadPicture(Cale_proiect & proiectul & "_gl" & u & ".agx")
'Next u

COD_CODON.TextRTF = ff1
out_put.Text = ff2
End Sub

Private Sub distanta_n_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
distanta_x.Value = 0
End Sub

Private Sub distanta_x_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
distanta_n.Value = 0
End Sub

Private Sub Fa_micuta_Click()
Me.WindowState = 1
End Sub

Private Sub Form_Resize()
If (Me.WindowState <> 1) Then

bara_date.Width = Me.Width - bara_date.Left - 100
Editare_mijloc.Width = (Me.Width - Dreapta_principal.Width) - stanga_principal.Width

COD_CODON.Left = 100
COD_CODON.Width = Editare_mijloc.Width - 200

out_put.Left = 100
out_put.Width = Editare_mijloc.Width - 200
out_put.Height = Editare_mijloc.Height - COD_CODON.Height - COD_CODON.Top - 50

iesire.Left = Top_miscare.Width - iesire.Width - 100
max_ecran.Left = iesire.Left - max_ecran.Width - 25
Fa_micuta.Left = max_ecran.Left - Fa_micuta.Width - 25

Despartitor_in_out.Width = COD_CODON.Width
Despartitor_in_out.Left = COD_CODON.Left

segment_total_gena.Width = COD_CODON.Width


Proprietati_stanga.Left = Dreapta_principal.Left
Proprietati_stanga.Height = Me.Height - Proprietati_stanga.Top - Anunt_jos.Height

If Proprietati_stanga.Height - Testeaza_CODUL.Top - 100 >= 300 Then
Testeaza_CODUL.Height = Proprietati_stanga.Height - Testeaza_CODUL.Top - 100
End If

Info_GEL.Left = stanga_principal.Left + 20
Info_GEL.Width = stanga_principal.Width - 40
aduna_LA_OUTPUT.Top = Info_GEL.Top + Info_GEL.Height + 100

If stanga_principal.Height - aduna_LA_OUTPUT.Top - 100 >= 300 Then
aduna_LA_OUTPUT.Height = stanga_principal.Height - aduna_LA_OUTPUT.Top - 100
End If

aduna_LA_OUTPUT.Left = Gelul.Left

Printeaza_compileaza.Width = Top_miscare.Width - Printeaza_compileaza.Left - (Top_miscare.Width - Fa_micuta.Left) - 10

End If
End Sub

Private Sub Gel_general_Click()
'Set Obj_de_colorat_principal = Gel_general
End Sub

Private Sub Gel_general_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Gel_vizivil = False Then
Exit Sub
End If

Select_gel1.Visible = False
Select_gel2.Visible = False
Y_baze_linie.x1 = Gelul.ScaleWidth

latime_gel2 = Principal.Gel_general.ScaleWidth
cate_enz = Mai_mult.enzime_de_test.ListCount 'de inlocuit list cu o variabila pt ca utilizatorul poate sterge din enzime cand vrea muschiul lui.
If cate_enz <= 0 Then cate_enz = 1
grosime_linii2 = latime_gel2 / cate_enz
'------------------------------------------------------------------------------
For R = 0 To Mai_mult.enzime_de_test.ListCount
DoEvents
lim11 = (grosime_linii2 * R) + 10
lim22 = (grosime_linii2 * R) + grosime_linii2 - 5

If X > lim11 And X < lim22 Then

Select_multi_test.Left = lim11 - 8
Select_multi_test.Width = lim22 - lim11 + 7

multi_gel_inc1.x1 = 0
multi_gel_inc1.x2 = Select_multi_test.Left
multi_gel_inc2.x1 = Select_multi_test.Left + Select_multi_test.Width
multi_gel_inc2.x2 = Principal.Gelul.ScaleWidth
'Select_multi_test.Height = 6
'Select_multi_test.Top = Y - (Select_multi_test.Height / 2)

'multi_gel_inc1.y1 = Y
'multi_gel_inc1.y2 = Y
'multi_gel_inc2.y1 = Y
'multi_gel_inc2.y2 = Y
'Y_baze_linie.y1 = multi_gel_inc1.y1
'Y_baze_linie.y2 = multi_gel_inc1.y2
'Y_baze_linie.y1 = multi_gel_inc2.y1
'Y_baze_linie.y2 = multi_gel_inc2.y2
On Error Resume Next 'asta e asa casa il compilez pt v1.2
Principal.Gelul.Picture = Lista_cu_geluri.Gel_nr_X(R).Image

am_x_segmente = Split(Lista_cu_geluri.Gel_nr_X(R).ToolTipText, "@")(0)
gel_enzima.Caption = "Enzima: " & Split(Lista_cu_geluri.Gel_nr_X(R).ToolTipText, "@")(1)

End If

Next R
'------------------------------------------------------------------------------
End Sub

Private Sub Gelul_Click()
'Set Obj_de_colorat_principal = Gelul
End Sub

Private Sub Gelul_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Gel_vizivil = False Then
Exit Sub
End If

If blocat = True Then
blocat = False
Else
blocat = True
End If

SC_out_put (vbCrLf & vbCrLf & "Fragment nr. [" & fragmentul_nr_x & "] - lungime: [" & _
Len(secvente_rerstrictate(fragmentul_nr_x)) & "b]" & Proportie_A_T_G_C_U(secvente_rerstrictate(fragmentul_nr_x), False) & vbCrLf & _
"------------------------------------------------------------------------" & vbCrLf & _
secvente_rerstrictate(fragmentul_nr_x) & vbCrLf & _
"------------------------------------------------------------------------" & vbCrLf)
End Sub



Private Sub Gelul_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If blocat = True Then Exit Sub

If (Y <= 0) Then Exit Sub
'-----------------------------------------------------------------------------------------------------
Y_baze_linie.y1 = Y
Y_baze_linie.y2 = Y
'-----------------------------------------------------------------------------------------------------
Select_multi_test.Height = 6
Select_multi_test.Top = Y - (Select_multi_test.Height / 2)
multi_gel_inc1.y1 = Y_baze_linie.y1
multi_gel_inc1.y2 = Y_baze_linie.y2
multi_gel_inc2.y1 = Y_baze_linie.y1
multi_gel_inc2.y2 = Y_baze_linie.y2
'-----------------------------------------------------------------------------------------------------

y_real = Int(((Principal.Gelul.ScaleHeight / 100) * (100 / Y)) * Principal.timpul.Value)
distanta_pe_GEL.Caption = "Migrare : " & y_real & "b - " & Y & "mm"

If Gel_vizivil = False Then
Y_baze_linie.x1 = Gelul.Width
Exit Sub
Else
Select_gel1.Visible = True
Select_gel2.Visible = True
End If
'-----------------------------------------------------------------------------------------------------
Y_baze_linie.x1 = Select_gel2.x1
'------------------------------------------------------------------------------
latime_gel = Principal.Gelul.ScaleWidth
grosime_linii = latime_gel / (am_x_segmente)
'------------------------------------------------------------------------------
For i = 0 To am_x_segmente - 1
DoEvents
lim1 = (grosime_linii * i) + 10
lim2 = (grosime_linii * i) + grosime_linii - 5
If X < lim1 And X > lim2 Then
DoEvents
fragmentul_nr_x = i + 1
Sunt_la_fragmentul.Caption = "Fragment : " & i + 1
Select_gel1.x1 = lim1 + 1
Select_gel1.x2 = lim1 + 1
Select_gel2.x1 = lim2 - 1
Select_gel2.x2 = lim2 - 1
'------------------------------------------------------------------------------
DoEvents
If Ordine_ambii_algoritmi(i + 1) = Empty Then GoTo 666
new_X2 = Ordine_ambii_algoritmi(i + 1)

If i = 0 Then
x1 = 0
Else
new_X1 = Ordine_ambii_algoritmi(i)
x1 = (Principal.segment_total_gena.ScaleWidth / Len(COD_CODON.Text)) * new_X1
End If


x2 = (Principal.segment_total_gena.ScaleWidth / Len(COD_CODON.Text)) * new_X2

select_segment_de_la_gel.Top = 2
select_segment_de_la_gel.Height = 20

select_segment_de_la_gel.Left = x1
select_segment_de_la_gel.Width = x2 - x1
'------------------------------------------------------------------------------
For R = 0 To Mai_mult.enzime_de_test.ListCount
DoEvents
lim11 = (grosime_linii2 * R) + 10
lim22 = (grosime_linii2 * R) + grosime_linii2 - 5

If X < lim11 And X > lim22 Then
If grosime_linii2 <= 0 Then GoTo 44
Select_multi_test.Top = 2
Select_multi_test.Height = 20
Select_multi_test.Left = lim11
Select_multi_test.Width = grosime_linii2
44:
End If

Next R
'------------------------------------------------------------------------------
666:
'------------------------------------------------------------------------------
End If
Next i
'-----------------------------------------------------------------------------------------------------
End Sub



Private Sub genereaza_nuc_Click()
deasupara_reprezentarii_grafice = False
Generator.Show
End Sub

Private Sub Iesire_Click()
tmpo = MsgBox("Doriti sa iesiti din program fara sa salvati proiectul ?", vbYesNo, "Applied Genetics")
If tmpo = 7 Then
Salveza_fisier_Click
Else
Unload Me
Unload Despre
Unload Intro
End
End If
End Sub

Private Sub isButton14_Click()
R = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & App.Path & "\ajutor.htm", vbMaximizedFocus)
End Sub

Private Sub cautare_secvente_Click()
deasupara_reprezentarii_grafice = False
COD_CODON.HideSelection = False
Actiuni.Show
Actiuni.actiuni_bloc_cod.Tab = 0
End Sub

Private Sub Master_timer_Timer()
DoEvents
'-------------------------------------------------------------------------------------------------------------------
de_la_distanta.Caption = Incepand_distanta_DEL.Value
dista.Caption = distanta_DEL.Value

txtStartLen.Text = Brut_dela.Value
txtEndLen.Text = Brut_la.Value

r_dela.Text = r_v_dela.Value
r_la.Text = r_v_la.Value

rep_txt_min.Text = rep_min.Value
rep_txt_max.Text = rep_max.Value

FSV.Caption = "Frame Shift: " & FrameShift_Val.Value

If (Incepand_distanta_DEL.Value >= dista.Caption) Then
d_n.Caption = "Distanta " & dista.Caption & "b ... " & de_la_distanta.Caption & "b"
Else
d_n.Caption = "Distanta " & de_la_distanta.Caption & "b ... " & dista.Caption & "b"
End If

d_x.Caption = "Distanta " & dista.Caption & "b"


'-------------------------------------------------------------------------------------------------------------------
timp_gelus.Caption = timpul.Value
'-------------------------------------------------------------------------------------------------------------------
COD_CODON_inaltime = Despartitor_in_out.Top - COD_CODON.Top + (Despartitor_in_out.Height / 2)
If (COD_CODON_inaltime < 1335) Then
Else
COD_CODON.Height = COD_CODON_inaltime
End If

out_put_inaltime = Editare_mijloc.Height - COD_CODON.Height - COD_CODON.Top - Despartitor_in_out.Height - 50
If (out_put_inaltime < 1335) Then
Else
out_put.Top = Despartitor_in_out.Top + Despartitor_in_out.Height - (Despartitor_in_out.Height / 2)
out_put.Height = out_put_inaltime
End If
'-------------------------------------------------------------------------------------------------------------------

   Dim s As String
   s = Format(Now, "hh:mm:ss ampm")
   If Left(s, 1) = "0" Then ' remove leading zero in hour if present
      s = Right(s, Len(s) - 1)
      ceasul.Value = Left(s, 4) & "E+" & Mid(s, 6, 2)
      'ceasul.Value = Left(s, 4) & " " & Mid(s, 9, 1) 'h:mm A" or "h:mm P"
   Else
      ceasul.Value = Left(s, 5) & "E+" & Mid(s, 7, 2)
      'ceasul.Value = Left(s, 5) & " " & Mid(s, 10, 1)   'hh:mm A" or "hh:mm P"
   End If

End Sub

Private Sub max_ecran_Click()
Exit Sub ' asta o sterg in versiunea 2 cand rezolv marirea si micsorarea reprezentarii vizuale. ;)
If Me.WindowState = 2 Then
Me.WindowState = 0
Else
Me.WindowState = 2
End If
End Sub

Private Sub optiuni_proprietati_arata_Click()
deasupara_reprezentarii_grafice = False
Mai_mult.Show
Mai_mult.op_tab.Tab = 0
End Sub

Private Sub Despartitor_in_out_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
aici_y_vechi = Y
End Sub

Private Sub Despartitor_in_out_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

If (Editare_mijloc.Height - COD_CODON.Height - COD_CODON.Top - Despartitor_in_out.Height - 50 < 1335) Then
If (Despartitor_in_out.Top - aici_y_vechi + Y < old_Despartitor_in_out) Then
Despartitor_in_out.Top = Despartitor_in_out.Top - aici_y_vechi + Y
Else
GoTo 1
End If
End If

If (Despartitor_in_out.Top - COD_CODON.Top + (Despartitor_in_out.Height / 2) < 1335) Then
If (Despartitor_in_out.Top - aici_y_vechi + Y > old_Despartitor_in_out) Then
Despartitor_in_out.Top = Despartitor_in_out.Top - aici_y_vechi + Y
Else
GoTo 1
End If
End If

Despartitor_in_out.Top = Despartitor_in_out.Top - aici_y_vechi + Y
old_Despartitor_in_out = Despartitor_in_out.Top
1:

End If
End Sub

Private Sub out_put_Click()
Set Obj_de_colorat_principal = out_put
End Sub


Private Sub Printare_Click()

bara_date.Max = Lista_cu_geluri.Gel_nr_X.Count + Lista_cu_geluri.secvente_inverse_nr.Count

If Printare.Caption <> "Compilare" Then
MsgBox "Programul nu a terminat de procesat datele !"
Exit Sub
End If

Printare.Caption = "Procesez ..."
    Dim Fifi As Integer
    Fifi = FreeFile

StartUpGDIPlus GdiPlusVersion

rString = Replace$(out_put.Text, vbCrLf, "<br>")

If Nume_proiect = "" Then Nume_proiect = "Proiect1"

Open App.Path & "\print.htm" For Output As #Fifi
Print #Fifi, "<html>"
Print #Fifi, "<body bgcolor='" & DectoWebCol(out_put.BackColor) & "'>"
Print #Fifi, "<center><br><table><td><img src='ico.jpg'></td><td><font size=7 color=" & DectoWebCol(out_put.ForeColor) & ">[" & Nume_proiect & "]</font></td></table><br><br>"
'----------------------------------------------------]insert gel general si interfata gena
If FExist(App.Path & "\gel.jpg") = True Then
Kill App.Path & "\gel.jpg"
End If

If FExist(App.Path & "\graf.jpg") = True Then
Kill App.Path & "\graf.jpg"
End If

Gel_general.Picture = Gel_general.Image
If SavePictureFromHDC(Gel_general.Picture, App.Path & "\gel.jpg") = True Then

ce_tip_img1 = "<img src='gel.jpg'>"
Else
ce_tip_img1 = nu_poza(Gel_general)
End If

segment_total_gena.Picture = segment_total_gena.Image 'ca sa fie jpg trebuie sa fie egale
If SavePictureFromHDC(segment_total_gena.Picture, App.Path & "\graf.jpg") = True Then
ce_tip_img2 = "<img src='graf.jpg'>"
Else
ce_tip_img2 = nu_poza(segment_total_gena)
End If

'-----------------------------------------------------------------------]insert SI.
Print #Fifi, "<table border='1' width='100%'><tr><td valign='middle' align='center'>" & ce_tip_img1 & "</td><td align='center'>" & ce_tip_img2 & "</td></tr>"

For t = 0 To Lista_cu_geluri.secvente_inverse_nr.Count - 1

bara_date.Value = bara_date.Value + 1

If t >= Lista_cu_geluri.secvente_inverse_nr.Count - 1 Then GoTo 33
If FExist(App.Path & "\SI" & t & ".jpg") = True Then
Kill App.Path & "\SI" & t & ".jpg"
End If

Lista_cu_geluri.secvente_inverse_nr(t).Picture = Lista_cu_geluri.secvente_inverse_nr(t).Image
If SavePictureFromHDC(Lista_cu_geluri.secvente_inverse_nr(t).Picture, App.Path & "\SI" & t & ".jpg") = True Then
ce_z = "<img src='SI" & t & ".jpg'>"
Else
ce_z = nu_poza(Lista_cu_geluri.secvente_inverse_nr(t))
End If


    Dim A9 As Variant
    'Dim i As Integer
    i = 1
    A9 = Parse(Lista_cu_geluri.secvente_inverse_nr(t).ToolTipText, "#")

For i = 1 To UBound(A9) - 1 '========================
A9(i) = Replace(A9(i), " ", "")

ea1 = Split(A9(i), "@")(0) ' locul de unde incepe secventa inversa.
ea2 = Split(A9(i), "@")(1)
ea3 = Split(A9(i), "@")(2) ' pentru dimensiune segment intre cele doua secvente.

ea2 = "<a alt='Lungime totala secventa: " & (Len(ea2) * 2) + ea3 & "b' " & _
"onclick='javascript:injectare(" & Chr(34) & "[" & ea2 & "]" & Chr(34) & ");' href='#'>" & ea2 & "</a>"

palindromZ = ""
If ea3 = "0" Then
palindromZ = "[PA]"
Else
palindromZ = "[SI]"
End If

ea_tot = ea_tot & "Nr:[" & i & "]" & palindromZ & ea2 & " de la [" & ea1 & "b]<br>"
22:
Next i '==========================================


Print #Fifi, "<tr><td><font color=" & DectoWebCol(out_put.ForeColor) & " face='Lucida Console'><b>" & ea_tot & "</b></font></td><td width=" & Lista_cu_geluri.secvente_inverse_nr(t).ScaleWidth & ">" & ce_z & "</td></tr>"
ea_tot = ""

33:
Next t

Print #Fifi, "</table><hr>"
ce_tip_img1 = Empty
ce_tip_img2 = Empty
'----------------------------------------------------]insert geluri
If Lista_cu_geluri.Gel_nr_X.Count <= 1 Then GoTo 2

Print #Fifi, "<table border='0' id='geluri'>"

For i = 1 To Lista_cu_geluri.Gel_nr_X.Count - 1

bara_date.Value = bara_date.Value + 1

Lista_cu_geluri.Gel_nr_X(i).Picture = Lista_cu_geluri.Gel_nr_X(i).Image
If SavePictureFromHDC(Lista_cu_geluri.Gel_nr_X(i).Picture, App.Path & "\sub_gel" & i & ".jpg") = True Then
ce_t = "<img src='sub_gel" & i & ".jpg'>"
Else
ce_t = nu_poza(Lista_cu_geluri.Gel_nr_X(i))
End If

dh1 = dh1 & "<td align='center'>" & ce_t & "</td>"
e1 = Split(Lista_cu_geluri.Gel_nr_X(i).ToolTipText, "@")(0)
e2 = Split(Lista_cu_geluri.Gel_nr_X(i).ToolTipText, "@")(1)
dh2 = dh2 & "<td><Center><font color='" & DectoWebCol(out_put.ForeColor) & "'>Gel nr:[" & i & "]<br> Contine [" & e1 & "] fragmente.<br> Enzima: [" & e2 & "]</font></center></td>"

v = v + 1
If v > 4 Then

Print #Fifi, "<tr>" & dh1 & "</tr>"
Print #Fifi, "<tr>" & dh2 & "</tr>"
dh2 = ""
dh1 = ""
v = 0
End If

Next i

Print #Fifi, "</table><hr>"

ce_t = Empty
2:


'----------------------------------------------------]insert secventa nucleotide
For w = 1 To Len(COD_CODON.Text) Step 80
ss = LCase(Mid(COD_CODON.Text, w, 80))
ddf = ddf & "</br><font color=" & DectoWebCol(out_put.ForeColor) & " face='Lucida Console'><b>" & ss & "</b></font>"
Next w
Print #Fifi, ddf
'----------------------------------------------------]insert text OUTPUT
Print #Fifi, "</center><hr><br><font color='" & DectoWebCol(out_put.ForeColor) & "'>" & Replace(rString, "--------------------------------------------------------------------------------------", "<hr>") & "</font>"
Print #Fifi, "</body>"
Print #Fifi, "</html>"
Close #Fifi
'----------------------------------------------------] insert script
  If FExist(App.Path & "\scr.db") = True Then
  Dim Filenr1 As Integer
  Filenr1 = FreeFile
  Dim BitArray1 As String
  Open App.Path & "\scr.db" For Binary As #Filenr1
  BitArray1 = Space(LOF(Filenr1))
  Get #Filenr1, , BitArray1
  Close #Filenr1
  End If
  
  If FExist(App.Path & "\print.htm") = True Then
  Dim Filenr2 As Integer
  Filenr2 = FreeFile
  Dim BitArray2 As String
  Open App.Path & "\print.htm" For Binary As #Filenr2
  BitArray2 = Space(LOF(Filenr2))
  Get #Filenr2, , BitArray2
  Close #Filenr2
  End If
  
If FExist(App.Path & "\print.htm") = True Then
Kill App.Path & "\print.htm"
End If
  
    Dim F As Integer
    F = FreeFile
    Open App.Path & "\print.htm" For Binary As #F
    Put #F, , BitArray1 & BitArray2
    Close #F
'----------------------------------------------------]
rString = Empty
Printeaza_compileaza.Caption = sTmp
R = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & App.Path & "\print.htm", vbMaximizedFocus)
Printare.Caption = "Compilare"
bara_date.Value = 0
bara_date.Max = 100
End Sub

Function nu_poza(ByVal obiect As Object) As String
Dim aPix() As String, lPX As Long, bR As Byte, Bg As Byte, bB As Byte
Dim lWidth As Long, oldPix As String, lCur As Long
' reset the file and add the basics

returneaza_l = returneaza_l & "<table border=0 cellpadding=0 cellspacing=0 bgcolor=ffffff>"
returneaza_l = returneaza_l & "<tr height=1>"
ReDim aPix(0 To obiect.ScaleWidth, 0 To obiect.ScaleHeight)
lHeight = 1
lWidth = 1
' setup oldpix here outside the loop to stop from having 1 more command every time, faster in the long run
lPX = GetPixel(obiect.hdc, 0, 0)
bR = lPX And 255
Bg = (lPX \ 256) And 255
bB = (lPX \ 65536) And 255
oldPix = rgbtohex(bR, Bg, bB)
For Y = 0 To obiect.ScaleHeight - 1
    For X = 0 To obiect.ScaleWidth - 1
        lPX = GetPixel(obiect.hdc, X, Y)
        bR = lPX And 255
        Bg = (lPX \ 256) And 255
        bB = (lPX \ 65536) And 255
        aPix(X, Y) = rgbtohex(bR, Bg, bB)
        If oldPix = aPix(X, Y) Then
            lWidth = lWidth + 1
            Else
            'paint a line, the width before the pixel color was changed
            'that way it makes lines not just pixel by pixel unless the line
            'is only 1 pixel long, the height always goes by 1 pixel tho
            returneaza_l = returneaza_l & "<td width=" & lWidth & " bgcolor=" & oldPix & "></td>"
            lWidth = 1
        End If
        oldPix = aPix(X, Y)
        lCur = lCur + 1
        Call UpdateProgress(lCur, ((obiect.ScaleHeight - 1) * (obiect.ScaleWidth - 1)))
    Next
    returneaza_l = returneaza_l & "<td width=" & lWidth & " bgcolor=" & oldPix & "></td>"
    lWidth = 1
    ' print the next height to start at
    If Y <> obiect.ScaleHeight - 1 Then
returneaza_l = returneaza_l & "</tr>"
returneaza_l = returneaza_l & "</table>"
returneaza_l = returneaza_l & "<table border=0 cellpadding=0 cellspacing=0 bgcolor=ffffff>"
returneaza_l = returneaza_l & "<tr height=1>"
    End If
Next
returneaza_l = returneaza_l & "</table></center><center>"
nu_poza = returneaza_l
End Function



Sub UpdateProgress(lProg As Long, lMax As Long)
Dim lTmp As Long, sTmp As String
lTmp = ((lProg * 100) / lMax)

For X = 0 To lTmp
    sTmp = sTmp & "||"
Next
DoEvents
'bara_date.Value = lTmp
Printeaza_compileaza.Caption = sTmp
DoEvents
End Sub

'not my function, got it offa pscode imlazy
'http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=49849&lngWId=1
Public Function rgbtohex(R As Byte, g As Byte, b As Byte) As String

'input format = 255,255,255

'get the r value
If R < 16 Then
hex1 = 0 & Hex(R)
Else
hex1 = Hex(R)
End If


'get the g value
If R < 16 Then
hex2 = 0 & Hex(g)
Else
hex2 = Hex(g)
End If


'get the b value
If b < 16 Then
hex3 = 0 & Hex(b)
Else
hex3 = Hex(b)
End If

rgbtohex = "#" & hex1 & hex2 & hex3
End Function




Private Sub reset_text_Click()
COD_CODON.SelStart = 0
COD_CODON.SelLength = Len(COD_CODON.Text)
COD_CODON.SelColor = RGB(255, 255, 255)
'-------------------------------
COD_CODON.SelBold = False
COD_CODON.SelItalic = False
COD_CODON.SelStrikeThru = False
COD_CODON.SelUnderline = False
'-------------------------------
COD_CODON.SelStart = 0
COD_CODON.SelLength = 0
COD_CODON.TextRTF = COD_CODON.Text

segment_total_gena.Cls
' de resetat si gelul de electroforeza la nivel variabile, bare si amd.
Gel_general.Cls
Gelul.Cls

out_put.Text = Empty
End Sub

Private Sub Cauta_posibilitati_Click()
Tab_proprietati.Tab = 0
End Sub


Private Sub COD_CODON_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
selectare_start.Caption = "0"
selectare_stop.Caption = "0"
Total_selectie.Caption = "0"
Ce_E.Caption = "Primer = 0b"

Textul_selectat.Caption = ""
Textul_selectat_coada = ""
Posibil_tm.Caption = "Tm = 0C°"
procent_GC.Caption = "Procentul G+C = 0%"
End Sub


Private Sub COD_CODON_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'deasupara_reprezentarii_grafice = False
'If Len(COD_CODON.Text) <= 0 Then Exit Sub

'plus_minus = CATENA_PLUS(COD_CODON.SelText)
plus_minus = COD_CODON.SelText
zulu_selectie = Int((Len(COD_CODON.Text) / Principal.segment_total_gena.ScaleWidth))

selectare_start.Caption = COD_CODON.SelStart
selectare_stop.Caption = COD_CODON.SelStart + COD_CODON.SelLength
Total_selectie.Caption = COD_CODON.SelLength

If (Len(COD_CODON.SelText) > 20) Then
Textul_selectat.Caption = Mid(COD_CODON.SelText, 1, 20) & "..."
Else
If (Len(COD_CODON.SelText) > 0) Then Textul_selectat.Caption = COD_CODON.SelText Else Textul_selectat.Caption = ""
End If

If (Len(COD_CODON.SelText) > 31) Then
Textul_selectat_coada.Caption = "..." & Mid(COD_CODON.SelText, Len(COD_CODON.SelText) - 10, Len(COD_CODON.SelText))
Else
Textul_selectat_coada.Caption = ""
End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If (Len(plus_minus) > zulu_selectie) Then

If (Len(plus_minus) < 31 And Len(plus_minus) > 18) Then
Ce_E.Caption = "Sonda = " & Total_selectie.Caption & "b"
End If

If (Len(plus_minus) <= 18) Then
Ce_E.Caption = "Primer = " & Total_selectie.Caption & "b"
End If

If (Len(plus_minus) < 31) Then
For n = 1 To Len(plus_minus)
d = LCase(Mid(plus_minus, n, 1))
If (d = "a") Then A = A + 1
If (d = "t") Then t = t + 1
If (d = "g") Then g = g + 1
If (d = "c") Then c = c + 1
If (d = "u") Then u = u + 1
Next n

total_nucleotideX = Len(plus_minus)
GC_la_suta = Int((100 / total_nucleotideX) * g) + Int((100 / total_nucleotideX) * c)
procent_GC.Caption = "Procentul G+C = " & GC_la_suta & "%"
Posibil_tm.Caption = formula_primer_PCR(A, c, g, t, GC_la_suta, total_nucleotideX) ' apeleaza calculul ;)

End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

If Tab_proprietati.Tab = 3 Then
'plus_minus
End If

End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'deasupara_reprezentarii_grafice = True
End Sub

Function formula_primer_PCR(ByVal a_TM As Variant, ByVal c_TM As Variant, ByVal g_TM As Variant, ByVal t_TM As Variant, ByVal GC_suta As Variant, ByVal n_total_x As Variant) As String
If Mai_mult.Formula_TM_primer_ADN(0).Value = 1 Then
formula_primer_PCR = "Tm = " & Val(Val(Mai_mult.nr_GC_calcul.Text) * (g_TM + c_TM) + Val(Mai_mult.nr_AT_calcul.Text) * (a_TM + t_TM)) & "C°"
Else
formula_primer_PCR = "Tm = " & Int(Val(Mai_mult.Nr_1.Text) + Val(Mai_mult.Nr_2.Text) * (Log(Val(Mai_mult.log_concentratie_Na.Text)) / Log(10)) + Val(Mai_mult.Nr_3.Text) * (GC_suta) - Val(Mai_mult.Nr_4.Text) / n_total_x) & "C°"
End If
End Function

Private Sub coloreaza_cuvantul_Click()
Tab_proprietati.Tab = 1
End Sub

Private Sub Form_Load()

cmd_prm = Empty
'On Error GoTo ClickError
If Command <> "%1" Then
If Command <> "" Then
cmd_prm = Command
Deschidere_fisier_Click
End If
ElseIf Command <> "" Then
End If
'MsgBox Command
'ClickError:
'    If Err.Number <> 0 Then Exit Sub
'-------------------------------------------------------------------------------------
'On Error GoTo 1

If FExist(App.Path & "\aso.dat") = True Then
  Filenr = FreeFile
  Dim BitArray As String
  Open App.Path & "\aso.dat" For Binary As #Filenr
  BitArray = Space(LOF(Filenr))
  Get #Filenr, , BitArray
  Close #Filenr
End If

If BitArray <> "[#AG#]" Then
Call AssociateMe(".pro", "AG Text", App.Path + "\", "AG")
Call AssociateMe2(".agx", "MS Text", App.Path + "\", "MS")
    Dim F As Integer
    F = FreeFile
    Open App.Path & "\aso.dat" For Binary As #F
    Put #F, , "[#AG#]"
    Close #F
End If


If GetSetting(App.ExeName, "Setari", "p1") = "" Then
COD_CODON.SelStart = 0
COD_CODON.SelLength = Len(COD_CODON.Text)
COD_CODON.SelColor = vbWhite
COD_CODON.SelStart = 0
Else
'--------------------------------------------------------------------------------------
COD_CODON.Font.Name = GetSetting(App.ExeName, "Setari", "p1")
COD_CODON.Font.size = GetSetting(App.ExeName, "Setari", "p2")
COD_CODON.Font.Bold = GetSetting(App.ExeName, "Setari", "p3")
COD_CODON.Font.Italic = GetSetting(App.ExeName, "Setari", "p4")
COD_CODON.Font.Underline = GetSetting(App.ExeName, "Setari", "p5")
COD_CODON.Font.Strikethrough = GetSetting(App.ExeName, "Setari", "p6")

COD_CODON.SelStart = 0
COD_CODON.SelLength = Len(COD_CODON.Text)
COD_CODON.SelColor = GetSetting(App.ExeName, "Setari", "p7")
COD_CODON.SelStart = 0

COD_CODON.BackColor = GetSetting(App.ExeName, "Setari", "p8")
'--------------------------------------------------------------------------------------
End If

If GetSetting(App.ExeName, "Setari", "o1") <> "" Then
out_put.FontName = GetSetting(App.ExeName, "Setari", "o1")
out_put.FontSize = GetSetting(App.ExeName, "Setari", "o2")
out_put.FontBold = GetSetting(App.ExeName, "Setari", "o3")
out_put.FontItalic = GetSetting(App.ExeName, "Setari", "o4")
out_put.FontUnderline = GetSetting(App.ExeName, "Setari", "o5")
out_put.FontStrikethru = GetSetting(App.ExeName, "Setari", "o6")
out_put.ForeColor = GetSetting(App.ExeName, "Setari", "o7")
out_put.BackColor = GetSetting(App.ExeName, "Setari", "o8")
'--------------------------------------------------------------------------------------
End If

If GetSetting(App.ExeName, "Culori1", "SESCRIE") <> "" Then
Dim intSettings As Integer
Dim ctr As Control

MySettings = GetAllSettings(App.ExeName, "Culori1")
MySettings2 = GetAllSettings(App.ExeName, "Culori2")

For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
   For Each ctr In Mai_mult.Controls
   If ctr.Name = MySettings(intSettings, 0) Then
   ctr.BackColor = MySettings(intSettings, 1)
   'ctr.ForeColor = MySettings2(intSettings, 1)
   End If
   Next ctr
Next intSettings

For intSettings = LBound(MySettings2, 1) To UBound(MySettings2, 1)
   For Each ctr In Mai_mult.Controls
   If ctr.Name = MySettings2(intSettings, 0) Then ctr.ForeColor = MySettings2(intSettings, 1)
   Next ctr
Next intSettings
End If


bara_de_delimitare_text = "--------------------------------------------------------------------------------------"

nucleotida(1) = "A"
nucleotida(2) = "T"
nucleotida(3) = "G"
nucleotida(4) = "C"
nucleotida(5) = "U"

Gel_vizivil = False

'Proprietati_Enzime.Visible = False
sel_manuala = Culoarea1.BackColor
Master_timer.Enabled = True

Select_gel1.Visible = False
Select_gel2.Visible = False

sterge_selectie = False
'fara_couloare.BackColor = COD_CODON.BackColor
'Call Segment_geneza(2, 20, vbWhite)
'MsgBox cmd_prm & " cmd_prm"
'-------------------------------------------------------|
Set cBrute = New clsBruteforce

With cBrute
    txtCharacterset = .CharacterSet
    txtStartLen = .StartLength
    txtEndLen = .EndLength
    txtStartCombo = .StartWord
    
    r_dela.Text = txtStartLen.Text
    r_la.Text = txtEndLen.Text
End With
'-------------------------------------------------------|
adresa_site_general = "https://github.com/Gagniuc/Applied-Genetics-1.2"

If FExist(App.Path & "\sit.dat") = True Then
  Filenr = FreeFile
  Open App.Path & "\sit.dat" For Binary As #Filenr
  BitArray = Space(LOF(Filenr))
  Get #Filenr, , BitArray
  Close #Filenr
  
  adresa_site_general = BitArray
End If
'-------------------------------------------------------|

If GetSetting(App.ExeName, "Codoni", "Notatie") = "" Then
amino_notatie_scurta = True
Call SaveSetting(App.ExeName, "Codoni", "Notatie", "1")
Else
Mai_mult.Notatie_amino.Value = GetSetting(App.ExeName, "Codoni", "Notatie")
If Mai_mult.Notatie_amino.Value = 1 Then
amino_notatie_scurta = True
Mai_mult.Notatie_amino.Caption = "Notatie scurta. Ex:[M,P,Q]"
Else
amino_notatie_scurta = False
Mai_mult.Notatie_amino.Caption = "Notatie lunga. Ex:[Met,Pro,Gln]"
End If
End If
'-------------------------------------------------------|

If GetSetting(App.ExeName, "Codoni", "codon0") = "" Then
For i = 0 To Mai_mult.codoni_aminoacidului.Count - 1
Call SaveSetting(App.ExeName, "Codoni", "codon" & i, Mai_mult.codoni_aminoacidului(i).Caption)
Next i
Else
For i = 0 To Mai_mult.codoni_aminoacidului.Count - 1
Mai_mult.codoni_aminoacidului(i).Caption = GetSetting(App.ExeName, "Codoni", "codon" & i)
Next i
End If
'-------------------------------------------------------|
If GetSetting(App.ExeName, "Formule", "f1") = "" And GetSetting(App.ExeName, "Formule", "c") = "" Then

Call SaveSetting(App.ExeName, "Formule", "f1", Mai_mult.nr_GC_calcul.Text)
Call SaveSetting(App.ExeName, "Formule", "f2", Mai_mult.nr_AT_calcul.Text)

Call SaveSetting(App.ExeName, "Formule", "l1", Mai_mult.log_concentratie_Na.Text)


Call SaveSetting(App.ExeName, "Formule", "n1", Mai_mult.Nr_1.Text)
Call SaveSetting(App.ExeName, "Formule", "n2", Mai_mult.Nr_2.Text)
Call SaveSetting(App.ExeName, "Formule", "n3", Mai_mult.Nr_3.Text)
Call SaveSetting(App.ExeName, "Formule", "n4", Mai_mult.Nr_4.Text)

Call SaveSetting(App.ExeName, "Formule", "c", "1")
Else '-------------------------------------------------------|

Mai_mult.nr_GC_calcul.Text = GetSetting(App.ExeName, "Formule", "f1")
Mai_mult.nr_AT_calcul.Text = GetSetting(App.ExeName, "Formule", "f2")

Mai_mult.log_concentratie_Na.Text = GetSetting(App.ExeName, "Formule", "l1")

Mai_mult.Nr_1.Text = GetSetting(App.ExeName, "Formule", "n1")
Mai_mult.Nr_2.Text = GetSetting(App.ExeName, "Formule", "n2")
Mai_mult.Nr_3.Text = GetSetting(App.ExeName, "Formule", "n3")
Mai_mult.Nr_4.Text = GetSetting(App.ExeName, "Formule", "n4")

zz_tmp = GetSetting(App.ExeName, "Formule", "c")
If zz_tmp = "1" Then
Mai_mult.Formula_TM_primer_ADN(0).Value = 1
Mai_mult.Formula_TM_primer_ADN(1).Value = 0
Else
Mai_mult.Formula_TM_primer_ADN(0).Value = 1
Mai_mult.Formula_TM_primer_ADN(1).Value = 0
End If

End If
'-------------------------------------------------------|
'Call ReplaceColor(GetPixel(segment_total_gena.hdc, 1, 1), Mai_mult.R_cul_grafic.BackColor, principal.segment_total_gena)
'-------------------------------------------------------|
deasupara_reprezentarii_grafice = True
afisat_pan = False
Stanga_af_gr.Left = -(Stanga_af_gr.Width)
Stanga_af_gr.Top = -1
'-------------------------------------------------------|
culoare1_gel_old = Mai_mult.GEL_CUL.ForeColor
culoare2_gel_old = Mai_mult.GEL_CUL.BackColor
If Command <> "%1" Then
If Command = "" Then
Gel_general.BackColor = Mai_mult.GEL_CUL.BackColor
Gelul.BackColor = Mai_mult.GEL_CUL.BackColor
Lista_cu_geluri.Gel_nr_X(0).BackColor = Mai_mult.GEL_CUL.BackColor

segment_total_gena.BackColor = Mai_mult.R_cul_grafic.BackColor
Segment_grafic.BorderColor = Mai_mult.R_cul_grafic.ForeColor
Else
'-'
End If
End If
'-------------------------------------------------------|
'Gaseste caracterele inproprii codului genetic si afiseaza secventa filtrata in cazul in care aceste caractere sunt depistate!
'Salturi directe la o anumita pozitie in secventa sau cautari de secvente definite de utilizator !
'Generare secvente ADN sau ARN !

'Genereaza secventa de la coada la cap. Ex: gaaatt va fi ttaaag

Revers.Left = Antisens.Left + Antisens.Width + 10
ARN_la_ADN.Left = Revers.Left + Revers.Width + 10
ADN_la_ARN.Left = ARN_la_ADN.Left + ARN_la_ADN.Width + 10
Adn_prot.Left = ADN_la_ARN.Left + ADN_la_ARN.Width + 10
End Sub


Private Sub Gaseste_Eruari_Click()
deasupara_reprezentarii_grafice = False
'00000000000000000000000000000000000000000000
Gasit_bruiaj = False
'00000000000000000000000000000000000000000000
COD_CODON.SelStart = 0 ' altfel coloreaza tot naspa
For R = 1 To Len(COD_CODON.Text)
DoEvents
'--------
Principal.bara_date.Value = (Principal.bara_date.Max / Len(COD_CODON.Text)) * R
'--------
nuc = Mid(COD_CODON.Text, R, 1)
'00000000000000000000000000000000000000000000
If (filtru_nucleotide(nuc) <> 1) Then
COD_CODON.SelColor = RGB(255, 0, 0)
COD_CODON.SelStart = R - 1
COD_CODON.SelLength = 1
Gasit_bruiaj = True
b = b + 1
Else
n = n + 1
j = j & UCase(nuc)
End If
'00000000000000000000000000000000000000000000
Next R
COD_CODON.SelColor = RGB(255, 255, 255)
'--------
If (Gasit_bruiaj = True) Then
da_nu_bruiaj = "Au fost gasite " & b & " caractere care nu fac parte din codul genetic."
out_put.Text = da_nu_bruiaj & vbCrLf & "Secventa filtrata[" & n & " b]:" & vbCrLf & out_put.Text & j & vbCrLf
Else
da_nu_bruiaj = "Nu au fost gasite caractere inproprii codului genetic."
out_put.Text = da_nu_bruiaj & vbCrLf
End If
'--------
'00000000000000000000000000000000000000000000
'--------
Principal.bara_date.Value = 0
'--------
deasupara_reprezentarii_grafice = True
End Sub

Function filtru_nucleotide(ByVal xx As String) As Variant
Dim gasit As Boolean
xx = LCase(xx)
For n = 1 To 5
If (xx = LCase(nucleotida(n))) Then filtru_nucleotide = 1: Exit Function
Next n
End Function

Private Sub test_Click()
    'MsgB "My Message" & vbCrLf & "Line2" & vbCrLf & "Line 3" & vbCrLf & "Line 4" & vbCrLf & "Line 5" & vbCrLf & "Line 6" & vbCrLf & "Line 7", mbxMessenger, mbxMoney, mbxYesNoCancel, mbxTrash, "My Title"
    'MsgB "Trial Message", mbxGradient, mbxMetallic, mbxOKOnly, mbxAlert, "Trial Message"
    'MsgB "Trial Message to chech if this message box changes its width according to length of message", mbxMessenger, mbxOlive, mbxPrintDontPrint, mbxPrint, "Print"
    
End Sub

Private Sub Salveza_fisier_Click()
    
'On Error Resume Next
Dim CurEnvFile As String
CommonDialog1.FileName = "Proiect1"
CommonDialog1.Filter = "Applied Genetics v1.2|*.pro"
CommonDialog1.ShowSave
CurEnvFile = CommonDialog1.FileName
If CommonDialog1.FileName = "" Or CommonDialog1.FileTitle = "" Then Exit Sub



Cale_proiect = Mid(CommonDialog1.FileName, 1, InStrRev(CommonDialog1.FileName, "\"))
proiectul = Mid(CommonDialog1.FileTitle, 1, InStrRev(CommonDialog1.FileTitle, ".") - 1)
SavePicture segment_total_gena.Image, Cale_proiect & proiectul & "_ag.agx"
SavePicture Gel_general.Image, Cale_proiect & proiectul & "_gl.agx"

Nume_proiect = proiectul
'For i = 0 To Lista_cu_geluri.Gel_nr_X.Count - 1
'SavePicture Lista_cu_geluri.Gel_nr_X(i).Image, Cale_proiect & proiectul & "_gl" & i & ".agx"
'Next i


'SavePicture segment_total_gena.Image, App.Path & "\" & proiectul & "gl2.agx"


'  Filenr = FreeFile
'  Dim BitArray As String
'  Open App.Path & "\ag.tmp" For Binary As #Filenr
'  BitArray = Space(LOF(Filenr))
'  Get #Filenr, , BitArray
'  Close #Filenr
    
'  Filenr = FreeFile
'  Dim BitArray2 As String
'  Open App.Path & "\gl1.tmp" For Binary As #Filenr
'  BitArray2 = Space(LOF(Filenr))
'  Get #Filenr, , BitArray2
'  Close #Filenr
  
'  Filenr = FreeFile
'  Dim BitArray3 As String
'  Open App.Path & "\gl2.tmp" For Binary As #Filenr
'  BitArray3 = Space(LOF(Filenr))
'  Get #Filenr, , BitArray3
'  Close #Filenr

'xxc = xxc & Gel_general.BackColor & ","
'xxc = xxc & Gelul.BackColor & ","

    Dim F As Integer
    F = FreeFile
    Open CurEnvFile For Binary As #F
    Put #F, , COD_CODON.TextRTF & "[#PROCESARE#]" & out_put.Text & "[#PIC#]" & Lista_cu_geluri.Gel_nr_X.Count & "[#NR_GEL#]"  '& BitArray2 & "[#GEL2#]" & BitArray3
    Close #F

'BitArray = Empty
'BitArray2 = Empty
'BitArray3 = Empty;)



Exit Sub
SaveCodeError:
MsgBox "Va rugam verificati calea pana la fisier !", , ""
End Sub

Private Sub segment_total_gena_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Segment_grafic.x1 = X
Segment_grafic.x2 = X

If deasupara_reprezentarii_grafice = False Then Exit Sub
COD_CODON.SetFocus

COD_CODON.SelStart = Int((Len(COD_CODON.Text) / Principal.segment_total_gena.ScaleWidth) * X)
zulu_selectie = Int((Len(COD_CODON.Text) / Principal.segment_total_gena.ScaleWidth))
If zulu_selectie < 1 Then COD_CODON.SelLength = 1 Else COD_CODON.SelLength = zulu_selectie
End Sub

Private Sub Tabel_culori_Click()
On Error GoTo bye
If Button = 1 Then
    Culoarea1.BackColor = Tabel_culori.Point(X, Y)
    'Obj_de_colorat_principal.BackColor = Tabel_culori.Point(X, Y)
End If
If Button = 2 Then
    Culoarea2.BackColor = Tabel_culori.Point(X, Y)
    'Obj_de_colorat_principal.ForeColor = Tabel_culori.Point(X, Y)
End If
bye:
End Sub

Private Sub Tabel_culori_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Culoarea1.BackColor = Tabel_culori.Point(X, Y)
On Error GoTo bye
If Button = 1 Then
    Culoarea1.BackColor = Tabel_culori.Point(X, Y)
    'Obj_de_colorat_principal.BackColor = Tabel_culori.Point(X, Y)
    sel_manuala = Culoarea1.BackColor
End If
If Button = 2 Then
    Culoarea2.BackColor = Tabel_culori.Point(X, Y)
    'Obj_de_colorat_principal.ForeColor = Tabel_culori.Point(X, Y)
End If
bye:
End Sub

Private Sub Tabel_culori_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo bye
If Button = 1 Then
    Culoarea1.BackColor = Tabel_culori.Point(X, Y)
    sel_manuala = Culoarea1.BackColor
End If
If Button = 2 Then
    Culoarea2.BackColor = Tabel_culori.Point(X, Y)
End If
bye:
End Sub


Private Sub Testeaza_CODUL_Click()
Debug.Print "[" & Timer & "] Testeaza_CODUL_Click ENTER - Tab=" & Tab_proprietati.Tab & " Caption=" & Testeaza_CODUL.Caption & " bStop=" & bStop
deasupara_reprezentarii_grafice = False
'Call Wav_Play(App.Path & "\st.wav")
'***************************************************************(*
If Tab_proprietati.Tab = 0 Then
COD_CODON.HideSelection = True

If Testeaza_CODUL.Caption = "Start" Then
Run_Primar.Caption = "Stop"
Testeaza_CODUL.Caption = "Stop"
DoEvents
Call ALEGE_ALGORITM_ENZIME(Principal)
Testeaza_CODUL.Caption = "Start"
Run_Primar.Caption = "Start"
Else
Testeaza_CODUL.Caption = "Start"
Run_Primar.Caption = "Start"
End If

End If
'***************************************************************(*
If Tab_proprietati.Tab = 1 Then
COD_CODON.HideSelection = True

If Testeaza_CODUL.Caption = "Start" Then
Run_Primar.Caption = "Stop"
Testeaza_CODUL.Caption = "Stop"
DoEvents
out_put.Text = Gena_codon_aminoacid(ADN_ARN_f(COD_CODON.Text), amino_notatie_scurta)
Testeaza_CODUL.Caption = "Start"
Run_Primar.Caption = "Start"
Else
Testeaza_CODUL.Caption = "Start"
Run_Primar.Caption = "Start"
End If
End If
'***************************************************************(*
If Tab_proprietati.Tab = 2 Then

deasupara_reprezentarii_grafice = False

COD_CODON.HideSelection = True
txtCharacterset.Enabled = Not txtCharacterset.Enabled
txtStartLen.Enabled = Not txtStartLen.Enabled
txtEndLen.Enabled = Not txtEndLen.Enabled
txtStartCombo.Enabled = Not txtStartCombo.Enabled
txtTarget.Enabled = Not txtTarget.Enabled

If Testeaza_CODUL.Caption = "Start" Then
Testeaza_CODUL.Caption = "Stop"
Run_Primar.Caption = "Stop"

For i = 1 To Lista_cu_geluri.secvente_inverse_nr.Count - 1
        Unload Lista_cu_geluri.secvente_inverse_nr(i)
Next

Lista_cu_geluri.secvente_inverse_nr(0).ToolTipText = ""
Lista_cu_geluri.secvente_inverse_nr(0).Width = segment_total_gena.ScaleWidth
Lista_cu_geluri.fer_seg1.Width = segment_total_gena.ScaleWidth
Lista_cu_geluri.fer_seg2.Width = segment_total_gena.ScaleWidth

Lista_cu_geluri.secvente_inverse_nr(0).BackColor = Mai_mult.R_cul_grafic.BackColor
Lista_cu_geluri.fer_seg1.BackColor = Mai_mult.R_cul_grafic.BackColor
Lista_cu_geluri.fer_seg2.BackColor = Mai_mult.R_cul_grafic.BackColor

Lista_cu_geluri.SI_derulare.Left = Lista_cu_geluri.fer_seg1.Left + Lista_cu_geluri.fer_seg1.ScaleWidth
Lista_cu_geluri.Width = (Lista_cu_geluri.SI_derulare.Left + Lista_cu_geluri.SI_derulare.Width + 10) * Screen.TwipsPerPixelX
Lista_cu_geluri.iesire.Left = ((Lista_cu_geluri.Width / Screen.TwipsPerPixelX) / 2) - (Lista_cu_geluri.iesire.Width / 2)

With cBrute
        .CharacterSet = txtCharacterset
        .StartLength = Val(txtStartLen)
        .EndLength = Val(txtEndLen)
        If Len(txtStartCombo) > 0 Then .StartWord = txtStartCombo
End With
    
bStop = False
intBenchmarkSec = 0
DummyLoop
Else
    Debug.Print "[" & Timer & "] Tab=2 STOP branch - setting bStop=True (was " & bStop & ")"
    Testeaza_CODUL.Caption = "Start"
    Run_Primar.Caption = "Start"
    bStop = True
    Debug.Print "[" & Timer & "] Tab=2 STOP branch - bStop is now " & bStop

If Principal.SI_grafic_step_by_step.Value = 1 Then
Lista_cu_geluri.Show
Else
Lista_cu_geluri.Hide
End If
    
End If
End If
'***************************************************************]
If Tab_proprietati.Tab = 3 Then
COD_CODON.HideSelection = True
txtCharacterset.Enabled = Not txtCharacterset.Enabled
txtTarget.Enabled = Not txtTarget.Enabled
r_dela.Enabled = Not r_dela.Enabled
r_la.Enabled = Not r_la.Enabled
'txtStartCombo.Enabled = Not txtStartCombo.Enabled


If Testeaza_CODUL.Caption = "Start" Then
Testeaza_CODUL.Caption = "Stop"
Run_Primar.Caption = "Stop"
    
With cBrute
        .CharacterSet = txtCharacterset
        .StartLength = Val(r_dela.Text)
        .EndLength = Val(r_la.Text)
        'If Len(txtStartCombo) > 0 Then .StartWord = txtStartCombo
End With
    
bStop = False
intBenchmarkSec = 0
Debug.Print "[" & Timer & "] Tab=3 calling DummyLoop, bStop=" & bStop & " StartLength=" & r_dela.Text & " EndLength=" & r_la.Text
DummyLoop
Debug.Print "[" & Timer & "] Tab=3 DummyLoop RETURNED, bStop=" & bStop
Else
    Debug.Print "[" & Timer & "] Tab=3 STOP branch - setting bStop=True (was " & bStop & ")"
    Testeaza_CODUL.Caption = "Start"
    Run_Primar.Caption = "Start"
    bStop = True
    Debug.Print "[" & Timer & "] Tab=3 STOP branch - bStop is now " & bStop
End If
End If
'***************************************************************]
deasupara_reprezentarii_grafice = True
End Sub

Private Sub Top_miscare_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Transpozoni_Click()
Tab_proprietati.Tab = 2
End Sub

'***************************************************************(*
Sub DummyLoop()
Dim strTarget As String
Dim lLoopIter As Long
lLoopIter = 0

' Fix: previne pornirea unei noi bucle cat timp una veche inca ruleaza.
' Testeaza_CODUL si Run_Primar sunt doua controale suprapuse care apeleaza
' acelasi handler -- daca un singur click e procesat de doua ori (o data
' de fiecare control), a doua invocare vede caption-ul deja "Start" si
' crede gresit ca e un click nou de pornire, repornind bucla imediat dupa
' ce a fost oprita. Cu flag-ul asta, a doua invocare e ignorata.
If bLoopRunning Then
    Debug.Print "[" & Timer & "] DummyLoop: bLoopRunning deja True -> Exit Sub (reintrare ignorata)"
    Exit Sub
End If
bLoopRunning = True
Debug.Print "[" & Timer & "] DummyLoop ENTER - Tab=" & Tab_proprietati.Tab & " bStop=" & bStop

strTarget = txtTarget

Do While bStop <> True
' Fix: DoEvents explicit la fiecare iteratie, ca butonul Stop sa fie
' mereu procesat -- inainte, DoEvents se intampla doar sporadic (in
' interiorul clsBruteforce.Combinations), asa ca uneori click-ul pe
' Stop nu era prins la timp.
DoEvents
If bStop = True Then
    Debug.Print "[" & Timer & "] DummyLoop: bStop=True detectat dupa DoEvents, iesim din Do"
    Exit Do
End If
assas = UCase(cBrute.BruteForce)
lLoopIter = lLoopIter + 1
If lLoopIter Mod 20000 = 0 Then
    Debug.Print "[" & Timer & "] DummyLoop iteratia " & lLoopIter & " assas=" & assas & " bStop=" & bStop
End If
'***************************************************************]
If Tab_proprietati.Tab = 2 Then
    If Len(assas) > Val(txtEndLen) - 1 Then
        cBrute.ForceCombinationCalc
        MsgBox "Procesare terminata !", vbApplicationModal + vbInformation, "Secvente inverse !"
        Testeaza_CODUL_Click
    End If
    
 If VER_PALINDROM_DISTANTA(assas) = True Then ' imi spune daca exista macar un polindrom./
 Principal.out_put.Text = Principal.out_put.Text & SECVENTE_INVERSE(Principal.COD_CODON, assas, COD_CODON.Text)
 Principal.out_put.SelStart = Len(Principal.out_put.Text)
 End If
End If
'***************************************************************]
'***************************************************************]
If Tab_proprietati.Tab = 3 Then
    If Len(assas) > Val(r_la.Text) - 1 Then
        cBrute.ForceCombinationCalc
        MsgBox "Procesare terminata !", vbApplicationModal + vbInformation, "Secvente REPETATE !"
        Testeaza_CODUL_Click
    End If
    
 If VER_REPETITIE_DISTANTA(assas) = True Then ' imi spune daca exista macar 2 repetitii
 Principal.out_put.Text = Principal.out_put.Text & REPETITII(Principal.COD_CODON, assas, COD_CODON.Text)
 Principal.out_put.SelStart = Len(Principal.out_put.Text)
 End If
End If
'***************************************************************]
Loop

Debug.Print "[" & Timer & "] DummyLoop EXIT bucla - total iteratii=" & lLoopIter & " bStop=" & bStop
cBrute.ResetStats
bLoopRunning = False
Debug.Print "[" & Timer & "] DummyLoop EXIT Sub - bLoopRunning=" & bLoopRunning

End Sub

Function VER_PALINDROM_DISTANTA(ByVal dd As String) As Boolean
VER_PALINDROM_DISTANTA = False
zz = InStr(UCase(COD_CODON.Text), dd)
If zz > 0 And zz < Len((COD_CODON.Text)) - Len(dd) Then
cc = Mid(UCase(COD_CODON.Text), zz + Len(dd), Len(COD_CODON.Text))
If InStr(cc, StrReverse(dd)) > 0 Then
VER_PALINDROM_DISTANTA = True
Else
VER_PALINDROM_DISTANTA = False
End If

End If
End Function

Function VER_REPETITIE_DISTANTA(ByVal dd As String) As Boolean
VER_REPETITIE_DISTANTA = False
codul_mod = Replace(UCase(COD_CODON.Text), dd, "*")

' Fix: cz nu se golea intre apeluri -> crestea nemarginit la fiecare
' iteratie din DummyLoop (tab S.R.), incetinind bucla progresiv pana
' parea ca butonul Stop nu mai raspunde.
cz = ""
For Yx = 1 To rep_min.Value
cz = cz & "*"
Next Yx

zz = InStr(codul_mod, cz)
If zz > 0 Then 'And zz < Len((codul_mod)) - Len(cz) Then
VER_REPETITIE_DISTANTA = True
Else
VER_REPETITIE_DISTANTA = False
End If

End Function

Private Sub cBrute_Statistics(lngTotalCombos As Long, lngCombinationsPerSec As Long, strCurrentPassword As String)
Static lngSeconds As Long

If Tab_proprietati.Tab = 2 Then
    lblComboPS = Format(lngCombinationsPerSec, "#,###,##0")
    lblCurrentCOmbo = strCurrentPassword
    lblComboLen = Len(strCurrentPassword)
    lblTotalCombo = Format(lngTotalCombos, "###,###,###,##0")
End If

If Tab_proprietati.Tab = 3 Then
    c_s_r.Caption = Format(lngCombinationsPerSec, "#,###,##0")
    c_c_r.Caption = strCurrentPassword
    c_len_r.Caption = Len(strCurrentPassword)
    c_t_r.Caption = Format(lngTotalCombos, "###,###,###,##0")
End If

End Sub
