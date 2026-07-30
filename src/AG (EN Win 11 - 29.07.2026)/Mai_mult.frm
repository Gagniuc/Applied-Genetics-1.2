VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Mai_mult 
   BorderStyle     =   0  'None
   Caption         =   "Options - properties"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab op_tab 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "RESTRICTION ENZYMES"
      TabPicture(0)   =   "Mai_mult.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "db_enzime"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "enzime_de_test"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "exp_enzimeq"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "PROGRAM COLORS"
      TabPicture(1)   =   "Mai_mult.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Culoarea1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Tabel_culori"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Culoarea2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Picture1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Expl_culori"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "PROGRAM TEXT"
      TabPicture(2)   =   "Mai_mult.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Ferestre_plat_forma"
      Tab(2).Control(1)=   "bg_txt1"
      Tab(2).Control(2)=   "Tabel_culori_font"
      Tab(2).Control(3)=   "bg_txt2"
      Tab(2).Control(4)=   "Picture2"
      Tab(2).Control(5)=   "Frame2"
      Tab(2).Control(6)=   "font_tip"
      Tab(2).Control(7)=   "font_marime_real"
      Tab(2).Control(8)=   "font_marime"
      Tab(2).Control(9)=   "Frame1"
      Tab(2).Control(10)=   "Mostra_text"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "CODONS"
      TabPicture(3)   =   "Mai_mult.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CODAMINIFORMA"
      Tab(3).Control(1)=   "codoni_explicatii_formal"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "PRIMERS"
      TabPicture(4)   =   "Mai_mult.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "primer_formal"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame primer_formal 
         Caption         =   "PRIMER FORMULA"
         Height          =   5055
         Left            =   120
         TabIndex        =   128
         Top             =   600
         Width           =   8415
         Begin VB.Frame Primer_platforma 
            Caption         =   "Formula type"
            Height          =   1575
            Index           =   0
            Left            =   240
            TabIndex        =   135
            Top             =   360
            Width           =   2295
            Begin VB.CheckBox Formula_TM_primer_ADN 
               Caption         =   "Recommended formula"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   137
               Top             =   960
               Width           =   1935
            End
            Begin VB.CheckBox Formula_TM_primer_ADN 
               Caption         =   "Approximate formula"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   136
               Top             =   480
               Value           =   1  'Checked
               Width           =   1935
            End
         End
         Begin VB.Frame Primer_platforma 
            Caption         =   "Explanations"
            Height          =   2775
            Index           =   1
            Left            =   240
            TabIndex        =   133
            Top             =   2040
            Width           =   2295
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   " Here you can change the calculation formula for primer sequences."
               Height          =   1695
               Left            =   120
               TabIndex        =   134
               Top             =   360
               Width           =   1935
            End
         End
         Begin VB.Frame alege_calcul_TM 
            Caption         =   "Recommended formula"
            Height          =   2775
            Index           =   1
            Left            =   2640
            TabIndex        =   130
            Top             =   2040
            Width           =   5535
            Begin VB.TextBox Nr_4 
               Height          =   285
               Left            =   4560
               TabIndex        =   154
               Text            =   "675"
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox Nr_3 
               Height          =   285
               Left            =   3240
               TabIndex        =   152
               Text            =   "0.41"
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox log_concentratie_Na 
               Height          =   285
               Left            =   2520
               TabIndex        =   147
               Text            =   "0.05"
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox Nr_2 
               Height          =   285
               Left            =   1440
               TabIndex        =   146
               Text            =   "16.6"
               Top             =   1200
               Width           =   495
            End
            Begin VB.TextBox Nr_1 
               Height          =   285
               Left            =   720
               TabIndex        =   144
               Text            =   "81.5"
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   "n => Total number of selected nucleotides."
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   157
               Top             =   2160
               Width           =   3255
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   "[Na+] => Na+ molar concentration"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   156
               Top             =   1920
               Width           =   3255
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   "/n"
               Height          =   255
               Index           =   6
               Left            =   5040
               TabIndex        =   155
               Top             =   1200
               Width           =   255
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   " * (%G+C) -"
               Height          =   255
               Index           =   5
               Left            =   3720
               TabIndex        =   153
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   "]) + "
               Height          =   255
               Index           =   4
               Left            =   3000
               TabIndex        =   151
               Top             =   1200
               Width           =   255
            End
            Begin VB.Label mic_nr 
               BackStyle       =   0  'Transparent
               Caption         =   "10"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Index           =   1
               Left            =   2325
               TabIndex        =   150
               Top             =   1335
               Width           =   135
            End
            Begin VB.Label mic_nr 
               BackStyle       =   0  'Transparent
               Caption         =   "10"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   238
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Index           =   0
               Left            =   1960
               TabIndex        =   149
               Top             =   600
               Width           =   165
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   " * (log   ["
               Height          =   255
               Index           =   3
               Left            =   1920
               TabIndex        =   148
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   " + "
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   145
               Top             =   1200
               Width           =   255
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   "Tm = "
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   143
               Top             =   1200
               Width           =   375
            End
            Begin VB.Label afisare_formula_recomandata_ADN 
               Caption         =   "Tm = 81.5 + 16.6 * (log    [Na+]) + 0.41 * (%G+C) - 675/n"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   131
               Top             =   480
               Width           =   4935
            End
         End
         Begin VB.Frame alege_calcul_TM 
            Caption         =   "Approximate formula"
            Height          =   1575
            Index           =   0
            Left            =   2640
            TabIndex        =   129
            Top             =   360
            Width           =   5535
            Begin VB.TextBox nr_AT_calcul 
               Height          =   285
               Left            =   3120
               TabIndex        =   141
               Text            =   "2"
               Top             =   960
               Width           =   375
            End
            Begin VB.TextBox nr_GC_calcul 
               Height          =   285
               Left            =   1920
               TabIndex        =   139
               Text            =   "4"
               Top             =   960
               Width           =   375
            End
            Begin VB.Label afisare_formula_standard_ADN 
               Caption         =   " * (a + t)"
               Height          =   255
               Index           =   2
               Left            =   3480
               TabIndex        =   142
               Top             =   960
               Width           =   615
            End
            Begin VB.Label afisare_formula_standard_ADN 
               Caption         =   " * (g + c) + "
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   140
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label8 
               Caption         =   "Tm ="
               Height          =   255
               Left            =   1440
               TabIndex        =   138
               Top             =   960
               Width           =   375
            End
            Begin VB.Label afisare_formula_standard_ADN 
               Caption         =   "Tm = 4 * (g + c) + 2 * (a + t)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   480
               TabIndex        =   132
               Top             =   360
               Width           =   2535
            End
         End
      End
      Begin VB.Frame codoni_explicatii_formal 
         Caption         =   "Explanations"
         Height          =   615
         Left            =   -74760
         TabIndex        =   125
         Top             =   5160
         Width           =   8295
         Begin VB.Label codoni_explicatii 
            BackStyle       =   0  'Transparent
            Caption         =   "Here you can change the codons associated with certain amino acids, as well as the notation type of the resulting amino acids."
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   8055
         End
      End
      Begin VB.Frame CODAMINIFORMA 
         Caption         =   "Amino acid - CODONS"
         Height          =   4575
         Left            =   -74760
         TabIndex        =   78
         Top             =   480
         Width           =   8295
         Begin VB.CheckBox Notatie_amino 
            Caption         =   "Short notation. Ex:[M,P,Q]"
            Height          =   255
            Left            =   360
            TabIndex        =   127
            Top             =   4200
            Width           =   3735
         End
         Begin VB.Frame mod_codoni_formal 
            Caption         =   "Edit codons"
            Height          =   1095
            Left            =   240
            TabIndex        =   119
            Top             =   480
            Visible         =   0   'False
            Width           =   3375
            Begin VB.TextBox inputul_codonilor 
               Height          =   285
               Left            =   120
               TabIndex        =   121
               Top             =   330
               Width           =   3135
            End
            Begin AG.isButton Mod_codoni 
               Height          =   300
               Left            =   480
               TabIndex        =   120
               Top             =   675
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               Icon            =   "Mai_mult.frx":008C
               Style           =   7
               Caption         =   "Change"
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
            Begin AG.isButton iesi_codoni 
               Height          =   300
               Left            =   1800
               TabIndex        =   122
               Top             =   680
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               Icon            =   "Mai_mult.frx":00A8
               Style           =   7
               Caption         =   "Exit"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   79
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":00C4
            Style           =   7
            Caption         =   "A\Ala"
            CaptionAlign    =   3
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   80
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":00E0
            Style           =   7
            Caption         =   "C\Cys"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   2
            Left            =   240
            TabIndex        =   81
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":00FC
            Style           =   7
            Caption         =   "D\Asp"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   3
            Left            =   240
            TabIndex        =   82
            Top             =   1560
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":0118
            Style           =   7
            Caption         =   "E\Glu"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   4
            Left            =   240
            TabIndex        =   83
            Top             =   1920
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":0134
            Style           =   7
            Caption         =   "F\Phe"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   5
            Left            =   240
            TabIndex        =   84
            Top             =   2280
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":0150
            Style           =   7
            Caption         =   "G\Gly"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   6
            Left            =   240
            TabIndex        =   85
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":016C
            Style           =   7
            Caption         =   "H\His"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   7
            Left            =   240
            TabIndex        =   86
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":0188
            Style           =   7
            Caption         =   "I\Ile"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   8
            Left            =   240
            TabIndex        =   87
            Top             =   3360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":01A4
            Style           =   7
            Caption         =   "K\Lys"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   9
            Left            =   240
            TabIndex        =   88
            Top             =   3720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":01C0
            Style           =   7
            Caption         =   "L\Leu"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   10
            Left            =   4200
            TabIndex        =   89
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":01DC
            Style           =   7
            Caption         =   "M\Met"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   11
            Left            =   4200
            TabIndex        =   90
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":01F8
            Style           =   7
            Caption         =   "N\Asn"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   12
            Left            =   4200
            TabIndex        =   91
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":0214
            Style           =   7
            Caption         =   "P\Pro"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   13
            Left            =   4200
            TabIndex        =   92
            Top             =   1560
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":0230
            Style           =   7
            Caption         =   "Q\Gln"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   14
            Left            =   4200
            TabIndex        =   93
            Top             =   1920
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":024C
            Style           =   7
            Caption         =   "R\Arg"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   15
            Left            =   4200
            TabIndex        =   94
            Top             =   2280
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":0268
            Style           =   7
            Caption         =   "S\Ser"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   16
            Left            =   4200
            TabIndex        =   95
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":0284
            Style           =   7
            Caption         =   "T\Thr"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   17
            Left            =   4200
            TabIndex        =   96
            Top             =   3000
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":02A0
            Style           =   7
            Caption         =   "V\Val"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   18
            Left            =   4200
            TabIndex        =   97
            Top             =   3360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":02BC
            Style           =   7
            Caption         =   "W\Trp"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   19
            Left            =   4200
            TabIndex        =   98
            Top             =   3720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":02D8
            Style           =   7
            Caption         =   "Y\Tyr"
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
         Begin AG.isButton amino_buton 
            Height          =   300
            Index           =   20
            Left            =   4200
            TabIndex        =   123
            Top             =   4080
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            Icon            =   "Mai_mult.frx":02F4
            Style           =   7
            Caption         =   "Z\[STOP]"
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
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "UAA,UAG,UGA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   5400
            TabIndex        =   124
            Top             =   4080
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "UAC,UAU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   5400
            TabIndex        =   118
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "UGG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   5400
            TabIndex        =   117
            Top             =   3360
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "GUA,GUC,GUG,GUU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   5400
            TabIndex        =   116
            Top             =   3000
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "ACA,ACC,ACG,ACU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   5400
            TabIndex        =   115
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "UCA,UCC,UCG,UCU,AGC,AGU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   5400
            TabIndex        =   114
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "CGA,CGC,CGG,CGU,AGA,AGG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   5400
            TabIndex        =   113
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "CAA,CAG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   5400
            TabIndex        =   112
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "CCA,CCC,CCG,CCU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   5400
            TabIndex        =   111
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "AAC,AAU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   5400
            TabIndex        =   110
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "AUG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   5400
            TabIndex        =   109
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "UUA,UUG,CUA,CUC,CUG,CUU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1440
            TabIndex        =   108
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "AAA,AAG "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1440
            TabIndex        =   107
            Top             =   3360
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "AUA,AUC,AUU "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1440
            TabIndex        =   106
            Top             =   3000
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "CAC,CAU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1440
            TabIndex        =   105
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "GGA,GGC,GGG,GGU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   1440
            TabIndex        =   104
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "UUC,UUU "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   103
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "GAA,GAG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   102
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "GAC,GAU "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   101
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "UGC,UGU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   100
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label codoni_aminoacidului 
            BackStyle       =   0  'Transparent
            Caption         =   "GCA,GCC,GCG,GCU"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   99
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.Frame Ferestre_plat_forma 
         Caption         =   "Select windows"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   70
         Top             =   600
         Width           =   2655
         Begin VB.OptionButton f_text 
            Caption         =   "OutPut  Window"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   72
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton f_text 
            Caption         =   "Nucleotide window"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   71
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Explanations"
         Height          =   855
         Left            =   -74880
         TabIndex        =   68
         Top             =   2880
         Width           =   2655
         Begin VB.Label total_enz 
            Alignment       =   2  'Center
            Caption         =   "Total restriction enzymes: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Input ENZYMES"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   61
         Top             =   3840
         Width           =   2655
         Begin VB.TextBox nume_e 
            Height          =   285
            Left            =   1440
            TabIndex        =   63
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox cod_e 
            Height          =   285
            Left            =   1440
            TabIndex        =   62
            Top             =   720
            Width           =   1095
         End
         Begin AG.isButton scrie_enz_db 
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Icon            =   "Mai_mult.frx":0310
            Style           =   7
            Caption         =   "Add"
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
         Begin AG.isButton sterge_enz_db 
            Height          =   375
            Left            =   1440
            TabIndex        =   65
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Icon            =   "Mai_mult.frx":032C
            Style           =   7
            Caption         =   "Del"
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
         Begin VB.Label exp 
            BackStyle       =   0  'Transparent
            Caption         =   "Enzyme name:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label exp 
            BackStyle       =   0  'Transparent
            Caption         =   "Recognition site:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Window Colors"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   56
         Top             =   480
         Width           =   5055
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   21
            Left            =   3120
            Top             =   840
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   3120
            Top             =   480
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   240
            Top             =   840
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   240
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Graphic display"
            Height          =   255
            Index           =   12
            Left            =   3240
            TabIndex        =   159
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label R_cul_grafic 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "Vertical line color."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   158
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "GEL simulation"
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   58
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label GEL_CUL 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "= _=_== ="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   840
            Width           =   2895
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Algorithm Colors"
         Height          =   3735
         Left            =   -74760
         TabIndex        =   25
         Top             =   1920
         Width           =   5055
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   3120
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   18
            Left            =   3120
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   17
            Left            =   3120
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   3120
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   15
            Left            =   3120
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   3120
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   3120
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   3120
            Top             =   720
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   3120
            Top             =   360
            Width           =   1815
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   240
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   240
            Top             =   2880
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   240
            Top             =   2520
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   240
            Top             =   2160
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   240
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   240
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   240
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   240
            Top             =   720
            Width           =   2895
         End
         Begin VB.Shape bordura 
            BorderColor     =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   240
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " deletion 5' - 3'"
            Height          =   255
            Index           =   5
            Left            =   3360
            TabIndex        =   55
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " insertion 5' - 3'"
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   54
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " normal 3' - 5'"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   53
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label start3normal 
            Alignment       =   2  'Center
            BackColor       =   &H00404080&
            Caption         =   "Xn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Label stop_ins5 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "Yn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1680
            TabIndex        =   51
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label start_ins5 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "Xn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " normal 5' - 3'"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   49
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " insertion 3' - 5'"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   48
            Top             =   2520
            Width           =   1575
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " deletion 5' - 3'"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   47
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label start5normal 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "Xn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   2880
            Width           =   2895
         End
         Begin VB.Label stop_ins3 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "Xn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label start_ins3 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "Yn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label start3 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Yn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label mijloc3 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "1...n"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   1080
            TabIndex        =   42
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label stop3 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "Xn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   2280
            TabIndex        =   41
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label stop5 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "Yn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   255
            Left            =   2280
            TabIndex        =   40
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label mijloc5 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "1..n"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   1080
            TabIndex        =   39
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label start5 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "Xn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Inverse sequences"
            Height          =   255
            Index           =   8
            Left            =   3360
            TabIndex        =   37
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Repeats"
            Height          =   255
            Index           =   9
            Left            =   3360
            TabIndex        =   36
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label exp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Genes"
            Height          =   255
            Index           =   10
            Left            =   3360
            TabIndex        =   35
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label SI1_culoare 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label SI2_culoare 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   255
            Left            =   2280
            TabIndex        =   33
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label R1_cul 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "n"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label R2_cul 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "n+1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   255
            Left            =   1680
            TabIndex        =   31
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Start_CUL 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Intron_cul 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "C n+1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   960
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Stop_cul 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "Stop"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2400
            TabIndex        =   28
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Exon_cul 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "C n+2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   360
            Width           =   735
         End
         Begin VB.Label SI_mid 
            Alignment       =   2  'Center
            BackColor       =   &H0080FFFF&
            Caption         =   "0 ... n b"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   1080
            TabIndex        =   26
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.PictureBox bg_txt1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71760
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   24
         Top             =   4320
         Width           =   375
      End
      Begin VB.PictureBox Tabel_culori_font 
         AutoSize        =   -1  'True
         Height          =   1755
         Left            =   -71880
         MousePointer    =   2  'Cross
         Picture         =   "Mai_mult.frx":0348
         ScaleHeight     =   1695
         ScaleWidth      =   3000
         TabIndex        =   23
         Top             =   2520
         Width           =   3060
      End
      Begin VB.PictureBox bg_txt2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71640
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   22
         Top             =   4440
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -71160
         ScaleHeight     =   345
         ScaleWidth      =   2265
         TabIndex        =   21
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   "Explanations"
         Height          =   3735
         Left            =   -74760
         TabIndex        =   20
         Top             =   1920
         Width           =   2655
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Here you can change the font type, size and color, as well as the background color for both windows, "
            Height          =   975
            Left            =   240
            TabIndex        =   75
            Top             =   1320
            Width           =   2175
         End
      End
      Begin VB.ListBox font_tip 
         Height          =   1815
         Left            =   -71880
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox font_marime_real 
         Height          =   285
         Left            =   -68640
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "real_font_size"
         Top             =   600
         Width           =   2175
      End
      Begin VB.ListBox font_marime 
         Height          =   1425
         Left            =   -68640
         TabIndex        =   16
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   -68640
         TabIndex        =   11
         Top             =   2520
         Width           =   2175
         Begin VB.CheckBox Check4 
            Caption         =   "Strikethru"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Underline"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Italic"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Bold"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame exp_enzimeq 
         Caption         =   "Explanations"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   2655
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "To delete a specific restriction enzyme from the list, click with the mouse on the row of interest and press the 'Del' button"
            Height          =   975
            Left            =   120
            TabIndex        =   77
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label exp_enzime 
            BackStyle       =   0  'Transparent
            Caption         =   "In this section you can test different restriction enzymes by DbClick on the list on the right."
            Height          =   855
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Expl_culori 
         Caption         =   "Explanations"
         Height          =   2655
         Left            =   -69600
         TabIndex        =   9
         Top             =   3000
         Width           =   3015
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"Mai_mult.frx":5628
            Height          =   1215
            Left            =   240
            TabIndex        =   76
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To change colors, click one of the segments on the left, then click the color box. "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   74
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -68880
         ScaleHeight     =   345
         ScaleWidth      =   2265
         TabIndex        =   8
         Top             =   2400
         Width           =   2295
      End
      Begin VB.PictureBox Culoarea2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -69360
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   7
         Top             =   2520
         Width           =   375
      End
      Begin VB.PictureBox Tabel_culori 
         AutoSize        =   -1  'True
         Height          =   1755
         Left            =   -69600
         MousePointer    =   2  'Cross
         Picture         =   "Mai_mult.frx":56E5
         ScaleHeight     =   1695
         ScaleWidth      =   3000
         TabIndex        =   6
         Top             =   600
         Width           =   3060
      End
      Begin VB.PictureBox Culoarea1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -69480
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   5
         Top             =   2400
         Width           =   375
      End
      Begin VB.ListBox enzime_de_test 
         Height          =   5130
         Left            =   -69120
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid db_enzime 
         Height          =   5175
         Left            =   -72120
         TabIndex        =   3
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   9128
         _Version        =   393216
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
      End
      Begin VB.Label Mostra_text 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Test text"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -71880
         TabIndex        =   19
         Top             =   4800
         Width           =   5415
      End
   End
   Begin AG.jcFrames Mai_multe_op 
      Height          =   7095
      Left            =   0
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   12515
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   "Options"
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
      Begin AG.isButton Iesire_o 
         Height          =   300
         Left            =   8400
         TabIndex        =   0
         Top             =   50
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         Icon            =   "Mai_mult.frx":A9C5
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
      Begin AG.isButton ok 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   6480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Icon            =   "Mai_mult.frx":A9E1
         Style           =   7
         Caption         =   "OK"
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
      Begin AG.isButton Aplica 
         Height          =   375
         Left            =   3480
         TabIndex        =   59
         Top             =   6480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Icon            =   "Mai_mult.frx":A9FD
         Style           =   7
         Caption         =   "Apply"
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
      Begin AG.isButton iesire 
         Height          =   375
         Left            =   5640
         TabIndex        =   60
         Top             =   6480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Icon            =   "Mai_mult.frx":AA19
         Style           =   7
         Caption         =   "Exit"
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
End
Attribute VB_Name = "Mai_mult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private rRed As Long, rBlue As Long, rGreen As Long

Private Type Pixel
    Pix As Long
    Top As Long
    Left As Long
    Right As Long
    Bottom As Long
End Type

Private Type RGBVals
    Red As Long
    Blue As Long
    Green As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim oldfindstart As Integer

Private hopa_index As Variant

Public Obj_de_colorat As Object

Dim ret As Integer
Public click_enzime_general As String




Private Sub amino_buton_Click(Index As Integer)
mod_codoni_formal.Top = amino_buton(Index).Top + amino_buton(Index).Height
mod_codoni_formal.Left = amino_buton(Index).Left
If mod_codoni_formal.Top + mod_codoni_formal.Height >= CODAMINIFORMA.Height Then
mod_codoni_formal.Top = amino_buton(Index).Top - mod_codoni_formal.Height
End If
mod_codoni_formal.Caption = "Edit codons for " & amino_buton(Index).Caption
inputul_codonilor.Text = codoni_aminoacidului(Index).Caption
mod_codoni_formal.Visible = True
'--------------------------------------------------------------------------------------]
hopa_index = Index
End Sub

Public Sub ReplaceColor(FindColor As Long, ReplaceColor As Long, PicBox As PictureBox)
'Replace a color in one picturebox with another
Dim newVal As Integer, h As Long, w As Long, K As Integer
Dim c As Long

PicBox.ScaleMode = 3

For h = 0 To PicBox.ScaleHeight
    For w = 0 To PicBox.ScaleWidth
        c = GetPixel(PicBox.hdc, w, h)
        If c = FindColor Then c = ReplaceColor
        SetPixel PicBox.hdc, w, h, c
    Next w
Next h
End Sub

Private Sub db_enzime_Click()
r_tmp_r = db_enzime.Row
click_enzime_general = db_enzime.TextMatrix(r_tmp_r, 0) & "[#]" & db_enzime.TextMatrix(r_tmp_r, 1)
End Sub

Private Sub db_enzime_DblClick()
enzime_de_test.AddItem db_enzime.TextMatrix(db_enzime.MouseRow, 1)
End Sub

Private Sub enzime_de_test_DblClick()
enzime_de_test.RemoveItem enzime_de_test.ListIndex
End Sub

Private Sub Exon_cul_Click()
Culoarea1.BackColor = Exon_cul.BackColor
Culoarea2.BackColor = Exon_cul.ForeColor
Set Obj_de_colorat = Exon_cul
End Sub

Private Sub Form_Load()

 Dim wq As Integer
    wq = FreeFile

db_enzime.TextMatrix(0, 0) = "nume"
db_enzime.TextMatrix(0, 1) = "cod"

db_enzime.ColWidth(0) = 1350
db_enzime.ColWidth(1) = 1350
i = 0
Open App.Path & "\AG_DB.db" For Input As #wq
While Not EOF(wq)
Line Input #wq, Temp_e

i = i + 1

nume_en = Split(Temp_e, "[#]")(0)
cod_en = Split(Temp_e, "[#]")(1)


   db_enzime.AddItem nume_en
   db_enzime.TextMatrix(i, 0) = nume_en
   db_enzime.TextMatrix(i, 1) = cod_en


Wend
Close #wq
db_enzime.RemoveItem (db_enzime.Rows)
total_enz.Caption = "Total restriction enzymes: " & i
i = 0
'888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
ret = frm
Debug.Print ret
For X = 1 To Screen.FontCount
font_tip.AddItem Screen.Fonts(X)
Next
For X = 5 To 72: font_marime.AddItem Str$(X): Next

For X = 0 To font_tip.ListCount - 1
 If Principal.COD_CODON.Font.Name = font_tip.List(X) Then
  font_tip.ListIndex = X
  Mostra_text.FontName = font_tip.List(X)
  Exit For
 End If
Next

For X = 0 To font_marime.ListCount - 1
 If Int(Val(Principal.COD_CODON.Font.size)) = Val(font_marime.List(X)) Then
  font_marime.ListIndex = X
  Mostra_text.FontSize = Val(font_marime.List(X))
  font_marime_real.Text = font_marime.List(X)
  Exit For
 End If
Next

If Principal.COD_CODON.Font.Bold = True Then
 Mostra_text.FontBold = True
 Check1.Value = 1
End If
If Principal.COD_CODON.Font.Italic = True Then
 Mostra_text.FontItalic = True
 Check2.Value = 1
End If
If Principal.COD_CODON.Font.Underline = True Then
 Mostra_text.FontUnderline = True
 Check3.Value = 1
End If
If Principal.COD_CODON.Font.Strikethrough = True Then
 Mostra_text.FontStrikethru = True
 Check4.Value = 1
End If

Mostra_text.BackColor = Principal.COD_CODON.BackColor
Set Obj_de_colorat = start3normal
End Sub

Private Sub Formula_TM_primer_ADN_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
If Formula_TM_primer_ADN(0).Value = 1 Then
Call SaveSetting(App.ExeName, "Formule", "c", "1")
alege_calcul_TM(0).Enabled = True
alege_calcul_TM(1).Enabled = False
Formula_TM_primer_ADN(1).Value = 0
Else
Call SaveSetting(App.ExeName, "Formule", "c", "0")
alege_calcul_TM(0).Enabled = False
alege_calcul_TM(1).Enabled = True
Formula_TM_primer_ADN(1).Value = 1
End If
End If

If Index = 1 Then
If Formula_TM_primer_ADN(1).Value = 1 Then
Call SaveSetting(App.ExeName, "Formule", "c", "0")
alege_calcul_TM(0).Enabled = False
alege_calcul_TM(1).Enabled = True
Formula_TM_primer_ADN(0).Value = 0
Else
Call SaveSetting(App.ExeName, "Formule", "c", "1")
alege_calcul_TM(0).Enabled = True
alege_calcul_TM(1).Enabled = False
Formula_TM_primer_ADN(0).Value = 1
End If
End If
End Sub

Private Sub GEL_CUL_Click()
Culoarea1.BackColor = GEL_CUL.BackColor
Culoarea2.BackColor = GEL_CUL.ForeColor
Set Obj_de_colorat = GEL_CUL
End Sub

Private Sub iesi_codoni_Click()
mod_codoni_formal.Visible = False
End Sub

Private Sub Iesire_Click()
deasupara_reprezentarii_grafice = True
Me.Hide
End Sub

Private Sub Iesire_o_Click()
deasupara_reprezentarii_grafice = True
Me.Hide
End Sub

Private Sub Intron_cul_Click()
Culoarea1.BackColor = Intron_cul.BackColor
Culoarea2.BackColor = Intron_cul.ForeColor
Set Obj_de_colorat = Intron_cul
End Sub

Private Sub Mai_multe_op_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub







Private Sub Mod_codoni_Click()

If Mid(inputul_codonilor.Text, Len(inputul_codonilor.Text), 1) = "," Or Mid(inputul_codonilor.Text, 1, 1) = "," Then
MsgBox "Do not use the ',' character at the beginning or end of the codon string !"
Exit Sub
End If
    Dim A As Variant
    Dim i As Integer
    i = 1
    A = Parse(inputul_codonilor.Text, ",")

    For i = 1 To UBound(A)
A(i) = Replace(A(i), " ", "")
'-------------------------------------------------------------------------------]
Dim zat As String
zat = A(i)
If Len(zat) > 3 Then
MsgBox "Codons cannot exceed 3 nucleotides !"
Exit Sub
End If

If Len(zat) < 3 Then
MsgBox "Codons cannot be smaller than three nucleotides !"
Exit Sub
End If
'-------------------------------------------------------------------------------]
For w = 1 To Len(A(i))
genue = LCase(Mid(A(i), w, 1))

If genue = "a" Or genue = "c" Or genue = "g" Or genue = "u" Then
Else
If genue = "t" Then
MsgBox "Thymine is a DNA component, replace 't' with 'u' !"
Else
MsgBox "Only the characters: a/c/g/u/ are accepted"
End If
Exit Sub
End If
Next w
'-------------------------------------------------------------------------------]
    Next i

codoni_aminoacidului(hopa_index).Caption = UCase(inputul_codonilor.Text)
Call SaveSetting(App.ExeName, "Codoni", "codon" & hopa_index, Mai_mult.codoni_aminoacidului(hopa_index).Caption)
mod_codoni_formal.Visible = False
End Sub

Private Sub Notatie_amino_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Notatie_amino.Value = 1 Then
Notatie_amino.Caption = "Notatie scurta. Ex:[M,P,Q]"
amino_notatie_scurta = True
Call SaveSetting(App.ExeName, "Codoni", "Notatie", "1")
Else
Notatie_amino.Caption = "Long notation. Ex:[Met,Pro,Gln]"
amino_notatie_scurta = False
Call SaveSetting(App.ExeName, "Codoni", "Notatie", "0")
End If
End Sub

Private Sub ok_Click()

If f_text(0).Value = True Then

Principal.COD_CODON.Font.Name = Mostra_text.FontName
Principal.COD_CODON.Font.size = Mostra_text.FontSize
Principal.COD_CODON.Font.Bold = Mostra_text.FontBold
Principal.COD_CODON.Font.Italic = Mostra_text.FontItalic
Principal.COD_CODON.Font.Underline = Mostra_text.FontUnderline
Principal.COD_CODON.Font.Strikethrough = Mostra_text.FontStrikethru
'--------------------------------------------------------------------------------------
Principal.COD_CODON.SelStart = 0
Principal.COD_CODON.SelLength = Len(Principal.COD_CODON.Text)
Principal.COD_CODON.SelColor = Mai_mult.Mostra_text.ForeColor 'RGB(255, 255, 255)
Principal.COD_CODON.SelStart = 0
'--------------------------------------------------------------------------------------
Principal.COD_CODON.BackColor = Mostra_text.BackColor



Call SaveSetting(App.ExeName, "Setari", "p1", Mostra_text.FontName)
Call SaveSetting(App.ExeName, "Setari", "p2", Mostra_text.FontSize)
Call SaveSetting(App.ExeName, "Setari", "p3", Mostra_text.FontBold)
Call SaveSetting(App.ExeName, "Setari", "p4", Mostra_text.FontItalic)
Call SaveSetting(App.ExeName, "Setari", "p5", Mostra_text.FontUnderline)
Call SaveSetting(App.ExeName, "Setari", "p6", Mostra_text.FontStrikethru)
Call SaveSetting(App.ExeName, "Setari", "p7", Mai_mult.Mostra_text.ForeColor)
Call SaveSetting(App.ExeName, "Setari", "p8", Mostra_text.BackColor)

Else
Principal.out_put.FontName = Mostra_text.FontName
Principal.out_put.FontSize = Mostra_text.FontSize
Principal.out_put.FontBold = Mostra_text.FontBold
Principal.out_put.FontItalic = Mostra_text.FontItalic
Principal.out_put.FontUnderline = Mostra_text.FontUnderline
Principal.out_put.FontStrikethru = Mostra_text.FontStrikethru
Principal.out_put.ForeColor = Mai_mult.Mostra_text.ForeColor 'RGB(255, 255, 255)
Principal.out_put.BackColor = Mostra_text.BackColor

Call SaveSetting(App.ExeName, "Setari", "o1", Mostra_text.FontName)
Call SaveSetting(App.ExeName, "Setari", "o2", Mostra_text.FontSize)
Call SaveSetting(App.ExeName, "Setari", "o3", Mostra_text.FontBold)
Call SaveSetting(App.ExeName, "Setari", "o4", Mostra_text.FontItalic)
Call SaveSetting(App.ExeName, "Setari", "o5", Mostra_text.FontUnderline)
Call SaveSetting(App.ExeName, "Setari", "o6", Mostra_text.FontStrikethru)
Call SaveSetting(App.ExeName, "Setari", "o7", Mai_mult.Mostra_text.ForeColor)
Call SaveSetting(App.ExeName, "Setari", "o8", Mostra_text.BackColor)

End If


Call SaveSetting(App.ExeName, "Formule", "f1", nr_GC_calcul.Text)
Call SaveSetting(App.ExeName, "Formule", "f2", nr_AT_calcul.Text)
Call SaveSetting(App.ExeName, "Formule", "l1", log_concentratie_Na.Text)
Call SaveSetting(App.ExeName, "Formule", "n1", Nr_1.Text)
Call SaveSetting(App.ExeName, "Formule", "n2", Nr_2.Text)
Call SaveSetting(App.ExeName, "Formule", "n3", Nr_3.Text)
Call SaveSetting(App.ExeName, "Formule", "n4", Nr_4.Text)

If GetPixel(Principal.segment_total_gena.hdc, 1, 1) <> R_cul_grafic.BackColor Or Principal.Segment_grafic.BorderColor <> R_cul_grafic.ForeColor Then
Principal.Segment_grafic.BorderColor = R_cul_grafic.ForeColor
'principal.anunt_schimbare_fundal.ForeColor = R_cul_grafic.ForeColor
'principal.anunt_schimbare_fundal.Left = (principal.segment_total_gena.ScaleWidth / 2) - (principal.anunt_schimbare_fundal.Width / 2)
'principal.anunt_schimbare_fundal.Visible = True
Call ReplaceColor(GetPixel(Principal.segment_total_gena.hdc, 1, 1), R_cul_grafic.BackColor, Principal.segment_total_gena)
Principal.segment_total_gena.Refresh
'principal.anunt_schimbare_fundal.Visible = False
End If

If culoare1_gel_old <> GEL_CUL.ForeColor Or culoare2_gel_old <> GEL_CUL.BackColor Then

Call ReplaceColor(culoare1_gel_old, GEL_CUL.ForeColor, Principal.Gel_general)
Call ReplaceColor(culoare2_gel_old, GEL_CUL.BackColor, Principal.Gel_general)
Principal.Gel_general.Refresh
Call ReplaceColor(culoare1_gel_old, GEL_CUL.ForeColor, Principal.Gelul)
Call ReplaceColor(culoare2_gel_old, GEL_CUL.BackColor, Principal.Gelul)
Principal.Gelul.Refresh

culoare1_gel_old = GEL_CUL.ForeColor
culoare2_gel_old = GEL_CUL.BackColor
End If


deasupara_reprezentarii_grafice = True
Me.Hide
End Sub

Private Sub Aplica_Click()
If f_text(0).Value = True Then

Principal.COD_CODON.Font.Name = Mostra_text.FontName
Principal.COD_CODON.Font.size = Mostra_text.FontSize
Principal.COD_CODON.Font.Bold = Mostra_text.FontBold
Principal.COD_CODON.Font.Italic = Mostra_text.FontItalic
Principal.COD_CODON.Font.Underline = Mostra_text.FontUnderline
Principal.COD_CODON.Font.Strikethrough = Mostra_text.FontStrikethru
'--------------------------------------------------------------------------------------
Principal.COD_CODON.SelStart = 0
Principal.COD_CODON.SelLength = Len(Principal.COD_CODON.Text)
Principal.COD_CODON.SelColor = Mai_mult.Mostra_text.ForeColor 'RGB(255, 255, 255)
Principal.COD_CODON.SelStart = 0
'--------------------------------------------------------------------------------------
Principal.COD_CODON.BackColor = Mostra_text.BackColor


Call SaveSetting(App.ExeName, "Setari", "p1", Mostra_text.FontName)
Call SaveSetting(App.ExeName, "Setari", "p2", Mostra_text.FontSize)
Call SaveSetting(App.ExeName, "Setari", "p3", Mostra_text.FontBold)
Call SaveSetting(App.ExeName, "Setari", "p4", Mostra_text.FontItalic)
Call SaveSetting(App.ExeName, "Setari", "p5", Mostra_text.FontUnderline)
Call SaveSetting(App.ExeName, "Setari", "p6", Mostra_text.FontStrikethru)
Call SaveSetting(App.ExeName, "Setari", "p7", Mai_mult.Mostra_text.ForeColor)
Call SaveSetting(App.ExeName, "Setari", "p8", Mostra_text.BackColor)

Else
Principal.out_put.FontName = Mostra_text.FontName
Principal.out_put.FontSize = Mostra_text.FontSize
Principal.out_put.FontBold = Mostra_text.FontBold
Principal.out_put.FontItalic = Mostra_text.FontItalic
Principal.out_put.FontUnderline = Mostra_text.FontUnderline
Principal.out_put.FontStrikethru = Mostra_text.FontStrikethru
Principal.out_put.ForeColor = Mai_mult.Mostra_text.ForeColor 'RGB(255, 255, 255)
Principal.out_put.BackColor = Mostra_text.BackColor

Call SaveSetting(App.ExeName, "Setari", "o1", Mostra_text.FontName)
Call SaveSetting(App.ExeName, "Setari", "o2", Mostra_text.FontSize)
Call SaveSetting(App.ExeName, "Setari", "o3", Mostra_text.FontBold)
Call SaveSetting(App.ExeName, "Setari", "o4", Mostra_text.FontItalic)
Call SaveSetting(App.ExeName, "Setari", "o5", Mostra_text.FontUnderline)
Call SaveSetting(App.ExeName, "Setari", "o6", Mostra_text.FontStrikethru)
Call SaveSetting(App.ExeName, "Setari", "o7", Mai_mult.Mostra_text.ForeColor)
Call SaveSetting(App.ExeName, "Setari", "o8", Mostra_text.BackColor)

End If

Call SaveSetting(App.ExeName, "Formule", "f1", nr_GC_calcul.Text)
Call SaveSetting(App.ExeName, "Formule", "f2", nr_AT_calcul.Text)
Call SaveSetting(App.ExeName, "Formule", "l1", log_concentratie_Na.Text)
Call SaveSetting(App.ExeName, "Formule", "n1", Nr_1.Text)
Call SaveSetting(App.ExeName, "Formule", "n2", Nr_2.Text)
Call SaveSetting(App.ExeName, "Formule", "n3", Nr_3.Text)
Call SaveSetting(App.ExeName, "Formule", "n4", Nr_4.Text)

If GetPixel(Principal.segment_total_gena.hdc, 1, 1) <> R_cul_grafic.BackColor Or Principal.Segment_grafic.BorderColor <> R_cul_grafic.ForeColor Then
Principal.Segment_grafic.BorderColor = R_cul_grafic.ForeColor
'principal.anunt_schimbare_fundal.ForeColor = R_cul_grafic.ForeColor
'principal.anunt_schimbare_fundal.Left = (principal.segment_total_gena.ScaleWidth / 2) - (principal.anunt_schimbare_fundal.Width / 2)
'principal.anunt_schimbare_fundal.Visible = True
Call ReplaceColor(GetPixel(Principal.segment_total_gena.hdc, 1, 1), R_cul_grafic.BackColor, Principal.segment_total_gena)
Principal.segment_total_gena.Refresh
'principal.anunt_schimbare_fundal.Visible = False
End If

End Sub

Private Sub R_cul_grafic_Click()
Culoarea1.BackColor = R_cul_grafic.BackColor
Culoarea2.BackColor = R_cul_grafic.ForeColor
Set Obj_de_colorat = R_cul_grafic
End Sub

Private Sub R1_cul_Click()
Culoarea1.BackColor = R1_cul.BackColor
Culoarea2.BackColor = R1_cul.ForeColor
Set Obj_de_colorat = R1_cul
End Sub

Private Sub R2_cul_Click()
Culoarea1.BackColor = R2_cul.BackColor
Culoarea2.BackColor = R2_cul.ForeColor
Set Obj_de_colorat = R2_cul
End Sub

Private Sub scrie_enz_db_Click()

nume_e.Text = Replace(nume_e.Text, "[#]", "")
cod_e.Text = Replace(cod_e.Text, "[#]", "")

If nume_e.Text = "" Or cod_e.Text = "" Then Exit Sub
'----------------------------------------------------
Dim opi As Variant
opi = 0
For i = 1 To Len(cod_e.Text)
genue = LCase(Mid(cod_e.Text, i, 1))
If genue = "*" Then opi = opi + 1
If opi >= 2 Then
MsgBox "The '*' character represents the cleavage point of the restriction enzyme" & vbCrLf & " si poate fi scris o singura data !"
Exit Sub
End If
Next i

For i = 1 To Len(cod_e.Text)
genue = LCase(Mid(cod_e.Text, i, 1))
If genue = "a" Or genue = "c" Or genue = "g" Or genue = "t" Or genue = "u" Or genue = "*" Then
Else
MsgBox "Only the characters: a/c/g/t/u/*/ are accepted"
Exit Sub
End If
Next i
'----------------------------------------------------
 Dim wy As Integer
    wy = FreeFile
i = 0
Open App.Path & "\AG_DB.db" For Input As #wy
While Not EOF(wy)
Line Input #wy, Temp_e

i = i + 1

nume_en = Split(Temp_e, "[#]")(0)
cod_en = Split(Temp_e, "[#]")(1)

If nume_e.Text = nume_en Then
MsgBox "There cannot be two enzymes with the same name in the database !"
Close #wy
Exit Sub
End If

If cod_e.Text = cod_en Then
MsgBox "There cannot be two identical sequences in the database !"
Close #wy
Exit Sub
End If

Wend
Close #wy
'----------------------------------------------------
i = 1
 Dim sq As Integer
 Dim F As Integer
    F = FreeFile
Open App.Path & "\AG_BD.tmp" For Output As #F
sq = FreeFile

Print #F, nume_e.Text & "[#]" & cod_e.Text

Open App.Path & "\AG_DB.db" For Input As #sq
While Not EOF(sq)
Line Input #sq, Temp_e

Print #F, Temp_e

Wend
Close #sq
Close #F

Kill (App.Path & "\AG_DB.db")
Call FileCopy(App.Path & "\AG_BD.tmp", App.Path & "\AG_DB.db")
'-----------------------------------------------------------------
'For i = 1 To db_enzime.Rows
3:
If db_enzime.Rows <= 2 Then GoTo 4
db_enzime.RemoveItem (db_enzime.Rows)
db_enzime.Refresh
GoTo 3
4:
'Next i
'-----------------------------------------------------------------
 Dim wq As Integer
    wq = FreeFile

db_enzime.TextMatrix(0, 0) = "nume"
db_enzime.TextMatrix(0, 1) = "cod"

db_enzime.ColWidth(0) = 1350
db_enzime.ColWidth(1) = 1350
i = 0
Open App.Path & "\AG_DB.db" For Input As #wq
While Not EOF(wq)
Line Input #wq, Temp_e

i = i + 1

nume_en = Split(Temp_e, "[#]")(0)
cod_en = Split(Temp_e, "[#]")(1)

   db_enzime.AddItem nume_en
   db_enzime.TextMatrix(i, 0) = nume_en
   db_enzime.TextMatrix(i, 1) = cod_en

Wend
Close #wq
db_enzime.RemoveItem (db_enzime.Rows)
total_enz.Caption = "Total restriction enzymes: " & i
i = 0

nume_e.Text = ""
cod_e.Text = ""
End Sub

Private Sub sterge_enz_db_Click()
'On Error Resume Next
'tmpo = MsgBox("Do you want to delete this enzyme ?", vbYesNo, "Applied Genetics - beta")
'If tmpo = 7 Then Exit Sub

i = 1
 Dim sq As Integer
 Dim F As Integer
    F = FreeFile
Open App.Path & "\AG_BD.tmp" For Output As #F
sq = FreeFile
Open App.Path & "\AG_DB.db" For Input As #sq
While Not EOF(sq)
Line Input #sq, Temp_e
comp_enz = Replace(Temp_e, click_enzime_general, "!")
If comp_enz = "!" Then
db_enzime.RemoveItem (i)
Else

Print #F, Temp_e

End If
i = i + 1

Wend
Close #sq
Close #F

total_enz.Caption = "Total restriction enzymes: " & i - 1
Kill (App.Path & "\AG_DB.db")
Call FileCopy(App.Path & "\AG_BD.tmp", App.Path & "\AG_DB.db")

End Sub

Private Sub SI_mid_Click()
Culoarea1.BackColor = SI_mid.BackColor
Culoarea2.BackColor = SI_mid.ForeColor
Set Obj_de_colorat = SI_mid
End Sub

Private Sub SI1_culoare_Click()
Culoarea1.BackColor = SI1_culoare.BackColor
Culoarea2.BackColor = SI1_culoare.ForeColor
Set Obj_de_colorat = SI1_culoare
End Sub

Private Sub SI2_culoare_Click()
Culoarea1.BackColor = SI2_culoare.BackColor
Culoarea2.BackColor = SI2_culoare.ForeColor
Set Obj_de_colorat = SI2_culoare
End Sub

Private Sub Start_CUL_Click()
Culoarea1.BackColor = Start_CUL.BackColor
Culoarea2.BackColor = Start_CUL.ForeColor
Set Obj_de_colorat = Start_CUL
End Sub

Private Sub start3normal_Click()
Culoarea1.BackColor = start3normal.BackColor
Culoarea2.BackColor = start3normal.ForeColor
Set Obj_de_colorat = start3normal
End Sub

Private Sub start5normal_Click()
Culoarea1.BackColor = start5normal.BackColor
Culoarea2.BackColor = start5normal.ForeColor
Set Obj_de_colorat = start5normal
End Sub


Private Sub start_ins3_Click()
Culoarea1.BackColor = start_ins3.BackColor
Culoarea2.BackColor = start_ins3.ForeColor
Set Obj_de_colorat = start_ins3
End Sub

Private Sub Stop_cul_Click()
Culoarea1.BackColor = Stop_cul.BackColor
Culoarea2.BackColor = Stop_cul.ForeColor
Set Obj_de_colorat = Stop_cul
End Sub

Private Sub stop_ins3_Click()
Culoarea1.BackColor = stop_ins3.BackColor
Culoarea2.BackColor = stop_ins3.ForeColor
Set Obj_de_colorat = stop_ins3
End Sub


Private Sub start_ins5_Click()
Culoarea1.BackColor = start_ins5.BackColor
Culoarea2.BackColor = start_ins5.ForeColor
Set Obj_de_colorat = start_ins5
End Sub
Private Sub stop_ins5_Click()
Culoarea1.BackColor = stop_ins5.BackColor
Culoarea2.BackColor = stop_ins5.ForeColor
Set Obj_de_colorat = stop_ins5
End Sub




Private Sub start3_Click()
Culoarea1.BackColor = start3.BackColor
Culoarea2.BackColor = start3.ForeColor
Set Obj_de_colorat = start3
End Sub
Private Sub mijloc3_Click()
Culoarea1.BackColor = mijloc3.BackColor
Culoarea2.BackColor = mijloc3.ForeColor
Set Obj_de_colorat = mijloc3
End Sub
Private Sub stop3_Click()
Culoarea1.BackColor = stop3.BackColor
Culoarea2.BackColor = stop3.ForeColor
Set Obj_de_colorat = stop3
End Sub



Private Sub start5_Click()
Culoarea1.BackColor = start5.BackColor
Culoarea2.BackColor = start5.ForeColor
Set Obj_de_colorat = start5
End Sub
Private Sub mijloc5_Click()
Culoarea1.BackColor = mijloc5.BackColor
Culoarea2.BackColor = mijloc5.ForeColor
Set Obj_de_colorat = mijloc5
End Sub
Private Sub stop5_Click()
Culoarea1.BackColor = stop5.BackColor
Culoarea2.BackColor = stop5.ForeColor
Set Obj_de_colorat = stop5
End Sub

Private Sub Tabel_culori_font_Click()
On Error GoTo bye
If Button = 1 Then
    bg_txt1.BackColor = Tabel_culori_font.Point(X, Y)
    Mostra_text.BackColor = Tabel_culori_font.Point(X, Y)
End If
If Button = 2 Then
    bg_txt2.BackColor = Tabel_culori_font.Point(X, Y)
    Mostra_text.ForeColor = Tabel_culori_font.Point(X, Y)
End If
bye:
End Sub


Private Sub Tabel_culori_font_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo bye
If Button = 1 Then
    bg_txt1.BackColor = Tabel_culori_font.Point(X, Y)
    Mostra_text.BackColor = Tabel_culori_font.Point(X, Y)
End If
If Button = 2 Then
    bg_txt2.BackColor = Tabel_culori_font.Point(X, Y)
    Mostra_text.ForeColor = Tabel_culori_font.Point(X, Y)
End If
bye:
End Sub

Private Sub Tabel_culori_font_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo bye
If Button = 1 Then
    bg_txt1.BackColor = Tabel_culori_font.Point(X, Y)
    Mostra_text.BackColor = Tabel_culori_font.Point(X, Y)
End If
If Button = 2 Then
    bg_txt2.BackColor = Tabel_culori_font.Point(X, Y)
    Mostra_text.ForeColor = Tabel_culori_font.Point(X, Y)
End If
bye:
End Sub


Private Sub Tabel_culori_Click()
On Error GoTo bye

 Call SaveSetting(App.ExeName, "Culori1", "SESCRIE", "1")


If Button = 1 Then
    Culoarea1.BackColor = Tabel_culori.Point(X, Y)
    Obj_de_colorat.BackColor = Tabel_culori.Point(X, Y)
    Call SaveSetting(App.ExeName, "Culori1", Obj_de_colorat.Name, Tabel_culori.Point(X, Y))
    Call SaveSetting(App.ExeName, "Culori2", Obj_de_colorat.Name, Obj_de_colorat.ForeColor)
End If
If Button = 2 Then
    Culoarea2.BackColor = Tabel_culori.Point(X, Y)
    Obj_de_colorat.ForeColor = Tabel_culori.Point(X, Y)
    Call SaveSetting(App.ExeName, "Culori1", Obj_de_colorat.Name, Obj_de_colorat.BackColor)
    Call SaveSetting(App.ExeName, "Culori2", Obj_de_colorat.Name, Tabel_culori.Point(X, Y))

End If
bye:
End Sub

Private Sub Tabel_culori_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo bye

 Call SaveSetting(App.ExeName, "Culori1", "SESCRIE", "1")

If Button = 1 Then
    Culoarea1.BackColor = Tabel_culori.Point(X, Y)
    Obj_de_colorat.BackColor = Tabel_culori.Point(X, Y)
    Call SaveSetting(App.ExeName, "Culori1", Obj_de_colorat.Name, Tabel_culori.Point(X, Y))
    Call SaveSetting(App.ExeName, "Culori2", Obj_de_colorat.Name, Obj_de_colorat.ForeColor)
End If
If Button = 2 Then
    Culoarea2.BackColor = Tabel_culori.Point(X, Y)
    Obj_de_colorat.ForeColor = Tabel_culori.Point(X, Y)
    Call SaveSetting(App.ExeName, "Culori1", Obj_de_colorat.Name, Obj_de_colorat.BackColor)
    Call SaveSetting(App.ExeName, "Culori2", Obj_de_colorat.Name, Tabel_culori.Point(X, Y))

End If
bye:
End Sub

Private Sub Tabel_culori_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo bye



If Button = 1 Then
 Call SaveSetting(App.ExeName, "Culori1", "SESCRIE", "1")
    Culoarea1.BackColor = Tabel_culori.Point(X, Y)
    Obj_de_colorat.BackColor = Tabel_culori.Point(X, Y)
    Call SaveSetting(App.ExeName, "Culori1", Obj_de_colorat.Name, Tabel_culori.Point(X, Y))
    Call SaveSetting(App.ExeName, "Culori2", Obj_de_colorat.Name, Obj_de_colorat.ForeColor)
End If
If Button = 2 Then
 Call SaveSetting(App.ExeName, "Culori1", "SESCRIE", "1")
    Culoarea2.BackColor = Tabel_culori.Point(X, Y)
    Obj_de_colorat.ForeColor = Tabel_culori.Point(X, Y)
    Call SaveSetting(App.ExeName, "Culori1", Obj_de_colorat.Name, Obj_de_colorat.BackColor)
    Call SaveSetting(App.ExeName, "Culori2", Obj_de_colorat.Name, Tabel_culori.Point(X, Y))

End If
bye:
End Sub

'88888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
Private Sub font_tip_Click()
Mostra_text.FontName = font_tip.List(font_tip.ListIndex)
End Sub

Private Sub font_marime_Click()
font_marime_real.Text = font_marime.List(font_marime.ListIndex)
Mostra_text.FontSize = Val(font_marime_real.Text)
End Sub

Private Sub font_marime_real_Change()

For X = 0 To font_marime.ListCount - 1
If Val(font_marime_real.Text) = Val(font_marime.List(X)) Then
  font_marime.ListIndex = X
  Mostra_text.FontSize = Val(font_marime_real.Text)
 Exit For
End If
Next
End Sub


Private Sub Check1_Click()
If Check1.Value = 1 Then
  Mostra_text.FontBold = True
Else
  Mostra_text.FontBold = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
  Mostra_text.FontItalic = True
Else
  Mostra_text.FontItalic = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
  Mostra_text.FontUnderline = True
Else
  Mostra_text.FontUnderline = False
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
  Mostra_text.FontStrikethru = True
Else
  Mostra_text.FontStrikethru = False
End If
End Sub
