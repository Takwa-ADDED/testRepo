VERSION 5.00
Object = "{9E6A409A-83E5-4437-9E06-0D39D3882522}#2.2#0"; "SToolBox.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Statistiques 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Statistiques Carburant"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16320
   LinkTopic       =   "Form1"
   ScaleHeight     =   10140
   ScaleWidth      =   16320
   Begin TabDlg.SSTab Tab_Satistiques 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Statistiques Carburant"
      TabPicture(0)   =   "Frm_Statistique.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Pic_ControlStatCarburant"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Satistiques Reparation"
      TabPicture(1)   =   "Frm_Statistique.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Pic_ControlStatR"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Satistiques Trafic"
      TabPicture(2)   =   "Frm_Statistique.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Pic_ControlStatFT"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Satistiques En/Hors Service"
      TabPicture(3)   =   "Frm_Statistique.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Pic_ControlStatPersonnel"
      Tab(3).ControlCount=   1
      Begin VB.PictureBox Pic_ControlStatR 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   120
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   32
         Top             =   360
         Width           =   14415
         Begin MSComctlLib.ListView List_DetailsRp 
            Height          =   1335
            Left            =   240
            TabIndex        =   46
            Top             =   720
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   2355
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView List_detailRp 
            Height          =   4455
            Left            =   120
            TabIndex        =   44
            Top             =   2280
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   7858
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.ComboBox Cbo_Vehicule 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_Statistique.frx":0070
            Left            =   1560
            List            =   "Frm_Statistique.frx":0072
            TabIndex        =   33
            Top             =   240
            Width           =   4095
         End
         Begin SToolBox.SCommand cmd_FindMatricule 
            Height          =   345
            Left            =   5760
            TabIndex        =   34
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_Statistique.frx":0074
            ButtonType      =   1
         End
         Begin MSComCtl2.DTPicker Dta_Fin 
            Height          =   375
            Left            =   9600
            TabIndex        =   35
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   106954753
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker Dta_Debut 
            Height          =   375
            Left            =   7080
            TabIndex        =   36
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   106954753
            CurrentDate     =   42860
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   9000
            TabIndex        =   40
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   6480
            TabIndex        =   39
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Véhicule:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   1350
         End
         Begin VB.Image Cmd_Find 
            Height          =   495
            Left            =   11280
            Picture         =   "Frm_Statistique.frx":03AE
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.PictureBox Pic_ControlStatFT 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   -74880
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   17
         Top             =   360
         Width           =   14415
         Begin ComctlLib.ListView Lsv_DetailsFT 
            Height          =   1095
            Left            =   0
            TabIndex        =   43
            Top             =   1080
            Width           =   12015
            _ExtentX        =   21193
            _ExtentY        =   1931
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.ComboBox cbo_VehiculeFT 
            Height          =   315
            ItemData        =   "Frm_Statistique.frx":10FB0
            Left            =   1080
            List            =   "Frm_Statistique.frx":10FB2
            TabIndex        =   20
            Top             =   120
            Width           =   2055
         End
         Begin VB.ComboBox cbo_ConducteurFT 
            Height          =   315
            ItemData        =   "Frm_Statistique.frx":10FB4
            Left            =   5160
            List            =   "Frm_Statistique.frx":10FB6
            TabIndex        =   19
            Top             =   120
            Width           =   2055
         End
         Begin VB.ComboBox cbo_DestinationFT 
            Height          =   315
            Left            =   9360
            TabIndex        =   18
            Top             =   120
            Width           =   2055
         End
         Begin SToolBox.SCommand Cmd_FindVehiculeFT 
            Height          =   315
            Left            =   3240
            TabIndex        =   24
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_Statistique.frx":10FB8
            ButtonType      =   1
         End
         Begin SToolBox.SCommand Cmd_FindConducteurFT 
            Height          =   315
            Left            =   7320
            TabIndex        =   25
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_Statistique.frx":112F2
            ButtonType      =   1
         End
         Begin SToolBox.SCommand Cmd_FindDestinationFT 
            Height          =   315
            Left            =   11520
            TabIndex        =   26
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_Statistique.frx":1162C
            ButtonType      =   1
         End
         Begin MSComCtl2.DTPicker cda_FinFT 
            Height          =   375
            Left            =   9000
            TabIndex        =   27
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   106954753
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker cda_Debutft 
            Height          =   375
            Left            =   6480
            TabIndex        =   28
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   106954753
            CurrentDate     =   42860
         End
         Begin SToolBox.SGrid grid_Ft 
            Height          =   5655
            Left            =   0
            TabIndex        =   31
            Top             =   2280
            Width           =   12015
            _ExtentX        =   21193
            _ExtentY        =   9975
            RowMode         =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   2
            DisableIcons    =   -1  'True
            MaxVisibleRows  =   0
         End
         Begin VB.Image Cmd_SearchFT 
            Height          =   495
            Left            =   11280
            Picture         =   "Frm_Statistique.frx":11966
            Stretch         =   -1  'True
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   8400
            TabIndex        =   30
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   5880
            TabIndex        =   29
            Top             =   600
            Width           =   600
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Vehicule"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Conducteur"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3720
            TabIndex        =   22
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Destination"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7800
            TabIndex        =   21
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.PictureBox Pic_ControlStatPersonnel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   -74880
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   8
         Top             =   360
         Width           =   14415
         Begin VB.ComboBox cbo_Conducteur 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_Statistique.frx":22568
            Left            =   1920
            List            =   "Frm_Statistique.frx":2256A
            TabIndex        =   13
            Top             =   240
            Width           =   3735
         End
         Begin SToolBox.SGrid grid_Service 
            Height          =   6735
            Left            =   0
            TabIndex        =   10
            Top             =   960
            Width           =   12015
            _ExtentX        =   21193
            _ExtentY        =   11880
            RowMode         =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   2
            DisableIcons    =   -1  'True
            MaxVisibleRows  =   0
         End
         Begin MSComCtl2.DTPicker cda_FinService 
            Height          =   375
            Left            =   9600
            TabIndex        =   11
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   106954753
            CurrentDate     =   42860
         End
         Begin MSComCtl2.DTPicker cda_DebutService 
            Height          =   375
            Left            =   7080
            TabIndex        =   12
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14737632
            Format          =   106954753
            CurrentDate     =   42860
         End
         Begin SToolBox.SCommand cmdFindConducteur 
            Height          =   345
            Left            =   5760
            TabIndex        =   14
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_Statistique.frx":2256C
            ButtonType      =   1
         End
         Begin VB.Image Cmd_SearchService 
            Height          =   495
            Left            =   11280
            Picture         =   "Frm_Statistique.frx":228A6
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conducteur:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   6480
            TabIndex        =   16
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   9000
            TabIndex        =   15
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.PictureBox Pic_ControlStatCarburant 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   -74880
         ScaleHeight     =   7815
         ScaleWidth      =   14415
         TabIndex        =   2
         Top             =   360
         Width           =   14415
         Begin MSComctlLib.ListView Lsv_Details 
            Height          =   1455
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   2566
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin SToolBox.SDateBox cda_fin 
            Height          =   285
            Left            =   9600
            TabIndex        =   42
            Top             =   240
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SDateBox cda_debut 
            Height          =   285
            Left            =   7080
            TabIndex        =   41
            Top             =   240
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
            Text            =   ""
         End
         Begin SToolBox.SGrid Grid_Carb 
            Height          =   5415
            Left            =   120
            TabIndex        =   37
            Top             =   2280
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   9551
            RowMode         =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            GroupRowForeColor=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Editable        =   -1  'True
            DisableIcons    =   -1  'True
            MaxVisibleRows  =   0
         End
         Begin VB.ComboBox cbo_Matricule 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Frm_Statistique.frx":334A8
            Left            =   1560
            List            =   "Frm_Statistique.frx":334AA
            TabIndex        =   3
            Top             =   240
            Width           =   4095
         End
         Begin SToolBox.SCommand cmdFindMatricule 
            Height          =   345
            Left            =   5760
            TabIndex        =   4
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   609
            BackStyle       =   0
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "Frm_Statistique.frx":334AC
            ButtonType      =   1
         End
         Begin VB.Image Cmd_Search 
            Height          =   495
            Left            =   11400
            Picture         =   "Frm_Statistique.frx":337E6
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Au :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   9000
            TabIndex        =   7
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Du :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   6480
            TabIndex        =   6
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Véhicule:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1350
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_Statistique.frx":443E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image CmdPrint 
      Height          =   495
      Left            =   11880
      Picture         =   "Frm_Statistique.frx":61B42
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label m 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statistiques"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image PicBox_Header 
      Height          =   1005
      Left            =   0
      Picture         =   "Frm_Statistique.frx":72744
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14535
   End
End
Attribute VB_Name = "Frm_Statistiques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim thekey As Integer
    Dim theshift As Integer
    Dim itmX
    Dim VCodeVehicle As String              'Code Vehicule
    Dim VCodeDrive  As String               'Code Conducteur
    Dim VCodeDestination  As String         'Code Destination
    Dim NAnomalieTotal As Integer           'Nombre Anomalie Total***
    Dim NAnomalieKm As Integer              'Nombre Anomalie Km***
    Dim NAnomalieDuree As Integer           'Nombre Anomalie Durée***


'~~~~~~~~~~~~~~~~~~~~
    'Mise en Forme~~~
'~~~~~~~~~~~~~~~~~~~~
Private Sub Form_Load()
    Me.WindowState = 2
    Tab_Satistiques.Tab = 0
    cda_debut.Text = "01/" & Month(Date) & "/" & Year(Date)
    cda_fin.Text = Date
    cda_DebutService.Value = "01/" & Month(Date) & "/" & Year(Date)
    cda_FinService.Value = Date
    cda_Debutft.Value = "01/" & Month(Date) & "/" & Year(Date)
    cda_FinFT.Value = Date
    Dta_Debut.Value = "01/" & Month(Date) & "/" & Year(Date)
    Dta_Fin.Value = Date
    
    
    
    Call Initgrid_Services
    Call Initgrid_FT
    Call Initgrid_Carb
    
    Cbo_Vehicule.AddItem "Tous", 0
    cbo_Matricule.AddItem "Tous", 0
    cbo_VehiculeFT.AddItem "Tous", 0
    cbo_ConducteurFT.AddItem ("Tous"), 0
    cbo_DestinationFT.AddItem ("Tous"), 0
    Call Affiche_Matricule_Combo(Cbo_Vehicule)
    Call Affiche_Matricule_Combo(cbo_Matricule)
    Call Affiche_Matricule_Combo(cbo_VehiculeFT)
    Call Affiche_Personnel_Combo(cbo_ConducteurFT)
    Call Affiche_Personnel_Combo(cbo_Conducteur)
    Call Affiche_Destination_Combo(cbo_DestinationFT)
    Cbo_Vehicule.ListIndex = 0
    cbo_Matricule.ListIndex = 0
    cbo_VehiculeFT.ListIndex = 0
    cbo_ConducteurFT.ListIndex = 0
    cbo_DestinationFT.ListIndex = 0
    VCodeVehicle = "  -  Tous"
    VCodeDrive = "  -  Tous"
    VCodeDestination = "  -  Tous"
    
End Sub
Private Sub Form_Resize()
    Dim WidthForm As Integer, HeightForm As Integer
    WidthForm = Me.Width
    HeightForm = Me.Height
        PicBox_Header.Width = WidthForm
        Tab_Satistiques.Width = WidthForm - 400
        CmdPrint.Left = WidthForm - 2000
        Pic_ControlStatCarburant.Width = Tab_Satistiques.Width - 200
        Pic_ControlStatPersonnel.Width = Tab_Satistiques.Width - 200
        Pic_ControlStatFT.Width = Tab_Satistiques.Width - 200
        Pic_ControlStatR.Width = Tab_Satistiques.Width - 200
        Cmd_SearchService.Left = WidthForm - 3000
        Cmd_SearchService.Top = 200
        Cmd_SearchFT.Left = WidthForm - 3000
        Cmd_SearchFT.Top = 200
        Cmd_Find.Left = WidthForm - 3000
        Cmd_Find.Top = 200
        Cmd_Search.Left = WidthForm - 3000
        Cmd_Search.Top = 200
        grid_Service.Width = Tab_Satistiques.Width - 200
        grid_Ft.Width = Tab_Satistiques.Width - 200
        Lsv_DetailsFT.Width = Tab_Satistiques.Width - 200
        List_DetailsRp.Width = Tab_Satistiques.Width - 200
        List_detailRp.Width = Tab_Satistiques.Width - 200
        Lsv_Details.Width = Tab_Satistiques.Width - 200
        Grid_Carb.Width = Tab_Satistiques.Width - 200
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo erreur
    Dim Msg
    Msg = "Voulez-vous vraiment quitter?"
    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbNo Then Cancel = True Else Unload Me
Exit Sub
erreur:
   MsgBox Err.Description, 48
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~
    'Initialise SGrid~~~
'~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Initgrid_Services()
    With grid_Service
        .Redraw = False
        .ClearSelection
        .ClearRows
        .Clear
        .HideGroupingBox = True
        .AllowGrouping = True
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .AddColumn "Numero", "", , , 500, , , , , , , CCLSortNumeric
        .AddColumn "Date", "Date", , , 100
        .AddColumn "Etat", "Etat", , , 0
        .AddColumn "HDebut", "Heure Sortie", , , 100
        .AddColumn "HFin", "Heure Entre", , , 100
        .AddColumn "DureTrafic", "Durée Trafic", , , 100
        .AddColumn "Activités", "Activités", , , 500
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
        .Redraw = True
    End With
End Sub
Private Sub Initgrid_FT()
    With grid_Ft
        .HideGroupingBox = True
        .AllowGrouping = True
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .AddColumn "DateFT", "Date", , , 80
        .AddColumn "Numero", "Numero", , , 60, False, , , , , , CCLSortNumeric
        .AddColumn "Matricule", "Matricule", , , 100
        .AddColumn "Conducteur", "Conducteur", , , 100
        .AddColumn "Destination", "Destination", , , 140
        .AddColumn "HeureS", "H.Sortie", , , 60
        .AddColumn "HeureE", "H.Entrée", , , 60
        .AddColumn "CPTS", "CPT.S", , , 60
        .AddColumn "CPTE", "CPT.E", , , 60
        .AddColumn "Distance", "Distance(KM)", , , 40
        .AddColumn "MaxK", "Max-Km", , , 60
        .AddColumn "DifK", "Dif Km", , , 60
        .AddColumn "Dure", "Durée(Heure)", , , 60
        .AddColumn "MaxD", "Max-Durée", , , 60
        .AddColumn "DifD", "Dif-Durée", , , 60 ', False
        .AddColumn "OS", "Op.Sortie)", , , 100
        .AddColumn "OE", "Op.Entre", , , 100
        .AddColumn "Null", ""
        .StretchLastColumnToFit = True
    End With
End Sub
Private Sub Initgrid_Carb()
    With Grid_Carb
        .Redraw = False
        .HideGroupingBox = True
        .AllowGrouping = True
        .GroupRowBackColor = vbWindowBackground
        .GroupRowForeColor = vbWindowText
        .GridLineColor = vbWindowBackground
        .GridFillLineColor = vbWindowBackground
        .GridLines = True
        .SelectionAlphaBlend = True
        .SelectionOutline = True
        .DrawFocusRectangle = False
        .AddColumn "Numero", "Pièce", , , 80, , , , , , , CCLSortNumeric
        .AddColumn "Vehicule", "Véhicule", , , 120
        .AddColumn "Date", "Date", , , 90
        .AddColumn "NbrL", "Nbr.Litres", , , 70
        .AddColumn "Montant", "Montant", , , 90
        .AddColumn "Compteur", "Compteur", , , 90
        .AddColumn "KmParc", "Km.Parcouru", , , 80
        .AddColumn "Consom", "Consomation/100Km", , , 80
        .AddColumn "Anomalie", "Anomalie.consomation", , , 80
        .AddColumn "NULL", ""
        .StretchLastColumnToFit = True
    .Redraw = True
    End With
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Afficher Liste (FindView)~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub cmdFindConducteur_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ConducteurE/H"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Cmd_FindConducteurFT_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "ConducteurStque"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Cmd_FindDestinationFT_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "DestinationE/H"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub Cmd_FindVehiculeFT_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeStqueTF"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub cmd_FindMatricule_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeStqueRp"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Private Sub cmdFindMatricule_Click()
On Error GoTo Err
    Unload FrmFind
    Unload FrmFind_Actif
    Unload FrmFind_Fils
    Unload Frm_FindView
    With Frm_FindView
        .StrSource = "VehiculeStqueCBr"
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'~~~~~~~~~~~~~~~~~
    'ControlBox~~~
'~~~~~~~~~~~~~~~~~
Private Sub grid_Service_ColumnClick(ByVal lCol As Long)
    Dim sTag As String, i As Long
    With grid_Service.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = grid_Service.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        grid_Service.ColumnTag(lCol) = sTag
        Select Case grid_Service.ColumnKey(lCol)
            Case "Conducteur"
                 .SortType(1) = CCLSortString
            Case "Etat"
                 .SortType(1) = CCLSortString
            Case "HDebut"
                 .SortType(1) = CCLSortDateHourAccuracy
            Case "HFin"
                 .SortType(1) = CCLSortDateHourAccuracy
            Case "Dure"
                 .SortType(1) = CCLSortDateHourAccuracy
        End Select
    End With
    Screen.MousePointer = vbHourglass
    grid_Service.Sort
    Screen.MousePointer = vbDefault
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Afficher Row***
Public Sub AfficheRow_Vehicule(ByVal VCode As String)
    Dim LObj_Find As New VEHICULE, Lrs_Find As New Recordset, cbo As ComboBox
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        If Frm_FindView.StrSource = "VehiculeStqueCBr" Then Set cbo = cbo_Matricule
        If Frm_FindView.StrSource = "VehiculeStqueRp" Then Set cbo = Cbo_Vehicule
        If Frm_FindView.StrSource = "VehiculeStqueTF" Then Set cbo = cbo_VehiculeFT
        If Not IsNull(Lrs_Find("Matricule")) Then
            cbo.Text = Lrs_Find("Matricule")
            VCode = Lrs_Find("code")
        End If
    Else
        MsgBox "Code introuvable", vbInformation
        cbo.SetFocus
        Exit Sub
    End If
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
    'Satistiques En/Hors Service***
Private Sub Cmd_SearchService_Click()
    Call Affiche_FP
End Sub
Public Sub Affiche_FP()
    Dim LObj_Find As New Traffic, Lrs_Find As Recordset
    Dim VdateD As Date, vDateF As Date
    Dim Conducteur As String, DESTINATION As String
    Dim YearTrafic As Integer, Name_Table As String
On Error GoTo Err
    Call Initgrid_Services
    Conducteur = cbo_Conducteur.Text
    VdateD = cda_DebutService.Value
    vDateF = cda_FinService.Value
    If cda_DebutService.Value > cda_FinService.Value Then
        MsgBox "Vérifier dates de recherche!...", vbExclamation, App.ProductName
        Exit Sub
    End If
    If Conducteur = "" Then
        MsgBox "Choisir un conducteur !...", vbExclamation, App.ProductName
        Exit Sub
    End If
    For YearTrafic = Year(VdateD) To Year(vDateF)
        Name_Table = "FicheTraffic"
        If YearTrafic < Year(Date) Then Name_Table = "FicheTraffic_" & YearTrafic
        Set Lrs_Find = LObj_Find.GETALL_STATISTIQUESSERVICES(ErrNumber, ErrDescription, ErrSourceDetail, Name_Table, VdateD, vDateF, Conducteur, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_Find = Nothing
    Next
    Dim Ass_Statistique
    If Not Lrs_Find.EOF Then
        grid_Service.Redraw = False
        While Not Lrs_Find.EOF
            If Not ((IsNull(Lrs_Find("HDebut"))) And (IsNull(Lrs_Find("HFin")))) Then Ass_Statistique = UCase(Format(Lrs_Find("DATEdebut"), "dddd-dd-mm-yyyy")) & "     |      Du: " & Lrs_Find("HDebut") & "   |=> Au: " & Lrs_Find("HFin") & "     ||      Durée : " & Lrs_Find("DUREE")
            With grid_Service
                If (Lrs_Find("Etat") = "En-Service") Then
                    .AddRow
                    If Not (IsNull(Lrs_Find("Ndisp"))) Then .CellDetails .Rows, .ColumnIndex("Numero"), Ass_Statistique, , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("Etat"))) Then .CellDetails .Rows, .ColumnIndex("Etat"), Lrs_Find("ETAT"), , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("HDebut"))) Then .CellDetails .Rows, .ColumnIndex("Date"), Lrs_Find("dateDebut"), , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("HDebut"))) Then .CellDetails .Rows, .ColumnIndex("HDebut"), Lrs_Find("Heuresortie"), , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("HFin"))) Then .CellDetails .Rows, .ColumnIndex("HFin"), Lrs_Find("Heureentre"), , , &HC0FFC0
                    If Not (IsNull(Lrs_Find("Heureentre"))) Then .CellDetails .Rows, .ColumnIndex("DureTrafic"), Lrs_Find("DUREETRAFIC"), , , &HC0FFC0
                    DESTINATION = ""
                    If Not (IsNull(Lrs_Find("Destination"))) Then DESTINATION = DESTINATION & " | " & Format(Lrs_Find("HeureSortie"), "hh:mm") & " Aller à " & Lrs_Find("Destination") & " par " & Lrs_Find("vehicule")
                    .CellDetails .Rows, .ColumnIndex("Activités"), DESTINATION, , , &HC0FFC0
                    .CellDetails .Rows, .ColumnIndex("Null"), "", , , &HC0FFC0
                End If
            End With
            Lrs_Find.MoveNext
            Ass_Statistique = ""
        Wend
        grid_Service.Redraw = True
        Set Lrs_Find = Nothing
        With grid_Service
            .GroupRowBackColor = RGB(251, 246, 206)
            .GroupRowForeColor = QBColor(12)
            .ColumnIsGrouped(1) = True
            .GroupRowForeColor = QBColor(10)
            .HideGroupingBox = True
            .AllowGrouping = True
        End With
    End If
    Set Lrs_Find = Nothing
    If grid_Service.Rows > 0 Then grid_Service.SelectedRow = 1
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
    'Satistiques FicheTrafic***
Public Sub Cmd_Searchft_Click()
    Dim VCode As String, CCode As String, DCode As String
    Dim LObj_V As New VEHICULE, LObj_C As New Conducteur, LObj_D As New DESTINATION
    Dim Lrs_V As New Recordset, Lrs_C As New Recordset, Lrs_D As New Recordset
On Error GoTo Err
        VCode = cbo_VehiculeFT.Text
        CCode = cbo_ConducteurFT.Text
        DCode = cbo_DestinationFT.Text
    If VCode <> "Tous" And VCode <> "" Then
        Set Lrs_V = LObj_V.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cbo_VehiculeFT.Text)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_V = Nothing
        If Not Lrs_V.EOF Then
            VCode = Lrs_V("code")
            Set Lrs_V = Nothing
        Else
            MsgBox "Vehicule invalide!..."
            Set Lrs_V = Nothing
            Exit Sub
        End If
    End If
    If CCode <> "Tous" And CCode <> "" Then
        Set Lrs_C = LObj_C.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, cbo_ConducteurFT.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_C = Nothing
        If Not Lrs_C.EOF Then
            CCode = Lrs_C("code")
            Set Lrs_C = Nothing
        Else
            MsgBox "Conducteur invalide!..."
            Set Lrs_C = Nothing
            Exit Sub
        End If
    End If
    If DCode <> "Tous" And DCode <> "" Then
        Set Lrs_D = LObj_D.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, cbo_DestinationFT.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_D = Nothing
        If Not Lrs_D.EOF Then
            DCode = Lrs_D("numero")
            Set Lrs_D = Nothing
        Else
            MsgBox "Déstination invalide!..."
            Set Lrs_D = Nothing
            Exit Sub
        End If
    End If
    If cda_Debutft.Value > cda_FinFT.Value Then
        MsgBox "Période de recherche invalide,..."
        Exit Sub
    End If
    If (cbo_ConducteurFT.Text = "") Or (cbo_VehiculeFT.Text = "") Or (cbo_DestinationFT.Text = "") Then
        MsgBox "Vérifier les informations de recherche,..."
        Exit Sub
    End If
    Call Affiche_FT(VCode, CCode, DCode, cda_Debutft.Value, cda_FinFT.Value)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub Affiche_FT(ByVal VEHICULE As String, _
                    ByVal Conducteur As String, _
                    ByVal DESTINATION As String, _
                    ByVal VdateD As String, _
                    ByVal vDateF As String)
                    
    Dim LObj_Find As New Traffic, Lrs_Find As New Recordset
    Dim YearTrafic As Integer, Name_Table As String, itmX As ListItem
    Dim Voyage As Long, NSecond As Long, Distance As Long
    Dim S As Long, m As Long, H As Long, x As Long, Y As Long, C As Long, Time As String
    Dim w As Long, V As Long, Q As Long, A As Long, T As String
    Dim MoyDur As Long, MoyDis As Long
        Voyage = 0
        NSecond = 0
        Distance = 0
        NAnomalieDuree = 0
        NAnomalieKm = 0
        NAnomalieTotal = 0
    grid_Ft.ClearRows
    For YearTrafic = Year(VdateD) To Year(vDateF)
        Name_Table = "FicheTraffic"
        If YearTrafic < Year(Date) Then Name_Table = "FicheTraffic_" & YearTrafic
        Set Lrs_Find = LObj_Find.GETALL_SUPERVISIONTRAFFICBYDATE(ErrNumber, ErrDescription, ErrSourceDetail, Name_Table, VdateD, vDateF, Conducteur, VEHICULE, DESTINATION, "Statistique", YearTrafic, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        grid_Ft.Redraw = False
        While Not Lrs_Find.EOF
            With grid_Ft
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Numero"), Lrs_Find("Numero")
                .CellDetails .Rows, .ColumnIndex("Matricule"), Lrs_Find("MatriculeVehic")
                .CellDetails .Rows, .ColumnIndex("Conducteur"), Lrs_Find("LibelleCond")
                .CellDetails .Rows, .ColumnIndex("Destination"), Lrs_Find("LibelleDest")
                .CellDetails .Rows, .ColumnIndex("DateFT"), Lrs_Find("DateSortie"), , , &H80FF80, &HFF0000
                .CellDetails .Rows, .ColumnIndex("HeureS"), Lrs_Find("HeureSortie")
                .CellDetails .Rows, .ColumnIndex("OS"), Lrs_Find("OperateurSortie")
                If Not IsNull(Lrs_Find("OperateurEntre")) Then .CellDetails .Rows, .ColumnIndex("OE"), Lrs_Find("OperateurEntre")
                If Not IsNull(Lrs_Find("DifK")) Then .CellDetails .Rows, .ColumnIndex("DifK"), Lrs_Find.Fields("DifK"), , , &H80C0FF, &HFF0000
                If Not IsNull(Lrs_Find("MaxDuree")) Then
                    If Lrs_Find.Fields("Duree") >= Lrs_Find.Fields("MaxDuree") Then
                        .CellDetails .Rows, .ColumnIndex("DifD"), Lrs_Find.Fields("DifD"), , , &H80C0FF, &HFF0000
                    Else
                        .CellDetails .Rows, .ColumnIndex("DifD"), "- " & Lrs_Find.Fields("DifDm"), , , &H80C0FF, &HFF0000
                    End If
                End If
                If Not IsNull(Lrs_Find("MaxCompteur")) And Not IsNull(Lrs_Find.Fields("Duree")) Then
                    If (Val(Lrs_Find.Fields("Kmt")) <= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("Dure"), Lrs_Find.Fields("Duree"), , , &H8080FF Else .CellDetails .Rows, .ColumnIndex("Dure"), Lrs_Find.Fields("Duree")
                    If (Val(Lrs_Find.Fields("Kmt")) <= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("MaxD"), Lrs_Find.Fields("MaxDuree"), , , &H80FFFF Else .CellDetails .Rows, .ColumnIndex("MaxD"), Lrs_Find.Fields("MaxDuree")
                    If (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") <= Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("Distance"), Lrs_Find.Fields("Kmt"), , , &H8080FF Else .CellDetails .Rows, .ColumnIndex("Distance"), Lrs_Find.Fields("Kmt")
                    If (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") <= Lrs_Find.Fields("MaxDuree")) Or (Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree")) Then .CellDetails .Rows, .ColumnIndex("MaxK"), Lrs_Find.Fields("MaxCompteur"), , , &H80FFFF Else .CellDetails .Rows, .ColumnIndex("MaxK"), Lrs_Find.Fields("MaxCompteur")
                End If
                If Not (IsNull(Lrs_Find("HeureENtre"))) Then .CellDetails .Rows, .ColumnIndex("HeureE"), Lrs_Find("HeureENtre")
                .CellDetails .Rows, .ColumnIndex("CPTS"), Lrs_Find("CompteurSortie")
                If Not (IsNull(Lrs_Find("HeureENtre"))) Then .CellDetails .Rows, .ColumnIndex("CPTE"), Lrs_Find("CompteurEntre")
                If Not IsNull(Lrs_Find("MaxCompteur")) And Not IsNull(Lrs_Find.Fields("Duree")) Then
                    If Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") <= Lrs_Find.Fields("MaxDuree") Then
                        NAnomalieKm = NAnomalieKm + 1
                    End If
                    If Val(Lrs_Find.Fields("Kmt")) <= Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree") Then
                        NAnomalieDuree = NAnomalieDuree + 1
                    End If
                    If Val(Lrs_Find.Fields("Kmt")) > Val(Lrs_Find.Fields("MaxCompteur")) And Lrs_Find.Fields("Duree") > Lrs_Find.Fields("MaxDuree") Then
                        NAnomalieDuree = NAnomalieDuree + 1
                        NAnomalieKm = NAnomalieKm + 1
                    End If
                End If
            End With
            ' Affiche LSv_details
            If Not (IsNull(Lrs_Find("HeureENtre"))) And Not IsNull(Lrs_Find.Fields("Kmt")) Then
                If Val(Lrs_Find.Fields("Kmt")) > 0 Then
                    Voyage = Voyage + 1
                    NSecond = NSecond + Lrs_Find("NSecond")
    '                Dure = Minute(Dure) + Minute(Lrs_Find("duree"))
                    Distance = Distance + Lrs_Find("kmt")
                End If
            End If
            Lrs_Find.MoveNext
        Wend
        grid_Ft.Redraw = True
    End If
    x = NSecond \ 60
    Y = x * 60
    S = NSecond - Y     'N° second "Z"
    H = x \ 60         'N° Heure
    m = x - (H * 60)   'N° Minute
    Time = CStr(H) & ":" & CStr(m) & ":" & CStr(S)
    'Lsv_toto
    Lsv_DetailsFT.ListItems.Clear
    Set itmX = Lsv_DetailsFT.ListItems.Add(, , CStr(Voyage))
    If Voyage > 0 Then
        MoyDur = NSecond \ Voyage
        w = MoyDur \ 60
        V = w * 60
        Q = MoyDur - V     'N° second "Z"
        C = w \ 60          'N° Heure
        A = w - (C * 60)   'N° Minute
        T = CStr(C) & ":" & CStr(A) & ":" & CStr(Q)
    End If
    itmX.SubItems(1) = CStr(Time)
    If Distance > 0 Then
        itmX.SubItems(2) = CStr(T)
    Else
        itmX.SubItems(2) = "Voyages égale à zéro"
    End If
    itmX.SubItems(3) = CStr(Distance)
    If Voyage > 0 Then
        MoyDis = Distance \ Voyage
    End If
    If Voyage > 0 Then
        itmX.SubItems(4) = CStr(MoyDis)
    Else
        itmX.SubItems(4) = "Voyages égale à zéro"
    End If
    Lrs_Find.Close
    Set Lrs_Find = Nothing
    Next
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Satistiques Reparation***
Private Sub Cmd_Find_Click()
    Dim VCode As String, LObj_V As New VEHICULE, Lrs_V As New Recordset
On Error GoTo Err
    If Dta_Debut.Value > Dta_Fin.Value Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    List_DetailsRp.ListItems.Clear
    List_detailRp.ListItems.Clear
    VCode = Cbo_Vehicule.Text
     If VCode <> "Tous" And VCode <> "" Then
        Set Lrs_V = LObj_V.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_V = Nothing
        If Not Lrs_V.EOF Then
            VCode = Lrs_V("code")
            Set Lrs_V = Nothing
        Else
            MsgBox "Vehicule introuvable!..."
            Set Lrs_V = Nothing
            Exit Sub
        End If
    End If
    If Cbo_Vehicule.Text = "Tous" Then
        Call AfficheDetailsRp_Tous(Dta_Debut.Value, Dta_Fin.Value)
    Else
        Call AfficheDetailsRp_ParVehicule(Cbo_Vehicule.Text, Dta_Debut.Value, Dta_Fin.Value)
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetailsRp_Tous(ByVal VdateD As Date, ByVal vDateF As Date)
    'variables globales
    Dim LOBJ_StatRp As PieceReparation
    Dim rs As New Recordset
    'variables DetailsP
    Dim TotHTBrut As Double
    Dim TotTTC As Double
    Dim Fcode As String
    Dim Qte As Double
    Dim PUHT As Double
    Dim Remise As Double
    Dim tva As Double
    Dim RP As Double
    Dim TotalG As Double
    Dim i
    'variables Details
    Dim nbRep As Double
    Dim Valeur As Double
    Dim MOeuvre As Double
On Error GoTo Err
    Set itmX = List_DetailsRp.ListItems.Add(, , "Tous")
    Set LOBJ_StatRp = New PieceReparation
    'nombre des reparations
    Set rs = LOBJ_StatRp.Get_SumNbrRepStatist(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        nbRep = 0
        nbRep = nbRep + rs("nbrRep")
        itmX.SubItems(1) = CStr(rs("nbrRep"))
        'TTC
        If Not IsNull(rs("Valeur")) Then
            Valeur = 0
            Valeur = Valeur + rs("valeur")
            itmX.SubItems(2) = CStr(Format(rs("valeur"), "#,##0.000"))
        Else
            itmX.SubItems(2) = "Valeur Null"
        End If
    End If
    rs.Close
    'Vehicule + nbr reparations par vehicule
    Set rs = LOBJ_StatRp.Get_NbrRepStatistGrpVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        While Not rs.EOF
            nbRep = 0
            nbRep = nbRep + rs("nbrRep")
            Set itmX = List_DetailsRp.ListItems.Add(, , CStr(rs("vehicule")))
            itmX.SubItems(1) = CStr(rs("nbrRep"))
        rs.MoveNext
        Wend
    End If
    rs.Close
    'Valeur Reparation par vehicule
    For i = 2 To List_DetailsRp.ListItems.Count
        Set rs = LOBJ_StatRp.Get_ValRepStatistVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, List_DetailsRp.ListItems(i), VdateD, vDateF)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        TotalG = 0
        If Not rs.EOF Then
            While Not rs.EOF
                TotTTC = 0
                TotHTBrut = 0
                Qte = 0
                PUHT = 0
                Remise = 0
                tva = 0
                RP = 0
                Qte = rs("Qte")
                PUHT = rs("PUHT")
                Remise = rs("Remise")
                tva = rs("tva")
                RP = rs("remisePiece")
                TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
                TotTTC = TotTTC + (TotHTBrut + (TotHTBrut * (tva / 100)))
                TotTTC = TotTTC - (TotTTC * RP / 100)
                TotalG = TotalG + TotTTC
                List_DetailsRp.ListItems(i).SubItems(2) = Format(TotalG, "#,##0.000")
                rs.MoveNext
            Wend
        End If
        rs.Close
    Next
    Valeur = 0
    For i = 2 To List_DetailsRp.ListItems.Count
        Valeur = Valeur + List_DetailsRp.ListItems(i).SubItems(2)
    Next
    Set itmX = List_DetailsRp.ListItems.Item(1)
    itmX.SubItems(2) = CStr(Format(Valeur, "#,##0.000"))
    'detailP
    Set rs = LOBJ_StatRp.Get_DetRepStatist(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
    If Not rs.EOF Then
        While Not rs.EOF
            TotTTC = 0
            TotHTBrut = 0
            Qte = 0
            PUHT = 0
            Remise = 0
            tva = 0
            Qte = rs("Qte")
            PUHT = rs("PUHT")
            Remise = rs("Remise")
            tva = rs("tva")
            RP = rs("remisePiece")
            TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
            TotTTC = TotHTBrut + (TotHTBrut * (tva / 100))
            TotTTC = TotTTC - (TotTTC * RP / 100)
               Set itmX = List_detailRp.ListItems.Add(, , rs("Numero"))
                itmX.SubItems(1) = rs("datePiece")
                itmX.SubItems(2) = rs("Vehicule")
                itmX.SubItems(3) = rs("Designation")
                itmX.SubItems(4) = rs("Qte")
                itmX.SubItems(5) = Format(TotTTC, "#,##0.000")
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set LOBJ_StatRp = Nothing
    List_detailRp.ColumnHeaders(3).Width = 1640
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetailsRp_ParVehicule(ByVal Matricule As String, ByVal VdateD As Date, ByVal vDateF As Date)
    'variables globales
    Dim LOBJ_StatRp As PieceReparation
    Dim rs As New Recordset
    Dim i As Integer
    'variables DetailsP
    Dim TotHTBrut As Double
    Dim TotTTC As Double
    Dim Fcode As String
    Dim Qte As Double
    Dim PUHT As Double
    Dim Remise As Double
    Dim tva As Double
    Dim RemiseP As Double
    'variables Details
    Dim nbRep As Double
    Dim Valeur As Double
On Error GoTo Err
    Set itmX = List_DetailsRp.ListItems.Add(, , Cbo_Vehicule.Text)
    Set LOBJ_StatRp = New PieceReparation
    'nombre des reparations
    Set rs = LOBJ_StatRp.Get_NbrRepStatistVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Matricule, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        nbRep = 0
        nbRep = nbRep + rs("nbrRep")
        itmX.SubItems(1) = CStr(rs("nbrRep"))
    End If
    rs.Close
    'detail P
    Set rs = LOBJ_StatRp.Get_PieceRepStatistVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Matricule, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        While Not rs.EOF
        TotTTC = 0
        TotHTBrut = 0
        Qte = 0
        PUHT = 0
        Remise = 0
        tva = 0
        RemiseP = 0
        Qte = rs("Qte")
        PUHT = rs("PUHT")
        Remise = rs("Remise")
        tva = rs("tva")
        RemiseP = rs("RemisePiece")
        TotHTBrut = FrmSaisiePieceReparation.Return_TotHT(Qte, PUHT, Remise)
        TotHTBrut = TotHTBrut + (TotHTBrut * (tva / 100))
        TotTTC = TotHTBrut - (RemiseP * TotHTBrut / 100)
                Set itmX = List_detailRp.ListItems.Add(, , rs("Numero"))
                itmX.SubItems(1) = rs("datePiece")
                itmX.SubItems(2) = rs("Vehicule")
                itmX.SubItems(3) = rs("Designation")
                itmX.SubItems(4) = rs("Qte")
                itmX.SubItems(5) = Format(TotTTC, "#,##0.000")
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set LOBJ_StatRp = Nothing
     List_detailRp.ColumnHeaders(3).Width = 0
    'Totale réparation
    Valeur = 0
    If List_detailRp.ListItems.Count > 0 Then
        For i = 1 To List_detailRp.ListItems.Count
            Valeur = Valeur + List_detailRp.ListItems(i).SubItems(5)
        Next
    End If
    Set itmX = List_DetailsRp.ListItems.Item(1)
    itmX.SubItems(2) = CStr(Valeur)
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
'Satistiques Carburant***
Private Sub Cmd_Search_Click()
    Dim VCode As String, LObj_V As New VEHICULE, Lrs_V As New Recordset
On Error GoTo Err
    If CDate(cda_debut.Text) > CDate(cda_fin.Text) Then
        MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
        Exit Sub
    End If
    Lsv_Details.ListItems.Clear
    Grid_Carb.ClearRows
    VCode = cbo_Matricule.Text
     If VCode <> "Tous" And VCode <> "" Then
        Set Lrs_V = LObj_V.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VCode)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LObj_V = Nothing
        If Not Lrs_V.EOF Then
            VCode = Lrs_V("code")
            Set Lrs_V = Nothing
        Else
            MsgBox "Vehicule introuvable!..."
            Set Lrs_V = Nothing
            Exit Sub
        End If
    End If
    If cbo_Matricule.Text = "Tous" Then
        Call AfficheDetails_Tous(cda_debut.Text, cda_fin.Text)
    Else
        Call AfficheDetails_ParVehicule(cbo_Matricule.Text, cda_debut.Text, cda_fin.Text)
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetails_Tous(ByVal VdateD As Date, ByVal vDateF As Date)
    Dim LOBJ_BC As BonCarburant
    Dim rs As New Recordset
    Dim rD As New Recordset
    Dim i
    Dim TLitre As Double
    Dim Valeur As Double
    Dim MaxC As Long
    Dim MinC As Long
    Dim NbKM As Long
    Dim KmCarburant As Double
On Error GoTo Err
    'detail P
    Set LOBJ_BC = New BonCarburant
    Set rs = LOBJ_BC.Get_StatistDetBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        Grid_Carb.Redraw = False
        While Not rs.EOF
            With Grid_Carb
                .AddRow
                .CellDetails .Rows, 1, rs("Numero")
                .CellDetails .Rows, .ColumnIndex("Vehicule"), rs("Matricule")
                .CellDetails .Rows, .ColumnIndex("Date"), rs("DateDoc")
                .CellDetails .Rows, .ColumnIndex("NbrL"), Format(rs("Litre"), "#,##0.00")
                .CellDetails .Rows, .ColumnIndex("Montant"), Format(rs("Litre") * rs("prixLitre"), "#,##0.000")
                .CellDetails .Rows, .ColumnIndex("Compteur"), rs("Compteur")
                'Get ancien compteur pour chaque voiture et chaque boncarb
                Set rD = LOBJ_BC.Get_AnComptCar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Numero"), rs("Vehicule"))
                If ErrNumber <> 0 Then
                    ErrNumber = 0
                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                    Exit Sub
                End If
                If Not rD.EOF Then
                    .CellDetails .Rows, .ColumnIndex("KmParc"), (Val(rs("Compteur")) - Val(rD("maxCpt")))
                    If (Val(rs("Compteur")) - Val(rD("maxCpt"))) <> 0 Then .CellDetails .Rows, .ColumnIndex("Consom"), Format((rs("Litre") * 100) / (Val(rs("Compteur")) - Val(rD("maxCpt"))), "#,##0.000")
                End If
                rD.Close
                If Not IsNull(rs("AnomalieConsom")) Then
                    If CDbl(rs("AnomalieConsom")) >= 2 Then
                        .CellDetails .Rows, .ColumnIndex("Anomalie"), Format(rs("AnomalieConsom"), "#,##0.00"), , , &H8080FF
                    Else
                        .CellDetails .Rows, .ColumnIndex("Anomalie"), Format(rs("AnomalieConsom"), "#,##0.00")
                    End If
                End If
            End With
        rs.MoveNext
        Wend
    Grid_Carb.ColumnWidth("Vehicule") = 120
    Grid_Carb.Redraw = True
    Else
        MsgBox "Pas de donnèes à visualiser !", vbInformation
        Exit Sub
    End If
    rs.Close
    'totaux Lsv Details
    'Parcourir Grid_carb
    Dim Valc As Double
    Dim NBL As Double
    Dim itm
    Dim Veh As String
    Dim ii
    Dim KmParcVeh As Double
    Dim TotLitVeh As Double
    Dim TotKm As Long
    Valc = 0
    NBL = 0
    TotKm = 0
    For i = 1 To Grid_Carb.Rows
        Valc = Valc + Grid_Carb.CellText(i, 5)
        NBL = NBL + Grid_Carb.CellText(i, 4)
        TotKm = TotKm + Grid_Carb.CellText(i, 7)
    Next
    Set itm = Lsv_Details.ListItems.Add(, , "Tous")
        itm.SubItems(1) = CStr(Format(Valc, "#,##0.000"))
        itm.SubItems(3) = CStr(Format(NBL, "#,##0.00"))
        itm.SubItems(4) = CStr(TotKm)
        If TotKm <> 0 Then itm.SubItems(5) = CStr(Format(NBL * 100 / TotKm, "#,##0.00"))
        
    'Details Lsv_details
    Set rs = LOBJ_BC.Get_StatistBC(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
        While Not rs.EOF
        'Lire et prix Litre
            TLitre = 0
            Valeur = 0
            KmParcVeh = 0
            TotLitVeh = 0
            TLitre = TLitre + rs("Litre")
            Valeur = Valeur + rs("Litre") * rs("Prix")
            For ii = 1 To Grid_Carb.Rows
                If Grid_Carb.CellText(ii, 2) = rs("Vehicule") Then
                    KmParcVeh = KmParcVeh + Grid_Carb.CellText(ii, 7)
                    TotLitVeh = TotLitVeh + Grid_Carb.CellText(ii, 4)
                End If
            Next
            'Consommation par 100 KM
            If Not (KmParcVeh = 0) Then
                KmCarburant = ((TotLitVeh * 100) / KmParcVeh)
            Else
                KmCarburant = 0
            End If
            Set itmX = Lsv_Details.ListItems.Add(, , CStr(rs("Vehicule")))
                itmX.SubItems(1) = CStr(Format(Valeur, "#,##0.000"))
                itmX.SubItems(2) = CStr(Format(rs("Prix"), "#,##0.000"))
                itmX.SubItems(3) = CStr(Format(TLitre, "#,##0.00"))
                itmX.SubItems(4) = CStr(KmParcVeh)
                itmX.SubItems(5) = CStr(Format(KmCarburant, "#,##0.000"))
            
            rs.MoveNext
        Wend
    End If
    rs.Close
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
Public Sub AfficheDetails_ParVehicule(ByVal Matricule As String, ByVal VdateD As Date, ByVal vDateF As Date)
    Dim LOBJ_BC As BonCarburant
    Dim rs As New Recordset
    Dim rD As New Recordset
    Dim i As Integer
    Dim itm
    Dim TLitre As Double
    Dim Valeur As Double
    Dim MaxC As Long
    Dim MinC As Long
    Dim NbKM As Long
    ''selection de code de Vehicule
    Dim CodV As String
On Error GoTo Err
    CodV = Return_CodVehicule(Matricule)
    'Remplissage de Grid Lsv_detailP
    Set LOBJ_BC = New BonCarburant
    Set rs = LOBJ_BC.Get_StatistBCVeh(ErrNumber, ErrDescription, ErrSourceDetail, CNB, VdateD, vDateF, CodV)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    If Not rs.EOF Then
    Grid_Carb.Redraw = False
        While Not rs.EOF
            With Grid_Carb
                .AddRow
                .CellDetails .Rows, 1, rs("Numero")
                .CellDetails .Rows, .ColumnIndex("Vehicule"), rs("Matricule")
                .CellDetails .Rows, .ColumnIndex("Date"), rs("DateDoc")
                .CellDetails .Rows, .ColumnIndex("NbrL"), Format(rs("Litre"), "#,##0.00")
                .CellDetails .Rows, .ColumnIndex("Montant"), Format(rs("Litre") * rs("prixLitre"), "#,##0.000")
                .CellDetails .Rows, .ColumnIndex("Compteur"), rs("Compteur")
                'Get ancien compteur pour chaque voiture et chaque boncarb
                Set rD = LOBJ_BC.Get_AnComptCar(ErrNumber, ErrDescription, ErrSourceDetail, CNB, rs("Numero"), CodV)
                If ErrNumber <> 0 Then
                    ErrNumber = 0
                    MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
                    Exit Sub
                End If
                If Not rD.EOF Then
                    .CellDetails .Rows, .ColumnIndex("KmParc"), (Val(rs("Compteur")) - Val(rD("maxCpt")))
                    If (Val(rs("Compteur")) - Val(rD("maxCpt"))) <> 0 Then .CellDetails .Rows, .ColumnIndex("Consom"), Format((rs("Litre") * 100) / (Val(rs("Compteur")) - Val(rD("maxCpt"))), "#,##0.000")
                End If
                rD.Close
                If Not IsNull(rs("AnomalieConsom")) Then
                    If CDbl(rs("AnomalieConsom")) >= 2 Then
                        .CellDetails .Rows, .ColumnIndex("Anomalie"), Format(rs("AnomalieConsom"), "#,##0.00"), , , &H8080FF
                    Else
                        .CellDetails .Rows, .ColumnIndex("Anomalie"), Format(rs("AnomalieConsom"), "#,##0.00")
                    End If
                End If
            End With
            rs.MoveNext
        Wend
        Grid_Carb.ColumnWidth("Vehicule") = 0
        Grid_Carb.Redraw = True
    Else
        MsgBox "Pas de donnèes à visualiser !", vbInformation
        Exit Sub
    End If
    rs.Close
    'Parcourir Grid_Carb
    Dim Valc As Double
    Dim NBL As Double
    Dim KmParcVeh As Long
    Dim KmCarburant As Double
    Valc = 0
    NBL = 0
    KmParcVeh = 0
    KmCarburant = 0
    For i = 1 To Grid_Carb.Rows
        Valc = Valc + Grid_Carb.CellText(i, 5)
        NBL = NBL + Grid_Carb.CellText(i, 4)
        KmParcVeh = KmParcVeh + Grid_Carb.CellText(i, 7)
    Next
    Set itm = Lsv_Details.ListItems.Add(, , CStr(Matricule))
        itm.SubItems(1) = CStr(Format(Valc, "#,##0.000"))
        itm.SubItems(3) = CStr(Format(NBL, "#,##0.00"))
        itm.SubItems(4) = KmParcVeh
        'consommation par 100 km
    If Not (itm.SubItems(4) = 0) Then
        KmCarburant = ((itm.SubItems(3) * 100) / itm.SubItems(4))
        itm.SubItems(5) = CStr(Format(KmCarburant, "#,##0.000"))
    Else
        itm.SubItems(5) = "zéro km!!"
    End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation, App.ProductName
End Sub
'____________________________________________________________________________________________________________________________________
'{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}{}
    'Print***
Private Sub CmdPrint_Click()
    Dim VCode As String
    Dim DATEDEBUT As Date
    Dim DateFin As Date
    Dim TotLitre As Double
    Dim nbrRep As Long
    Dim Total As Double
    Dim J
On Error GoTo Err
'Imprimer statistique carburant
If Tab_Satistiques.Tab = 0 Then
    If Grid_Carb.Rows = 0 Then
        MsgBox "Pas de données à imprimer .", vbInformation
        Exit Sub
    End If
    
    DATEDEBUT = cda_debut.Text
    DateFin = cda_fin.Text
    If DATEDEBUT > DateFin Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    VCode = cbo_Matricule.Text
    TotLitre = 0
    Total = 0
    If MsgBox("Imprimer statistiques carburant   ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
        If VCode = "Tous" Then
            For J = 1 To Lsv_Details.ListItems.Count
                If (Lsv_Details.ListItems(J)) = "Tous" Then
                    TotLitre = Lsv_Details.ListItems(J).ListSubItems(3)
                    Total = Lsv_Details.ListItems(J).ListSubItems(1)
                End If
            Next
        Else
            TotLitre = Lsv_Details.ListItems(1).ListSubItems(3)
            Total = Lsv_Details.ListItems(1).ListSubItems(1)
        End If
        Call Frm_Rpt_Apercus.PrintOutAndApercu_StatCarb(0, DATEDEBUT, DateFin, VCode, LStr_NameUser, TotLitre, Total)
        Frm_Rpt_Apercus.Show
    End If
 'Imprimer statistique réparation
ElseIf Tab_Satistiques.Tab = 1 Then
    If List_detailRp.ListItems.Count = 0 Then
        MsgBox "Pas de données à imprimer .", vbInformation
        Exit Sub
    End If
    DATEDEBUT = Dta_Debut.Value
    DateFin = Dta_Fin.Value
     If DATEDEBUT > DateFin Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    VCode = Cbo_Vehicule.Text
    nbrRep = 0
    Total = 0
    If MsgBox("Imprimer statistiques réparation   ", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
        If VCode = "Tous" Then
            For J = 1 To List_DetailsRp.ListItems.Count
                If (List_DetailsRp.ListItems(J)) = "Tous" Then
                    nbrRep = List_DetailsRp.ListItems(J).ListSubItems(1)
                    Total = List_DetailsRp.ListItems(1).ListSubItems(2)
                End If
            Next
        Else
            nbrRep = List_DetailsRp.ListItems(1).ListSubItems(1)
            Total = List_DetailsRp.ListItems(1).ListSubItems(2)
        End If
        Call Frm_Rpt_Apercus.PrintOutAndApercu_StatRep(0, DATEDEBUT, DateFin, VCode, LStr_NameUser, nbrRep, Total)
        Frm_Rpt_Apercus.Show
    End If
 'Imprimer statistique traffic
ElseIf Tab_Satistiques.Tab = 2 Then
    If grid_Ft.Rows = 0 Then
        MsgBox "Pas de données à imprimer.", vbInformation
        Exit Sub
    End If
    DATEDEBUT = cda_Debutft.Value
    DateFin = cda_FinFT.Value
     If DATEDEBUT > DateFin Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    If MsgBox("Imprimer statistique traffic ?", vbYesNo + vbDefaultButton1 + vbInformation, App.ProductName) = vbYes Then
        Call Frm_Rpt_Apercus.PrintOutAndApercu_AnomalieTrafic(0, DATEDEBUT, DateFin, VCodeDrive, VCodeVehicle, VCodeDestination, LStr_NameUser, CStr(NAnomalieKm), CStr(NAnomalieDuree), CStr(NAnomalieTotal), False)
        Frm_Rpt_Apercus.Show
    End If
'Imprimer statistique Services
ElseIf Tab_Satistiques.Tab = 3 Then
    If grid_Service.Rows = 0 Then
        MsgBox "Pas de données à imprimer.", vbInformation
        Exit Sub
    End If
    DATEDEBUT = cda_DebutService.Value
    DateFin = cda_FinService.Value
     If DATEDEBUT > DateFin Then
       MsgBox "Vérifier les dates saisies ! ", vbInformation, App.ProductName
       Exit Sub
    End If
    If MsgBox("Imprimer statistique service ?", vbYesNo + vbDefaultButton1 + vbInformation, App.ProductName) = vbYes Then
        Call Frm_Rpt_Apercus.PrintOutAndApercu_StatService(0, DATEDEBUT, DateFin, cbo_Conducteur.Text, LStr_NameUser)
        Frm_Rpt_Apercus.Show
    End If
End If
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
'=================================================
Public Sub AfficheRow_Conducteur(ByVal VCode As String)
    Dim LOBJ_Cond As New Conducteur
    Dim Lrs_Cond As New Recordset
    Dim cboPers As ComboBox
On Error GoTo Err
    Set Lrs_Cond = LOBJ_Cond.GetRow_Conducteur_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LOBJ_Cond = Nothing
    If Not Lrs_Cond.EOF Then
        If Frm_FindView.StrSource = "ConducteurStque" Then Set cboPers = cbo_ConducteurFT
        If Frm_FindView.StrSource = "ConducteurE/H" Then Set cboPers = cbo_Conducteur
        If Not IsNull(Lrs_Cond("Libelle")) Then
            cboPers.Text = Lrs_Cond("Libelle")
            VCodeDrive = Lrs_Cond("code")
        End If
    End If
    Lrs_Cond.Close
    Set Lrs_Cond = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Sub AfficheRow_Destination(ByVal VCode As String)
    Dim LObj_Find As New DESTINATION
    Dim Lrs_Find As New Recordset
    Dim cboDest As ComboBox
On Error GoTo Err
    Set Lrs_Find = LObj_Find.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, VCode, CNB)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Sub
    End If
    Set LObj_Find = Nothing
    If Not Lrs_Find.EOF Then
        'Charge
        If Frm_FindView.StrSource = "DestinationE/H" Then Set cboDest = cbo_DestinationFT
        If Not IsNull(Lrs_Find("Libelle")) Then
            cboDest.Text = Lrs_Find("Libelle")
            VCodeDestination = Lrs_Find("Numero")
        End If
    End If
    Lrs_Find.Close
    Set Lrs_Find = Nothing
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub cbo_ConducteurFT_Click()
    Dim LOBJ_Cond As New Conducteur
    Dim Lrs_Cond As Recordset
    If cbo_ConducteurFT.ListIndex = 0 Then
        VCodeDrive = "  -  Tous"
    Else
        Set Lrs_Cond = LOBJ_Cond.GetRow_Conducteur_ByLibelle(ErrNumber, ErrDescription, ErrSourceDetail, cbo_ConducteurFT.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set LOBJ_Cond = Nothing
        
        If Not Lrs_Cond.EOF Then VCodeDrive = Lrs_Cond("Code")
    End If
End Sub
Private Sub cbo_VehiculeFT_Click()
    Dim Lobj_Vehicule As New VEHICULE
    Dim Lrs_Vehicule As Recordset
    If cbo_VehiculeFT.ListIndex = 0 Then
        VCodeVehicle = "  -  Tous"
    Else
         Set Lrs_Vehicule = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, cbo_VehiculeFT.Text)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set Lobj_Vehicule = Nothing
        
        If Not Lrs_Vehicule.EOF Then VCodeVehicle = Lrs_Vehicule("Code")
        Set Lrs_Vehicule = Nothing
    End If
End Sub
Private Sub cbo_DestinationFT_Click()
    Dim Lobj_Dest As New DESTINATION
    Dim Lrs_Dest As Recordset
    If cbo_DestinationFT.ListIndex = 0 Then
        VCodeDestination = "  -  Tous"
    Else
        Set Lrs_Dest = Lobj_Dest.GetRow_Destination_ByCode(ErrNumber, ErrDescription, ErrSourceDetail, cbo_DestinationFT.Text, CNB)
        If ErrNumber <> 0 Then
            ErrNumber = 0
            MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
            Exit Sub
        End If
        Set Lobj_Dest = Nothing
        
        If Not Lrs_Dest.EOF Then VCodeDestination = Lrs_Dest("Numero")
    End If
End Sub
Private Sub Cbo_Conducteur_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_Conducteur.Text)
        For i = 0 To cbo_Conducteur.ListCount - 1
            If Left(cbo_Conducteur.List(i), start) = cbo_Conducteur.Text Then
                cbo_Conducteur.Text = cbo_Conducteur.List(i)
            End If
        Next
        cbo_Conducteur.SelStart = start
        cbo_Conducteur.SelLength = Len(cbo_Conducteur.Text)
    End If
End Sub
Private Sub Cbo_Conducteur_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub cbo_vehiculeft_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_VehiculeFT.Text)
        For i = 0 To cbo_VehiculeFT.ListCount - 1
            If Left(cbo_VehiculeFT.List(i), start) = cbo_VehiculeFT.Text Then
                cbo_VehiculeFT.Text = cbo_VehiculeFT.List(i)
            End If
        Next
        cbo_VehiculeFT.SelStart = start
        cbo_VehiculeFT.SelLength = Len(cbo_VehiculeFT.Text)
    End If
End Sub
Private Sub cbo_Vehiculeft_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub cbo_ConducteurFT_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_ConducteurFT.Text)
        For i = 0 To cbo_ConducteurFT.ListCount - 1
            If Left(cbo_ConducteurFT.List(i), start) = cbo_ConducteurFT.Text Then
                cbo_ConducteurFT.Text = cbo_ConducteurFT.List(i)
            End If
        Next
        cbo_ConducteurFT.SelStart = start
        cbo_ConducteurFT.SelLength = Len(cbo_ConducteurFT.Text)
    End If
End Sub
Private Sub Cbo_Conducteurft_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub cbo_destinationft_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_DestinationFT.Text)
        For i = 0 To cbo_DestinationFT.ListCount - 1
            If Left(cbo_DestinationFT.List(i), start) = cbo_DestinationFT.Text Then
                cbo_DestinationFT.Text = cbo_DestinationFT.List(i)
            End If
        Next
        cbo_DestinationFT.SelStart = start
        cbo_DestinationFT.SelLength = Len(cbo_DestinationFT.Text)
    End If
End Sub
Private Sub cbo_destinationft_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub grid_ft_ColumnClick(ByVal lCol As Long)
    Dim sTag As String
    Dim i As Long
    With grid_Ft.SortObject
        .Clear
        .SortColumn(1) = lCol
        sTag = grid_Ft.ColumnTag(lCol)
        If (sTag = "") Then
            sTag = "DESC"
            .SortOrder(1) = CCLOrderAscending
        Else
            sTag = ""
            .SortOrder(1) = CCLOrderDescending
        End If
        grid_Ft.ColumnTag(lCol) = sTag
        Select Case grid_Ft.ColumnKey(lCol)
            Case "Matricule"
                 .SortType(1) = CCLSortString
            Case "Conducteur"
                 .SortType(1) = CCLSortString
            Case "Destination"
                 .SortType(1) = CCLSortString
            Case "DateFT"
                 .SortType(1) = CCLSortDate
            Case "HeureS"
                 .SortType(1) = CCLSortDateHourAccuracy
            Case "HeureE"
                 .SortType(1) = CCLSortDateHourAccuracy
            Case "CPTS"
                 .SortType(1) = CCLSortNumeric
            Case "CPTE"
                 .SortType(1) = CCLSortNumeric
            Case "Distance"
                 .SortType(1) = CCLSortNumeric
            Case "Dure"
                 .SortType(1) = CCLSortDateHourAccuracy
        End Select
    End With
    Screen.MousePointer = vbHourglass
    grid_Ft.Sort
    Screen.MousePointer = vbDefault
End Sub
Private Sub grid_ft_DblClick(ByVal lRow As Long, ByVal lCol As Long)
On Error GoTo Err
    With Frm_MajStatFT
        .selectFT (grid_Ft.CellText(grid_Ft.SelectedRow, grid_Ft.ColumnIndex("Numero")))
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub Cbo_Vehicule_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(Cbo_Vehicule.Text)
        For i = 0 To Cbo_Vehicule.ListCount - 1
            If Left(Cbo_Vehicule.List(i), start) = Cbo_Vehicule.Text Then
                Cbo_Vehicule.Text = Cbo_Vehicule.List(i)
            End If
        Next
        Cbo_Vehicule.SelStart = start
        Cbo_Vehicule.SelLength = Len(Cbo_Vehicule.Text)
    End If
End Sub
Private Sub List_detailRp_DblClick()
    Dim VCode
    Dim i As Integer
On Error GoTo Err
    i = List_detailRp.SelectedItem.Index
    VCode = List_detailRp.ListItems.Item(i)
    'ViderZone (frm)
    With FrmConsultPieceReception       'FrmPieceReparation
        .AfficheRow (VCode)
        .Show
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Private Sub Cbo_Vehicule_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub Cbo_Vehicule_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub cbo_Matricule_Change()
    Dim i As Integer, start As Integer
    Dim ShiftDown As Boolean
    Dim CtrlDown As Boolean
    Dim AltDown As Boolean
    
    ShiftDown = (theshift And vbShiftMask) > 0
    CtrlDown = (theshift And vbCtrlMask) > 0
    AltDown = (theshift And vbAltMask) > 0
    If thekey = vbKeyLeft Or thekey = vbKeyRight Or thekey = vbKeyUp Or thekey = vbKeyDown _
        Or thekey = vbKeyBack Or thekey = vbKeyDelete Or ShiftDown Or AltDown Or CtrlDown Then
    Else
        start = Len(cbo_Matricule.Text)
        For i = 0 To cbo_Matricule.ListCount - 1
            If Left(cbo_Matricule.List(i), start) = cbo_Matricule.Text Then
                cbo_Matricule.Text = cbo_Matricule.List(i)
            End If
        Next
        cbo_Matricule.SelStart = start
        cbo_Matricule.SelLength = Len(cbo_Matricule.Text)
    End If
End Sub
Private Sub Cbo_Matricule_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub
Private Sub cbo_Matricule_KeyUp(KeyCode As Integer, Shift As Integer)
    thekey = KeyCode
    theshift = Shift
End Sub
Private Sub Grid_Carb_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    Dim VCode
On Error GoTo Err
    VCode = Grid_Carb.CellText(lRow, 1)
    With frmConsultBC
        .AfficheRow (VCode)
        .Show vbModal
    End With
Exit Sub
Err:
    MsgBox Err.Description, vbInformation
End Sub
Public Function Return_CodVehicule(ByVal Matricule As String) As String
    Dim Lobj_Vehicule As New VEHICULE
    Dim Lrs_Vehicule As Recordset
On Error GoTo Err
    Set Lrs_Vehicule = Lobj_Vehicule.GetVehByCodeMat(ErrNumber, ErrDescription, ErrSourceDetail, CNB, Matricule)
    If ErrNumber <> 0 Then
        ErrNumber = 0
        MsgBox ErrDescription & vbCr & ErrSourceDetail, vbQuestion
        Exit Function
    End If
    Set Lobj_Vehicule = Nothing
    If Not Lrs_Vehicule.EOF Then
        'Charge
        If Not IsNull(Lrs_Vehicule("Code")) Then
            Return_CodVehicule = CStr(Lrs_Vehicule("Code"))
        End If
    End If
    Lrs_Vehicule.Close
    Set Lrs_Vehicule = Nothing
Exit Function
Err:
    MsgBox Err.Description, vbExclamation, App.ProductName
End Function



