VERSION 5.00
Begin VB.Form frmconnexionx 
   BorderStyle     =   0  'None
   Caption         =   "Connexion"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmconnexionx.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3960
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4800
      ScaleHeight     =   375
      ScaleWidth      =   180
      TabIndex        =   9
      Top             =   2640
      Width           =   173
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4800
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4800
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4800
      ScaleHeight     =   375
      ScaleWidth      =   1095
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4320
      Top             =   3240
   End
   Begin VB.TextBox txtpasse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtutilisateur 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   6600
      Picture         =   "frmconnexionx.frx":17569
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bizerte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   5280
      TabIndex        =   5
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Centra Nord"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1830
   End
   Begin VB.Label cmdconnexion 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   6120
      Picture         =   "frmconnexionx.frx":17873
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   6480
      Picture         =   "frmconnexionx.frx":184B5
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   6720
      Picture         =   "frmconnexionx.frx":190F7
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   6000
      Picture         =   "frmconnexionx.frx":19401
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmconnexionx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flap
Dim Y

Private Sub cmdconnexion_Click()

On Error GoTo Err
    
    Call CONNEXION_DIRECT
    
    Exit Sub
Err:
   MsgBox Err.Description, vbInformation
    
End Sub

Private Sub Form_Load()

On Error GoTo Err
    Call PROG_EXISTE
    txtutilisateur.Text = GetSetting("CentraNord", "GestParc", "CONNECTIONDEFAUT", "")
    Y = 0
   
   txtutilisateur.Text = GetIni("connexion", "utilisateur")

    Exit Sub
Err:
MsgBox Err.Description, vbInformation
End Sub

Public Sub PROG_EXISTE()
           
    Dim rep As VbMsgBoxResult, Lfvar As String
    Dim AppName As String
    Lfvar = Chr$(13) + Chr$(10)

    If App.PrevInstance Then
            If MsgBox("Le logiciel [" & App.Title + "] ne peut pas être chargé." + Lfvar + _
                        "car il existe déjà en mémoire." + Lfvar + Lfvar + _
                        "Arrêtez  ce proccessus.", vbYesNo + vbExclamation, "Logiciel est déjà en mémoire.") = vbYes Then
                        
                        Call KillProcessus(App.Title)
                        Call KillProcessus(App.Title)
            Else
                Unload Me
                End
            End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteIni "connexion", "utilisateur", txtutilisateur.Text
End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub Timer1_Timer()
Call runtop
End Sub

Private Sub txtpasse_GotFocus()
txtpasse.SelStart = 0
txtpasse.SelLength = Len(txtutilisateur.Text)
End Sub

Private Sub txtpasse_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo erreur

    If KeyCode = vbKeyReturn Then
        Call cmdconnexion_Click
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    Exit Sub
erreur:
MsgBox Err.Description, vbInformation

'End Sub

txtutilisateur.SelStart = 0
txtutilisateur.SelLength = Len(txtutilisateur.Text)
'End Sub

   If KeyCode = vbKeyReturn Then
        SendKeys "{tab}"
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub txtutilisateur_GotFocus()
txtutilisateur.SelStart = 0
End Sub


Private Sub CONNEXION_DIRECT()

Dim i As Integer
Dim LInt_Code As Integer

'Connexion à la base
Dim strConnect As String
Dim w, k
Dim SQL As String
Dim rs As New ADODB.Recordset
On Error GoTo Err
Call LOBJ_CON.Connect(ErrNumber, ErrDescription, ErrSourceDetail, CNB, _
    GetSetting("CentraNord", "GestParc", "dbServer"), _
    GetSetting("CentraNord", "GestParc", "dbName"), _
    GetSetting("CentraNord", "GestParc", "dbUser"), _
    GetSetting("CentraNord", "GestParc", "dbPwd"), _
    GetSetting("CentraNord", "GestParc", "dbServer2"), _
    GetSetting("CentraNord", "GestParc", "dbPwd2"))

    If ErrNumber <> 0 Then
        MsgBox ErrDescription & vbNewLine & ErrSourceDetail, vbQuestion
        ErrNumber = 0
        Exit Sub
    End If

    SaveSetting "CentraNord", "GestParc", "CONNECTIONDEFAUT", ""

    'Verifier si l'utilisateur est actif
    w = UCase(txtpasse.Text)
    k = ""
    For i = 1 To Len(w)
        k = k & Asc(Mid(w, i, 1))
    Next
    
    rs.CursorLocation = adUseClient
    SQL = "select * from Utilisateur where MP like " & SQLText(UCase(k)) & " and NOMPRN like '" & txtutilisateur.Text & "'"
    rs.Open SQL, CNB, adOpenKeyset, adLockPessimistic
    If Not rs.EOF Then
        
        LInt_UserId = rs.Fields("Code")
        LStr_NameUser = rs.Fields("NomPrn")
        Timer1.Interval = 0
        Unload Me
        Frm_Main.Show
    Else
        MsgBox "Mot passe incorrect .?  ", vbExclamation
        With txtpasse
            .SetFocus
            .SelStart = 0
            .SelLength = Len(txtpasse.Text)
        End With
    End If

    Set rs = Nothing
Exit Sub
Err:
    MsgBox Err.Number & vbCr & Err.Description & vbCr & Err.Source, vbExclamation, App.ProductName

End Sub
Private Sub runtop()
    ' Avance l'animation d'un cadre.
    Y = Y + 1: If Y = 4 Then Y = 0
    Image2.Picture = Image1(Y).Picture
End Sub
