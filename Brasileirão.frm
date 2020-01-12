VERSION 5.00
Begin VB.Form Brasileirão 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brasileirão"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton BtnBuscaTime 
      Caption         =   "Buscar Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox TxtTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "dsad"
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite aqui o nome do time"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Brasileirão Série A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1900
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   1965
      Picture         =   "Brasileirão.frx":0000
      Top             =   120
      Width           =   2490
   End
End
Attribute VB_Name = "Brasileirão"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnBuscaTime_Click()
    If TxtTime.Text = "palmeiras" Or TxtTime.Text = "Palmeiras" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "santos" Or TxtTime.Text = "Santos" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "flamengo" Or TxtTime.Text = "Flamengo" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "internacional" Or TxtTime.Text = "Internacional" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "atletico" Or TxtTime.Text = "Atletico" Or TxtTime.Text = "atlético" Or TxtTime.Text = "Atlético" Or TxtTime.Text = "atletico mg" Or TxtTime.Text = "Atletico mg" Or TxtTime.Text = "atlético mg" Or TxtTime.Text = "Atlético Mg" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "goias" Or TxtTime.Text = "Goias" Or TxtTime.Text = "goiás" Or TxtTime.Text = "Goiás" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "botafogo" Or TxtTime.Text = "Botafogo" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "bahia" Or TxtTime.Text = "Bahia" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "athletico" Or TxtTime.Text = "Athletico" Or TxtTime.Text = "athletico pr" Or TxtTime.Text = "Athletico PR" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "ceara" Or TxtTime.Text = "Ceara" Or TxtTime.Text = "ceará" Or TxtTime.Text = "Ceará" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "fortaleza" Or TxtTime.Text = "Fortaleza" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "vasco" Or TxtTime.Text = "Vasco" Or TxtTime.Text = "vasco da gama" Or TxtTime.Text = "Vasco da Gama" Or TxtTime.Text = "Vasco Da Gama" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "fluminense" Or TxtTime.Text = "Fluminense" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "chapecoense" Or TxtTime.Text = "Chapecoense" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "cruzeiro" Or TxtTime.Text = "Cruzeiro" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "csa" Or TxtTime.Text = "CSA" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    ElseIf TxtTime.Text = "avai" Or TxtTime.Text = "Avai" Or TxtTime.Text = "avaí" Or TxtTime.Text = "Avaí" Then
    Time = TxtTime.Text
    FrmEstatisticas.Show
    Else
        MsgBox "Time não Encontrado", vbInformation, "Erro no Time"
    End If
End Sub

Private Sub BtnSair_Click()
    Unload Me
End Sub


