VERSION 5.00
Begin VB.Form FrmEstatisticas 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estatisticas do time"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnFechar 
      Caption         =   "Fechar"
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
      Left            =   3960
      TabIndex        =   28
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   2895
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   9495
      Begin VB.Label LblStManteve 
         BackStyle       =   0  'Transparent
         Caption         =   "Manteve a posição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   83
         Top             =   2325
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label LblStSula 
         BackStyle       =   0  'Transparent
         Caption         =   "Fase de grupos da Copa Sul-Americana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   82
         Top             =   2330
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label LblStQuali 
         BackStyle       =   0  'Transparent
         Caption         =   "Qualificatórias da Copa Libertadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   81
         Top             =   2330
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label LblStLiberta 
         BackStyle       =   0  'Transparent
         Caption         =   "Fase de grupos da Copa Libertadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   62
         Top             =   2325
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label LblStRebaixamento 
         BackStyle       =   0  'Transparent
         Caption         =   "Rebixamaneto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   61
         Top             =   2325
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   3000
         TabIndex        =   60
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label LblRebaixa 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   57
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label LblSula 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   56
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label LblLiber 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   55
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label LblCampeao 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   54
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label LblRebeixamento 
         BackStyle       =   0  'Transparent
         Caption         =   "Rebaixamento:"
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
         Left            =   7680
         TabIndex        =   32
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label LblS 
         BackStyle       =   0  'Transparent
         Caption         =   "Classificação para Sul-Americana:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         TabIndex        =   31
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label LblL 
         BackStyle       =   0  'Transparent
         Caption         =   "Classificação para a Libertadores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   30
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label LblC 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance de ser campeão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Probabilidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   20
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label LblBueno 
      BackStyle       =   0  'Transparent
      Caption         =   "Ricardo Bueno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   84
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LblNeves 
      BackStyle       =   0  'Transparent
      Caption         =   "Thiago Neves"
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
      Left            =   2160
      TabIndex        =   80
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label LblGalhardo 
      BackStyle       =   0  'Transparent
      Caption         =   "Thiago Galhardo"
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
      Left            =   2040
      TabIndex        =   79
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label LblPedro 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedro"
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
      Left            =   2760
      TabIndex        =   78
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblPedrinho 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedrinho"
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
      Left            =   2500
      TabIndex        =   77
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPaolo 
      BackStyle       =   0  'Transparent
      Caption         =   "Paolo Guerrero"
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
      Left            =   2040
      TabIndex        =   76
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LblMexi 
      BackStyle       =   0  'Transparent
      Caption         =   "Mexi Lopez"
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
      Left            =   2330
      TabIndex        =   75
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LblMarcelo 
      BackStyle       =   0  'Transparent
      Caption         =   "Marcelo Cirino"
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
      Left            =   2160
      TabIndex        =   74
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LblKayke 
      BackStyle       =   0  'Transparent
      Caption         =   "Kayke"
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
      Left            =   2640
      TabIndex        =   73
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblGabriel 
      BackStyle       =   0  'Transparent
      Caption         =   "Gabriel"
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
      Left            =   2640
      TabIndex        =   72
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LblEverton 
      BackStyle       =   0  'Transparent
      Caption         =   "Everton"
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
      Left            =   2590
      TabIndex        =   71
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LblEveraldo 
      BackStyle       =   0  'Transparent
      Caption         =   "Everaldo"
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
      Left            =   2520
      TabIndex        =   70
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblSasha 
      BackStyle       =   0  'Transparent
      Caption         =   "Eduardo Sasha"
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
      Left            =   2040
      TabIndex        =   69
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label LblChara 
      BackStyle       =   0  'Transparent
      Caption         =   "Chará"
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
      Left            =   2680
      TabIndex        =   68
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label LblBruno 
      BackStyle       =   0  'Transparent
      Caption         =   "Bruno Henrique"
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
      Left            =   2085
      TabIndex        =   67
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label LblArtur 
      BackStyle       =   0  'Transparent
      Caption         =   "Artur"
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
      Left            =   2805
      TabIndex        =   66
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LblAndreLuis 
      BackStyle       =   0  'Transparent
      Caption         =   "Andre Luis"
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
      Left            =   2400
      TabIndex        =   65
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblAlexandre 
      BackStyle       =   0  'Transparent
      Caption         =   "Alexandre Pato"
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
      Left            =   2115
      TabIndex        =   64
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LblAlexSantana 
      BackStyle       =   0  'Transparent
      Caption         =   "Alex Santana"
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
      Left            =   2270
      TabIndex        =   63
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   3960
      Picture         =   "Estatisticas.frx":0000
      Top             =   120
      Width           =   1725
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Posi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   59
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label LblPosiJ 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   58
      Top             =   4470
      Width           =   1215
   End
   Begin VB.Label LblAvai 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4470
      TabIndex        =   53
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label LblCSA 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4460
      TabIndex        =   52
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label LblCruzeiro 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   51
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label LblChapeco 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   50
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label LblFluminense 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      TabIndex        =   49
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label LblVasco 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   48
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label LblFortaleza 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   47
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label LblCeara 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4050
      TabIndex        =   46
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label LblPR 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3790
      TabIndex        =   45
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label LblGremio 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4220
      TabIndex        =   44
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label LblCorinthians 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   43
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label LblSP 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4040
      TabIndex        =   42
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label LblBahia 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4350
      TabIndex        =   41
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label LblBotafogo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4130
      TabIndex        =   40
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label LblGoias 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4340
      TabIndex        =   39
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LblAtletico 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4240
      TabIndex        =   38
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label LblInternacional 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3870
      TabIndex        =   37
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label LblSantos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4290
      TabIndex        =   36
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label LblPosicao 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7560
      TabIndex        =   35
      Top             =   4395
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Posição do time na tabela:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   34
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label LblPalmeiras 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4050
      TabIndex        =   33
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label LblGols 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4590
      TabIndex        =   27
      Top             =   4005
      Width           =   495
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Gols:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   26
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label LblCV 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   8640
      TabIndex        =   25
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label LblCA 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Cartões vermelhos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Carotões amarelos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   22
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Principal artilheiro do time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   21
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1170
      Left            =   2280
      Picture         =   "Estatisticas.frx":793B
      Top             =   3840
      Width           =   1620
   End
   Begin VB.Label LblAproveitamento 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label LblSG 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label LblGC 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label LblGF 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8640
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label LblE 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label LblD 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label LblV 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label LblPartidas 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label LblPontos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Aproveitamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo de gols:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gols Contra:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Gols a favor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Empates:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Derrotas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Vitórias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Partidas Jogadas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label LblFlamengo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4050
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pontos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "FrmEstatisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If Time = "palmeiras" Or Time = "Palmeiras" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Palmeiras.jpg")
        LblPalmeiras.Caption = "Palmeiras"
        LblPontos.Caption = 25
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 8
        LblV.ForeColor = &HFF00&
        LblD.Caption = 0
        LblD.ForeColor = &HFF00&
        LblE.Caption = 1
        LblGF.Caption = 18
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 2
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 16
        LblSG.ForeColor = &HFF00&
        LblAproveitamento.Caption = "92,6%"
        LblCA.Caption = 5
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\BrunoHenrique.jpg")
        LblBruno.Visible = True
        LblGols.Caption = 4
        LblPosiJ.Caption = "Volante"
        LblPosicao.Caption = "Primeiro Lugar"
        LblCampeao.Caption = "48,5%"
        LblLiber.Caption = "92,9%"
        LblSula.Caption = "6,4%"
        LblRebaixa.Caption = "0,053%"
        LblStLiberta.Visible = True
    ElseIf Time = "santos" Or Time = "Santos" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Santos.jpg")
        LblSantos.Caption = "Santos"
        LblPontos.Caption = 20
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 6
        LblV.ForeColor = &HFF00&
        LblD.Caption = 1
        LblD.ForeColor = &HFF&
        LblE.Caption = 2
        LblGF.Caption = 12
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 7
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 5
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "74,1%"
        LblCA.Caption = 4
        LblCV.Caption = 1
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\EduardoSasha.jpg")
        LblSasha.Visible = True
        LblGols.Caption = 4
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Segundo Lugar"
        LblCampeao.Caption = "14,8%"
        LblLiber.Caption = "73,0%"
        LblSula.Caption = "22,1%"
        LblRebaixa.Caption = "0,77%"
        LblStLiberta.Visible = True
    ElseIf Time = "flamengo" Or Time = "Flamengo" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Flamengo.jpg")
        LblFlamengo.Caption = "Flamengo"
        LblPontos.Caption = 17
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 5
        LblV.ForeColor = &HFF00&
        LblD.Caption = 2
        LblD.ForeColor = &HFF&
        LblE.Caption = 2
        LblGF.Caption = 15
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 9
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 6
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "63%"
        LblCA.Caption = 4
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\Gabriel.jpg")
        LblGabriel.Visible = True
        LblGols.Caption = 5
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Terceiro Lugar"
        LblCampeao.Caption = "5,7%"
        LblLiber.Caption = "49,8%"
        LblSula.Caption = "35,6%"
        LblRebaixa.Caption = "3,3%"
        LblStLiberta.Visible = True
    ElseIf Time = "internacional" Or Time = "Internacional" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Internacional.jpg")
        LblInternacional.Caption = "Internacional"
        LblPontos.Caption = 16
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 5
        LblV.ForeColor = &HFF00&
        LblD.Caption = 3
        LblD.ForeColor = &HFF&
        LblE.Caption = 1
        LblGF.Caption = 13
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 8
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 5
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "59,3%"
        LblCA.Caption = 0
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\PaoloGuerrero.jpg")
        lblPaolo.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Quarto Lugar"
        LblCampeao.Caption = "5,0%"
        LblLiber.Caption = "47%"
        LblSula.Caption = "35,7%"
        LblRebaixa.Caption = "4,5%"
        LblStLiberta.Visible = True
    ElseIf Time = "atletico" Or Time = "Atlético" Or Time = "Atletico" Or Time = "Atlético Mg" Or Time = "atlético" Or Time = "atletico mg" Or Time = "atlético mg" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Atletico.jpg")
        LblAtletico.Caption = "Atlético"
        LblPontos.Caption = 16
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 5
        LblV.ForeColor = &HFF00&
        LblD.Caption = 3
        LblD.ForeColor = &HFF&
        LblE.Caption = 1
        LblGF.Caption = 14
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 11
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 3
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "59,3%"
        LblCA.Caption = 0
        LblCV.Caption = 1
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\Chara.jpg")
        LblChara.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Quinto Lugar"
        LblCampeao.Caption = "4,3%"
        LblLiber.Caption = "43,9%"
        LblSula.Caption = "37%"
        LblRebaixa.Caption = "5,2%"
        LblStQuali.Visible = True
    ElseIf Time = "goias" Or Time = "Goias" Or Time = "góias" Or Time = "Góias" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Goias.jpg")
        LblGoias.Caption = "Goiás"
        LblPontos.Caption = 15
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 8
        LblV.Caption = 5
        LblV.ForeColor = &HFF00&
        LblD.Caption = 3
        LblD.ForeColor = &HFF&
        LblE.Caption = 0
        LblGF.Caption = 11
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 8
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 3
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "62,5%"
        LblCA.Caption = 0
        LblCV.Caption = 1
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\Kayke.jpg")
        LblKayke.Visible = True
        LblGols.Caption = 4
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Sexto Lugar"
        LblCampeao.Caption = "5,5%"
        LblLiber.Caption = "46,1%"
        LblSula.Caption = "35,4%"
        LblRebaixa.Caption = "5,1%"
        LblStQuali.Visible = True
    ElseIf Time = "botafogo" Or Time = "Botafogo" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Botafogo.jpg")
        LblBotafogo.Caption = "Botafogo"
        LblPontos.Caption = 15
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 5
        LblV.ForeColor = &HFF00&
        LblD.Caption = 4
        LblD.ForeColor = &HFF&
        LblE.Caption = 0
        LblGF.Caption = 8
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 8
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 0
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "55,6%"
        LblCA.Caption = 0
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\AlexSantana.jpg")
        LblAlexSantana.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Volante"
        LblPosicao.Caption = "Sétimo Lugar"
        LblCampeao.Caption = "3,3%"
        LblLiber.Caption = "37,3%"
        LblSula.Caption = "37,8%"
        LblRebaixa.Caption = "7,8%"
        LblStSula.Visible = True
    ElseIf Time = "bahia" Or Time = "Bahia" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Bahia.jpg")
        LblBahia.Caption = "Bahia"
        LblPontos.Caption = 14
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 4
        LblV.ForeColor = &HFF00&
        LblD.Caption = 3
        LblD.ForeColor = &HFF&
        LblE.Caption = 2
        LblGF.Caption = 11
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 11
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 0
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "51,9%"
        LblCA.Caption = 0
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\Artur.jpg")
        LblArtur.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Oitavo Lugar"
        LblCampeao.Caption = "3,4%"
        LblLiber.Caption = "39,3%"
        LblSula.Caption = "38,6%"
        LblRebaixa.Caption = "6,3%"
        LblStSula.Visible = True
    ElseIf Time = "sao paulo" Or Time = "Sao paulo" Or Time = "São Paulo" Or Time = "são paulo" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\SaoPaulo.jpg")
        LblSP.Caption = "São Paulo"
        LblPontos.Caption = 14
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 3
        LblV.ForeColor = &HFF00&
        LblD.Caption = 1
        LblD.ForeColor = &HFF&
        LblE.Caption = 5
        LblGF.Caption = 8
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 5
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 3
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "51,9%"
        LblCA.Caption = 0
        LblCV.Caption = 3
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\AlexandrePato.jpg")
        LblAlexandre.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Nono Lugar"
        LblCampeao.Caption = "1,9%"
        LblLiber.Caption = "29%"
        LblSula.Caption = "40,6%"
        LblRebaixa.Caption = "9,8%"
        LblStSula.Visible = True
    ElseIf Time = "corinthians" Or Time = "Corinthians" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Corinthians.jpg")
        LblCorinthians.Caption = "Corinthians"
        LblPontos.Caption = 12
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 8
        LblV.Caption = 3
        LblV.ForeColor = &HFF00&
        LblD.Caption = 2
        LblD.ForeColor = &HFF&
        LblE.Caption = 3
        LblGF.Caption = 7
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 5
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 2
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "50%"
        LblCA.Caption = 0
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\Pedrinho.jpg")
        LblPedrinho.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Meia Central"
        LblPosicao.Caption = "Décimo Lugar"
        LblCampeao.Caption = "2,8%"
        LblLiber.Caption = "33,6%"
        LblSula.Caption = "38,7%"
        LblRebaixa.Caption = "9,1%"
        LblStSula.Visible = True
    ElseIf Time = "gremio" Or Time = "Gremio" Or Time = "Grêmio" Or Time = "grêmio" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Corinthians.jpg")
        LblGremio.Caption = "Grêmio"
        LblPontos.Caption = 11
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 3
        LblV.ForeColor = &HFF00&
        LblD.Caption = 4
        LblD.ForeColor = &HFF&
        LblE.Caption = 2
        LblGF.Caption = 10
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 11
        LblGC.ForeColor = &HFF&
        LblSG.Caption = -1
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "40,7%"
        LblCA.Caption = 0
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\Everton.jpg")
        LblEverton.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Décimo primeiro Lugar"
        LblCampeao.Caption = "1,3%"
        LblLiber.Caption = "21,8%"
        LblSula.Caption = "36,9%"
        LblRebaixa.Caption = "16,9%"
        LblStSula.Visible = True
    ElseIf Time = "athletico" Or Time = "Athletico" Or Time = "athletico pr" Or Time = "Athletico PR" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\AthleticoPR.jpg")
        LblPR.Caption = "Athletico-PR"
        LblPontos.Caption = 10
        LblPontos.ForeColor = &HFF&
        LblPartidas.Caption = 9
        LblV.Caption = 3
        LblV.ForeColor = &HFF00&
        LblD.Caption = 5
        LblD.ForeColor = &HFF&
        LblE.Caption = 1
        LblGF.Caption = 13
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 12
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 1
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "37%"
        LblCA.Caption = 0
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\MarceloCirino.jpg")
        LblMarcelo.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Décimo Segundo Lugar"
        LblCampeao.Caption = "0,83%"
        LblLiber.Caption = "17,1%"
        LblSula.Caption = "34,2%"
        LblRebaixa.Caption = "22,3%"
        LblStSula.Visible = True
    ElseIf Time = "ceara" Or Time = "Ceará" Or Time = "ceara Sc" Or Time = "Ceara Sc" Or Time = "ceará sc" Or Time = "Ceará SC" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\CearaSC.jpg")
        LblCeara.Caption = "Ceará"
        LblPontos.Caption = 10
        LblPontos.ForeColor = &HFF&
        LblPartidas.Caption = 9
        LblV.Caption = 3
        LblV.ForeColor = &HFF00&
        LblD.Caption = 5
        LblD.ForeColor = &HFF&
        LblE.Caption = 1
        LblGF.Caption = 10
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 9
        LblGC.ForeColor = &HFF&
        LblSG.Caption = 1
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "37%"
        LblCA.Caption = 0
        LblCV.Caption = 0
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\ThiagoGalhardo.jpg")
        LblGalhardo.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Meia Central"
        LblPosicao.Caption = "Décimo Terceiro Lugar"
        LblCampeao.Caption = "0,61%"
        LblLiber.Caption = "13,5%"
        LblSula.Caption = "31,4%"
        LblRebaixa.Caption = "27,1%"
        LblStManteve.Visible = True
    ElseIf Time = "fortaleza" Or Time = "Fortaleza" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Fortaleza.jpg")
        LblFortaleza.Caption = "Fortaleza"
        LblPontos.Caption = 10
        LblPontos.ForeColor = &HFF00&
        LblPartidas.Caption = 9
        LblV.Caption = 3
        LblV.ForeColor = &HFF00&
        LblD.Caption = 5
        LblD.ForeColor = &HFF&
        LblE.Caption = 1
        LblGF.Caption = 8
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 13
        LblGC.ForeColor = &HFF&
        LblSG.Caption = -5
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "37%"
        LblCA.Caption = 0
        LblCV.Caption = 2
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\AndreLuis.jpg")
        LblAndreLuis.Visible = True
        LblGols.Caption = 2
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Décimo Quarto Lugar"
        LblCampeao.Caption = "0,64%"
        LblLiber.Caption = "14,5%"
        LblSula.Caption = "32,7%"
        LblRebaixa.Caption = "25%"
        LblStManteve.Visible = True
    ElseIf Time = "vasco" Or Time = "Vasco" Or Time = "vasco da gama" Or Time = "Vasco da Gama" Or Time = "Vasco Da Gama" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Vasco.jpg")
        LblVasco.Caption = "Vasco da Gama"
        LblPontos.Caption = 9
        LblPontos.ForeColor = &HFF&
        LblPartidas.Caption = 9
        LblV.Caption = 2
        LblV.ForeColor = &HFF00&
        LblD.Caption = 4
        LblD.ForeColor = &HFF&
        LblE.Caption = 3
        LblGF.Caption = 8
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 14
        LblGC.ForeColor = &HFF&
        LblSG.Caption = -6
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "33,3%"
        LblCA.Caption = 0
        LblCV.Caption = 1
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\MexiLopez.jpg")
        LblMexi.Visible = True
        LblGols.Caption = 2
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Décimo Quinto Lugar"
        LblCampeao.Caption = "0,34%"
        LblLiber.Caption = "9.7%"
        LblSula.Caption = "9,7%"
        LblRebaixa.Caption = "28,4%"
        LblStManteve.Visible = True
    ElseIf Time = "fluminense" Or Time = "Fluminense" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Fluminense.jpg")
        LblFluminense.Caption = "Fluminense"
        LblPontos.Caption = 8
        LblPontos.ForeColor = &HFF&
        LblPartidas.Caption = 9
        LblV.Caption = 2
        LblV.ForeColor = &HFF00&
        LblD.Caption = 5
        LblD.ForeColor = &HFF&
        LblE.Caption = 2
        LblGF.Caption = 13
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 16
        LblGC.ForeColor = &HFF&
        LblSG.Caption = -3
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "29,6%"
        LblCA.Caption = 5
        LblCV.Caption = 3
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\Pedro.jpg")
        LblPedro.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Décimo Sexto Lugar"
        LblCampeao.Caption = "0,31%"
        LblLiber.Caption = "9,1%"
        LblSula.Caption = "27,3%"
        LblRebaixa.Caption = "34,5%"
        LblStManteve.Visible = True
    ElseIf Time = "chapecoense" Or Time = "Chapecoense" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Chapecoense.jpg")
        LblChapeco.Caption = "Chapecoense"
        LblPontos.Caption = 8
        LblPontos.ForeColor = &HFF&
        LblPartidas.Caption = 9
        LblV.Caption = 2
        LblV.ForeColor = &HFF00&
        LblD.Caption = 5
        LblD.ForeColor = &HFF&
        LblE.Caption = 2
        LblGF.Caption = 10
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 14
        LblGC.ForeColor = &HFF&
        LblSG.Caption = -4
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "29,6%"
        LblCA.Caption = 0
        LblCV.Caption = 1
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\Everaldo.jpg")
        LblEveraldo.Visible = True
        LblGols.Caption = 5
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Décimo Sétimo Lugar"
        LblCampeao.Caption = "0,26%"
        LblLiber.Caption = "8,2%"
        LblSula.Caption = "25,6%"
        LblRebaixa.Caption = "37,5%"
        LblStRebaixamento.Visible = True
    ElseIf Time = "cruzeiro" Or Time = "Cruzeiro" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Cruzeiro.jpg")
        LblCruzeiro.Caption = "Cruzeiro"
        LblPontos.Caption = 8
        LblPontos.ForeColor = &HFF&
        LblPartidas.Caption = 9
        LblV.Caption = 2
        LblV.ForeColor = &HFF00&
        LblD.Caption = 5
        LblD.ForeColor = &HFF&
        LblE.Caption = 2
        LblGF.Caption = 9
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 16
        LblGC.ForeColor = &HFF&
        LblSG.Caption = -7
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "29,6%"
        LblCA.Caption = 0
        LblCV.Caption = 2
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\ThiagoNeves.jpg")
        LblNeves.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Meia Central"
        LblPosicao.Caption = "Décimo Oitavo Lugar"
        LblCampeao.Caption = "0,25%"
        LblLiber.Caption = "7,8%"
        LblSula.Caption = "25,4%"
        LblRebaixa.Caption = "37,7%"
        LblStRebaixamento.Visible = True
    ElseIf Time = "csa" Or Time = "CSA" Or Time = "Csa" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\CSA.jpg")
        LblCSA.Caption = "CSA"
        LblPontos.Caption = 6
        LblPontos.ForeColor = &HFF&
        LblPartidas.Caption = 9
        LblV.Caption = 1
        LblV.ForeColor = &HFF00&
        LblD.Caption = 5
        LblD.ForeColor = &HFF&
        LblE.Caption = 3
        LblGF.Caption = 3
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 15
        LblGC.ForeColor = &HFF&
        LblSG.Caption = -12
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "22,2%"
        LblCA.Caption = 0
        LblCV.Caption = 1
        Image2.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Jogadores\RicardoBueno.jpg")
        LblBueno.Visible = True
        LblGols.Caption = 3
        LblPosiJ.Caption = "Atacante"
        LblPosicao.Caption = "Décimo Nono Lugar"
        LblCampeao.Caption = "0,083%"
        LblLiber.Caption = "4%"
        LblSula.Caption = "17,8%"
        LblRebaixa.Caption = "52,2%"
        LblStRebeixamento.Visible = True
    ElseIf Time = "avai" Or Time = "Avai" Or Time = "avaí" Or Time = "Avaí" Then
        Image1.Picture = LoadPicture("C:\Users\Alex\Documents\Visual Basic\Exercícios\Times\Avai.jpg")
        LblAvai.Caption = "Avaí"
        LblPontos.Caption = 4
        LblPontos.ForeColor = &HFF&
        LblPartidas.Caption = 9
        LblV.Caption = 0
        LblV.ForeColor = &HFF00&
        LblD.Caption = 5
        LblD.ForeColor = &HFF&
        LblE.Caption = 4
        LblGF.Caption = 4
        LblGF.ForeColor = &HFF00&
        LblGC.Caption = 11
        LblGC.ForeColor = &HFF&
        LblSG.Caption = -7
        LblSG.ForeColor = &HFF&
        LblAproveitamento.Caption = "14,8%"
        LblCA.Caption = 4
        LblCV.Caption = 0
        Image2.Visible = False
        LblGols.Caption = ""
        LblPosiJ.Caption = ""
        LblPosicao.Caption = "Vigésimo Lugar"
        LblCampeao.Caption = "0,035%"
        LblLiber.Caption = "2,2%"
        LblSula.Caption = "12,6%"
        LblRebaixa.Caption = "62,5%"
        LblStRebeixamento.Visible = True
    End If
End Sub
Private Sub BtnFechar_Click()
    Unload Me
End Sub

