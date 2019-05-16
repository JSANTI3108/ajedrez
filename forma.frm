VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Juego de Ajedrez"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15375
   LinkTopic       =   "Form2"
   ScaleHeight     =   9045
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000C&
      FillColor       =   &H00000080&
      ForeColor       =   &H8000000F&
      Height          =   8895
      Left            =   7920
      ScaleHeight     =   8835
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000000&
         Caption         =   "INICIAR JUEGO"
         Height          =   1095
         Left            =   600
         TabIndex        =   2
         Top             =   3120
         Width           =   3855
      End
   End
   Begin MSComctlLib.ImageList listaImagenes 
      Left            =   14040
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":3052
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":60A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":90F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":C148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":F19A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":FA2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":12A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":15AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":16F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":19F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":1D0C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":20118
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":2316A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":2682C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":2A046
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":2D198
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":302EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":3333C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":3638E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":393E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":3C432
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":3EE00
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":41E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":44F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":47FB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   8895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   64
         Left            =   6720
         Picture         =   "forma.frx":4B008
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   63
         Left            =   5760
         Picture         =   "forma.frx":4E086
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   62
         Left            =   4800
         Picture         =   "forma.frx":51104
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   61
         Left            =   3840
         Picture         =   "forma.frx":54182
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   60
         Left            =   2880
         Picture         =   "forma.frx":57200
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   59
         Left            =   1920
         Picture         =   "forma.frx":5A27E
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   58
         Left            =   960
         Picture         =   "forma.frx":5D2FC
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   57
         Left            =   0
         Picture         =   "forma.frx":6037A
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   56
         Left            =   6720
         Picture         =   "forma.frx":633F8
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   55
         Left            =   5760
         Picture         =   "forma.frx":66476
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   54
         Left            =   4800
         Picture         =   "forma.frx":694F4
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   53
         Left            =   3840
         Picture         =   "forma.frx":6C572
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   52
         Left            =   2880
         Picture         =   "forma.frx":6F5F0
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   51
         Left            =   1920
         Picture         =   "forma.frx":7266E
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   50
         Left            =   960
         Picture         =   "forma.frx":756EC
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   49
         Left            =   0
         Picture         =   "forma.frx":7876A
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   48
         Left            =   6720
         Picture         =   "forma.frx":7B7E8
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   47
         Left            =   5760
         Picture         =   "forma.frx":7E866
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   46
         Left            =   4800
         Picture         =   "forma.frx":818E4
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   45
         Left            =   3840
         Picture         =   "forma.frx":84962
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   44
         Left            =   2880
         Picture         =   "forma.frx":879E0
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   43
         Left            =   1920
         Picture         =   "forma.frx":8AA5E
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   42
         Left            =   960
         Picture         =   "forma.frx":8DADC
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   41
         Left            =   0
         Picture         =   "forma.frx":90B5A
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   40
         Left            =   6720
         Picture         =   "forma.frx":93BD8
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   39
         Left            =   5760
         Picture         =   "forma.frx":96C56
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   38
         Left            =   4800
         Picture         =   "forma.frx":99CD4
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   37
         Left            =   3840
         Picture         =   "forma.frx":9CD52
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   36
         Left            =   2880
         Picture         =   "forma.frx":9FDD0
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   35
         Left            =   1920
         Picture         =   "forma.frx":A2E4E
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   34
         Left            =   960
         Picture         =   "forma.frx":A5ECC
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   33
         Left            =   0
         Picture         =   "forma.frx":A8F4A
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   32
         Left            =   6720
         Picture         =   "forma.frx":ABFC8
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   31
         Left            =   5760
         Picture         =   "forma.frx":AF046
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   30
         Left            =   4800
         Picture         =   "forma.frx":B20C4
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   29
         Left            =   3840
         Picture         =   "forma.frx":B5142
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   28
         Left            =   2880
         Picture         =   "forma.frx":B81C0
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   27
         Left            =   1920
         Picture         =   "forma.frx":BB23E
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   26
         Left            =   960
         Picture         =   "forma.frx":BE2BC
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   25
         Left            =   0
         Picture         =   "forma.frx":C133A
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   17
         Left            =   0
         Picture         =   "forma.frx":C43B8
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   18
         Left            =   960
         Picture         =   "forma.frx":C7436
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   19
         Left            =   1920
         Picture         =   "forma.frx":CA4B4
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   20
         Left            =   2880
         Picture         =   "forma.frx":CD532
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   21
         Left            =   3840
         Picture         =   "forma.frx":D05B0
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   22
         Left            =   4800
         Picture         =   "forma.frx":D362E
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         DragMode        =   1  'Automatic
         Height          =   1005
         Index           =   23
         Left            =   5760
         Picture         =   "forma.frx":D66AC
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   24
         Left            =   6720
         Picture         =   "forma.frx":D972A
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   16
         Left            =   6720
         Picture         =   "forma.frx":DC7A8
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   15
         Left            =   5760
         Picture         =   "forma.frx":DF826
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   14
         Left            =   4800
         Picture         =   "forma.frx":E28A4
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   13
         Left            =   3840
         Picture         =   "forma.frx":E5922
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   12
         Left            =   2880
         Picture         =   "forma.frx":E89A0
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   11
         Left            =   1920
         Picture         =   "forma.frx":EBA1E
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   10
         Left            =   960
         Picture         =   "forma.frx":EEA9C
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   9
         Left            =   0
         Picture         =   "forma.frx":F1B1A
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   8
         Left            =   6720
         Picture         =   "forma.frx":F4B98
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   7
         Left            =   5760
         Picture         =   "forma.frx":F7C16
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   6
         Left            =   4800
         Picture         =   "forma.frx":FAC94
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   5
         Left            =   3840
         Picture         =   "forma.frx":FDD12
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   4
         Left            =   2880
         Picture         =   "forma.frx":100D90
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   3
         Left            =   1920
         Picture         =   "forma.frx":103E0E
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   2
         Left            =   960
         Picture         =   "forma.frx":106E8C
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   1
         Left            =   0
         Picture         =   "forma.frx":109F0A
         Top             =   0
         Width           =   1035
      End
   End
   Begin VB.Image imagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Index           =   0
      Left            =   13920
      Picture         =   "forma.frx":10CF88
      Top             =   1920
      Width           =   1035
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' VARIABLES GLOBALES
Dim celdaSeleccionada As Integer
Dim tablero(1 To 8, 1 To 8) As Integer

Dim puedeJugar As Boolean



' Valor de las piezas
Private Enum Pieza
    Vacio
    Rey
    Reina
    Torre
    Caballo
    Alfil
    Peon
End Enum


' Indice de las piezas en el listado de imagenes
Private Enum Piezas
    Nada
    Blanco
    NEGRO
    Alfil_Blanco_Negro
    Alfil_Blanco_Blanco
    Alfil_Negro_Negro
    Alfil_Negro_Blanco
    Peon_Blanco_Negro
    Peon_Blanco_Blanco
    Peon_Negro_Negro
    Peon_Negro_Blanco
    Caballo_Blanco_Blanco
    Caballo_Blanco_Negro
    Caballo_Negro_Blanco
    Caballo_Negro_Negro
    Torre_Blanco_Negro
    Torre_Blanco_Blanco
    Torre_Negro_Blanco
    Torre_Negro_Negro
    Reina_Blanco_Negro
    Reina_Blanco_Blanco
    Reina_Negro_Blanco
    Reina_Negro_Negro
    Rey_Blanco_Blanco
    Rey_Blanco_Negro
    Rey_Negro_Blanco
    Rey_Negro_Negro

End Enum

Private Sub inicioJuego()

Call llenarTablero
Call Tablero_Inicial

'Debe ser TRUE solo si es tu TURNO!!!
puedeJugar = True

End Sub

Private Sub Form_Load()

Call Tablero_Previo
puedeJugar = False

'Dim x As Integer
''Dim resto As Integer
''
''Dim celda As Integer
''
''
''celda = 0
'
'piezaMoviendo = 0
'
'For x = 1 To 64
'
''    imagen(x).Picture = Nothing
''    If Celda = 0 Then
''        imagen(x).Picture = listaImagenes.ListImages(Piezas.Blanco).Picture
''        Celda = 1
''    Else
''        imagen(x).Picture = listaImagenes.ListImages(Piezas.Negro).Picture
''        Celda = 0
''    End If
''    If (x Mod 8) = 0 Then
''        Celda = IIf(Celda = 0, 1, 0)
''    End If
'    imagen(x).DragMode = 1
'    imagen(x).Picture = listaImagenes.ListImages(IIf(EsCeldaBlanca(x), Piezas.Blanco, Piezas.Negro)).Picture
'Next


End Sub

'Inicio del Juego
Public Sub Command1_Click()
	Call inicioJuego
End Sub








'Orden Piezas negras en Tablero
Private Sub Tablero_Inicial()

Dim z As Integer
Dim x As Integer
Dim subIndice As Integer

subIndice = 0

For z = 1 To 8
    For x = 1 To 8
        imagen(subIndice + x).DragMode = 1
        imagen(subIndice + x).Picture = listaImagenes.ListImages(funcionPieza(tablero(z, x), subIndice + x)).Picture
    Next
    subIndice = subIndice + 8
Next

'        imagen(62).Picture = listaImagenes.ListImages(Piezas.Alfil_Negro_Blanco).Picture
'        imagen(59).Picture = listaImagenes.ListImages(Piezas.Alfil_Negro_Negro).Picture
'        imagen(63).Picture = listaImagenes.ListImages(Piezas.Caballo_Negro_Blanco).Picture
'        imagen(58).Picture = listaImagenes.ListImages(Piezas.Caballo_Negro_Negro).Picture
'
'        Dim x As Integer
'        For x = 9 To 16
'            imagen(x).Picture = listaImagenes.ListImages((IIf(x Mod 2 = 1, Piezas.Peon_Blanco_Negro, Piezas.Peon_Blanco_Blanco))).Picture
'        Next
'        For x = 49 To 56
'            imagen(x).Picture = listaImagenes.ListImages((IIf(x Mod 2 = 1, Piezas.Peon_Negro_Blanco, Piezas.Peon_Negro_Negro))).Picture
'        Next
'
''Orden Piezas Blancas en Tablero
'
'        imagen(3).Picture = listaImagenes.ListImages(Piezas.Alfil_Blanco_Blanco).Picture
'        imagen(6).Picture = listaImagenes.ListImages(Piezas.Alfil_Blanco_Negro).Picture

End Sub

Private Sub Tablero_Previo()

Dim x As Integer
Dim subIndice As Integer
subIndice = 0
For z = 1 To 8
    For x = 1 To 8
        imagen(subIndice + x).DragMode = 0
        imagen(subIndice + x).Picture = listaImagenes.ListImages(IIf(EsCeldaBlanca(subIndice + x), Piezas.Blanco, Piezas.NEGRO)).Picture
    Next
    subIndice = subIndice + 8
Next

End Sub

' CAMBIAR NOMBRE  DE FUNCION
Private Function funcionPieza(valorDeLaPieza As Integer, indiceDeLaCelda As Integer) As Integer

    '' LOGICA que indice de pieza es
    If (Abs(valorDeLaPieza) = Pieza.Vacio) Then
        funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Blanco, Piezas.NEGRO)
    ElseIf (Abs(valorDeLaPieza) = Pieza.Alfil) Then
         
        If valorDeLaPieza > 0 Then
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Alfil_Blanco_Blanco, Piezas.Alfil_Blanco_Negro)
        Else
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Alfil_Negro_Blanco, Piezas.Alfil_Negro_Negro)
        End If
        
    ElseIf (Abs(valorDeLaPieza) = Pieza.Peon) Then
    
        If valorDeLaPieza > 0 Then
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Peon_Blanco_Blanco, Piezas.Peon_Blanco_Negro)
        Else
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Peon_Negro_Blanco, Piezas.Peon_Negro_Negro)
        End If
    ElseIf (Abs(valorDeLaPieza) = Pieza.Caballo) Then
         
        If valorDeLaPieza > 0 Then
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Caballo_Blanco_Blanco, Piezas.Caballo_Blanco_Negro)
        Else
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Caballo_Negro_Blanco, Piezas.Caballo_Negro_Negro)
        End If
    ElseIf (Abs(valorDeLaPieza) = Pieza.Torre) Then
         
        If valorDeLaPieza > 0 Then
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Torre_Blanco_Blanco, Piezas.Torre_Blanco_Negro)
        Else
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Torre_Negro_Blanco, Piezas.Torre_Negro_Negro)
        End If
    
    ElseIf (Abs(valorDeLaPieza) = Pieza.Reina) Then
         
        If valorDeLaPieza > 0 Then
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Reina_Blanco_Blanco, Piezas.Reina_Blanco_Negro)
        Else
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Reina_Negro_Blanco, Piezas.Reina_Negro_Negro)
        End If
    ElseIf (Abs(valorDeLaPieza) = Pieza.Rey) Then
         
        If valorDeLaPieza > 0 Then
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Rey_Blanco_Blanco, Piezas.Rey_Blanco_Negro)
        Else
            funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.Rey_Negro_Blanco, Piezas.Rey_Negro_Negro)
        End If
        
'    Else
'        funcionPieza = 1
    End If

End Function


'  CALCULO CELDA BLANCA O NEGRA
Private Function EsCeldaBlanca(celda As Integer) As Boolean
    If celda Mod 2 = 1 Then
        EsCeldaBlanca = True
    Else
        EsCeldaBlanca = False
    End If

    If (celda >= 9 And celda <= 16) Or (celda >= 25 And celda <= 32) Or (celda >= 41 And celda <= 48) Or (celda >= 57 And celda <= 64) Then
        EsCeldaBlanca = Not EsCeldaBlanca
    End If

End Function

' FUNCION VALIDA TORRE
 Private Function validaTorre(x As Integer, y As Integer, x1 As Integer, y1 As Integer) As Boolean
    Dim fichas As Integer
    Dim z As Integer
    
    fichas = 0
    validaTorre = False
    
    If tablero(x, y) < 0 Then
    
        If (x = x1) Or (y = y1) Then
            
            If x = x1 Then
            ' se mueve en las filas

                If (y < y1) Then
                
                    For z = (y + 1) To (y1 - 1)
                        If (tablero(x, z) <> 0) Then
                            fichas = fichas + 1
                        End If
                    Next
                     
                    
                Else
                
                    For z = (y1 + 1) To (y - 1)
                        If (tablero(x, z) <> 0) Then
                            fichas = fichas + 1
                        End If
                    Next
                
                End If
                
            Else
            
                If (x < x1) Then
                
                    For z = (x + 1) To (x1 - 1)
                        If (tablero(z, y) <> 0) Then
                            fichas = fichas + 1
                        End If
                    Next
                     
                    
                Else
                   
                    For z = (x1 + 1) To (x - 1)
                        If (tablero(z, y) <> 0) Then
                            fichas = fichas + 1
                        End If
                    Next

                End If

                '  MsgBox ("desde " & CStr(x) & " al " & CStr(x1))
                ' se mueve en las columnas
            End If
            
            If fichas = 0 Then
                validaTorre = True
            End If
        End If
        
        ' FIN MOVIMIENTO NEGRO
    Else
    
        validaTorre = True
    
    End If
    
   
End Function

'FUNCION VALIDA CABALLO
 Private Function validaCaballo(x As Integer, y As Integer, x1 As Integer, y1 As Integer) As Boolean
    Dim fichas As Integer
    Dim z As Integer
    
    fichas = 0
    validaCaballo = False
    
'          If tablero(x, y) Then

        If ((x1 = x - 1) And (y1 = y - 2)) Or ((y1 = y + 1) And (x1 = x + 2)) Then
            validaCaballo = True
          ElseIf ((x1 = x + 1) And (y1 = y - 2)) Then
            validaCaballo = True
          ElseIf ((x1 = x - 2) And (y1 = y - 1)) Then
            validaCaballo = True
          ElseIf ((x1 = x + 2) And (y1 = y - 1)) Then
             validaCaballo = True
          ElseIf ((x1 = x - 1) And (y1 = y + 2)) Then
            validaCaballo = True
          ElseIf ((x1 = x - 1) And (y1 = y + 2)) Then
            validaCaballo = True
          ElseIf ((x1 = x - 2) And (y1 = y + 1)) Then
            validaCaballo = True
          ElseIf ((x1 = x + 2) And (y1 = y + 1)) Then
             validaCaballo = True
'          ElseIf ((x1 = x - 1) And (y1 = y - 2)) Then
'             validaCaballo = True
             
        End If
   

              ' se mueve en las columnas
'            End If
'
'            If fichas = 0 Then
'                validaCaballo = False
'
'        End If

'         FIN MOVIMIENTO NEGRO
'    Else
'
'        validaCaballo = True
'
'    End If
'    End If

End Function

'FUNCION VALIDA ALFIL

 Private Function validaAlfil(x As Integer, y As Integer, x1 As Integer, y1 As Integer) As Boolean
    Dim fichas As Integer
    Dim z As Integer
    
    fichas = 0
    validaAlfil = False
    
          If tablero(x, y) Then

        If ((x1 = y - 1)) Or (y1 = y + 1) Then
            validaAlfil = True
'                                      If ((x1 = x + 1 And y1 = y - 1)) Or (y1 = y + 1) Then
'                                          validaAlfil = True
          ElseIf ((x1 = y + 1)) Or (y1 = y - 1) Then
            validaAlfil = True
            

        End If
   
End If

End Function
'FUNCION VALIDA MOVIMIENTO PEON
Private Function validaPeon(x As Integer, y As Integer, x1 As Integer, y1 As Integer) As Boolean

    If tablero(x, y) < 0 Then
        ' MOVIMIENTO NEGRO
        If (y = y1) And ((x = (x1 + 1)) Or (7 = (x1 + 2))) Then
            'ESTA VACIA LA POSICION ?
            If Abs(tablero(x1, y1)) = Pieza.Vacio Then
                validaPeon = True
            End If
            
        ' CAPTURAR CON NEGRO
        ElseIf (y1 = (y - 1) Or y1 = (y + 1)) And (x = (x1 + 1)) Then
            If tablero(x1, y1) > 0 Then
                validaPeon = True
            End If
        End If
        ' FIN MOVIMIENTO NEGRO
        
    Else
'        ' MOVIMIENTO CON BLANCO

          If (y = y1) And ((x = (x1 - 1)) Or (2 = (x1 - 2))) Then
            ' esta vacia la posicion ?
            If Abs(tablero(x1, y1)) = Pieza.Vacio Then
                validaPeon = True
            End If
            
        ' CAPTURAR BLANCO
        ElseIf (y1 = (y + 1) Or y1 = (y + 1)) And (x = (x1 + 1)) Then
            If tablero(x1, y1) > 0 Then
                validaPeon = True
            End If
        End If

'        ' FIN MOVIMIENTO BLANCO
    End If
End Function

' Ocurre al tomar la pieza
Private Sub imagen_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    If puedeJugar Then
        If (celdaSeleccionada = 0) Then
            celdaSeleccionada = Index
        End If
    End If
End Sub


' Ocurre al soltar la pieza
Private Sub imagen_DragDrop(posicionNueva As Integer, Source As Control, x As Single, y As Single)

'Call inicioJuego

    If puedeJugar Then
    
        Dim piezaValor As Integer
        Dim piezaX As Integer
        Dim piezaY As Integer
        
        Dim nuevaPiezaValor As Integer
        Dim nuevaX As Integer
        Dim nuevaY As Integer
        
        piezaX = Int(celdaSeleccionada / 8) + 1
        piezaY = Int(celdaSeleccionada Mod 8)
        
        ' DESBORDE
        If (piezaY = 0) Then
            piezaY = 8
            piezaX = piezaX - 1
        End If
        piezaValor = tablero(piezaX, piezaY)
        
        
        nuevaX = Int(posicionNueva / 8) + 1
        nuevaY = Int(posicionNueva Mod 8)
        ' DESBORDE
        If (nuevaY = 0) Then
            nuevaY = 8
            nuevaX = nuevaX - 1
        End If
        nuevaPiezaValor = tablero(nuevaX, nuevaY)
        
        
        'MsgBox ("Drag Over (" & CStr(celdaSeleccionada) & ")(" & CStr(Index) & ")(" & CStr(piezaValor) & ")(" & CStr(piezaX) & ")(" & CStr(piezaY) & ") ")
        
        
        Dim puedeMover As Boolean
        
        puedeMover = False
        
        
        If Abs(piezaValor) = Pieza.Peon Then
        
            If validaPeon(piezaX, piezaY, nuevaX, nuevaY) Then
                puedeMover = True
            End If

        ElseIf Abs(piezaValor) = Pieza.Torre Then
        
            If validaTorre(piezaX, piezaY, nuevaX, nuevaY) Then
                puedeMover = True
            End If
        
       
        ElseIf Abs(piezaValor) = Pieza.Caballo Then
        
            If validaCaballo(piezaX, piezaY, nuevaX, nuevaY) Then
                puedeMover = True
            End If
            
       ElseIf Abs(piezaValor) = Pieza.Alfil Then
        
            If validaAlfil(piezaX, piezaY, nuevaX, nuevaY) Then
                puedeMover = True
            End If
        
    End If
        If puedeMover Then
            If TerminarJugada(piezaX, piezaY, nuevaX, nuevaY) Then
                Call moverPieza(piezaX, piezaY, nuevaX, nuevaY)
                ' repintar
                Call Tablero_Inicial
            End If
        End If
    End If
    
    celdaSeleccionada = 0
    
End Sub





Private Sub llenarTablero()

Dim x  As Integer
Dim z  As Integer

tablero(1, 1) = Pieza.Torre
tablero(1, 2) = Pieza.Caballo
tablero(1, 3) = Pieza.Alfil
tablero(1, 4) = Pieza.Rey
tablero(1, 5) = Pieza.Reina
tablero(1, 6) = Pieza.Alfil
tablero(1, 7) = Pieza.Caballo
tablero(1, 8) = Pieza.Torre


For x = 1 To 8
    tablero(2, x) = Pieza.Peon
Next

For z = 3 To 6
    For x = 1 To 8
        tablero(z, x) = Pieza.Vacio
    Next
Next

For x = 1 To 8
    tablero(7, x) = Pieza.Peon * -1
Next

tablero(8, 1) = Pieza.Torre * -1
tablero(8, 2) = Pieza.Caballo * -1
tablero(8, 3) = Pieza.Alfil * -1
tablero(8, 4) = Pieza.Rey * -1
tablero(8, 5) = Pieza.Reina * -1
tablero(8, 6) = Pieza.Alfil * -1
tablero(8, 7) = Pieza.Caballo * -1
tablero(8, 8) = Pieza.Torre * -1



End Sub


Private Sub moverPieza(x As Integer, y As Integer, x1 As Integer, y1 As Integer)
    ' posicionar la pieza en la posicion final
    tablero(x1, y1) = tablero(x, y)
    ' eliminar la pieza en la posicion inicial
    tablero(x, y) = Pieza.Vacio
End Sub

Private Function TerminarJugada(x As Integer, y As Integer, x1 As Integer, y1 As Integer) As Boolean
    If tablero(x1, y1) = Pieza.Vacio Then
        TerminarJugada = True
    ElseIf (tablero(x1, y1) > 0 And tablero(x, y) < 0) Or (tablero(x1, y1) < 0 And tablero(x, y) > 0) Then
        TerminarJugada = True
    Else
        TerminarJugada = False
    End If
End Function
