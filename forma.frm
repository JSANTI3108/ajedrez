VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Juego de Ajedrez"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13290
   LinkTopic       =   "Form2"
   ScaleHeight     =   9045
   ScaleWidth      =   13290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreset 
      Caption         =   "Reinicia Juego"
      Height          =   735
      Left            =   8520
      TabIndex        =   21
      Top             =   8160
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   8280
      TabIndex        =   2
      Top             =   240
      Width           =   4935
      Begin VB.Frame nuevaPieza 
         Caption         =   "Seleccione su pieza"
         Height          =   1455
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Image imagenNueva 
            BorderStyle     =   1  'Fixed Single
            Height          =   1005
            Index           =   3
            Left            =   3360
            Picture         =   "forma.frx":0000
            Top             =   240
            Width           =   1035
         End
         Begin VB.Image imagenNueva 
            BorderStyle     =   1  'Fixed Single
            Height          =   1005
            Index           =   2
            Left            =   2280
            Picture         =   "forma.frx":307E
            Top             =   240
            Width           =   1035
         End
         Begin VB.Image imagenNueva 
            BorderStyle     =   1  'Fixed Single
            Height          =   1005
            Index           =   1
            Left            =   1200
            Picture         =   "forma.frx":60FC
            Top             =   240
            Width           =   1035
         End
         Begin VB.Image imagenNueva 
            BorderStyle     =   1  'Fixed Single
            Height          =   1005
            Index           =   0
            Left            =   120
            Picture         =   "forma.frx":917A
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame FrmJugador 
         Height          =   1575
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   5880
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Frame Frame6 
            Height          =   1095
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1095
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "TURNO"
               Height          =   195
               Left            =   240
               TabIndex        =   15
               Top             =   240
               Width           =   615
            End
            Begin VB.Label turno 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   1
               Left            =   240
               TabIndex        =   14
               Top             =   480
               Width           =   615
            End
         End
         Begin VB.Label jugador 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "jugador1"
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   19
            Top             =   360
            Width           =   1845
         End
         Begin VB.Label color 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BLANCO"
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   18
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label Label6 
            Caption         =   "jugador"
            Height          =   255
            Left            =   1440
            TabIndex        =   17
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "color"
            Height          =   255
            Left            =   1440
            TabIndex        =   16
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Frame FrmJugador 
         Height          =   1575
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Frame Frame5 
            Height          =   1095
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1095
            Begin VB.Label turno 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   0
               Left            =   240
               TabIndex        =   9
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "TURNO"
               Height          =   195
               Left            =   240
               TabIndex        =   7
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Label Label4 
            Caption         =   "color"
            Height          =   255
            Left            =   1440
            TabIndex        =   11
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "jugador"
            Height          =   255
            Left            =   1440
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
         Begin VB.Label color 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BLANCO"
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   8
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label jugador 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "jugador1"
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   5
            Top             =   360
            Width           =   1845
         End
      End
      Begin VB.CommandButton cmdInicio 
         BackColor       =   &H80000000&
         Caption         =   "INICIAR JUEGO"
         Height          =   1095
         Left            =   480
         TabIndex        =   3
         Top             =   4560
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000C&
      FillColor       =   &H00000080&
      ForeColor       =   &H8000000F&
      Height          =   8895
      Left            =   7920
      ScaleHeight     =   8835
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   120
      Width           =   135
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
            Picture         =   "forma.frx":C1F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":F24A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":1229C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":152EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":18340
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":1B392
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":1BC24
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":1EC76
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":21CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":2311A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":2616C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":292BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":2C310
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":2F362
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":32A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":3623E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":39390
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":3C4E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":3F534
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":42586
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":455D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":4862A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":4AFF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":4E04A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":5115C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "forma.frx":541AE
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
         Picture         =   "forma.frx":57200
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   63
         Left            =   5760
         Picture         =   "forma.frx":5A27E
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   62
         Left            =   4800
         Picture         =   "forma.frx":5D2FC
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   61
         Left            =   3840
         Picture         =   "forma.frx":6037A
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   60
         Left            =   2880
         Picture         =   "forma.frx":633F8
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   59
         Left            =   1920
         Picture         =   "forma.frx":66476
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   58
         Left            =   960
         Picture         =   "forma.frx":694F4
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   57
         Left            =   0
         Picture         =   "forma.frx":6C572
         Top             =   7560
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   56
         Left            =   6720
         Picture         =   "forma.frx":6F5F0
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   55
         Left            =   5760
         Picture         =   "forma.frx":7266E
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   54
         Left            =   4800
         Picture         =   "forma.frx":756EC
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   53
         Left            =   3840
         Picture         =   "forma.frx":7876A
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   52
         Left            =   2880
         Picture         =   "forma.frx":7B7E8
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   51
         Left            =   1920
         Picture         =   "forma.frx":7E866
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   50
         Left            =   960
         Picture         =   "forma.frx":818E4
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   49
         Left            =   0
         Picture         =   "forma.frx":84962
         Top             =   6480
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   48
         Left            =   6720
         Picture         =   "forma.frx":879E0
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   47
         Left            =   5760
         Picture         =   "forma.frx":8AA5E
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   46
         Left            =   4800
         Picture         =   "forma.frx":8DADC
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   45
         Left            =   3840
         Picture         =   "forma.frx":90B5A
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   44
         Left            =   2880
         Picture         =   "forma.frx":93BD8
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   43
         Left            =   1920
         Picture         =   "forma.frx":96C56
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   42
         Left            =   960
         Picture         =   "forma.frx":99CD4
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   41
         Left            =   0
         Picture         =   "forma.frx":9CD52
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   40
         Left            =   6720
         Picture         =   "forma.frx":9FDD0
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   39
         Left            =   5760
         Picture         =   "forma.frx":A2E4E
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   38
         Left            =   4800
         Picture         =   "forma.frx":A5ECC
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   37
         Left            =   3840
         Picture         =   "forma.frx":A8F4A
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   36
         Left            =   2880
         Picture         =   "forma.frx":ABFC8
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   35
         Left            =   1920
         Picture         =   "forma.frx":AF046
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   34
         Left            =   960
         Picture         =   "forma.frx":B20C4
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   33
         Left            =   0
         Picture         =   "forma.frx":B5142
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   32
         Left            =   6720
         Picture         =   "forma.frx":B81C0
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   31
         Left            =   5760
         Picture         =   "forma.frx":BB23E
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   30
         Left            =   4800
         Picture         =   "forma.frx":BE2BC
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   29
         Left            =   3840
         Picture         =   "forma.frx":C133A
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   28
         Left            =   2880
         Picture         =   "forma.frx":C43B8
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   27
         Left            =   1920
         Picture         =   "forma.frx":C7436
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   26
         Left            =   960
         Picture         =   "forma.frx":CA4B4
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   25
         Left            =   0
         Picture         =   "forma.frx":CD532
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   17
         Left            =   0
         Picture         =   "forma.frx":D05B0
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   18
         Left            =   960
         Picture         =   "forma.frx":D362E
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   19
         Left            =   1920
         Picture         =   "forma.frx":D66AC
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   20
         Left            =   2880
         Picture         =   "forma.frx":D972A
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   21
         Left            =   3840
         Picture         =   "forma.frx":DC7A8
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   22
         Left            =   4800
         Picture         =   "forma.frx":DF826
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         DragMode        =   1  'Automatic
         Height          =   1005
         Index           =   23
         Left            =   5760
         Picture         =   "forma.frx":E28A4
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   24
         Left            =   6720
         Picture         =   "forma.frx":E5922
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   16
         Left            =   6720
         Picture         =   "forma.frx":E89A0
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   15
         Left            =   5760
         Picture         =   "forma.frx":EBA1E
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   14
         Left            =   4800
         Picture         =   "forma.frx":EEA9C
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   13
         Left            =   3840
         Picture         =   "forma.frx":F1B1A
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   12
         Left            =   2880
         Picture         =   "forma.frx":F4B98
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   11
         Left            =   1920
         Picture         =   "forma.frx":F7C16
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   10
         Left            =   960
         Picture         =   "forma.frx":FAC94
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   9
         Left            =   0
         Picture         =   "forma.frx":FDD12
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   8
         Left            =   6720
         Picture         =   "forma.frx":100D90
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   7
         Left            =   5760
         Picture         =   "forma.frx":103E0E
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   6
         Left            =   4800
         Picture         =   "forma.frx":106E8C
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   5
         Left            =   3840
         Picture         =   "forma.frx":109F0A
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   4
         Left            =   2880
         Picture         =   "forma.frx":10CF88
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   3
         Left            =   1920
         Picture         =   "forma.frx":110006
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   2
         Left            =   960
         Picture         =   "forma.frx":113084
         Top             =   0
         Width           =   1035
      End
      Begin VB.Image imagen 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   1
         Left            =   0
         Picture         =   "forma.frx":116102
         Top             =   0
         Width           =   1035
      End
   End
   Begin VB.Image imagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Index           =   0
      Left            =   13920
      Picture         =   "forma.frx":119180
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
Dim tempX1 As Integer
Dim tempY1 As Integer
Dim tempX2 As Integer
Dim tempY2 As Integer

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
    BLANCO
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

cmdInicio.Visible = False

 
FrmJugador(0).Visible = True
FrmJugador(1).Visible = True
cmdreset.Visible = True


color(0).Caption = "PIEZAS BLANCAS"
color(1).Caption = "PIEZAS NEGRAS"


jugador(0).Caption = "PIEZAS BLANCAS"
jugador(1).Caption = "PIEZAS NEGRAS"

turno(0).Caption = "X"
turno(1).Caption = ""


Call llenarTablero
Call Tablero_Inicial

'Debe ser TRUE solo si es tu TURNO!!!
puedeJugar = True

End Sub

Private Function esTurnoBlancas() As Boolean
    esTurnoBlancas = (turno(0).Caption = "X")
End Function


Private Sub cmdInicio_Click()
Call inicioJuego
End Sub

Private Sub cmdreset_Click()
Call Form_Load
cmdreset.Visible = False
cmdInicio.Visible = True
FrmJugador(0).Visible = False
FrmJugador(1).Visible = False
End Sub

Private Sub Form_Load()

Call Tablero_Previo
puedeJugar = False

End Sub


'Orden Piezas negras en Tablero
Private Sub Tablero_Inicial()

Dim z As Integer
Dim X As Integer
Dim subIndice As Integer

subIndice = 0

For z = 1 To 8
    For X = 1 To 8
        imagen(subIndice + X).DragMode = 1
        imagen(subIndice + X).Picture = listaImagenes.ListImages(funcionPieza(tablero(z, X), subIndice + X)).Picture
    Next
    subIndice = subIndice + 8
Next

End Sub

Private Sub Tablero_Previo()

Dim X As Integer
Dim z As Integer
Dim subIndice As Integer
subIndice = 0
For z = 1 To 8
    For X = 1 To 8
        imagen(subIndice + X).DragMode = 0
        imagen(subIndice + X).Picture = listaImagenes.ListImages(IIf(EsCeldaBlanca(subIndice + X), Piezas.BLANCO, Piezas.NEGRO)).Picture
    Next
    subIndice = subIndice + 8
Next

End Sub

' CAMBIAR NOMBRE  DE FUNCION
Private Function funcionPieza(valorDeLaPieza As Integer, indiceDeLaCelda As Integer) As Integer

    '' LOGICA que indice de pieza es
    If (Abs(valorDeLaPieza) = Pieza.Vacio) Then
        funcionPieza = IIf(EsCeldaBlanca(indiceDeLaCelda), Piezas.BLANCO, Piezas.NEGRO)
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
Private Function validaRey(FilaX As Integer, ColumnaY As Integer, x1 As Integer, y1 As Integer) As Boolean

If tablero(FilaX, ColumnaY) <> 0 Then
 If ((FilaX + 1 >= x1) And (FilaX - 1 <= x1)) And ((ColumnaY + 1 >= y1) And (ColumnaY - 1 <= y1)) Then
    
    If False Then
        ' VALIDAR QUE NO ESTE ATACADA LA CELDA
    End If
    
    If Abs(tablero(x1, y1)) = Pieza.Vacio Then
        validaRey = True
    Else
        If (tablero(FilaX, ColumnaY) < 0 And tablero(x1, y1) > 0) Or (tablero(FilaX, ColumnaY) > 0 And tablero(x1, y1) < 0) Then
            validaRey = True
        End If
    End If
End If
End If

End Function
' FUNCION VALIDA TORRE
 Private Function validaTorre(FilaX As Integer, ColumnaY As Integer, x1 As Integer, y1 As Integer) As Boolean
    Dim fichas As Integer
    Dim z As Integer
    
    fichas = 0
    validaTorre = False
    
    If tablero(FilaX, ColumnaY) < 0 Then
    
        If (FilaX = x1) Or (ColumnaY = y1) Then
            
            If FilaX = x1 Then
            ' se mueve en las filas

                If (ColumnaY < y1) Then
                    For z = (ColumnaY + 1) To (y1 - 1)
                        If (tablero(FilaX, z) <> 0) Then
                            fichas = fichas + 1
                        End If
                    Next
                    
                Else
                    For z = (y1 + 1) To (ColumnaY - 1)
                        If (tablero(FilaX, z) <> 0) Then
                            fichas = fichas + 1
                        End If
                    Next
                End If
                
            Else
            
                If (FilaX < x1) Then
                    For z = (FilaX + 1) To (x1 - 1)
                        If (tablero(z, ColumnaY) <> 0) Then
                            fichas = fichas + 1
                        End If
                    Next
                     
                    
                Else
                   
                    For z = (x1 + 1) To (FilaX - 1)
                        If (tablero(z, ColumnaY) <> 0) Then
                            fichas = fichas + 1
                        End If
                    Next

                End If
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
 Private Function validaCaballo(FilaX As Integer, ColumnaY As Integer, x1 As Integer, y1 As Integer) As Boolean
    Dim fichas As Integer
    Dim z As Integer
    
    fichas = 0
    validaCaballo = False
    
        If FilaX = x1 - 2 Then ' fila de mas abajo
            If ColumnaY = y1 - 1 Or ColumnaY = y1 + 1 Then
                validaCaballo = True
            End If
        ElseIf FilaX = x1 - 1 Then ' fila penultima de abajo
            If ColumnaY = y1 - 2 Or ColumnaY = y1 + 2 Then
                validaCaballo = True
            End If
        ElseIf FilaX = x1 + 1 Then ' fila superior
            If ColumnaY = y1 - 2 Or ColumnaY = y1 + 2 Then
                validaCaballo = True
            End If
        ElseIf FilaX = x1 + 2 Then ' fila de mas arriba
            If ColumnaY = y1 - 1 Or ColumnaY = y1 + 1 Then
                validaCaballo = True
            End If
        End If
   
   
End Function

'FUNCION VALIDA ALFIL
Private Function validaAlfil(FilaX As Integer, ColumnaY As Integer, x1 As Integer, y1 As Integer) As Boolean
    Dim fichas As Integer
    Dim z As Integer
    Dim distanciaX As Integer
    Dim distanciaY As Integer

    Dim multiplicadorX As Integer
    Dim multiplicadorY As Integer

    
    fichas = 0
    validaAlfil = False
    
    distanciaX = Abs(FilaX - x1)
    distanciaY = Abs(ColumnaY - y1)
        
    If (distanciaX = distanciaY) Then
    
        multiplicadorX = IIf(FilaX > x1, -1, 1)
        multiplicadorY = IIf(ColumnaY > y1, -1, 1)

        Debug.Print "Alfil INICIO desde"; FilaX; " Hasta"; ColumnaY
        Debug.Print "Alfil TERMINO desde"; x1; " Hasta"; y1

        For z = 1 To (distanciaX - 1)
            Debug.Print "Fila valida celda ("; (FilaX + (z * multiplicadorX)); ", "; (ColumnaY + (z * multiplicadorY)); ")"
            
            If tablero(FilaX + (z * multiplicadorX), ColumnaY + (z * multiplicadorY)) <> Pieza.Vacio Then
                Exit Function
            End If
        Next

        validaAlfil = True
    End If
    
End Function
'FUNCION VALIDA REINA
Private Function validaReina(FilaX As Integer, ColumnaY As Integer, x1 As Integer, y1 As Integer) As Boolean
    validaReina = False
    Dim mitad As Boolean
    
    mitad = validaAlfil(FilaX, ColumnaY, x1, y1)
    
    If Not mitad Then
        mitad = validaTorre(FilaX, ColumnaY, x1, y1)
        If Not mitad Then
            'No Valido
        Else
            validaReina = True
        End If
    Else
        validaReina = True
   End If
End Function
'FUNCION VALIDA MOVIMIENTO PEON
Private Function validaPeon(FilaX As Integer, ColumnaY As Integer, x1 As Integer, y1 As Integer) As Boolean

    ' MOVIMIENTO NEGRO
    If esNegra(FilaX, ColumnaY) Then
    
        If (ColumnaY = y1) And ((FilaX = (x1 + 1)) Or (7 = (x1 + 2))) Then
            'ESTA VACIA LA POSICION ?
            If Abs(tablero(x1, y1)) = Pieza.Vacio Then
                validaPeon = True
            End If
        ' CAPTURAR CON NEGRO
        ElseIf (y1 = (ColumnaY - 1) Or y1 = (ColumnaY + 1)) And (FilaX = (x1 + 1)) Then
            If tablero(x1, y1) > 0 Then
                validaPeon = True
            End If
        End If
        ' FIN MOVIMIENTO NEGRO
        
      Else
        ' MOVIMIENTO CON BLANCO
         
        If (ColumnaY = y1) And ((FilaX = (x1 - 1)) Or ((x1 - 2) = 2)) Then
        ' esta vacia la posicion ?
            If Abs(tablero(x1, y1)) = Pieza.Vacio Then
                validaPeon = True
            End If
        ' CAPTURAR BLANCO
        ElseIf (y1 = (ColumnaY - 1) Or y1 = (ColumnaY + 1)) And (FilaX = (x1 - 1)) Then
            If tablero(x1, y1) < 0 Then
                validaPeon = True
            End If
        End If
        ' FIN MOVIMIENTO BLANCO
    End If
    
End Function

' Ocurre al tomar la pieza
Private Sub imagen_DragOver(Index As Integer, Source As Control, FilaX As Single, ColumnaY As Single, State As Integer)
    If puedeJugar Then
        If (celdaSeleccionada = 0) Then
            celdaSeleccionada = Index
        End If
    End If
End Sub


' Ocurre al soltar la pieza
Private Sub imagen_DragDrop(posicionNueva As Integer, Source As Control, FilaX As Single, ColumnaY As Single)
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
        
        
    If piezaValor > 0 Then ' que la pieza es blanca
        If Not esTurnoBlancas() Then
            MsgBox ("Turno Negras")
            celdaSeleccionada = 0
            Exit Sub
        End If
    ElseIf piezaValor < 0 Then ' Juega Negras
        If esTurnoBlancas() Then
            MsgBox ("Turno Blancas")
            celdaSeleccionada = 0
            Exit Sub
        End If
    Else ' Juega Negras
        MsgBox ("Movimiento Invalido")
        celdaSeleccionada = 0
        Exit Sub
    End If
         
        nuevaX = Int(posicionNueva / 8) + 1
        nuevaY = Int(posicionNueva Mod 8)
        ' DESBORDE
        If (nuevaY = 0) Then
            nuevaY = 8
            nuevaX = nuevaX - 1
        End If
        nuevaPiezaValor = tablero(nuevaX, nuevaY)
               
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
        ElseIf Abs(piezaValor) = Pieza.Reina Then
        
            If validaReina(piezaX, piezaY, nuevaX, nuevaY) Then
                puedeMover = True
            End If
        ElseIf Abs(piezaValor) = Pieza.Rey Then
             If validaRey(piezaX, piezaY, nuevaX, nuevaY) Then
                puedeMover = True
            End If
        End If
      
    
        If puedeMover Then
            ' DEBE SER FUNCION
            If Abs(piezaValor) = Pieza.Peon Then
                If piezaValor > 0 And (nuevaX = 8 Or nuevaX = 1) Then
                    nuevaPieza.Visible = True
                    imagenNueva(0).Picture = listaImagenes.ListImages(Piezas.Reina_Blanco_Blanco).Picture
                    imagenNueva(1).Picture = listaImagenes.ListImages(Piezas.Alfil_Blanco_Negro).Picture
                    imagenNueva(2).Picture = listaImagenes.ListImages(Piezas.Caballo_Blanco_Blanco).Picture
                    imagenNueva(3).Picture = listaImagenes.ListImages(Piezas.Torre_Blanco_Negro).Picture
                    puedeJugar = False
                ElseIf piezaValor < 0 And (nuevaX = 8 Or nuevaX = 1) Then
                    nuevaPieza.Visible = True
                    imagenNueva(0).Picture = listaImagenes.ListImages(Piezas.Reina_Negro_Blanco).Picture
                    imagenNueva(1).Picture = listaImagenes.ListImages(Piezas.Alfil_Negro_Negro).Picture
                    imagenNueva(2).Picture = listaImagenes.ListImages(Piezas.Caballo_Negro_Blanco).Picture
                    imagenNueva(3).Picture = listaImagenes.ListImages(Piezas.Torre_Negro_Negro).Picture
                    
                    puedeJugar = False
                End If
            ' DEBE SER FUNCION
            End If
            
            If TerminarJugada(piezaX, piezaY, nuevaX, nuevaY) Then
                Call moverPieza(piezaX, piezaY, nuevaX, nuevaY)
                ' repintar
                Call Tablero_Inicial
            End If
        End If
          celdaSeleccionada = 0
      End Sub
 
' Para depurar
Private Sub llenarTableroDebug()

    Dim X  As Integer
    Dim z  As Integer
    
    tablero(1, 6) = Pieza.Alfil
    
    For X = 1 To 8
        tablero(2, X) = Pieza.Peon
    Next
    
    tablero(8, 3) = Pieza.Alfil * -1
End Sub


Private Sub llenarTablero()
    
    Dim FilaX  As Integer
    Dim z  As Integer
    
    tablero(1, 1) = Pieza.Torre
    tablero(1, 2) = Pieza.Caballo
    tablero(1, 3) = Pieza.Alfil
    tablero(1, 4) = Pieza.Rey
    tablero(1, 5) = Pieza.Reina
    tablero(1, 6) = Pieza.Alfil
    tablero(1, 7) = Pieza.Caballo
    tablero(1, 8) = Pieza.Torre
    
    
 For FilaX = 1 To 8
        tablero(2, FilaX) = Pieza.Peon
    Next
    
    For z = 3 To 6
        For FilaX = 1 To 8
            tablero(z, FilaX) = Pieza.Vacio
        Next
    Next
    
 For FilaX = 1 To 8
        tablero(7, FilaX) = Pieza.Peon * -1
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

Private Sub moverPieza(FilaX As Integer, ColumnaY As Integer, x1 As Integer, y1 As Integer)
    If puedeJugar Then
        ' posicionar la pieza en la posicion final
        tablero(x1, y1) = tablero(FilaX, ColumnaY)
        ' eliminar la pieza en la posicion inicial
        tablero(FilaX, ColumnaY) = Pieza.Vacio
        
        If esTurnoBlancas() Then
            turno(0).Caption = ""
            turno(1).Caption = "X"
        Else
            turno(0).Caption = "X"
            turno(1).Caption = ""
        End If
    Else
        tempX1 = FilaX
        tempY1 = ColumnaY
        tempX2 = x1
        tempY2 = y1
    End If
End Sub

Private Function TerminarJugada(FilaX As Integer, ColumnaY As Integer, x1 As Integer, y1 As Integer) As Boolean
    If tablero(x1, y1) = Pieza.Vacio Then
        TerminarJugada = True
    ElseIf (tablero(x1, y1) > 0 And tablero(FilaX, ColumnaY) < 0) Or (tablero(x1, y1) < 0 And tablero(FilaX, ColumnaY) > 0) Then
        TerminarJugada = True
    Else
        TerminarJugada = False
    End If
End Function


Private Sub imagenNueva_Click(Index As Integer)
    puedeJugar = True
    Call moverPieza(tempX1, tempY1, tempX2, tempY2)
    ' obtener la pieza elegida
    
    If (Index = 0) Then
        tablero(tempX2, tempY2) = Pieza.Reina * IIf(tablero(tempX2, tempY2) < 0, -1, 1)
    End If
    If (Index = 1) Then
        tablero(tempX2, tempY2) = Pieza.Alfil * IIf(tablero(tempX2, tempY2) < 0, -1, 1)
    End If
    If (Index = 2) Then
        tablero(tempX2, tempY2) = Pieza.Caballo * IIf(tablero(tempX2, tempY2) < 0, -1, 1)
    End If
    If (Index = 3) Then
        tablero(tempX2, tempY2) = Pieza.Torre * IIf(tablero(tempX2, tempY2) < 0, -1, 1)
    End If
    ' repintar
    Call Tablero_Inicial
    
    nuevaPieza.Visible = False
End Sub

Private Function esBlanca(FilaX As Integer, ColumnaY As Integer) As Boolean
    esBlanca = (tablero(FilaX, ColumnaY) > 0)
End Function

Private Function esNegra(FilaX As Integer, ColumnaY As Integer) As Boolean
    esNegra = (tablero(FilaX, ColumnaY) < 0)
End Function

