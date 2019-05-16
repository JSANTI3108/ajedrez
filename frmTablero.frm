VERSION 5.00
Begin VB.Form frmTablero 
   Caption         =   "Ajedrez"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   LinkTopic       =   "Form2"
   ScaleHeight     =   8040
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Acciones"
      Height          =   7815
      Left            =   8040
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "Jugar"
         Height          =   855
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   945
         Index           =   5
         Left            =   2880
         Picture         =   "frmTablero.frx":0000
         Top             =   120
         Width           =   945
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   170
         Left            =   6720
         Picture         =   "frmTablero.frx":29BE
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   190
         Left            =   6720
         Picture         =   "frmTablero.frx":5940
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Index           =   61
         Left            =   6720
         Picture         =   "frmTablero.frx":88C2
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   60
         Left            =   0
         Picture         =   "frmTablero.frx":B784
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   186
         Left            =   0
         Picture         =   "frmTablero.frx":E706
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   166
         Left            =   0
         Picture         =   "frmTablero.frx":11688
         Top             =   5880
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   64
         Left            =   6720
         Picture         =   "frmTablero.frx":146CA
         Top             =   6840
         Width           =   1035
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         DragMode        =   1  'Automatic
         Height          =   1020
         Index           =   63
         Left            =   5760
         Picture         =   "frmTablero.frx":17748
         Top             =   6840
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   150
         Left            =   4800
         Picture         =   "frmTablero.frx":1A78A
         Top             =   6840
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Index           =   151
         Left            =   3840
         Picture         =   "frmTablero.frx":1D7CC
         Top             =   6840
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   152
         Left            =   2880
         Picture         =   "frmTablero.frx":208CE
         Top             =   6840
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   153
         Left            =   1920
         Picture         =   "frmTablero.frx":23850
         Top             =   6840
         Width           =   990
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1080
         Index           =   155
         Left            =   0
         Picture         =   "frmTablero.frx":26792
         Top             =   6840
         Width           =   1065
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   154
         Left            =   960
         Picture         =   "frmTablero.frx":29E04
         Top             =   6840
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   156
         Left            =   6720
         Picture         =   "frmTablero.frx":2CE46
         Top             =   5880
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   157
         Left            =   5760
         Picture         =   "frmTablero.frx":2FDC8
         Top             =   5880
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   158
         Left            =   4800
         Picture         =   "frmTablero.frx":32E0A
         Top             =   5880
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   159
         Left            =   3840
         Picture         =   "frmTablero.frx":35D8C
         Top             =   5880
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   160
         Left            =   2880
         Picture         =   "frmTablero.frx":38DCE
         Top             =   5880
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   162
         Left            =   960
         Picture         =   "frmTablero.frx":3BD50
         Top             =   5880
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   161
         Left            =   1920
         Picture         =   "frmTablero.frx":3ECD2
         Top             =   5880
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   171
         Left            =   5760
         Picture         =   "frmTablero.frx":41D14
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   200
         Left            =   4800
         Picture         =   "frmTablero.frx":44C96
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   172
         Left            =   3840
         Picture         =   "frmTablero.frx":47C18
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   173
         Left            =   2880
         Picture         =   "frmTablero.frx":4AB9A
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   174
         Left            =   1920
         Picture         =   "frmTablero.frx":4DB1C
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   176
         Left            =   0
         Picture         =   "frmTablero.frx":50A9E
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   175
         Left            =   960
         Picture         =   "frmTablero.frx":53A20
         Top             =   4920
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   180
         Left            =   6720
         Picture         =   "frmTablero.frx":569A2
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   181
         Left            =   5760
         Picture         =   "frmTablero.frx":59924
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   182
         Left            =   4800
         Picture         =   "frmTablero.frx":5C8A6
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   183
         Left            =   3840
         Picture         =   "frmTablero.frx":5F828
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   184
         Left            =   2880
         Picture         =   "frmTablero.frx":627AA
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   185
         Left            =   960
         Picture         =   "frmTablero.frx":6572C
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   29
         Left            =   1920
         Picture         =   "frmTablero.frx":686AE
         Top             =   3960
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   191
         Left            =   5760
         Picture         =   "frmTablero.frx":6B630
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   27
         Left            =   4800
         Picture         =   "frmTablero.frx":6E5B2
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   193
         Left            =   3840
         Picture         =   "frmTablero.frx":71534
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   194
         Left            =   2880
         Picture         =   "frmTablero.frx":744B6
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   195
         Left            =   1920
         Picture         =   "frmTablero.frx":77438
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   23
         Left            =   0
         Picture         =   "frmTablero.frx":7A3BA
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   196
         Left            =   960
         Picture         =   "frmTablero.frx":7D33C
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   21
         Left            =   6720
         Picture         =   "frmTablero.frx":802BE
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   20
         Left            =   5760
         Picture         =   "frmTablero.frx":83240
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   19
         Left            =   4800
         Picture         =   "frmTablero.frx":861C2
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   18
         Left            =   3840
         Picture         =   "frmTablero.frx":89144
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   17
         Left            =   2880
         Picture         =   "frmTablero.frx":8C0C6
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   16
         Left            =   960
         Picture         =   "frmTablero.frx":8F048
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   15
         Left            =   1920
         Picture         =   "frmTablero.frx":91FCA
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   14
         Left            =   5760
         Picture         =   "frmTablero.frx":94F4C
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Index           =   13
         Left            =   4800
         Picture         =   "frmTablero.frx":9638E
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   12
         Left            =   3840
         Picture         =   "frmTablero.frx":99250
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Index           =   11
         Left            =   2880
         Picture         =   "frmTablero.frx":9A692
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   10
         Left            =   1920
         Picture         =   "frmTablero.frx":9D554
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   9
         Left            =   0
         Picture         =   "frmTablero.frx":9E996
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Index           =   8
         Left            =   960
         Picture         =   "frmTablero.frx":9FDD8
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   7
         Left            =   6720
         Picture         =   "frmTablero.frx":A2C9A
         Top             =   120
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   1
         Left            =   5760
         Picture         =   "frmTablero.frx":A5C1C
         Top             =   120
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   6
         Left            =   4800
         Picture         =   "frmTablero.frx":A8B9E
         Top             =   120
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   3
         Left            =   3840
         Picture         =   "frmTablero.frx":ABBE0
         Top             =   120
         Width           =   1005
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1005
         Index           =   4
         Left            =   0
         Picture         =   "frmTablero.frx":AEC22
         Top             =   120
         Width           =   1035
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1020
         Index           =   2
         Left            =   1920
         Picture         =   "frmTablero.frx":B1CA0
         Top             =   120
         Width           =   1020
      End
      Begin VB.Image img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1050
         Index           =   0
         Left            =   960
         Picture         =   "frmTablero.frx":B2522
         Top             =   120
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmTablero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim piezaMoviendo As Integer

Private Sub Form_Load()
    piezaMoviendo = 0
End Sub

' Ocurre al soltar la pieza
Private Sub img_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    If piezaMoviendo = 63 Then
    
        If Index = 100 Or Index = 200 Then
            'MsgBox ("Drag Drop " & CStr(Index) & " - " & CStr(piezaMoviendo))
            'img(Index).Picture = img(piezaMoviendo).Picture
            img(Index).Picture = img(50).Picture
            
            
            img(63).Picture = img(40).Picture
        End If
    End If
    
    piezaMoviendo = 0
End Sub

' Ocurre al tomar la pieza
Private Sub img_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
     'MsgBox ("Drag Over " & CStr(Index))
    If (piezaMoviendo = 0) Then
        piezaMoviendo = Index
    End If
End Sub
