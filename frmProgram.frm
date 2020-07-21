VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmProgram 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abecedario por Voz  v1.0"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmProgram.frx":0000
   LinkTopic       =   "c"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAcercade 
      BackColor       =   &H8000000C&
      Caption         =   "&Acerca"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   27
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   9
      ToolTipText     =   "Pizarron de letras"
      Top             =   2770
      Width           =   855
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   190
         TabIndex        =   10
         ToolTipText     =   "Pizarron de letras"
         Top             =   50
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdAbc 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   50
      Picture         =   "frmProgram.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Abecedario Completo"
      Top             =   50
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H8000000C&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   960
      ScaleHeight     =   4575
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   0
      Width           =   615
      Begin VB.PictureBox picabc 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   2
         Top             =   0
         Width           =   615
         Begin VB.CommandButton cmdletra 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   0
            Width           =   615
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4575
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   480
      X2              =   600
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   480
      X2              =   360
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   465
      X2              =   480
      Y1              =   2280
      Y2              =   2640
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Letra Indicada"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   840
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Abc for Windows"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   6375
      Left            =   1920
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8281
      _cy             =   11245
   End
End
Attribute VB_Name = "frmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*
'/* código fuente creado
'/* by Martin Abecedario
'/* gráfico mediantre botones
'/* código creado  para youtube
'/*

Private Sub cmdAbc_Click()
reproducir_Audio "abc.mp3"
End Sub

Private Sub cmdAcercade_Click()
frmAcercade.Show 1
End Sub

Private Sub Form_Load()
abc ' cargar abecedario
End Sub

Private Sub abc()
Dim Virtual_abc, _
Virtualx_abc _
, i, v, c As Byte         ' crea memoria para los controles de 0 a 255 digitos
                          ' numericos
Dim col As New Collection ' crea una colecion de elementos en memoria

For Virtualx_abc = 1 To 26 ' recorre los botones para crear el abcedario

Virtual_abc = Virtual_abc + 1             'incrementa de 1 en 1 la varialbe
Load cmdletra(Virtual_abc)                'carga los controles
     cmdletra(Virtual_abc).Visible = True ' visibiliza los controles
     cmdletra(Virtual_abc).Top = 511 * Virtual_abc 'crea el rango de lienzo
     picabc.Height = 2280 * Virtual_abc   'establece la altura maxima
     With VScroll1                        'del lenzo
     .Min = 0
     .Max = -Virtual_abc + 3
     End With
     Next
                     '
    For i = 65 To 90 ' recorre caracteres de 65 a 90 en formato numerico
     col.Add Chr(i)  ' tranforma los caracteres de formato numerico a letras
    
    If i = 78 Then '
      col.Add "Ñ"  ' cuando el char acill reprecenta un caracter i=78
    End If         ' indefinido carga la letra ñ para evitar error
  
   Next i          'termina el recorrido totoal

   For v = 0 To 26                           '
    c = v + 1                                ' icrementa el tipo de letra
    cmdletra(v).Caption = col.Item(c)        ' optiene el valor espesifico
    cmdletra(v).BackColor = pintarABC()      ' pinta el abcdario de colores
Next v                                       ' aleatoriamente
End Sub

Private Sub VScroll1_Change()                '
picabc.Top = VScroll1.Value * 411            ' desplaza el abecedario
End Sub                                      ' al cambiar
                                             '
Private Sub VScroll1_Scroll()                ' desplaza el avezedario
VScroll1_Change                              ' al desplazar el control
End Sub                                      '

Private Function pintarABC()                 'genera lores aleatoriamente
pintarABC = RGB(Rnd(1) * 255, _
Rnd(1) * 255, Rnd(1) * 255)
End Function

Private Sub cmdok_Click()                    '
Unload Me                                    'descargar el formulario
End Sub                                      '

Private Sub reproducir_Audio(ByVal X_url As String)
On Error GoTo nose
    With wmp
        .URL = X_url
        .Controls.play
    End With
nose:
End Sub
Private Sub cmdletra_Click(Index As Integer) 'muestra un mensaje en pantalla
reproducir_Audio cmdletra.Item(Index).Caption & ".mp3"
Picture1.BackColor = cmdletra(Index).BackColor
Label3.Caption = cmdletra(Index).Caption
End Sub

