VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver trozo fichero"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   240
      Width           =   9735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Primeravez As Boolean

Private Sub Form_Activate()
    If Primeravez Then
        Primeravez = False
        AbrirFich
    End If
End Sub

Private Sub Form_Load()
    Primeravez = True
End Sub



Private Sub AbrirFich()
Dim NF  As Integer
Dim C As String
Dim L As Integer
Dim Fin As Boolean
    
    L = 0
    NF = FreeFile
    Open Me.Tag For Input As #NF
    While Not Fin
        Line Input #NF, Cad
        Text1.Text = Text1.Text & Cad & vbCrLf
        L = L + 1
        If EOF(NF) Then
            Fin = True
        Else
            If L > 40 Then Fin = True
        End If
    Wend
    Close #NF
End Sub

