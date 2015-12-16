VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recupera desde DUMPs"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10170
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Transforma fichero YOG"
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ver trozo fichero"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Adelante"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   9375
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   840
      Picture         =   "Form1.frx":08CA
      ToolTipText     =   "buscar"
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":12CC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   3720
      Width           =   9855
   End
   Begin VB.Label Label4 
      Caption         =   "*  Si no indicas la tabla recuperas la BD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Top             =   3360
      Width           =   8655
   End
   Begin VB.Label Label4 
      Caption         =   "*  Si marcas ""ver trozo fichero"" ves las primeras lineas del fichero para comprobar las comillas ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   3000
      Width           =   8655
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Tabla"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "BD"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    If Text1.Text = "" Then Exit Sub
    
    If Dir(Text1.Text, vbArchive) = "" Then
        MsgBox "No existe el fichero: " & Text1.Text, vbCritical
        Exit Sub
    End If
    
    If Me.Check1.Value = 1 Then
        Form2.Tag = Text1.Text
        Form2.Show vbModal
        Exit Sub
    End If
    
    If Check2.Value = 1 Then
        If MsgBox("Hay que editar el fichero y añadir al final la linea: /*Data */" & vbCrLf & "¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
    
        If Text2.Text = "" Then
            If Text3.Text = "" Then
                MsgBox "Si no indica la BD tiene que indicar la tabla", vbExclamation
                Exit Sub
            End If
        Else
            If Text3.Text = "" Then
                If MsgBox("Desea recuperar toda la BD?", vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbNo Then Exit Sub
            End If
        End If
    End If
    If Dir(App.Path & "\recu.txt", vbArchive) <> "" Then Kill App.Path & "\recu.txt"
    Screen.MousePointer = vbHourglass
    Me.Command1.Enabled = False
    Me.Refresh
    If Me.Check2.Value = 0 Then
        'Sacamos el trozo de la tabla correspondiente
        PRocesaFichero
    Else
        'Transformamos un fichero generado por el YOG a un fichero
        ProcesaFicheroBULK
    End If
    Me.Command1.Enabled = True
    Screen.MousePointer = vbDefault
    Text1.Text = ""
End Sub



Private Sub PRocesaFichero()
Dim NF As Integer
Dim NSal As Integer
Dim T As Long
Dim Llevo As Long
Dim Fin As Boolean
Dim CadenaTabla As String
Dim EstoyEnBD As Boolean
Dim EstoyEnTabla As Boolean
Dim MiBus As String
Dim Escribe As Boolean
Dim Por As Integer
Dim PorAnt As Integer
Dim T1 As Single
Dim Cad As String

    On Error GoTo Epr

    NF = FreeFile
    Open Text1.Text For Input As #NF
    T = FileLen(Text1.Text)
    
    NSal = FreeFile
    Open App.Path & "\recu.txt" For Output As #NSal
    
    '`ztotalctaconce`
    Label3.Caption = "Buscando"
    Label3.Refresh
    Fin = False
    If Text2.Text = "" Then
        'No hace falta buscar la BD.
        EstoyEnBD = True
        MiBus = "CREATE TABLE " & Text3.Text
    Else
        'SI, que busco la bd
        MiBus = "USE " & Text2.Text
        EstoyEnBD = False
    End If
    
    
    Escribe = False
    T1 = Timer
    PorAnt = 0
    While Not Fin
        
        Line Input #NF, Cad
        Llevo = Llevo + Len(Cad)
        Por = Round((Llevo / T), 2) * 100
        If PorAnt <> Por Then
            
            Label2.Caption = CStr(Por) & "%"
            Label2.Refresh
            PorAnt = Por
            DoEvents
        End If
        
        If Escribe Then
            If InStr(1, Cad, MiBus) > 0 Then
                'FINALIZADA. Cerramos y a casa
                Close #NSal
                Close #NF
                If Text3.Text <> "" Then
                    MsgBox "FIN para datos tabla:" & Text3.Text, vbInformation
                Else
                    Cad = "FIN para datos BD: " & Text2.Text & vbCrLf & vbCrLf & vbCrLf
                    Cad = Cad & "Debera añadir las lineas de comprobacion de claves:"
                    Cad = Cad & vbCrLf & "          SET FOREIGN_KEY_CHECKS=0;"
                    MsgBox Cad, vbInformation
                End If
                Exit Sub
            Else
                Print #NSal, Cad
            End If
            If EOF(NF) Then Fin = True
        Else
            'Estoy situandome todavia
            If InStr(1, Cad, MiBus) > 0 Then
                'OK he encontrado lo que estaba buscando
                If Not EstoyEnBD Then
                    Cad = "Linea: " & Cad & vbCrLf & vbCrLf
                    Cad = Cad & "Es esta?"
                    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                        'AHora buscare CREATE TABLE nomtabla
                        Label3.Caption = "Dentro BD"
                        Label3.Refresh
                        If Text3.Text = "" Then
                            'NO busco ninguna tabla, con lo cual deberia buscar
                            'porbalemnete
                            MiBus = "CREATE DATABASE"
                            Escribe = True
                        Else
                            MiBus = "CREATE TABLE " & Text3.Text
                        End If
                        EstoyEnBD = True
                        EstoyEnTabla = False
                    End If
                Else
                    Label3.Caption = "TABLA"
                    Label3.Refresh
                    Escribe = True
                    MiBus = "CREATE TABLE "
                End If
            End If
        End If
        
        If EOF(NF) Then Fin = True
    Wend
    Close #NF
    Close #NSal
    MsgBox "FIN FICHERO", vbExclamation
    Exit Sub
Epr:
    MsgBox Err.Description
End Sub




Private Sub ProcesaFicheroBULK()
Dim NF As Integer
Dim NSal As Integer
Dim T1 As Long
Dim Llevo As Long
Dim Fin As Boolean
Dim CadenaTabla As String
Dim EstoyEnBD As Boolean
Dim EstoyEnTabla As Boolean
Dim MiBus As String
Dim Escribe As Boolean
Dim Por As Integer
Dim PorAnt As Integer
Dim Cadena As String
Dim PosValue As Integer
Dim Cad As String
Dim t2 As Single

    On Error GoTo Epr

    NF = FreeFile
    Open Text1.Text For Input As #NF
    T1 = FileLen(Text1.Text)
    
    NSal = FreeFile
    Open App.Path & "\recu.txt" For Output As #NSal
    
    '`ztotalctaconce`
    Label3.Caption = "Generando fichero con multiinsert"
    Label3.Refresh
    Fin = False
    
    
    
    Escribe = True
    t2 = Timer
    PorAnt = 0
    Cadena = ""
    MiBus = "/*Data for the table"
    While Not Fin
        
        Line Input #NF, Cad
        Debug.Print Cad
        Llevo = Llevo + Len(Cad)
        Por = Round((Llevo / T1), 2) * 100
        If PorAnt <> Por Then
            
            Label2.Caption = CStr(Por) & "%"
            Label2.Refresh
            PorAnt = Por
            DoEvents
            If Timer - t2 > 5 Then
                DoEvents
                t2 = Timer
            End If
        End If
        
        
        
        If InStr(1, Cad, MiBus) > 0 Then
            If Escribe Then
                
            
                Escribe = False
                Cadena = ""
                Print #NSal, Cad
                MiBus = "/*Table structure"
                EstoyEnTabla = False   'Tengo que buscar el INSERT INTO
            Else
                'Procesamos la cadena
                Cadena = CadenaTabla & Cadena & ";"
                If Len(Cadena) = 1 Then Cadena = ""
                Print #NSal, Cadena
                Print #NSal,
                Print #NSal, Cad
                Cadena = ""
                CadenaTabla = ""
                MiBus = "/*Data for the table"
                Escribe = True
            End If
        
        Else
            If Escribe Then
                Print #NSal, Cad
            Else
        
                'Estoy procesando cadena
                If Len(Cad) = 0 Then
                    
                Else
                    If Not EstoyEnTabla Then
                        PosValue = InStr(1, Cad, ") VALUES (", vbTextCompare)
                        If PosValue = 0 Then
                            If MsgBox("NO se encuentra : ) VALUES ( " & vbCrLf & Cad & vbCrLf & vbCrLf & "¿Continuar?", vbQuestion + vbYesNo) = vbNo Then
                                Fin = True
                            Else
                            
                                Cadena = CadenaTabla & Cadena & ";"
                                If Len(Cadena) = 1 Then Cadena = ""
                                Print #NSal, Cadena
                                Print #NSal,
                                Print #NSal, Cad
                                CadenaTabla = ""
                                MiBus = "/*Data for the table"
                                Escribe = True
                            End If

                        Else
                            CadenaTabla = Mid(Cad, 1, PosValue + 8)
                            PosValue = PosValue + 8
                            
                            Cad = Mid(Cad, PosValue)
                            
                            Cadena = Mid(Cad, 1, Len(Cad) - 1)
                            EstoyEnTabla = True
                        End If
                    Else
                            
                            Cad = Mid(Cad, PosValue)
                            Cad = Mid(Cad, 1, Len(Cad) - 1)
                            If Len(Cadena) > 0 Then Cadena = Cadena & ","
                            Cadena = Cadena & Cad
                            If Len(Cadena) > 400000 Then
                                Cadena = CadenaTabla & Cadena & ";"
                                Print #NSal, Cadena
                                Cadena = ""
                            End If
                            
                    End If
                End If
                
                
                
                
            End If
        End If
        If EOF(NF) Then Fin = True
    Wend
    Close #NF
    Close #NSal
    MsgBox "FIN FICHERO", vbExclamation
    Exit Sub
Epr:
    MsgBox Err.Description
End Sub


Private Sub Image1_Click()
    If Not Me.Command1.Enabled Then Exit Sub
    Me.CommonDialog1.CancelError = False
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileTitle <> "" Then Me.Text1.Text = Me.CommonDialog1.FileName
    
End Sub

