VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Menu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desencriptar  Access 97"
   ClientHeight    =   2130
   ClientLeft      =   5340
   ClientTop       =   3450
   ClientWidth     =   4275
   Icon            =   "frm_Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FR_Norma19 
      Height          =   1365
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   4215
      Begin VB.TextBox TXT_Clave 
         Height          =   315
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   930
         Width           =   1995
      End
      Begin VB.CommandButton BTN_ListadoNorma19 
         Caption         =   "Descifrar Clave"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Petición del fichero de entrada y desencriptado de la clave"
         Top             =   450
         Width           =   3945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave :"
         Height          =   195
         Left            =   1500
         TabIndex        =   4
         Top             =   990
         Width           =   495
      End
   End
   Begin VB.CommandButton BTN_Salir 
      Caption         =   "&Salida"
      Height          =   405
      Left            =   3360
      TabIndex        =   0
      Top             =   1620
      Width           =   885
   End
   Begin MSComDlg.CommonDialog EligeFichero 
      Left            =   90
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.mdb"
      DialogTitle     =   "Elija base de datos con clave"
      FilterIndex     =   1
      InitDir         =   "C:\"
      Orientation     =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Kike'99"
      Height          =   195
      Left            =   690
      TabIndex        =   5
      Top             =   1770
      Width           =   525
   End
End
Attribute VB_Name = "frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub BTN_ListadoNorma19_Click()
   Dim IFicheroLibre As Long
   Dim StAux As String
   Dim StTira As String
   Dim ICont As Integer
   Dim B1 As Integer
   Dim B2 As Integer
   Dim R As Integer
   Dim StPass As String
   On Error GoTo Interrupcion
   
   EligeFichero.ShowOpen
   If EligeFichero.FileName = "" Then Exit Sub
         
   If Dir(EligeFichero.FileName) = "" Then
      MsgBox "el fichero no existe", vbInformation, "Desencripta Access"
      Exit Sub
   ElseIf UCase(Right(EligeFichero.FileName, 4)) <> ".MDB" Then
      MsgBox "el fichero no es una base de datos access.", vbInformation, "Desencripta Access"
      Exit Sub
   ElseIf FileLen(EligeFichero.FileName) < 1024 Then
      MsgBox "el fichero no es una base de datos access.", vbInformation, "Desencripta Access"
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   DoEvents: DoEvents
   
   IFicheroLibre = FreeFile()
   Open EligeFichero.FileName For Binary Access Read As #IFicheroLibre
   StAux = Input(66, #IFicheroLibre)
   StTira = Input(13, #IFicheroLibre)
   Close #IFicheroLibre
   StPass = ""
   For ICont = 1 To 13
      Select Case ICont
         Case 1: B1 = CInt(&H86)
         Case 2: B1 = CInt(&HFB)
         Case 3: B1 = CInt(&HEC)
         Case 4: B1 = CInt(&H37)
         Case 5: B1 = CInt(&H5D)
         Case 6: B1 = CInt(&H44)
         Case 7: B1 = CInt(&H9C)
         Case 8: B1 = CInt(&HFA)
         Case 9: B1 = CInt(&HC6)
         Case 10: B1 = CInt(&H5E)
         Case 11: B1 = CInt(&H28)
         Case 12: B1 = CInt(&HE6)
         Case 13: B1 = CInt(&H13)
      End Select
      
      R = B1 Xor Asc(Mid(StTira, ICont, 1))
      StAux = Chr(R)
      StPass = StPass & StAux
   Next
   TXT_Clave = StPass
   Screen.MousePointer = 0
   Exit Sub
Interrupcion:
   Screen.MousePointer = 0
   MsgBox "Error al abrir o manejar el fichero " & EligeFichero.FileName & ".", vbExclamation, "Error con fichero"
   
End Sub

Private Sub BTN_Salir_Click()
   Unload Me
   End
End Sub

Private Sub Form_Load()
   Me.Left = Screen.Width / 2 - Me.Width / 2
   Me.Top = Screen.Height / 2 - Me.Height / 2
   DoEvents: DoEvents
   Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm_Menu = Nothing
End Sub
