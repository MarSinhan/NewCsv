VERSION 5.00
Begin VB.Form frmOpenCsv 
   Caption         =   "CSV•\Ž¦Form"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin VB.ListBox lstCsvEnd 
      Height          =   2760
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   9615
   End
End
Attribute VB_Name = "frmOpenCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim inNum As Integer
    Dim strBuf As String
    Dim cnt As Integer
    
    inNum = FreeFile
    Open App.Path & "\" & csvFileName & ".csv" For Input As #inNum
    
    cnt = 0
    Do Until EOF(1)
        Line Input #inNum, strBuf
        If strBuf = "" Then
            Exit Do
        End If
        lstCsvEnd.AddItem strBuf
        
        If cnt >= 15 Then
            lstCsvEnd.RemoveItem 0
        End If
        cnt = cnt + 1
    Loop
    
    Close #inNum
        
End Sub
