VERSION 5.00
Begin VB.Form frmAddItm 
   Caption         =   "���Y�i�ǉ�Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.CommandButton getYosanItemBtn 
      Caption         =   "Get!Add!"
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox valText 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   2  '��
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox thText 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   2  '��
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cmbItem 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "�w�����i"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "�ǉ���"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�\�Z�p���i���"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddItm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub getYosanItemBtn_Click()
    Dim inputFileNumber As Long
    Dim outputFileNumber As Long
    Dim non0Wat1Cafe2flg As String
    Dim atmarkItemName As String
    Dim strBuf As String
    Dim cnt As Long
    Dim isThereYosanRecord As Boolean
    
    If "" = Trim(cmbItem.Text) Then
        Exit Sub
    ElseIf "" = Trim(thText.Text) Then
        Exit Sub
    ElseIf "" = Dir(App.Path & "\" & csvFileName & ".csv") Then
        MsgBox "���̃t�@�C������csv�t�@�C����������܂���B", vbCritical + vbOKOnly, "�t�@�C�����m�F"
        Exit Sub
    End If
    
    
    non0Wat1Cafe2flg = ""
    For cnt = 1 To UBound(isAtmarkFlg)
        atmarkItemName = isAtmarkFlg(cnt)
        
        If atmarkItemName = "" Then
            Exit For
        End If
        If Left(cmbItem.Text, Len(atmarkItemName)) = atmarkItemName Then
            non0Wat1Cafe2flg = atmarkItemName
            Exit For
        End If
    Next cnt
    
    If non0Wat1Cafe2flg = "" Then
        MsgBox "���̑��̂̍݌ɂ������ł��B"
        Exit Sub
    End If
    
    OpenInOutFile inputFileNumber, outputFileNumber
    
    isThereYosanRecord = False
    Do Until EOF(1)
        Line Input #inputFileNumber, strBuf
        If Left(strBuf, 2) = "�݌�" And Mid(strBuf, 3, Len(non0Wat1Cafe2flg)) = non0Wat1Cafe2flg Then
            isThereYosanRecord = True
        End If
        If strBuf = "" And isThereYosanRecord = False Then
            If MsgBox("�u" & non0Wat1Cafe2flg & "�v�͑��݂��܂���" & vbCrLf & _
                    "�V�����s��ǉ����܂����H", vbOKCancel) = vbOK Then
                strBuf = "�݌�" & non0Wat1Cafe2flg & ",0�@" & _
                        Mid(cmbItem.Text, Len(non0Wat1Cafe2flg) + 2) & ",\0"
                SpecialPrint outputFileNumber, strBuf, non0Wat1Cafe2flg, "+"
                strBuf = ""
            End If
            isThereYosanRecord = True
        End If
        SpecialPrint outputFileNumber, strBuf, non0Wat1Cafe2flg, "+"
    Loop
    
    Close #outputFileNumber
    Close #inputFileNumber
    
    FileNameConverter csvFileName
    
    cmbItem.ListIndex = -1
    thText.Text = ""
    valText.Text = ""
End Sub

Private Sub thText_Change()
    thText.Text = retNumericText(thText.Text)
End Sub


Private Sub valText_Change()
    valText.Text = retNumericText(valText.Text)
End Sub
