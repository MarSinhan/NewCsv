VERSION 5.00
Begin VB.Form frmAddItm 
   Caption         =   "資産品追加Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton getYosanItemBtn 
      Caption         =   "Get!Add!"
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox valText 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox thText 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox cmbItem 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "購入価格"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "追加個数"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "予算用商品種別"
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
        MsgBox "そのファイル名のcsvファイルが見つかりません。", vbCritical + vbOKOnly, "ファイル未確認"
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
        MsgBox "その俗称の在庫が無効です。"
        Exit Sub
    End If
    
    OpenInOutFile inputFileNumber, outputFileNumber
    
    isThereYosanRecord = False
    Do Until EOF(1)
        Line Input #inputFileNumber, strBuf
        If Left(strBuf, 2) = "在庫" And Mid(strBuf, 3, Len(non0Wat1Cafe2flg)) = non0Wat1Cafe2flg Then
            isThereYosanRecord = True
        End If
        If strBuf = "" And isThereYosanRecord = False Then
            If MsgBox("「" & non0Wat1Cafe2flg & "」は存在しません" & vbCrLf & _
                    "新しい行を追加しますか？", vbOKCancel) = vbOK Then
                strBuf = "在庫" & non0Wat1Cafe2flg & ",0ｺ@" & _
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
