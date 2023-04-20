VERSION 5.00
Begin VB.Form frmCsvIn 
   Caption         =   "CSV入力Form"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton btnOpenCsv 
      Caption         =   "OpenCSV"
      Height          =   375
      Left            =   7080
      TabIndex        =   18
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton btnNewType 
      Caption         =   "newType変換"
      Height          =   375
      Left            =   9720
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton delBackUp 
      Caption         =   "DeleteBackUp"
      Height          =   375
      Left            =   6840
      TabIndex        =   16
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   840
      Top             =   1680
   End
   Begin VB.CommandButton csvButton 
      Caption         =   "csv!csv!Input!"
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton btnAddItem 
      Caption         =   "資産品追加"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton btnListD 
      Caption         =   "リストデータ入力"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox eatCostText 
      Alignment       =   1  '右揃え
      Height          =   270
      IMEMode         =   2  'ｵﾌ
      Left            =   11280
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox typeComBox 
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
      IMEMode         =   1  'ｵﾝ
      ItemData        =   "frmCsvIn.frx":0000
      Left            =   9120
      List            =   "frmCsvIn.frx":0002
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
   Begin VB.ComboBox payMethComBox 
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
      IMEMode         =   1  'ｵﾝ
      Left            =   6720
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox valueText 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   2  'ｵﾌ
      Left            =   4680
      MaxLength       =   9
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox itemNameComBox 
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
      IMEMode         =   1  'ｵﾝ
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.ComboBox dateComBox 
      Height          =   300
      Left            =   360
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblCsvName 
      Caption         =   "csvName"
      Height          =   495
      Left            =   9480
      TabIndex        =   14
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "(食費)"
      Height          =   255
      Left            =   11400
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "商品種別"
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "支払方法"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "価格"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "商品名"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "日付"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmCsvIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long
                                                  
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, _
                                                          ByVal dwPriorityClass As Long) As Long

Private Function PutMoneyCsv() As String
    Dim retStr As String
    
    retStr = ""
    retStr = retStr & frmCsvIn.dateComBox.Text
    retStr = retStr & ","
    retStr = retStr & frmCsvIn.itemNameComBox.Text
    retStr = retStr & ","
    retStr = retStr & "\" & frmCsvIn.valueText.Text
    retStr = retStr & ","
    retStr = retStr & frmCsvIn.payMethComBox.Text
    retStr = retStr & ","
    retStr = retStr & frmCsvIn.typeComBox.Text
    
    If True = IsNumeric(eatCostText.Text) Then
        retStr = retStr & ","
        retStr = retStr & frmCsvIn.eatCostText.Text
    End If
    
    PutMoneyCsv = retStr
    
End Function

Private Sub btnNewType_Click()
    Dim inNum As Integer
    Dim outNum As Integer
    Dim strBuf As String
    Dim afterBuf As String
    Dim wkBuf As String
    Dim refBuf As String
    Dim refBuf2 As String
    Dim cnt As Integer
    Dim cha As String
    Dim calcFlg As Boolean
    Dim stringFlg As Boolean
    Dim strSecondFlg As Boolean
    Dim kakko As Integer
    
    inNum = FreeFile
    Open App.Path & "\" & csvFileName & ".csv" For Input As #inNum
    
    outNum = FreeFile
    Open App.Path & "\" & csvFileName & "new.csv" For Output As #outNum
    
    Do Until EOF(1)
        afterBuf = ""
        Line Input #inNum, strBuf
        calcFlg = False
        stringFlg = False
        Do While InStr(strBuf, "DATEDIF") > 0
            kakko = 0
            wkBuf = "DATEDIF"
            cnt = InStr(strBuf, "DATEDIF")
            cnt = cnt + Len("DATEDIF")
            Do
                cha = Mid(strBuf, cnt, 1)
                If cha = "(" Then
                    kakko = kakko + 1
                ElseIf cha = ")" Then
                    kakko = kakko - 1
                End If
                wkBuf = wkBuf & cha
                If kakko = 0 Then
                    Exit Do
                End If
                cnt = cnt + 1
            Loop
            strBuf = Replace(strBuf, wkBuf, "ゴメス")
            
            kakko = 1
            cnt = InStr(wkBuf, "DATEDIF")
            cnt = cnt + Len("DATEDIF") + 1
            refBuf = ""
                Do
                    cha = Mid(wkBuf, cnt, 1)
                    If cha = "(" Then
                        kakko = kakko + 1
                    ElseIf cha = ")" Then
                        kakko = kakko - 1
                    ElseIf cha = "," And kakko = 1 Then
                        Exit Do
                    End If
                    refBuf = refBuf & cha
                    cnt = cnt + 1
                Loop
                refBuf2 = refBuf
                cnt = cnt + 1
                refBuf = ""
                Do
                    cha = Mid(wkBuf, cnt, 1)
                    If cha = "(" Then
                        kakko = kakko + 1
                    ElseIf cha = ")" Then
                        kakko = kakko - 1
                    ElseIf cha = "," And kakko = 1 Then
                        Exit Do
                    End If
                    refBuf = refBuf & cha
                    cnt = cnt + 1
                Loop
            wkBuf = "DAYS(" & refBuf & "," & refBuf2 & ")"
            
            strBuf = Replace(strBuf, "ゴメス", wkBuf)
        Loop
                
        For cnt = 1 To Len(strBuf)
            cha = Mid(strBuf, cnt, 1)
            If cha = """" And cnt < Len(strBuf) Then
                If Mid(strBuf, cnt + 1, 1) = """" Then
                    stringFlg = Not (stringFlg)
                    cha = """"""
                    cnt = cnt + 1
                Else
                    calcFlg = Not (calcFlg)
                End If
            ElseIf cha = "," And calcFlg = True And stringFlg = False Then
                cha = ";"
            ElseIf cha = "\" And stringFlg = False Then
                cha = ""
            End If
            afterBuf = afterBuf & cha
        Next cnt
        Print #outNum, afterBuf
    Loop
    
    Close #outNum
    Close #inNum
End Sub

Private Sub btnOpenCsv_Click()
    frmOpenCsv.Show
End Sub

Private Sub csvButton_Click()
    Dim inputFileNumber As Long
    Dim outputFileNumber As Long
    Dim backUpFileName As String
    Dim tempFileNumber As Long
    Dim strBuf As String
    Dim wkStrBuf As String
    Dim isCsvOutputComplete As Boolean
    Dim cnt As Long
    Dim checkWatCafe As String
    Dim checkPayMeth As String
    Dim non0Wat1Cafe2flg As String
    Dim atmarkItemName As String


    If False = IsDate(frmCsvIn.dateComBox.Text) Then
        Exit Sub
    ElseIf "" = Trim(frmCsvIn.itemNameComBox.Text) Then
        Exit Sub
    ElseIf "" = Trim(frmCsvIn.payMethComBox.Text) Then
        Exit Sub
    ElseIf "" = Trim(frmCsvIn.typeComBox.Text) Then
        Exit Sub
    ElseIf "" = Trim(frmCsvIn.valueText.Text) Then
        Exit Sub
    ElseIf "" = Dir(App.Path & "\" & csvFileName & ".csv") Then
        MsgBox "そのファイル名のcsvファイルが見つかりません。", vbCritical + vbOKOnly, "ファイル未確認"
        Exit Sub
    End If
    
    OpenInOutFile inputFileNumber, outputFileNumber
        
    checkWatCafe = frmCsvIn.itemNameComBox.Text
    checkPayMeth = frmCsvIn.payMethComBox.Text
    
    non0Wat1Cafe2flg = ""
    For cnt = 1 To UBound(isAtmarkFlg)
        atmarkItemName = isAtmarkFlg(cnt)
        
        If atmarkItemName = "" Then
            Exit For
        End If
        If Left(checkWatCafe, Len(atmarkItemName)) = atmarkItemName And checkPayMeth = "在庫" Then
            non0Wat1Cafe2flg = atmarkItemName
            Exit For
        End If
    Next cnt
    
    If non0Wat1Cafe2flg = "" Then
        non0Wat1Cafe2flg = isNormalPayFlg
    End If
        
    isCsvOutputComplete = False
    Do Until EOF(1)
        Line Input #inputFileNumber, strBuf
        
            
        
        If (retYosanNameCheckBool(strBuf) = True Or strBuf = "") And isCsvOutputComplete = False Then
            wkStrBuf = strBuf
            strBuf = PutMoneyCsv
            If strBuf <> "" Then
                SpecialPrint outputFileNumber, strBuf, non0Wat1Cafe2flg, "-"
                isCsvOutputComplete = True
            End If
            SpecialPrint outputFileNumber, wkStrBuf, non0Wat1Cafe2flg, "-"
        Else
            SpecialPrint outputFileNumber, strBuf, non0Wat1Cafe2flg, "-"
        End If
    Loop
    
    Close #outputFileNumber
    Close #inputFileNumber
    
    FileNameConverter csvFileName
    
    btnNewType_Click
    
    frmCsvIn.itemNameComBox.Text = ""
    frmCsvIn.valueText.Text = ""
    frmCsvIn.payMethComBox.Text = ""
    frmCsvIn.typeComBox.Text = ""
    frmCsvIn.eatCostText.Text = ""
End Sub



Private Sub btnAddItem_Click()
    frmAddItm.Show
End Sub


Private Sub btnListD_Click()
    frmListD.Show
End Sub

Private Sub dateComBox_Change()
    dateComBox.Text = Format(dateComBox.Text, "mm/dd")
End Sub



Private Sub delBackUp_Click()
    Dim cnt As Long
    Dim killerFileName As String
    Dim killerDataFileName As String

    If vbOK = MsgBox("支配下のバックアップファイルを削除しますが？", vbOKCancel + vbExclamation, "削除の確認") Then
        For cnt = 9 To 1 Step -1
            If cnt > 1 Then
                killerFileName = Dir(App.Path & "\" & csvFileName & ".ba" & cnt)
                killerDataFileName = Dir(App.Path & "\" & dataFileName & ".ba" & cnt)
            ElseIf cnt = 1 Then
                killerFileName = Dir(App.Path & "\" & csvFileName & ".bak")
                killerDataFileName = Dir(App.Path & "\" & dataFileName & ".bak")
            End If
            
            If killerFileName <> "" Then
                Kill App.Path & "\" & killerFileName
            End If
            If killerDataFileName <> "" Then
                Kill App.Path & "\" & killerDataFileName
            End If
        Next
        MsgBox "削除しますた。", vbOKOnly, "完了"
    End If
End Sub



Private Sub eatCostText_Change()
    eatCostText.Text = retNumericText(eatCostText.Text)
End Sub

Private Sub Form_Load()
    Dim id As Long
    theEndOfProgramFlg = False
    
    SetPriorityClass Me.hWnd, &H8000&
    
    id = Me.ScaleMode
    Me.ScaleMode = vbPixels
    MoveWindow itemNameComBox.hWnd, itemNameComBox.Left, itemNameComBox.Top, itemNameComBox.Width, 800, 1&
    MoveWindow payMethComBox.hWnd, payMethComBox.Left, payMethComBox.Top, payMethComBox.Width, 800, 1&
    MoveWindow typeComBox.hWnd, typeComBox.Left, typeComBox.Top, typeComBox.Width, 800, 1&
    Me.ScaleMode = id

    dayListSetting

    DoComboListAddParam 1
    DoComboListAddParam 2

     frmCsvIn.dateComBox.ListIndex = frmCsvIn.dateComBox.ListCount - dayRange - 1
   

End Sub

Private Sub Form_Unload(Cancel As Integer)
    theEndOfProgramFlg = True
    Unload frmAddItm
    Unload frmListD
End Sub





Private Sub itemNameComBox_Click()
    Dim boxValue As String
    
    boxValue = itemNameComBox.Text

    If InStr(boxValue, "@") > 0 Then
        valueText.Text = Val(Right(boxValue, Len(boxValue) - InStr(boxValue, "@")))
        payMethComBox.Text = "在庫"
        Timer1.Enabled = True
    End If
    
End Sub



Private Sub Timer1_Timer()
    itemNameComBox.Text = Left(itemNameComBox.Text, InStr(itemNameComBox.Text, "@") - 1)
    Timer1.Enabled = False
End Sub

Private Sub valueText_Change()
    valueText.Text = retNumericText(valueText.Text)
End Sub
