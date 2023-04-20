VERSION 5.00
Begin VB.Form frmListD 
   Caption         =   "リストデータ入力Form"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton btnSaveData 
      Caption         =   "リスト保存"
      Height          =   255
      Left            =   10200
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton btnDelItem 
      Caption         =   "削除"
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton btnEditList 
      Caption         =   "編集"
      Height          =   375
      Left            =   8640
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton btnInputList 
      Caption         =   "入力"
      Height          =   255
      Left            =   8640
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtIndex 
      Alignment       =   2  '中央揃え
      Height          =   270
      IMEMode         =   2  'ｵﾌ
      Left            =   6960
      TabIndex        =   7
      Top             =   2745
      Width           =   1095
   End
   Begin VB.TextBox txtValueName 
      Height          =   270
      IMEMode         =   1  'ｵﾝ
      Left            =   4440
      TabIndex        =   6
      Top             =   2760
      Width           =   2295
   End
   Begin VB.ComboBox cmbListName 
      Height          =   300
      Left            =   2160
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ListBox kindList 
      Height          =   1860
      Left            =   9240
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.ListBox payMethList 
      Height          =   1860
      Left            =   7560
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox nullList 
      Height          =   1860
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox itemNameList 
      Height          =   1860
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.ListBox dayList 
      Height          =   1860
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "番号"
      Height          =   255
      Left            =   6960
      TabIndex        =   15
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "入力/編集内容"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "変更対象リスト"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblFOD 
      Caption         =   "dataFileName"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   3240
      Width           =   2415
   End
End
Attribute VB_Name = "frmListD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnDelItem_Click()
    If cmbListName.Text <> "" And txtIndex.Text <> "" Then
        If cmbListName.ListIndex = itemListIndex Then
            itemNameList.RemoveItem Val(txtIndex.Text) - 1
            itemNameList.ListIndex = -1
            txtIndex.Text = ""
            saveDataChgFlg = True
        ElseIf cmbListName.ListIndex = methListIndex Then
            payMethList.RemoveItem Val(txtIndex.Text) - 1
            payMethList.ListIndex = -1
            txtIndex.Text = ""
            saveDataChgFlg = True
        ElseIf cmbListName.ListIndex = kindListIndex Then
            kindList.RemoveItem Val(txtIndex.Text) - 1
            kindList.ListIndex = -1
            txtIndex.Text = ""
            saveDataChgFlg = True
        End If
    End If
End Sub

Private Sub btnEditList_Click()
    If cmbListName.Text <> "" And txtValueName.Text <> "" And txtIndex.Text <> "" Then
        If cmbListName.ListIndex = itemListIndex Then
            itemNameList.AddItem txtValueName.Text, Val(txtIndex.Text)
            itemNameList.RemoveItem Val(txtIndex.Text) - 1
            itemNameList.ListIndex = Val(txtIndex.Text) - 1
            saveDataChgFlg = True
        ElseIf cmbListName.ListIndex = methListIndex Then
            payMethList.AddItem txtValueName.Text, Val(txtIndex.Text)
            payMethList.RemoveItem Val(txtIndex.Text) - 1
            payMethList.ListIndex = Val(txtIndex.Text) - 1
            saveDataChgFlg = True
        ElseIf cmbListName.ListIndex = kindListIndex Then
            kindList.AddItem txtValueName.Text, Val(txtIndex.Text) - 1
            kindList.RemoveItem Val(txtIndex.Text) - 1
            kindList.ListIndex = Val(txtIndex.Text) - 1
            saveDataChgFlg = True
        End If
    End If
End Sub

Private Sub btnInputList_Click()
    If cmbListName.Text <> "" And txtValueName.Text <> "" And txtIndex.Text <> "" Then
        If cmbListName.ListIndex = itemListIndex Then
            itemNameList.AddItem txtValueName.Text, Val(txtIndex.Text)
            saveDataChgFlg = True
        ElseIf cmbListName.ListIndex = methListIndex Then
            payMethList.AddItem txtValueName.Text, Val(txtIndex.Text)
            saveDataChgFlg = True
        ElseIf cmbListName.ListIndex = kindListIndex Then
            kindList.AddItem txtValueName.Text, Val(txtIndex.Text)
            saveDataChgFlg = True
        End If
    End If
End Sub

Private Sub allListSelCancel(ByRef ctrl As Object)
    Dim ind As Integer

    ind = ctrl.ListIndex
    
    If ctrl.Name <> "dayList" Then
        dayList.ListIndex = -1
    End If
    If ctrl.Name <> "itemNameList" Then
        itemNameList.ListIndex = -1
    End If
    If ctrl.Name <> "nullList" Then
        nullList.ListIndex = -1
    End If
    If ctrl.Name <> "payMethList" Then
        payMethList.ListIndex = -1
    End If
    If ctrl.Name <> "kindList" Then
        kindList.ListIndex = -1
    End If
    
    ctrl.ListIndex = ind
    txtIndex = ind + 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub btnSaveData_Click()
    If saveDataChgFlg = True Then
        If MsgBox("リストデータの変更を保存しますか?", vbYesNo) = vbYes Then
            listEditSaveReflesh
            FileNameConverter dataFileName
            DoComboListAddParam 2
            saveDataChgFlg = False
        End If
    Else
        MsgBox "何も変更されていないようだす。"
    End If
End Sub

Private Sub dayList_Click()
    cmbListName.ListIndex = 0
    allListSelCancel dayList
End Sub


Private Sub Form_Load()
    Dim cnt As Long
    Dim dateCnt As Date
    Dim endDate As Date
    Dim listCnt As Long
    
    saveDataChgFlg = False
        
    dayListSetting
    
    frmListD.cmbListName.Clear
    frmListD.cmbListName.AddItem ""
    frmListD.cmbListName.AddItem "itemNameList"
    frmListD.cmbListName.AddItem "payMethList"
    frmListD.cmbListName.AddItem "kindList"
    
    frmListD.lblFOD.Caption = dataFileName & ".csv"
    
        'コンボボックス
        
        DoComboListAddParam 2
        
    


End Sub

Private Sub Form_Unload(Cancel As Integer)
    If saveDataChgFlg = True Then
        If MsgBox("リストデータの変更を保存しますか?", vbYesNo) = vbYes Then
            listEditSaveReflesh
            FileNameConverter dataFileName
            DoComboListAddParam 2
            saveDataChgFlg = False
        End If
    End If
    If theEndOfProgramFlg = False Then
        DoComboListAddParam 2
    End If
End Sub

Private Sub itemNameList_Click()
    cmbListName.ListIndex = itemListIndex
    allListSelCancel itemNameList
End Sub

Private Sub kindList_Click()
    cmbListName.ListIndex = kindListIndex
    allListSelCancel kindList
End Sub

Private Sub nullList_Click()
    cmbListName.ListIndex = 0
    allListSelCancel nullList
End Sub

Private Sub payMethList_Click()
    cmbListName.ListIndex = methListIndex
    allListSelCancel payMethList
End Sub

Private Sub txtIndex_Change()
    txtIndex.Text = retNumericText(txtIndex.Text)
End Sub
