Attribute VB_Name = "baseModule"
Option Explicit

Public Const csvInputSheetName = "csv入力シート"
Public Const forListDataSheetName = "リスト用データ"
Public Const yosanAddSheetName = "資産品追加"
Public Const fileNameSetCellPosition = "G2"
Public csvFileName As String
Public Const maxColumn = 5      'リスト用データ列幅
Public Const fileNameSetColumn = 1 'ファイル名の列番号
Public Const dateColumn = 1    '購入日付の列番号
Public Const itemColumn = 2    '品名の列番号
Public Const valueColumn = 3   '支払額の列番号
Public Const payMethColumn = 4 '支払方法の列番号
Public Const typeColumn = 5    '出費カテゴリーの列番号
Public Const monthRange = 1    '前の設定可能日付の幅（ヶ月）
Public Const dayRange = 3       '後の設定可能日付の幅(日)
Public Const dataFileName = "fileOfData"    'リスト用事前記録データファイルのファイル名
Public Const itemListIndex = 1  '商品名リストのリスト番号
Public Const methListIndex = 2  '支払方法リストのリスト番号
Public Const kindListIndex = 3  '商品種別リストのリスト番号
Public Const isNormalPayFlg = ""
Public saveDataChgFlg As Boolean 'リストが更新され未保存かどうかのフラグ
Public theEndOfProgramFlg As Boolean
Public isAtmarkFlg As Variant


Public Sub OpenInOutFile(ByRef inNum As Long, ByRef outNum As Long)
    inNum = FreeFile
    Open App.Path & "\" & csvFileName & ".csv" _
        For Input As #inNum
    
    outNum = FreeFile
    Open App.Path & "\" & csvFileName & ".outcsv" _
        For Output As #outNum
End Sub


Public Sub DoComboListAddParam(ByVal column As Long)
    Dim cnt As Long
    Dim listEnd As Long
    Dim ctrlObj As Object
    Dim ctrlList As Object
    Dim ctrlObj2nd As Object
    Dim strWk(1 To 15) As String
    Dim strWkMake As Long
    Dim strBuf As String
    Dim strListData As String
    Dim inNum As Integer
    Dim i As Integer
    

    If column <> valueColumn Then
    
        Select Case column
            Case dateColumn
                Set ctrlObj = frmCsvIn.dateComBox
        End Select
    
        If column = dateColumn Then
            listEnd = frmListD.dayList.ListCount
            For cnt = 0 To listEnd - 1
                ctrlObj.AddItem Format(DateValue(frmListD.dayList.List(cnt)), "mm/dd")
            Next cnt
            

        Else
            inNum = FreeFile
            Open App.Path & "\" & dataFileName & ".csv" For Input As #inNum
            cnt = 1
            Do Until EOF(1)
                Line Input #inNum, strBuf
                If cnt = fileNameSetColumn Then
                    csvFileName = strBuf
                    frmCsvIn.lblCsvName.Caption = csvFileName & ".csv"
                ElseIf cnt <> valueColumn And cnt <= maxColumn Then
                    Select Case cnt
                        Case itemColumn
                            Set ctrlObj = frmCsvIn.itemNameComBox
                            Set ctrlList = frmListD.itemNameList
                        Case payMethColumn
                            Set ctrlObj = frmCsvIn.payMethComBox
                            Set ctrlList = frmListD.payMethList
                        Case typeColumn
                            Set ctrlObj = frmCsvIn.typeComBox
                            Set ctrlList = frmListD.kindList
                    End Select
                    
                    ctrlList.Clear
                    For i = 1 To Len(strBuf)
                        strListData = ""
                        Do While Mid(strBuf, i, 1) <> "," And i <= Len(strBuf)
                            strListData = strListData & Mid(strBuf, i, 1)
                            i = i + 1
                        Loop
                        ctrlList.AddItem strListData
                    Next i
                End If
                cnt = cnt + 1
            Loop
            
            Close #inNum
            
            For i = itemColumn To maxColumn
                Select Case i
                    Case itemColumn
                        Set ctrlObj = frmCsvIn.itemNameComBox
                        Set ctrlList = frmListD.itemNameList
                    Case payMethColumn
                        Set ctrlObj = frmCsvIn.payMethComBox
                        Set ctrlList = frmListD.payMethList
                    Case typeColumn
                        Set ctrlObj = frmCsvIn.typeComBox
                        Set ctrlList = frmListD.kindList
                End Select
                If i <> valueColumn Then
                    listEnd = ctrlList.ListCount
                    ctrlObj.Clear
                    For cnt = 1 To listEnd
                        ctrlObj.AddItem ctrlList.List(cnt - 1)
                    Next cnt
                    ctrlObj.AddItem "", 0
                End If
            Next i
            
            For cnt = 1 To UBound(strWk)
                strWk(cnt) = ""
            Next cnt
                
            listEnd = frmListD.itemNameList.ListCount
            
            strWkMake = 1
            frmAddItm.cmbItem.Clear
            For cnt = 1 To listEnd
                strBuf = frmListD.itemNameList.List(cnt - 1)
                If InStr(strBuf, "@") > 0 Then
                    frmAddItm.cmbItem.AddItem strBuf
                    strWk(strWkMake) = Left(strBuf, InStr(strBuf, "@") - 1)
                    strWkMake = strWkMake + 1
                End If
            Next cnt
            frmAddItm.cmbItem.AddItem "", 0
            isAtmarkFlg = Array(, strWk(1), strWk(2), strWk(3), strWk(4), strWk(5), strWk(6), strWk(7), _
                    strWk(8), strWk(9), strWk(10), strWk(11), strWk(12), strWk(13), strWk(14), strWk(15))
        End If
        
        
    End If
End Sub

Public Sub SpecialPrint( _
        ByVal fileNumber As Long, ByVal strBuf As String, ByVal non0Wat1Cafe2flg As String, flg As String)
    Dim digNum As Long
    Dim beforeItemNum As String
    Dim afterItemNum As String
    Dim beforeValue As String
    Dim afterValue As String
    Dim num As String
    Dim value As String
    Dim singl As Long
    Dim cnt As Integer
    Dim searchIndex As String
        
    If non0Wat1Cafe2flg <> isNormalPayFlg Then
        If Left(strBuf, 2 + Len(non0Wat1Cafe2flg)) = ("在庫" & non0Wat1Cafe2flg) Then
            digNum = InStr(strBuf, "ｺ") - InStr(strBuf, ",") - 1
            num = Mid(strBuf, 2 + Len(non0Wat1Cafe2flg) + 2, digNum)
            beforeItemNum = num & "ｺ"
            If flg = "+" Then
                num = Val(num) + Val(frmAddItm.thText.Text)
                afterItemNum = Val(num) & "ｺ"
            Else
                num = Val(num) - 1
                afterItemNum = Val(num) & "ｺ"
            End If
            
            
            
            value = Mid(strBuf, InStr(strBuf, "\") + 1)
            beforeValue = "\" & value
            If flg = "+" Then
                singl = Val(Mid(frmAddItm.cmbItem.Text, Len(non0Wat1Cafe2flg) + 2))
                If frmAddItm.valText.Text <> "" Then
                    searchIndex = non0Wat1Cafe2flg & "@" & Val(singl)
                    For cnt = 0 To frmListD.itemNameList.ListCount - 1
                        If frmListD.itemNameList.List(cnt) = searchIndex Then
                            Exit For
                        End If
                    Next cnt
                    beforeItemNum = beforeItemNum & "@" & Val(singl)
                    value = Val(value) + Val(frmAddItm.valText.Text)
                    singl = Val(Int(Val(value) / Val(num) + 0.5))
                    value = Val(singl) * Val(num)
                    searchIndex = non0Wat1Cafe2flg & "@" & Val(singl)
                    frmListD.itemNameList.AddItem searchIndex, cnt + 1
                    frmListD.itemNameList.RemoveItem cnt
                    afterItemNum = afterItemNum & "@" & Val(singl)
                    frmListD.Show
                    saveDataChgFlg = True
                Else
                    value = Val(singl) * Val(num)
                End If
                afterValue = "\" & Val(value)
            Else
                value = Val(value) - Val(frmCsvIn.valueText.Text)
                afterValue = "\" & Val(value)
            End If
            strBuf = Replace(strBuf, beforeItemNum, afterItemNum)
            strBuf = Replace(strBuf, beforeValue, afterValue)
        End If
    End If
        
    Print #fileNumber, strBuf
End Sub

Public Sub FileNameConverter(changeFileName As String)
    Dim cnt As Long

    If "" <> Dir(App.Path & "\" & changeFileName & ".bak") Then
        If "" <> Dir(App.Path & "\" & changeFileName & ".ba9") Then
            Kill App.Path & "\" & changeFileName & ".ba9"
        End If
        For cnt = 8 To 2 Step -1
            If "" <> Dir(App.Path & "\" & changeFileName & ".ba" & cnt) Then
                Name App.Path & "\" & changeFileName & ".ba" & cnt As _
                        App.Path & "\" & changeFileName & ".ba" & (cnt + 1)
            End If
        Next
        Name App.Path & "\" & changeFileName & ".bak" As App.Path & "\" & changeFileName & ".ba2"
    End If
        
    Name App.Path & "\" & changeFileName & ".csv" As App.Path & "\" & changeFileName & ".bak"
    Name App.Path & "\" & changeFileName & ".outcsv" As App.Path & "\" & changeFileName & ".csv"
    
End Sub

Public Function retYosanNameCheckBool(ByVal strBuf As String) As Boolean
    Dim cnt As Long
    
    For cnt = 1 To UBound(isAtmarkFlg)
        If isAtmarkFlg(cnt) = "" Then
            retYosanNameCheckBool = False
            Exit For
        End If
        If Left(strBuf, Len("在庫" & isAtmarkFlg(cnt))) = ("在庫" & isAtmarkFlg(cnt)) Then
            retYosanNameCheckBool = True
            Exit For
        End If
    Next cnt
End Function

Public Function retNumericText(ByVal strNum As String) As String
    Dim rp As Integer
    Dim mdValue As String
    
    rp = 1
    Do While rp <= Len(strNum)
        mdValue = Mid(strNum, rp, 1)
        If Asc(mdValue) < Asc("0") Or Asc(mdValue) > Asc("9") Then
            If rp = 1 Then
                strNum = Mid(strNum, 2, Len(strNum) - 1)
            ElseIf rp = Len(strNum) Then
                strNum = Mid(strNum, 1, Len(strNum) - 1)
            Else
                strNum = Mid(strNum, 1, rp - 1) & Mid(strNum, rp + 1, Len(strNum) - rp)
            End If
            
            rp = rp - 1
        End If
        
        rp = rp + 1
    Loop
    
    retNumericText = strNum
End Function

Public Sub listEditSaveReflesh()
    Dim outNum As Integer
    Dim cnt As Integer
    Dim i As Integer
    Dim strBuf As String
    Dim ctrlObj As Object
    
    outNum = FreeFile
    Open App.Path & "\" & dataFileName & ".outcsv" For Output As #outNum
    
    For cnt = 1 To maxColumn
        If cnt = fileNameSetColumn Then
            Print #outNum, csvFileName
        Else
            If cnt = itemColumn Then
                Set ctrlObj = frmListD.itemNameList
            ElseIf cnt = valueColumn Then
                Print #outNum, "0"
            ElseIf cnt = payMethColumn Then
                Set ctrlObj = frmListD.payMethList
            ElseIf cnt = typeColumn Then
                Set ctrlObj = frmListD.kindList
            End If
            
            If cnt <> valueColumn Then
                strBuf = ""
                For i = 0 To ctrlObj.ListCount - 1
                    If i <> 0 Then
                        strBuf = strBuf & ","
                    End If
                    strBuf = strBuf & ctrlObj.List(i)
                Next i
                Print #outNum, strBuf
            End If
        End If
    Next cnt
    
    Close #outNum
End Sub

Public Sub dayListSetting()
    Dim cnt As Long
    Dim dateCnt As Date
    Dim endDate As Date
    Dim listCnt As Long
    
    frmListD.dayList.Clear
            '前1か月と本日の日付リスト追加
    cnt = 1
    dateCnt = Date
    dateCnt = DateAdd("m", -1 * monthRange, dateCnt)
    Do While dateCnt <= Date
        frmListD.dayList.AddItem dateCnt
        cnt = cnt + 1
        dateCnt = dateCnt + 1
    Loop
    
    '後1か月の日付リスト追加
    dateCnt = Date + 1
    endDate = DateAdd("d", dayRange, Date)
    Do While dateCnt <= endDate
        frmListD.dayList.AddItem dateCnt
        cnt = cnt + 1
        dateCnt = dateCnt + 1
    Loop
End Sub

Private Sub Main()

End Sub
