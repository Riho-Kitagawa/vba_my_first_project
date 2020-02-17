Attribute VB_Name = "Module1"

Option Explicit

Sub sample6()
    Dim file As String                      '開いたAnalysisのファイル名前
    Dim i As Long: i = 1                'セル番地
    Dim filePath As String             'ファイルのパス
    Dim opnAls As Workbook       '開いたAnalysisブック
    Dim WS As Worksheet             'ワークシート
    Dim find_name As String        'Sheet1を探す変数
    Dim flg As Boolean                 'Sheet1があるかどうかのフラグ
    Dim fileNum As Long             'Analysisのファイル番号
    Dim newSheet As Worksheet '新ブック
    Dim newSheetName As String  '新ブックの名前
    Dim rptName As String            'レポート帳票名
    Dim techKey() As String
    Dim prmptName() As String
    Dim alsWs As Worksheet
    Dim j As Long
    Dim lstNum As Long
    
    
    'Application.ScreenUpdating = False
    filePath = ThisWorkbook.Worksheets("Sheet1").Range("E3").Value & "\" 'E3でファイルパスを取得
    file = Dir(filePath & "*.xlsx") 'パス配下拡張子「.xlsx」の最初のファイル名を取得
    find_name = "Sheet2"
    flg = False
    
       '-----------------フォルダ内のブック繰り返し開く処理-----------------
    
    Do While file <> ""
    
        '-----------------このブックにSheet2があるかどうか-----------------
        
            For Each WS In ThisWorkbook.Worksheets
            'wsの名前がfind_name("Sheet2")と同じだったら
                If WS.name = find_name Then flg = True 'flgをtrueに
            Next WS
            
        '-----------------Sheet2があれば---------------------------------
        fileNum = Left(file, 4)     'ファイル名の頭４桁を取得、つまり管理番号
            If flg = True Then
                Dim p, s, d1, d2 As Long
                s = ThisWorkbook.Worksheets("Sheet2").Cells(1, 1).End(xlDown).row
                d1 = ThisWorkbook.Worksheets("Sheet2").Cells(1, 1).Value
                If d1 = fileNum Then
                     GoTo Label1  '開かずに飛ぶ
                End If
            
                For p = 2 To s
                d2 = ThisWorkbook.Worksheets("Sheet2").Cells(p, 1).Value
                If d2 = fileNum Then
                    GoTo Label1 '開かずに飛ぶ
                End If
                Next p
            End If

        flg = False
        Set opnAls = Workbooks.Open(filePath & file) 'Analysis該当ブックを開きopnAlsに入れる
        'Set opnAls = Workbooks.Open(filePath & "0914_ｴﾗｽﾚﾝ販売実績⑤国追加_? 宇浩様_販売管理 .xlsx") 'Analysis該当ブックを開きopnAlsに入れる
        file = opnAls.name          'Analysis該当ブック(opnAls)のファイル名を取得
'        fileNum = Left(file, 4)     'ファイル名の頭４桁を取得

         
                'rptName = opnAls.Worksheets("Sheet1").Cells(1, 1).Value '該当ファイルのA1の値を取得
                Dim lResult As Variant
                lResult = Application.Run("SAPGetSourceInfo", "DS_1", "DataSourceName")
                Dim Worksheet As Worksheet, flag As Boolean
                flag = False
                                
                Dim del As String
                del = "_" ' アンダースコアを探す
                Dim a As String
                a = InStrRev(file, del)  '最後の"_"の位置を取得
                Dim b As String
                b = Len(file)    'file名の長さを取得
                Dim c As String
                c = b - a  'file名から最後のアンスコの位置を引く
                
                Debug.Print InStrRev(file, del)
                Dim fileLastNm As String
                fileLastNm = Right(file, c) 'rightで最後のアンスコの後ろを取得
                
                For Each Worksheet In ThisWorkbook.Worksheets
                    If Worksheet.name = "Sheet2" Then flag = True
                Next Worksheet
                If flag = False Then
                    Set newSheet = ThisWorkbook.Worksheets.Add '新シート作成しnewBookに代入
                    newSheet.name = "Sheet2" '新シートの名前を取得
                    newSheetName = newSheet.name
                End If
                    
                Set newSheet = ThisWorkbook.Worksheets(1)

                If newSheet.Cells(i, 2).Value = "" And newSheet.Cells(i, 1).Value = "" And newSheet.Cells(i, 3).Value = "" Then  '新規作成したnewBookのシート１のA1の値がnullだったら
                    newSheet.Cells(i, 2).Value = lResult 'rptNameを新ブックのA1に出力
                    newSheet.Cells(i, 1).Value = Format(fileNum, "0000")
                    newSheet.Cells(i, 3).Value = fileLastNm
                Else
                    i = newSheet.Cells(1, 1).End(xlDown).row + 1
                    Debug.Print i
                    newSheet.Cells(i, 2).Value = lResult 'rptNameを新ブックのA1に出力
                    newSheet.Cells(i, 1).Value = Format(fileNum, "0000")
                    newSheet.Cells(i, 3).Value = fileLastNm
                End If

                i = i + 1 'インクリメントして、一行下に出力する。
'            On Error Resume Next
            Set alsWs = opnAls.Worksheets(1)
'
'          'newSheetの最後の行を取得
           lstNum = getLastRow(newSheet)
            
           Debug.Print lstNum
           
'            Dim flagflag As Boolean
'            flagflag = False
'            'Sheet2のA列がファイル名にあったら
'            Dim p As Long
'            Dim a As String
'            For p = 1 To lstNum
'                a = alsWs.Cells(p, 1).Value
'                If a = file Then flagflag = True
'            Next p
            
'           Dim testArray As Variant
'           With newSheet
'               testArray = .Range(.Cells(1, 1), .Cells(lstNum, 1)).Value
'           End With
           

           
           
'           Dim result As Variant
'           result = Filter(testArray, fileNum)
'
'           If (UBound(result) <> -1) Then
'               Debug.Print fileNum & "を含む配列は存在する"
'           Else
'                Debug.Print fileNum & "を含む配列は存在しません"
'           End If
           
'           Dim varResult
'           varResult = Filter(testArray, fileNum)
'
'           If UBound(varResult) <> -1 Then
'                Debug.Print file & "は配列内に存在する"
'           Else
'                Debug.Print file & "は配列内に存在しない"
'           End If
'ここから
'           Debug.Print UBound(testArray) '6
'           Debug.Print LBound(testArray)   '1
'           'newBook.Worksheets("Sheet1").Range("C1", Cells(LBound(testArray, 1), UBound(testArray, 1))).Value = testArray
'           Dim k As Long
'           Dim l As Long
'           For k = 3 To lstNum - 1
'               newBook.Worksheets("Sheet1").Cells(1, k).Value = testArray(l)
'           Next k
'ここまでやっている途中


        DoEvents
        Workbooks(file).Close SaveChanges:=False  '立ち上げたAnalysisブックを保存しないで閉じる
Label1:
        file = Dir
        ThisWorkbook.Save
    Loop
    
    'Call find_NA
    
    'A1をキーにして昇順にする
    Call Range("A:C").Sort( _
    Key1:=Range("A1"), _
    Order1:=xlAscending)
    
    MsgBox "終わりでござる"
    'Application.ScreenUpdating = True
End Sub

'帳票ブロックの最終行を取得するメソッド
Function getLastRow(WS As Worksheet, Optional CheckCol As Long = 1) As Long
    getLastRow = WS.Cells(1, CheckCol).End(xlDown).row
End Function



Public Function find_NA()
    Dim filePath As String
    Dim fileNum As String
    Dim newSheet As Worksheet
    Dim lstNum As Long
    Dim i As Long
    filePath = ThisWorkbook.Worksheets("Sheet1").Range("E3").Value & "\" 'E3でファイルパスを取得
    Set newSheet = ThisWorkbook.Worksheets("Sheet2")
    lstNum = getLastRow(newSheet)
    For i = 1 To lstNum
        If IsError(newSheet.Cells(i, 2).Value) Then
            fileNum = Cells(i, 1).Value
            Workbooks.Open fileName:=filePath & fileNum & "*.xlsx"
            Dim lResult As Variant
            lResult = Application.Run("SAPGetSourceInfo", "DS_1", "DataSourceName")
            
            
        End If
    Next i
End Function
