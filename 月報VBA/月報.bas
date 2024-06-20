Option Explicit
'作成
Sub Report()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim thisBK As Workbook
    Dim areaBK As Workbook
    Dim BK As Workbook
    
    'メインブック
    Set thisBK = ThisWorkbook
    
    'メインブック、表シートの格納
    Dim AreaSh As Worksheet
    Set AreaSh = thisBK.Worksheets("表")
    
    
    '作成日格納
    Dim toDay As Date
    
        If thisBK.Worksheets("メイン").Cells(2, 6).Value <> "" Then
            toDay = thisBK.Worksheets("メイン").Cells(2, 6).Value
        Else
            MsgBox "作成日の取得に失敗しました。" & vbCrLf & _
                   "メインシート、最終作成月日に月末日付を入力後、再実行してください。", vbExclamation
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
    '***** 本当に作成してよいか判定
    
    Dim eom As Boolean
        eom = False
        
        '月末判定
        If Month(toDay) = 2 Then
            '2月
            Select Case Day(toDay)
                Case 28, 29
                    eom = True
            End Select
        Else
            '他の月
            Select Case Day(toDay)
                Case 30, 31
                    eom = True
            End Select
        End If
        
    Dim areaCheck As Boolean
        areaCheck = False
        
        '書き込み完了判定
        If thisBK.Worksheets("メイン").Cells(26, 16).Value = thisBK.Worksheets("メイン").Cells(26, 16).Value Then
            areaCheck = True
        End If
        
        '条件が整っているか
        If eom And areaCheck Then
        Else
            '整ってない
            MsgBox "作成の準備が整っていないようです。" & vbCrLf & _
                   "作成日が月末か、全の作成が完了するまでは実行できません。", vbExclamation
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
    
    '***** 続行
    
    'ファイルがあるか精査する ==========================
    
    Call BookCreate(thisBK, toDay)
    
    '格納処理 ==========================================
    
    Dim R As Long
    Dim area As Variant
    Dim areaRng As Range
    Dim areaDic As Object
    Set areaDic = CreateObject("Scripting.Dictionary")

        R = AreaSh.Cells(AreaSh.Rows.Count, 1).End(xlUp).Row
    
        AreaSh.Activate
    
    Set areaRng = AreaSh.Range(AreaSh.Cells(2, 1), AreaSh.Cells(R, 1))
    
        '計上するを辞書に格納する
        For Each area In areaRng
            areaDic.Add area.Value, area.Offset(0, 1).Value
        Next area
    
    Set areaRng = Nothing
    Set AreaSh = Nothing
    
    'ブックを順に開き、情報を取得していく ============
    
    'ファイルナンバー
    Dim num As Integer
        num = searchmonth(toDay)
    
    'ファイルパス
    Dim areaPth As String
    
    'ファイル取得列
    Dim areaCol As Long
        areaCol = areaBookGetCol(toDay)
        
    'ファイルsh
    Dim AreaSh As Worksheet
                
        Set AreaSh = Nothing
    
    '昨年売上用配列
    Dim lastRevenue() As Currency '8分類　中分類+合計
    ReDim lastRevenue(7)
        
        
    '***** 本処理 *****
    
        'ファイルを開いて変数に格納
        Set BK = Workbooks.Open(thisBK.path & "\\" & Year(toDay) & ".xlsx")
        
        '日付はいってなかったら入れる
        If BK.Worksheets("支部").Cells(3, 4).Value = "" Then BK.Worksheets("支部").Cells(3, 4).Value = Year(toDay) & "/1/1"
        
    
    'ファイルを1つずつ開ける
    For Each area In areaDic
    
        '****** ファイルにシート格納、なければコピーする
          On Error Resume Next
    
            Set AreaSh = BK.Worksheets(area)
        
          On Error GoTo 0
            
            'ないとき
            If AreaSh Is Nothing Then
                BK.Worksheets("").Visible = True
                BK.Worksheets("").Copy after:=BK.Worksheets(BK.Sheets.Count)
                Set AreaSh = BK.Worksheets(BK.Sheets.Count)
                AreaSh.name = area
                AreaSh.Cells(3, 2).Value = area
                BK.Worksheets("").Visible = False
            End If
    
        
        '今年ファイルパス
        areaPth = thisBK.path & "\" & areaDic.Item(area) & "\" & Year(toDay) & "" & area & num & ".xlsx"
        '今年ファイルを開いて変数格納
        Set areaBK = Workbooks.Open(areaPth)
        
        '==================== データ転記 ====================
        
        Call outputAreaData(areaBK, areaCol, AreaSh, toDay)
            
            '洋ナシ
            Set AreaSh = Nothing
        
        '==================== 月間構成比作成 ====================
        
        Call salesRatioMonth(BK, area, areaBK, areaCol, toDay)
                
            '今年ファイル、保存しないで閉じる
            areaBK.Close False
            Set areaBK = Nothing
        
        
        '***** 去年のブックを開く
        
            '昨年ファイルパス
            areaPth = thisBK.path & "\" & areaDic.Item(area) & "\" & Year(toDay) - 1 & "" & area & num & ".xlsx"
            
            On Error Resume Next
            
                '昨年ファイルを開いて変数格納
                Set areaBK = Workbooks.Open(areaPth)
                
            On Error GoTo 0
            
            '去年ないとき
            If areaBK Is Nothing Then
                GoTo noStore '次の
            End If

        '==================== 昨年売上取得 ====================
        
        '分類毎合計の取得、合算
        lastRevenue = areaRevenueGet(areaBK, area, areaCol, lastRevenue)
        
            '去年ファイル、保存しないで閉じる
            areaBK.Close False
            Set areaBK = Nothing
noStore:
    Next area '次
        
        Set areaDic = Nothing
        
        
    '***** 転記完了 *****
        
    '================== 支部合計計算 ====================
    
    Call allAreaCalc(BK, toDay)
    
    '================== 支部昨対計算 ====================
    
    Call allStoreYoYCalc(BK, toDay, lastRevenue)
    
    '================= 年間構成比転記 ===================
    
    Call yearRatio(BK, toDay)
    
    '***** 終わりの処理
    
        BK.Worksheets("支部").Activate
        
        'ファイルを保存して終了
        BK.Close True
        Set BK = Nothing
        
        thisBK.Worksheets("メイン").Activate
        Set thisBK = Nothing

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "処理が完了しました。", vbInformation

End Sub

'新規ブック作成
Private Sub BookCreate(ByRef thisBK As Workbook, toDay As Date)

Application.ScreenUpdating = False

    'パス格納用
    Dim pth As String
        pth = thisBK.path
                
    Dim yer As Integer
        yer = Year(toDay)
                    
    'フォルダ内にファイルがあるか探す
    If Dir(pth & "\\" & yer & ".xlsx") = "" Then
        
        'フォルダの有無を監視
        On Error GoTo createErr
        
        'なければ、原本ファイルをコピーし、ブックとして新規作成 -> コピー元path , 新規ファイルpath
        FileCopy pth & "\" & "原本.xlsx", pth & "\\" & yer & ".xlsx"
        
        On Error GoTo 0
    End If
    
    '処理完了
    Exit Sub
    
'※エラー対処
'ファイル、フォルダが見つからなかったら ================================
createErr:

        MsgBox "「原本.xlsx」または「フォルダ」が既定の場所に見つかりませんでした。" & vbCrLf & _
               "「原本.xlsx」や「フォルダ」の場所、名前を確認してから処理をやり直してください。", vbCritical
        
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        
        End  'プログラム中断
        
End Sub

'ブックからの取得列を返す
Private Function areaBookGetCol(ByRef toDay As Date) As Long

    Dim c As Long
    
        Select Case Month(toDay)
            Case 1, 4, 7, 10
                c = 8
            Case 2
                c = 12
            Case 5, 8, 11
                c = 13
            Case 3, 6, 9, 12
                c = 18
        End Select
    
    '戻り値
    areaBookGetCol = c

End Function

'ファイルの情報をファイルに転記
Private Sub outputAreaData(ByRef areaBK As Workbook, ByRef areaCol As Long, ByRef sh As Worksheet, toDay As Date)

    'areaBK ファイル、areaCol 取得列、sh シート、mon 作成月
       
    '出力列格納
    Dim Col As Integer
        Col = Month(toDay) + 3
        
    Dim areaR As Long
        areaR = 140      '行から
        
    Dim R As Long
        R = 5
        
    Dim i As Integer '0-7 中分類+合計、8分類
    Dim j As Integer '分類内、9項目
    
    
    'ファイル、シート格納
    Dim areaws As Worksheet
    Set areaws = areaBK.Worksheets(sh.name)
    
        'データ転記
        For i = 0 To 7
            For j = 0 To 8
            
                sh.Cells(R, Col).Value = areaws.Cells(areaR, areaCol).Value
                
                If j = 0 Then
                    R = R + 2
                Else
                    R = R + 1
                End If
                
                areaR = areaR + 1
            
            Next j '次の項目
        Next i '次の分類
        
        '処理日入力
        sh.Cells(94, Col).Value = Date
        
    Set areaws = Nothing

End Sub

'月間構成比作成
Private Sub salesRatioMonth(ByRef BK As Workbook, ByRef area As Variant, ByRef areaBK As Workbook, ByRef areaCol As Long, ByRef toDay As Date)

    'bk ファイル、area 名、areabk ファイル、areacol 取得列、today 作成日

    Dim MonSh As Worksheet
    Set MonSh = BK.Worksheets("月間構成比")
    
    'ファイルsh、最終行
    Dim R As Long
        R = MonSh.Cells(MonSh.Rows.Count, 2).End(xlUp).Row
        
    '合計列、最終行
    Dim sr As Long
        sr = MonSh.Cells(MonSh.Rows.Count, 7).End(xlUp).Row
        
    '******** 先月分かを精査
    
        '日付が今月分ではなかったら
        If MonSh.Cells(1, 5).Value <> toDay Then
            '日付を今月分に
            MonSh.Cells(1, 5).Value = toDay
        
            '範囲セルを一掃する     range("A3:Br") ：支店列
            MonSh.Range(MonSh.Cells(3, 1), MonSh.Cells(R + 1, 2)).Value = ""
            
            '合計列                 range("G6:Gsr")
            MonSh.Range(MonSh.Cells(6, 7), MonSh.Cells(sr + 1, 7)).Value = ""

        End If
    
    '******** 今月、現分　書き込み処理
                
    'ファイルsh、書き込み開始行の再取得
        R = MonSh.Cells(MonSh.Rows.Count, 2).End(xlUp).Row + 1
       sr = MonSh.Cells(MonSh.Rows.Count, 7).End(xlUp).Row + 1

    '名取得
    Dim areaName As String
        areaName = area
        
        '合計列に名を入力
        MonSh.Cells(sr, 7).Value = areaName

    Dim storeName As String
    Dim Ws As Worksheet
        
        
        'ファイルのシートを１つずつ回す
        For Each Ws In areaBK.Worksheets
        
            storeName = Ws.name
            
            '支店シートならば、合計を取得転記
            If storeName <> areaName And storeName <> "支店" Then
                With MonSh
                    .Cells(R, 1).Value = areaName '名
                    .Cells(R, 2).Value = storeName '支店名
                    .Cells(R, 3).Value = Ws.Cells(203, areaCol).Value '売上金額
                End With
                R = R + 1
            End If
        Next Ws
        
    Set MonSh = Nothing
        
End Sub

'昨年の項目合計を取得、計算
Private Function areaRevenueGet(ByRef areaBK As Workbook, ByRef area As Variant, ByRef areaCol As Long, revenue() As Currency) As Currency()

    'areaBK ファイル、area 名、areaCol 取得列、revenue 昨年合計合算配列

    Dim areaws As Worksheet
    Set areaws = areaBK.Worksheets(area)
    
    Dim i As Integer
    Dim R As Long
        R = 140     '行
    
        For i = 0 To UBound(revenue)
            '合算
            revenue(i) = revenue(i) + areaws.Cells(R, areaCol).Value
            R = R + 9
        Next i
    
    Set areaws = Nothing
    
    '戻り
    areaRevenueGet = revenue()

End Function

'支部シートの計算
Private Sub allAreaCalc(ByRef BK As Workbook, ByRef toDay As Date)
    
    'BK ファイル、taDay 作成日
        
    '取得、出力列
    Dim Col As Long
        Col = Month(toDay) + 3

    '合計取得、合算
    Dim cateCalc(7, 3) As Currency '8分類4項目
    
    '項目行
    Dim R As Long
        
    Dim i As Integer
    Dim j As Integer
    
    Dim Ws As Worksheet
    Dim shName As String
    
    ' ********************* 金額取得、合算
    
        'ファイル、シート全回し
        For Each Ws In BK.Worksheets
        
            shName = Ws.name
            R = 5 'から
            
            'シートならば、
            If shName <> "支部" And shName <> "" And shName <> "月間構成比" And shName <> "年間構成比" Then
            
                '分類を回す
                For i = 0 To UBound(cateCalc, 1)
                    '項目を回す
                    For j = 0 To UBound(cateCalc, 2)
                        
                        cateCalc(i, j) = cateCalc(i, j) + Ws.Cells(R, Col).Value
                        
                        '行の加算
                        Select Case j
                            Case 0, 2       '売上、客数の取得時
                                R = R + 3
                            Case 1, 3       '点数、粗利額の取得時
                                R = R + 2
                        End Select
                    
                    Next j '次の項目
                Next i '次の分類
            End If
        Next Ws '次のシート
        
    ' ********************* 支部シートに出力
    
    Set Ws = BK.Worksheets("支部")
    
        R = 5
        
        For i = 0 To UBound(cateCalc, 1)
            For j = 0 To UBound(cateCalc, 2)
                
                Ws.Cells(R, Col).Value = cateCalc(i, j)
                
                    Select Case j
                        Case 0, 2       '売上、客数の出力時
                            R = R + 3
                        Case 1, 3       '点数、粗利額の出力時
                            R = R + 2
                    End Select
            Next j
        Next i
        
    Set Ws = Nothing
        
End Sub

'支部の昨対比を計算する
Private Sub allStoreYoYCalc(ByRef BK As Workbook, ByRef toDay As Date, lastRevenue() As Currency)

    'bk ファイル、today 作成日、lastrevenue 昨年中分類売上

    Dim Ws As Worksheet
    Set Ws = BK.Worksheets("支部")
    
    Dim Col As Long
        Col = Month(toDay) + 3  '出力対象：C列以降
        
    Dim i As Integer
    Dim R As Long
        R = 5 '計算項目行   出力昨対行 = 7
    
    Dim val As Currency

        For i = 0 To UBound(lastRevenue)
                        
            '0除算ならばゼロで書き込む
            If lastRevenue(i) = 0 Then
                Ws.Cells(R + 2, Col).Value = 0
            Else
                '今年分の売上取得
                val = Ws.Cells(R, Col).Value
                Ws.Cells(R + 2, Col).Value = val / lastRevenue(i)
            End If
            
            R = R + 10
            
        Next i '次の分類
    
        '処理日入力
        Ws.Cells(85, Col).Value = Date
    
    Set Ws = Nothing
    
End Sub

'ファイル年間構成比処理
Private Sub yearRatio(ByRef BK As Workbook, ByRef toDay As Date)

    Dim MonSh As Worksheet
    Dim YearSh As Worksheet
    
    Set MonSh = BK.Worksheets("月間構成比")
    Set YearSh = BK.Worksheets("年間構成比")

    Dim Col As Long
        Col = Month(toDay) + 2  '出力対象：B列以降
        
        '構成比取得のため再計算
        MonSh.Calculate
        
        MonSh.Activate
    
    '****** 月間シートから支店名、支部構成比取得
       
    '入れ物
    Dim storeDic As Object
    Set storeDic = CreateObject("Scripting.Dictionary")
    
    '月間シート最終行
    Dim R As Long
        R = MonSh.Cells(MonSh.Rows.Count, 2).End(xlUp).Row
        
    Dim store As Variant
    Dim storeRng As Range
    
        Set storeRng = MonSh.Range(MonSh.Cells(3, 2), MonSh.Cells(R, 2))
           
            For Each store In storeRng
                '値の取得
                storeDic.Add store.Value, store.Offset(0, 3).Value
                
            Next store
        
        Set storeRng = Nothing

            
    '****** 年間シートに出力
    
        YearSh.Activate
    
    '新支店チェック
    Dim newStore As Boolean
        newStore = False
        
    Dim rng As Variant
        
    '年間シート最終行
        R = YearSh.Cells(YearSh.Rows.Count, 2).End(xlUp).Row
        
        '支店範囲
        Set storeRng = YearSh.Range(YearSh.Cells(3, 2), YearSh.Cells(R, 2))
   
        
            '取得支店を順に回す
            For Each store In storeDic
            
                '範囲に支店があるか判定
                If WorksheetFunction.CountIf(YearSh.Columns(2), store) > 0 Then
                               
                    'B列、支店名でフィルターをかける
                    YearSh.Columns(2).AutoFilter Field:=1, Criteria1:=store
                    
                        For Each rng In storeRng.SpecialCells(xlCellTypeVisible)
                            '支店の構成比出力
                            If rng.Value = store Then
                                YearSh.Cells(rng.Row, Col).Value = storeDic(store)
                                Exit For '次の支店
                            End If
                        Next rng
                
                'なければ、新支店入力
                Else
                    R = R + 1
                    YearSh.Cells(R, 2).Value = store                 '支店名
                    YearSh.Cells(R, Col).Value = storeDic(store) '構成比
                    newStore = True 'フラグ立て
                End If
                
            Next store
            
        'フィルター解除
        If YearSh.AutoFilterMode = True Then
            YearSh.AutoFilterMode = False
        End If
         
        Set storeRng = Nothing
        
            
    '****** 新支店入力があった時、表の並べ替えを行う
        
        If newStore Then
        
            '並べ替えのために再計算
            YearSh.Calculate
            
            '名格納
            Dim storeArr() As String
            Dim Ws As Worksheet
            Dim shName As String
            Dim i As Integer
            
                ReDim storeArr(BK.Worksheets.Count - 5)
                i = 0
                
                For Each Ws In BK.Worksheets
                
                    shName = Ws.name
                    '主要シートじゃなければ
                    If shName <> "支部" And shName <> "" And shName <> "月間構成比" And shName <> "年間構成比" Then
                        storeArr(i) = shName
                        i = i + 1
                    End If
                Next Ws
            
                    
            '最終行再取得
            R = YearSh.Cells(YearSh.Rows.Count, 2).End(xlUp).Row
            
            '並べ替え範囲
            Set storeRng = YearSh.Range(YearSh.Cells(3, 1), YearSh.Cells(R, Col))
            
                '並べ替える
                With YearSh.Sort
                    .SortFields.Clear
                    .SortFields.Add Key:=YearSh.Columns(1), SortOn:=xlSortOnValues, CustomOrder:=Join(storeArr, ",")
                    .SetRange storeRng
                    .Header = xlNo
                    .Apply
                End With
                   
            Set storeRng = Nothing
            Erase storeArr
        End If
        
    Set storeDic = Nothing
    Set MonSh = Nothing
    Set YearSh = Nothing

End Sub
