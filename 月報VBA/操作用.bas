Option Explicit
'メイン処理
Sub main()

'画面更新オフ
Application.ScreenUpdating = False
'Excel関数、自動計算オフ
Application.Calculation = xlCalculationManual


    Dim thisBK As Workbook
    
    Dim tws As Worksheet
    Dim thisS As Worksheet
    Dim data_ws As Worksheet
    Dim store_ws As Worksheet
        
    Set thisBK = ThisWorkbook
    Set tws = thisBK.Worksheets("メイン")
    Set thisS = thisBK.Worksheets("支店コード")
    
    
    Dim R As Long
        
    'メインブックシートのフィルタ判定と解除
        Call searchFilter(thisBK)

    
    '下準備 **************************************************************************
            
    'メインシートの必要個所に入力があるか判定 ==============================
        
        If tws.Cells(2, 1).Value = "" Then
            MsgBox "作成支部が入力されていません。入力してから実行してください。", vbCritical
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
        If tws.Cells(9, 6).Value = "" Then
            MsgBox "作成日付が入力されていません。入力してから実行してください。", vbCritical
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
    'データシートに不足がないか判定 ==============================
    
    '抽出シート名判定用
    Dim data_name() As Variant
        data_name = Array("分類", "粗利")
    
    'Each用
    Dim area As Variant
    Dim store As Variant
    Dim Ws As Worksheet
    Dim dt As Variant
        
    Dim shCheck(3) As Boolean
    Dim s As Integer
        s = 0
    
        For Each dt In data_name
            For Each Ws In thisBK.Worksheets
                'シートが存在したらフラグを立てる
                If Ws.name = dt Then
                    shCheck(s) = True
                    s = s + 1
                    Exit For
                End If
            Next Ws
        Next dt
                
        'シート精査
        For Each dt In shCheck
            If dt = False Then
                MsgBox "処理用のデータシートが必要数そろっていません。" & vbCrLf & "確認してから実行してください。", vbCritical
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                Exit Sub
            End If
        Next dt
        

    '出力支部取得準備 ==========================================
    
    tws.Activate
    
        'メインシート、入力支部最終行を格納
        R = tws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    '入力支部範囲を格納
    Dim areaRng As Range
    Set areaRng = tws.Range(tws.Cells(2, 1), tws.Cells(R, 1)) 'A2以降
    
     'Dictionaryオブジェクトの宣言
    Dim areaDic As Object
    Set areaDic = CreateObject("Scripting.Dictionary")
    
    
    '処理日付
    Dim toDay As Date
    
            '作成日付格納
            toDay = tws.Cells(9, 6).Value
    
    
    '処理ファイル数値の格納　引数：date
    Dim num As Integer
        num = search_month(toDay)

    
        'Dicに書き込む支部を入れる
        For Each area In areaRng
            If area.Value <> "" Then
            '初期化、要素の追加     .Add キー,値
             areaDic.Add area.Value, area.Offset(0, 1).Value    '支部名,コード+支部頭
            End If
        Next area
    
    
    '既存書き込み支部の判定 =======================================
    
    '支部範囲の数値入れる用
    Dim rc() As Integer
        rc() = areaCheck(tws, toDay)
    
    '処理済み支部を書き込む範囲
    Set areaRng = tws.Range(tws.Cells(3, 6), tws.Cells(6, 8))
        
        
    '本処理 **************************************************************************
        
    '=========== 取得した入力支部を元に、処理を開始する ===========
     
    '新規判定用
    Dim newcheck As Boolean  'book用
    Dim newStore As Boolean  '新支店用
    
        newcheck = False
        newStore = False
    
    '支部ブック用
    Dim areaBK As Workbook
    Dim areaInfo(2) As String
    
    '書き出し列
    Dim c As Long
    
    '支店用配列
    Dim stores() As String
    
            
    'エラーチェック
    Dim errcheck As Boolean
    
    Dim filePth As String
    
        R = 0
        

    '出力する支部を１つずつ回すfor　area = dicのキー
    For Each area In areaDic
    
        '既存ファイル精査      引数、Item(area):コード+支部頭文字 , area:支部名 , ファイルナンバー
        newcheck = BookCreate(areaDic.Item(area), area, num, toDay)
        
        filePth = thisBK.path & "\" & areaDic.Item(area) & "\" & Year(toDay) & "_" & area & num & ".xlsx"
                
        '支部ファイルを開いて変数格納
        Set areaBK = Workbooks.Open(filePth)

'エラー戻り
newCreateAreaSheet:

                '新規bookだったとき、支部シート名とセルを、該当支部名に変更する
                If newcheck Then
                    areaBK.Worksheets("支部").name = area
                    areaBK.Worksheets(area).Cells(3, 2).Value = area
                    'コピー用ひな型シートを表示
                    areaBK.Worksheets("支店").Visible = True
                End If
        
        'メインブックのメインシートをactive化しておく
        tws.Activate

        '新規支部bookが未完成のまま閉じられた時の監視
        On Error GoTo nothingArea
        
        '書き込み列を判定して、列を格納　　引数：書き込み日付 ,支部ブック
        c = out_col(toDay, areaBK.Worksheets(area))
        
        On Error GoTo 0
        
        
        '支店コードシートをアクティブ化する
        thisS.Activate
        
        '処理支部でフィルタ   支部列 D列
        thisS.Cells(1, 1).AutoFilter Field:=4, Criteria1:=area
        
        '支部支店を格納
        stores() = storeNames(thisS)
        
        
        '========================== 支店毎のデータを書き込んでいく処理
        
        
        '支店配列を１つずつ回す
        For Each store In stores
        
'エラー戻り
newCreateStoreSheet:

            '新規book/新規新支店アリならば、支店シートの作成,その他入力
            If newcheck Or newStore Then
                'areaBk.Activate
                areaBK.Worksheets("支店").Copy after:=areaBK.Worksheets(areaBK.Sheets.Count)
                Set store_ws = areaBK.Worksheets(areaBK.Sheets.Count)
                store_ws.name = store
                store_ws.Cells(3, 2).Value = store
                thisBK.Activate
            Else
               'エラーキャッチ  try-catch
               On Error GoTo storeWsError
               
                '支部ブック、転記用シートの格納
                Set store_ws = areaBK.Worksheets(store)
                    'メモ：周期途中で新支店が加わった時はエラーになるため、対処へ
               
               'エラー監視続行
               On Error GoTo 0
                
            End If
            
            
            '******* データ用配列を回して、支店の各数値を転記
            For Each dt In data_name
                
                'メインブック、抽出用シートの格納
                Set data_ws = thisBK.Worksheets(dt)
                
                '粗利シートかを判定
                If InStr(data_ws.name, "粗利") > 0 Then
                    '粗利シート：支店でフィルタ   支店列 A列
                    data_ws.Cells(1, 1).AutoFilter Field:=1, Criteria1:=store
                Else
                    '分類/支店昨対シート：支店でフィルタ   支店列 C列
                    data_ws.Cells(1, 1).AutoFilter Field:=3, Criteria1:=store
                End If
            
                '転記処理  引数：各シート参照、書き込み列
                errcheck = outData(store_ws, data_ws, c)
                
                '情報がないとき、処理を中断する
                If errcheck Then
                    '支部bookは保存せず閉じる
                    Application.DisplayAlerts = False
                    areaBK.Close False
                    Application.DisplayAlerts = True
                    Application.ScreenUpdating = True
                    Application.Calculation = xlCalculationAutomatic
                    Exit Sub
                End If
                
                Set data_ws = Nothing
                
                
            Next dt '次の分類データシートへ
            
            '処理日の記述
            store_ws.Cells(230, c).Value = Date
            
            Set store_ws = Nothing
                            
        Next store  'storesのforEach　次の支店シートへ
        
        
        '=================== 支部シート合計の計算処理 ===================
        
        
        '支部ws、各分類合計の計算
        Call areaCalc(areaBK, area, c)
                
        
        '★昨対の入力
        '昨年bookの記述ならば、昨対は不要。今年度の処理であれば、昨対処理を呼び出す
        If Year(toDay) = Year(Date) Then
            Call lastYearCalc(areaBK, area, c)
        ElseIf toDay = Year(Date) - 1 & "12/31" And Month(Date) = 1 Then
            '1月時、12月末の処理
            Call lastYearCalc(areaBK, area, c)
        End If
        
                                
        '新規book or 新支店加入時は、ひな型シートを非表示に
        If newcheck Or newStore Then areaBK.Worksheets("支店").Visible = False
        
        'シート表示は支部シートを先頭に
        areaBK.Worksheets(area).Activate
        
            'areaBkの情報を取得しておく
            areaInfo(0) = area
            areaInfo(1) = areaBK.name
            areaInfo(2) = areaBK.path

        '支部ブックを閉じる
        areaBK.Close savechanges:=True
        
        Erase stores
        
        
        '送信用bookの作成 =========================
        
        If Year(toDay) = Year(Date) Then
            Call areaBookCopy(areaInfo, toDay)
        ElseIf toDay = Year(Date) - 1 & "12/31" And Month(Date) = 1 Then
            '1月時、12月末の処理
            Call areaBookCopy(areaInfo, toDay)
        End If
       
        Set areaBK = Nothing
        Erase areaInfo
        
        '完了支部の書き込み =====================
        areaRng(rc(0), rc(1)) = area
        tws.Cells(2 + R, 1).Value = ""
        
            rc(1) = rc(1) + 1
                R = R + 1
                
            If rc(1) = 4 Then
                rc(1) = 1
                rc(0) = rc(0) + 1
            End If
    
    Next area  'areaDic 次の支部へ
    
        Erase rc
        Erase data_name
        Set areaRng = Nothing
    
    '終わりの処理　*********************************************************
    
    'メインブックシートのフィルタ判定と解除
    Call searchFilter(thisBK)
    
    tws.Activate
    
    '作成日の入力
    tws.Cells(2, 6).Value = toDay
    
    '支部表上の、全支部処理が終了したら指定日付クリア
    If tws.Cells(26, 17).Value = tws.Cells(27, 17).Value Then tws.Cells(9, 6).Value = ""
                          
                          
    Set areaDic = Nothing
    Set thisBK = Nothing
    Set tws = Nothing
    Set thisS = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    
    MsgBox "処理が完了しました。", vbInformation
    
    '本処理を終了
    Exit Sub

'※エラー対処
'新支店時、エラー処理 ========================================
storeWsError:
                '新規支店が周期途中にできたとき
                newStore = True
                areaBK.Worksheets("支店").Visible = True
                err.Clear
                
                     Application.ScreenUpdating = False
                
                'メイン処理に戻る
                Resume newCreateStoreSheet
                

'bookに支部シートがなかったとき、エラー処理 ========================================
nothingArea:
                newcheck = True
                err.Clear
                
                    Application.ScreenUpdating = False
                
                Resume newCreateAreaSheet
End Sub

'書き込み支部のアドレスを返す
Private Function areaCheck(ByRef tws As Worksheet, toDay As Date) As Integer()

    
    '配列要素用
    Dim i As Integer
    Dim j As Integer
    Dim v As Variant
    
    Dim rc(1) As Integer
    Dim outRng As Range
    
    Set outRng = tws.Range(tws.Cells(3, 6), tws.Cells(6, 8))
    
    i = 1
    j = 1
    
    rc(0) = i
    rc(1) = j
    
     '書き込み支部を入力
     If tws.Cells(2, 6).Value <> toDay Then
        '日付が違ったら
        
        '既存支部は全て空白に
            outRng.Value = ""
        
     Else
        '日付が一緒だったら
        '既に処理済みの支部を残すため、入力可能range範囲を確かめる
        For Each v In outRng
                                     'i 縦　j 横　[i][j] 4,3まで
            If v.Value <> "" Then
                '記入済みであれば
                j = j + 1
                If j = 4 Then
                    j = 1
                    i = i + 1
                End If
            Else
                '空白を見つけたら
                rc(0) = i
                rc(1) = j
                Exit For 'ループ抜ける
            End If
        Next v
     End If
    
    '戻り値
    areaCheck = rc()
    
    Set outRng = Nothing

End Function

'新規で月ブック作成
Function BookCreate(area As Variant, name As Variant, num As Integer, toDay As Date) As Boolean

Application.ScreenUpdating = False

    'パス格納用
    Dim pth As String
        pth = ThisWorkbook.path
        
    Dim check As Boolean
        check = False
        
    Dim yer As Integer
        
        yer = Year(toDay)
        
            
    '支部フォルダ内に支部ファイルがあるか探す
    If Dir(pth & "\" & area & "\" & yer & "_" & name & num & ".xlsx") = "" Then
        
        'フォルダの有無を監視
        On Error GoTo createErr
        
        'なければ、原本ファイルをコピーし、支部ブックとして新規作成 -> コピー元path , 新規ファイルpath
        FileCopy pth & "\" & "原本.xlsx", pth & "\" & area & "\" & yer & "_" & name & num & ".xlsx"
        check = True
        
        On Error GoTo 0
    End If
    
    '判定を返す
    BookCreate = check
    Exit Function
    
'※エラー対処
'支部フォルダが見つからなかったら ================================
createErr:

        MsgBox "「原本.xlsx」または「" & area & "フォルダ」が既定の場所に見つかりませんでした。" & vbCrLf & _
               "「原本.xlsx」や「" & area & "フォルダ」の場所、名前を確認してから処理をやり直してください。", vbCritical
        
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        
        End  'プログラム中断
        
End Function

'書き込み列を返す関数
Private Function out_col(out_date As Date, ByRef aws As Worksheet) As Long

    '書き込み日付　out_date
    '支部シート　aws
            
    '書き込み用book、シートをアクティブ化
    aws.Activate
    
    '月セルが空白のとき、日付を入力する処理
    If aws.Cells(3, 4).Value = "" Then
        Dim setDate As String
        Select Case Month(out_date)
            Case 1, 2, 3
                setDate = Year(out_date) & "/1/1"
            Case 4, 5, 6
                setDate = Year(out_date) & "/4/1"
            Case 7, 8, 9
                setDate = Year(out_date) & "/7/1"
            Case 10, 11, 12
                setDate = Year(out_date) & "/10/1"
            End Select
        
        aws.Cells(3, 4).Value = CDate(setDate)
        aws.Calculate
    End If

    
    Dim dayRng As Range
  
        '入力月判定し、日付範囲を格納
        Select Case Month(out_date)
            Case Month(aws.Cells(3, 4).Value)
                Set dayRng = aws.Range(Cells(4, 4), Cells(4, 8))
            Case Month(aws.Cells(3, 9).Value)
                Set dayRng = aws.Range(Cells(4, 9), Cells(4, 13))
            Case Month(aws.Cells(3, 14).Value)
                Set dayRng = aws.Range(Cells(4, 14), Cells(4, 18))
        End Select
        
        
    '書き込み日付の判定
    'Dim d_arr() As Variant
    '    d_arr = Array(7, 14, 21, 28)
        
    Dim out_d As Integer
    
        If Day(out_date) < 8 Then
            out_d = 7
        ElseIf Day(out_date) < 15 Then
            out_d = 14
        ElseIf Day(out_date) < 22 Then
            out_d = 21
        ElseIf Day(out_date) < 29 Then
            out_d = 28
        Else
            Dim dEnd As Variant
                dEnd = DateSerial(Year(out_date), Month(out_date) + 1, 0)
                dEnd = CStr(dEnd)
                dEnd = Right(dEnd, 2)
                out_d = CInt(dEnd)
        End If
        
               
    Dim d As Variant
    Dim Col As Long
        
        On Error GoTo dayRangeErr
        
        '入力期間日付の判定
        For Each d In dayRng
            '値が一致したセルの列を返す
            If d.Value = out_d Then
                    Col = d.Column
                    Exit For
            End If
        Next d
        
        On Error GoTo 0
        
        '日付指定に誤りがあるとき
        If Col = 0 Then
        
            MsgBox "出力列の取得でエラーが発生しました。" & vbCrLf & _
                   "＜手動用＞に作成したい週の翌日日付で入力し、再実行してください。" & vbCrLf & _
                   "処理を中断します。", vbCritical
            
            '支部bookを保存せず閉じる
            Application.DisplayAlerts = False
            aws.Parent.Close False
            Application.DisplayAlerts = True
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        
            End  'プログラム全終了（中断）
        
        End If
                
  
    '出力列を返す
    out_col = Col
    Exit Function

'エラー対処
'range参照の取得漏れ ==================================================
dayRangeErr:
            
            MsgBox "出力月の判定でエラーが発生しました。" & vbCrLf & _
                   "＜手動用＞に作成したい週の翌日日付で入力し、再実行してください。" & vbCrLf & _
                   "処理を中断します。", vbCritical
            
            '支部bookを保存せず閉じる
            Application.DisplayAlerts = False
            aws.Parent.Close False
            Application.DisplayAlerts = True
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        
            End  'プログラム全終了（中断）
    
End Function

'ファイル数値判定
Function search_month(mydate As Date) As Integer
        
    Dim m As Integer
    Dim this_m As Integer

        m = Month(mydate)
    
    '該当ファイル数値判定
    Select Case m
        Case 1, 2, 3
            this_m = 1
        Case 4, 5, 6
            this_m = 2
        Case 7, 8, 9
            this_m = 3
        Case 10, 11, 12
            this_m = 4
    End Select

    search_month = this_m

End Function

'支部範囲の支店名配列を返す関数
Private Function storeNames(ByRef Ws As Worksheet) As String()
    
    'ws 支店コードsheet
    
    Dim R As Long
    Dim rng As Range
    Dim var As Variant
    
    Dim stores() As String
    Dim i As Integer

        
        '範囲の、最終行を格納
        R = Ws.Cells(Rows.Count, 5).End(xlUp).Row

        '支店名範囲を格納
        Set rng = Ws.Range(Ws.Cells(2, 5), Ws.Cells(R, 5)) 'E開始行:E最終行
    
    '要素数を入れる
    ReDim stores(rng.SpecialCells(xlCellTypeVisible).Count - 1)
    
    i = 0
    
    '範囲を回して支店名を配列に格納     可視セルのみで回す指定
    For Each var In rng.SpecialCells(xlCellTypeVisible)
    
        stores(i) = var.Value
                
        i = i + 1
        
    Next var
    
    storeNames = stores()
    
    Set rng = Nothing

End Function

'シートデータ転記処理
Function outData(ByRef store_ws As Worksheet, ByRef data_ws As Worksheet, ByRef outCol As Long) As Boolean

    'store_ws 書き込み用支店ws , data_ws 抽出ws , outCol 書き出し列
        
    Dim c As Long '最終列
    Dim R As Long '最終行
    
    Dim i As Integer
    Dim y As Long
    
    Dim outRng As Range
    Dim rng As Variant
    
    Dim str As String
        
            
    data_ws.Activate
    
    '最終行の格納
    R = data_ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    i = 1
    
    
    If InStr(data_ws.name, "粗利") > 0 Then
    
    '===================== 粗利シート時の処理 =====================
    
    'rangeで抽出シート側範囲格納
    
        '最終列の格納
        c = data_ws.Cells(1, Columns.Count).End(xlToLeft).Column
        '粗利シート、数値のみの格納
        Set outRng = data_ws.Range(data_ws.Cells(2, 2), data_ws.Cells(R, c)) 'range("B2:cr")
        
        Dim val As String
        Dim j As Integer
            j = 0
        
        For Each rng In outRng.SpecialCells(xlCellTypeVisible)
                                      
                        
            '支店がなかったら
            If VarType(rng.Value) = vbString Then
                MsgBox "「" & store_ws.name & "」の[" & data_ws.name & "]データが見つかりませんでした。" & vbCrLf & _
                       "確認後、この支部の処理からやり直してください。", vbCritical
                outData = True
                Exit Function
            End If
            
            '****** データ書き込み
            
            If i = 1 Then
                '書き出し行の判定
                val = data_ws.Cells(1, 2 + j).Value
                str = Left(val, InStr(val, "_") - 1)
                
                Select Case str
                            Case "食品"
                                y = 66
                            Case Else
                                '余分な列、不備があったら continue
                                GoTo nothingCase
                        End Select
            End If
            
            If rng.Value = "" Then
                store_ws.Cells(y, outCol).Value = 0
            Else
                store_ws.Cells(y, outCol).Value = rng.Value
            End If
            
            If i = 2 Then
                i = 1
            Else
                i = i + 1
                y = y + 1
            End If
nothingCase:
            j = j + 1 '次の項目精査

        Next rng
        
    Else
    
    '===================== 累計シート時の処理 =====================
    
    'rangeで抽出シート側範囲格納

        '最終列の格納
        c = data_ws.Cells(2, Columns.Count).End(xlToLeft).Column
        '分類/昨対シートならば、分類判定のためにA列も含める
        Set outRng = data_ws.Range(data_ws.Cells(2, 1), data_ws.Cells(R, c)) 'range("A2:cr")
                

        For Each rng In outRng.SpecialCells(xlCellTypeVisible)
        
            'iは1 - 10 をローテする
            Select Case i
            
                Case 1
                
                    str = rng.Value
                    
                    Select Case str
                        Case "食品"
                            y = 59
                        Case Else
                        'フィルターがかからず、データがないとき
                            MsgBox "「" & store_ws.name & "」の[" & data_ws.name & "]データが見つかりませんでした。" & vbCrLf & _
                            "確認後、この支部の処理からやり直してください。", vbCritical
                            outData = True
                            Exit Function
                    End Select  'case 1内のcase終わり
                
                Case 2, 3
                    '何もしない
                
                Case 4, 6, 7, 8, 9, 10
                    'データ出力
                        store_ws.Cells(y, outCol).Value = rng.Value
                        y = y + 1
                Case 5
                    '昨対の出力
                        store_ws.Cells(y, outCol).Value = rng.Value / 100
                        y = y + 1
            End Select
            
            'iの初期化
            If i = 10 Then
                i = 1
            Else
                i = i + 1
            End If
            
        Next rng
        
    End If
        
    '問題なく終了
    outData = False
    
End Function

'支部シート分類計算
Sub areaCalc(ByRef areaBK As Workbook, ByRef area As Variant, ByRef outCol As Long)

    'areaBk 支部ブック参照、area 処理支部名参照、outCol 書き込み列参照
    
'Excel関数再計算
Application.Calculate
     
    Dim Ws As Worksheet
    Dim str As String
    
    Dim i As Integer
    Dim j As Integer
    Dim y As Long '行加算用
    
    Dim val As String
    
    
    '数値を入れる二次元配列
    Dim cate_scores(22, 3) As Currency
    
    
        '支部ブックの全シートを１つずつ回す
        For Each Ws In areaBK.Worksheets
                    
            str = Ws.name  'シートネーム格納
              y = 5        '取得開始行
            
            '支部名、支店ひな型シート以外なら
            If str <> area And str <> "支店" Then
            
                
            '分類0-21回転  2,2,3,2
                
                '各分類の値を加算する       (二次元配列なので,引数1)
                For i = 0 To UBound(cate_scores, 1)
                                                 
                                
                    '加算対象セル 4つ　売上、点数、客数、粗利額
                    For j = 0 To 3
                    
                        val = Ws.Cells(y, outCol).Value
                        
                        '各分類の合計を加算する
                        If val <> "" Then
                            cate_scores(i, j) = cate_scores(i, j) + Ws.Cells(y, outCol).Value
                        End If
                        
                        If j <> 2 Then
                           y = y + 2
                        Else
                           y = y + 3
                        End If
                        
                    Next j '次のセルへ
                                   
                Next i '次の分類へ
            End If  'シート判定
        Next Ws  '次のシートへ
        
        '================ 支部シートに書き出し
       
    '支部シート格納
    Set Ws = areaBK.Worksheets(area)
        
        y = 5 '書き出し行
        
        For i = 0 To UBound(cate_scores, 1)
        
                
                '売上、粗利額を順に書き出し
                For j = 0 To 3
                        
                    Ws.Cells(y, outCol).Value = cate_scores(i, j)
                        
                    If j <> 2 Then
                        y = y + 2
                    Else
                        y = y + 3
                    End If
                        
                Next j '次のセルへ

        Next i '次の分類へ
        
        '処理日の記述
        Ws.Cells(221, outCol).Value = Date
        
Set Ws = Nothing
                
End Sub

'送信用の支部ブックを作成する
Private Sub areaBookCopy(areaInfo() As String, ByRef out_date As Date)

    'areaInfo(0) 支部名、areaInfo(1) 支部book名、areaInfo 支部bookパス、outDay 作成日
    
    '送信用日付を判定する ==================================
                
        Dim out_d As Integer
                
        If Day(out_date) < 8 Then
            out_d = 7
        ElseIf Day(out_date) < 15 Then
            out_d = 14
        ElseIf Day(out_date) < 22 Then
            out_d = 21
        ElseIf Day(out_date) < 29 Then
            out_d = 28
        Else
            Dim dEnd As Variant
                dEnd = DateSerial(Year(out_date), Month(out_date) + 1, 0)
                dEnd = CStr(dEnd)
                dEnd = Right(dEnd, 2)
                out_d = CInt(dEnd)
        End If
           
        
        Dim dateStr As String
            '日付を文字列として格納
            dateStr = Year(out_date) & "." & Month(out_date) & "." & out_d
        
    
    'ファイルコピー処理 ====================================
    
    Dim sendFileName As String
        '送信用ファイル名を格納     '支部名
        sendFileName = "【" & areaInfo(0) & "】報告書" & dateStr & ".xlsx"
    
        '作成した支部ファイルをコピーし、送信用ファイルとして新規作成 -> 今期作成した支部ファイルpath名 , 送信ファイルpath名（絶対パス）
        FileCopy areaInfo(2) & "\" & areaInfo(1), areaInfo(2) & "\" & sendFileName
        
            'areaInfo(2) : path　、　areaInfo(1) : 支部ブック名

End Sub

'フィルター判定と解除
Private Sub searchFilter(ByRef thisBK As Workbook)

    Dim Ws As Worksheet

    For Each Ws In thisBK.Worksheets
        If Ws.AutoFilterMode = True Then
            Ws.AutoFilterMode = False  'フィルタされてたら解除
        End If
    Next Ws '次シート

End Sub

'昨対bookを開いて、今年度の昨対を計算する
Private Sub lastYearCalc(ByRef areaBK As Workbook, ByRef area As Variant, ByRef outCol As Long)

    'areaBK 今年支部book参照、area 支部名、outCol 書き出し列

    Dim Ws As Worksheet
    Dim yer As Integer
    Dim nameStr As String
    Dim fileName As String
    
        fileName = areaBK.name
        
        '年度の切り出し     yyyy
        yer = CInt(Left(fileName, InStr(fileName, "_") - 1)) - 1
        
        '名前の切り出し     name.xlsx
        nameStr = Right(fileName, Len(fileName) - InStr(fileName, "_"))
        
    
    Dim filePth As String
        
        filePth = areaBK.path & "\" & yer & "_" & nameStr
        
    Dim lastBK As Workbook
        
        '前期bookの有無監視
        On Error GoTo nothingLastBook
        
        '昨対用ファイルを開いて変数格納
        Set lastBK = Workbooks.Open(filePth)
        
        On Error GoTo 0
    
    Dim str As String
    Dim i As Integer
    Dim y As Long '行加算
    
    'Dictionary 各支店の売上情報を入れる
    Dim storeDic As Object
    Set storeDic = CreateObject("Scripting.Dictionary")
    
    '売上入れる配列 支部ws = 25項目 , 支店ws = 1項目
    Dim revenues() As Currency
    
    
    '============== 前期のbookから、各シートの売上額を取得する ==============
            
    'シートごとに売上格納処理
    For Each Ws In lastBK.Worksheets
    
        
        str = Ws.name
        
        'ひな型,支部シート以外ならば、支店合計のみrevenueに取得
        If str <> "支店" And str <> area Then
             
             '1項目
             ReDim revenues(0)
   
             y = 212 '支店合計行
                
                    revenues(0) = Ws.Cells(y, outCol).Value
                
        '支部シートならば、全項目revenueに取得
        ElseIf str = area Then
               
                '24項目
                ReDim revenues(23)
               
                y = 5 '一般食品計行から
                
               '分類項、24回転
                For i = 0 To UBound(revenues)
                    revenues(i) = Ws.Cells(y, outCol).Value
                    y = y + 9
                Next i  '次の項目へ
        End If
        
        '辞書に格納
        storeDic.Add str, revenues

    Next Ws  'lastBk.次のシートへ
    
    
    '昨対用bookを保存せず閉じる
    Application.DisplayAlerts = False
    lastBK.Close False
    Application.DisplayAlerts = True

    '昨対bookは洋ナシ
    Set lastBK = Nothing
    
    
    '======================== 今期bookに各シートの昨対比を計算する ========================
    
    Dim sd As Variant
    Dim sale As Currency
    
    '支店辞書を１つずつ回す
    For Each sd In storeDic
        
        '前期で閉店した支店があったとき用に try-catch
        On Error GoTo closedStore
        
            '支店シートset
            Set Ws = areaBK.Worksheets(sd)
        
        On Error GoTo 0
        
        str = Ws.name
        
        '動的配列の初期化
        Erase revenues
        
        'itemを配列に入れ直す
        revenues = storeDic(sd)
        
        '支部シートのとき
        If str = area Then
        
            '一般食品売上行から
            y = 5
        
            '各項目と計算
            For i = 0 To UBound(revenues)
            
                If y = 212 Then Ws.Cells(y, outCol).Calculate
            
                '0で除算はできないので、配列要素を判定しておく
                If revenues(i) = 0 Then
                    '0除算であれば、0%で入力する
                    Ws.Cells(y + 1, outCol).Value = 0
                Else
                    sale = Ws.Cells(y, outCol).Value '今期売上の取得
                    Ws.Cells(y + 1, outCol).Value = sale / revenues(i) '今期売上/前期売上
                End If
                
                y = y + 9

            Next i '次の項目
            
        '支店シートの時
        ElseIf str <> area And str <> "支店" Then
        
            '売上行
            y = 212

            '支店合計のみ計算
                '0で除算はできないので、配列要素を判定しておく
                If revenues(0) = 0 Then
                    '0除算であれば、0%で入力する
                    Ws.Cells(y + 1, outCol).Value = 0
                Else
                    sale = Ws.Cells(y, outCol).Value '今期売上の取得
                    Ws.Cells(y + 1, outCol).Value = sale / revenues(0) '今期売上/前期売上
                End If
        End If
closeNext:
    Next sd '次の支店へ

Set Ws = Nothing
Set storeDic = Nothing
Erase revenues

Exit Sub 'エラーなし、当処理を正常終了

'※エラー対処
'前周期のbookが見つからなかったとき ===================================================
nothingLastBook:

        '今年からの新支部book時など
        err.Clear
        Application.ScreenUpdating = False
        
        Exit Sub  '昨対処理を抜ける

'前期には支店があったのに、今回なかったとき ===========================================
closedStore:

        err.Clear
        Application.ScreenUpdating = False
        GoTo closeNext

End Sub

