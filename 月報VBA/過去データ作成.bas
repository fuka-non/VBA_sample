Option Explicit
'過去のデータをまとめて作る処理
Sub createLastYearBook()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    Dim thisBK As Workbook
    Set thisBK = ThisWorkbook
        
    Dim dataBK As Workbook
    Dim areaBK As Workbook
    Set areaBK = Nothing

    Dim tws As Worksheet
    Dim thisS As Worksheet
    Dim data_ws As Worksheet
    Dim store_ws As Worksheet
        
    Set tws = thisBK.Worksheets("まとめ用")
    Set thisS = thisBK.Worksheets("支店コード")


    'データフォルダーを選択してもらう ============================================
    Dim folderPath As String
    
      With Application.FileDialog(msoFileDialogFolderPicker)
       '
        .InitialFileName = thisBK.path
        .Title = "データファイルが入ったフォルダーを選択"
       
       'キャンセル時は処理を抜ける
       If .Show = 0 Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
       End If
       'フォルダパス格納
       folderPath = .SelectedItems(1)
      End With
         
            
    'まとめ用シートの必要個所に入力があるか判定 ==============================
        
        If tws.Cells(2, 1).Value = "" Then
            MsgBox "作成が入力されていません。入力してから実行してください。", vbCritical
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
        
    '入力された転記の格納 ========================================================
    
    Dim r As Long
        'まとめ用シート、入力最終行を格納
        r = tws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    '入力範囲を格納
    Dim areaRng As Range
    Set areaRng = tws.Range(tws.Cells(2, 1), tws.Cells(r, 1)) 'A2以降
    
    Dim area As Variant
    
     'Dictionaryオブジェクトの宣言
    Dim areaDic As Object
    Set areaDic = CreateObject("Scripting.Dictionary")
            
        'Dicに書き込むを入れる
        For Each area In areaRng
            If area.Value <> "" Then
            '初期化、要素の追加     .Add キー,値
             areaDic.Add area.Value, area.Offset(0, 1).Value    '名,コード+頭
            End If
        Next area
    
    '処理済みを書き込む範囲
    Set areaRng = tws.Range(tws.Cells(3, 6), tws.Cells(6, 8))
    Dim rng As Variant
    
                
    'フォルダー内ファイルを格納 ============================
    
    Dim dateSplit() As String
    Dim tmp As String
    Dim toDay As Date
    
    Dim fso, file, files
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set files = fso.GetFolder(folderPath).files
    
    Dim num As Integer

            
    '=========== 取得した入力に処理開始 ===========
     
    '新規判定用
    Dim newcheck As Boolean  'book用
    Dim newStore As Boolean  '
    
        newcheck = False
        
    '書き出し列
    Dim c As Long
    
    '支店用辞書
    Dim stores As Object
    Set stores = CreateObject("Scripting.Dictionary")
    
    Dim store As Variant
    Dim storeDate As Date
    
    '抽出シート名判定用
    Dim data_name() As Variant
        data_name = Array("分類", "")
    Dim dt As Variant
            
    'エラーチェック
    Dim errcheck As Boolean
    
    'areaファイルパス
    Dim areaPth As String
    
    Dim areaFileName As String
        
        r = 0
    
            
    '出力するを１つずつ回す
    For Each area In areaDic
    
           
        '処理でフィルタ   列 D列
        thisS.Cells(1, 1).AutoFilter Field:=4, Criteria1:=area
        thisS.Activate
    
        '支店を格納
        Set stores = storeDatas(thisS)

                
            '====================== フォルダ内を１つずつ回す ======================
            
            For Each file In files

                '********* データファイルを開く
                                
                'データファイルを開いて変数格納
                Set dataBK = Workbooks.Open(file)
                
                'データファイル内シートの体裁を整える
                Call fixUpData(dataBK)
                
                'ファイル名から転記日付を取得する
                dateSplit = Split(dataBK.name, ".")
                tmp = dateSplit(0) & "/" & dateSplit(1) & "/" & dateSplit(2)
                toDay = CDate(tmp)
                               
                               
                '********* ファイルを開く
                
                     '作成周期判定
                    num = search_month(toDay)
                
                    '既存のファイル精査
                    newcheck = bookCreate(areaDic.Item(area), area, num, toDay)
                    
                    'ファイル名の格納
                    areaFileName = Year(toDay) & "_" & area & num & ".xlsx"
                    
                        'ファイルが開かれているか精査
                        If areaBKcheck(areaFileName) = False Then
                            
                            '開いているファイルが既にあれば、保存してとじる
                            If Not (areaBK Is Nothing) Then areaBK.Close True
                            
                            '開かれていなければファイルを開く
                            areaPth = thisBK.path & "\" & areaDic.Item(area) & "\" & Year(toDay) & "_" & area & num & ".xlsx"
                            'ファイルを開いて変数格納
                            Set areaBK = Workbooks.Open(areaPth)
                            
                        End If
               
'エラー戻り
newCreateAreaSheet:

                '新規bookだったとき、シート名とセルを、該当名に変更する
                If newcheck Then
                    areaBK.Worksheets("").name = area
                    areaBK.Worksheets(area).Cells(3, 2).Value = area
                    'コピー用ひな型シートを表示
                    areaBK.Worksheets("支店").Visible = True
                End If
        

        '新規bookが未完成のまま閉じられた時の監視
        On Error GoTo nothingArea2
        
            '書き込み列を判定して、列を格納　　引数：書き込み日付 ,ブック
            c = lastOut_col(toDay, areaBK.Worksheets(area))
        
        On Error GoTo 0
        
        
        '=============== 支店毎のデータを書き込んでいく処理 ===============
               
        '格納支店を１つずつ回す
        For Each store In stores
            
            '支店の営業開始日を格納
            storeDate = stores(store)
            newStore = False
        
'エラー戻り
newCreateStoreSheet:

            '新規book/新店アリ、かつ、データ起算日より支店営業日が古ければ、支店シートの作成,支店名入力
            If (newcheck Or newStore) And (storeDate < toDay) Then
                areaBK.Worksheets("支店").Copy after:=areaBK.Worksheets(areaBK.Sheets.Count)
                Set store_ws = areaBK.Worksheets(areaBK.Sheets.Count)
                store_ws.name = store
                store_ws.Cells(3, 2).Value = store
                
            '新規bookではない、データあり
            ElseIf storeDate < toDay Then
               'エラーキャッチ  try-catch
               On Error GoTo storeWsError2
               
                    'ブック、転記用シートの格納
                    Set store_ws = areaBK.Worksheets(store)
                    'メモ：周期途中で新店が加わった時はエラーになるため、対処へ
               
               'エラー監視続行
               On Error GoTo 0
                                   
            '新支店、このデータ起算日時点では営業していない
            Else
                GoTo nextStore  '次の支店へ
            End If
               
            
            '******* データ用配列を回して、支店の各数値を転記
            For Each dt In data_name
                
                'データブック、抽出用シートの格納
                Set data_ws = dataBK.Worksheets(dt)
                
                'シートかを判定
                If InStr(data_ws.name, "") > 0 Then
                    'シート：支店でフィルタ   支店列 A列
                    data_ws.Cells(1, 1).AutoFilter Field:=1, Criteria1:=store
                Else
                    '分類シート：支店でフィルタ   支店列 C列
                    data_ws.Cells(1, 1).AutoFilter Field:=3, Criteria1:=store
                End If
            
                    '転記処理  引数：各シート参照、書き込み列
                    errcheck = outData(store_ws, data_ws, c)
                
                '情報がないとき、処理を中断する
                If errcheck Then
                    'bookは保存せず閉じる
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
                
nextStore:
        Next store  '次の支店シートへ
        
        '=================== シート合計の計算処理(データ毎処理) ===================
               
        'ws、各分類合計の計算
        Call areaCalc(areaBK, area, c)

        '******************* データファイル切り替え
        
            'データファイルを閉じる
            Application.DisplayAlerts = False
            dataBK.Close False
            Application.DisplayAlerts = True
                                
            '情報リセット
            Erase dateSplit
            Set dataBK = Nothing
            
        Next file '次のデータファイル
        
        '=================== 転記完了 ===================
        
        'ひな型シートを非表示に
        If areaBK.Worksheets("支店").Visible = True Then areaBK.Worksheets("支店").Visible = False
        
        'シート表示はシートを先頭に
        areaBK.Worksheets(area).Activate
        
        'ブックを保存して閉じる
        areaBK.Close savechanges:=True
        
        Set stores = Nothing
        Set areaBK = Nothing
              
            tws.Activate

        '完了の書き込み =====================
        If tws.Cells(2, 6).Value <> Date Then
            areaRng.Value = ""
            tws.Cells(2, 6).Value = Date
        End If
        
        For Each rng In areaRng
            If rng.Value = "" Then
                rng.Value = area
                Exit For
            End If
        Next rng
        
        tws.Cells(2 + r, 1).Value = ""
        r = r + 1
        
    Next area '次の
    
    
    '終わりの処理　*********************************************************
    
    thisS.AutoFilterMode = False  'フィルタ解除
    
    Erase data_name
    Set areaRng = Nothing
    Set areaDic = Nothing
    Set thisBK = Nothing
    Set tws = Nothing
    Set thisS = Nothing
    Set fso = Nothing
    Set files = Nothing

    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
    
    MsgBox "処理が完了しました。", vbInformation

    '本処理を終了
    Exit Sub

'※エラー対処
'新店時、エラー処理 ========================================
storeWsError2:
                '新規支店が周期途中にできたとき
                newStore = True
                areaBK.Worksheets("支店").Visible = True
                err.Clear
                
                     Application.ScreenUpdating = False
                
                'エラー監視続行
                On Error GoTo 0
                
                '本処理に戻る
                Resume newCreateStoreSheet
                

'bookにシートがなかったとき、エラー処理 ========================================
nothingArea2:
                newcheck = True
                err.Clear
                
                    Application.ScreenUpdating = False
        
                On Error GoTo 0
                
                Resume newCreateAreaSheet
End Sub

'範囲の支店名、開店日を返す関数
Private Function storeDatas(ByRef ws As Worksheet) As Object
    
    'ws 支店コードsheet
    
    Dim r As Long
    Dim rng As Range
    Dim var As Variant
    
    Dim stores As Object
    Set stores = CreateObject("Scripting.Dictionary")

        
        '範囲の、最終行を格納
        r = ws.Cells(Rows.Count, 5).End(xlUp).Row

        '支店名範囲を格納
        Set rng = ws.Range(ws.Cells(2, 5), ws.Cells(r, 5)) 'E開始行:E最終行
    
        
    '範囲を回して支店名を配列に格納     可視セルのみで回す指定
    For Each var In rng.SpecialCells(xlCellTypeVisible)
    
        stores.Add var.Value, var.Offset(0, 1).Value
                       
    Next var
    
    '戻り値
    Set storeDatas = stores
    
    Set rng = Nothing

End Function

'ブックが開かれているかチェック
Private Function areaBKcheck(areaFileName As String) As Boolean

    ' 開いているすべてのブックを走査
    Dim wb As Workbook
    For Each wb In Workbooks
        ' ブック名が一致したらTrueを返す
        If wb.name = areaFileName Then
            areaBKcheck = True
            Exit Function
        End If
    Next wb
    
End Function

'書き込むデータの体裁を整える
Private Sub fixUpData(ByRef dataBK As Workbook)

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


    Dim ws As Worksheet
    Dim str As String
    
    Dim cate As Variant
    Dim cateArr As Variant
                
    Dim lastCol As Long
    Dim j As Long
    
    
    '============================== シートの体裁をそろえる
    
    'データブックのシートすべてを１枚ずつ処理
    For Each ws In dataBK.Worksheets
    
        str = ws.name

        '********************** シートは、列削除処理へ
        If InStr(str, "") > 0 Then

            'セル結合解除
            ws.UsedRange.UnMerge

            '1行目～12行目まで削除
            ws.Range("1:12").Delete shift:=xlUp


            '最終列が可変なので、シート毎に最終行を格納する
            lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

                '最終列から逆ループさせる
                For j = lastCol To 1 Step -1

                    'j列セルの文字列を格納
                    str = ws.Cells(1, j).Value

                '文字列の精査
                Select Case str <> ""
                    '含まれていたらなにもしない
                    Case str Like "支店?", str Like "*高", str Like "*率"

                    '条件以外は列削除
                    Case Else
                        ws.Columns(j).Delete

                End Select

                '空白列時の処理
                If str = "" Then ws.Columns(j).Delete

                Next j '次の列へ
                        
        End If
            
            ws.Columns("A").AutoFit
            
    Next ws '次のシートへ
    
    
    '============================== 分類シートを作る
                
    Dim r As Long
    Dim c As Long
    
    Dim lastR As Long
    Dim startR As Long
    Dim i As Integer
               
    Dim bun As Worksheet
    Set bun = dataBK.Sheets.Add(dataBK.Worksheets(1))
        
        bun.name = "分類"
        cateArr = Array("大分類", "中分類")
        
        lastR = 0
       startR = 1
            i = 1
            j = 0

        '大/中分類シートから値を持ってくる
        For Each cate In cateArr
        
            Set ws = dataBK.Worksheets(cate)
            
            r = ws.Cells(Rows.Count, 1).End(xlUp).Row
            c = ws.Cells(2, Columns.Count).End(xlToLeft).Column
            
                lastR = lastR + r - j
            
            bun.Range(bun.Cells(startR, 1), bun.Cells(lastR, c)).Value = _
            ws.Range(ws.Cells(i, 1), ws.Cells(r, c)).Value
        
            startR = lastR + 1
            If i = 1 Then
                i = 2
                j = 1
            End If
        Next cate
        
        
        '列幅調整
        bun.Columns.AutoFit
        
        Set ws = Nothing

End Sub

'書き込み列を返す関数
Private Function lastOut_col(out_date As Date, ByRef aws As Worksheet) As Long

    '書き込み日付　out_date
    'シート　aws
            
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
    Dim col As Long
        
        On Error GoTo dayRangeErr
        
        '入力期間日付の判定
        For Each d In dayRng
            '値が一致したセルの列を返す
            If d.Value = out_d Then
                    col = d.Column
                    Exit For
            End If
        Next d
        
        On Error GoTo 0
        
        '日付指定に誤りがあるとき
        If col = 0 Then
        
            MsgBox "出力列の取得でエラーが発生しました。" & vbCrLf & _
                   "処理を中断します。", vbCritical
            
            'bookを保存せず閉じる
            Application.DisplayAlerts = False
            aws.Parent.Close False
            Application.DisplayAlerts = True
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        
            End  'プログラム全終了（中断）
        
        End If
                
  
    '出力列を返す
    lastOut_col = col
    Exit Function

'エラー対処
'range参照の取得漏れ ==================================================
dayRangeErr:
            
            MsgBox "出力月の判定でエラーが発生しました。" & vbCrLf & _
                   "処理を中断します。", vbCritical
            
            'bookを保存せず閉じる
            Application.DisplayAlerts = False
            aws.Parent.Close False
            Application.DisplayAlerts = True
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        
            End  'プログラム全終了（中断）
    
End Function


