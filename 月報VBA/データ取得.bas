Option Explicit
'書き込むデータをメインブックに取得する
Sub dataInput()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


    Dim thisBK As Workbook
    Set thisBK = ThisWorkbook
    Dim ws As Worksheet
    Dim str As String
    
    Dim cate As Variant
    Dim cateArr As Variant
        cateArr = Array("分類", "粗利中分類")
        
    '============================== メインブックにデータシートが残っているなら消去する
       
     For Each cate In cateArr
        For Each ws In thisBK.Worksheets
           If ws.name = cate Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
                Exit For
           End If
        Next ws
    Next cate
    
    Erase cateArr
    
    '============================== データファイルの取得 ==============================
    
        ChDir thisBK.path
    
    Dim openFile As String
        
    'データファイルを選んでもらう
    openFile = Application.GetOpenFilename("Microsoft Excelブック,*.xlsx")
    
        If openFile = "False" Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
    
    Dim dataBK As Workbook
    Set dataBK = Workbooks.Open(openFile)
        
    Dim lastCol As Long
    Dim j As Long
    
    
    '============================== シートの体裁をそろえる
    
    'データブックのシートすべてを１枚ずつ処理
    For Each ws In dataBK.Worksheets
    
        str = ws.name

        '********************** 粗利シートは、列削除処理へ
        If InStr(str, "粗利") > 0 Then

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
                    Case str Like "店", str Like "*粗利高", str Like "*粗利率"

                    '条件以外は列削除
                    Case Else
                        ws.Columns(j).Delete

                End Select

                '空白列、不要列だった時の処理
                If str = "" Or InStr(str, "用") > 0 Then ws.Columns(j).Delete

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
        
        
    '============================== メインブックにデータを持ってくる
    
    For Each ws In dataBK.Worksheets
    
        str = ws.name
    
        If str = "分類" Or InStr(str, "粗利") > 0 Then
            'データシートならば、メインブックにシートコピー
            dataBK.Worksheets(str).Copy after:=thisBK.Worksheets(thisBK.Sheets.Count)
        End If
    Next ws
    
        'データファイルを閉じる
        Application.DisplayAlerts = False
        dataBK.Close False
        Application.DisplayAlerts = True
    
    thisBK.Worksheets("メイン").Activate
    
    Set dataBK = Nothing
    Set thisBK = Nothing
    Set ws = Nothing
    Set bun = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "データの取得が完了しました。", vbInformation

End Sub

'メインブックのデータシートを削除する
Sub sheetDelete()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ans As VbMsgBoxResult
    
        ans = MsgBox("このファイルに取り込まれたデータシートを削除します。" & vbCrLf & "よろしいですか？", vbYesNo + vbExclamation, "シートの削除")
    
        If ans = vbNo Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
   
    Dim thisBK As Workbook
    Set thisBK = ThisWorkbook
    Dim ws As Worksheet
    
    Dim cate As Variant
    Dim cateArr As Variant
        cateArr = Array("分類", "粗利")
        
    '============================== メインブックのデータシートを消去する
       
     For Each cate In cateArr
        For Each ws In thisBK.Worksheets
           If ws.name = cate Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
                Exit For
           End If
        Next ws
    Next cate
    
    Erase cateArr
    Set thisBK = Nothing

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
