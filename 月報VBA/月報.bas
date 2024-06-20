Option Explicit
'�쐬
Sub Report()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim thisBK As Workbook
    Dim areaBK As Workbook
    Dim BK As Workbook
    
    '���C���u�b�N
    Set thisBK = ThisWorkbook
    
    '���C���u�b�N�A�\�V�[�g�̊i�[
    Dim AreaSh As Worksheet
    Set AreaSh = thisBK.Worksheets("�\")
    
    
    '�쐬���i�[
    Dim toDay As Date
    
        If thisBK.Worksheets("���C��").Cells(2, 6).Value <> "" Then
            toDay = thisBK.Worksheets("���C��").Cells(2, 6).Value
        Else
            MsgBox "�쐬���̎擾�Ɏ��s���܂����B" & vbCrLf & _
                   "���C���V�[�g�A�ŏI�쐬�����Ɍ������t����͌�A�Ď��s���Ă��������B", vbExclamation
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
    '***** �{���ɍ쐬���Ă悢������
    
    Dim eom As Boolean
        eom = False
        
        '��������
        If Month(toDay) = 2 Then
            '2��
            Select Case Day(toDay)
                Case 28, 29
                    eom = True
            End Select
        Else
            '���̌�
            Select Case Day(toDay)
                Case 30, 31
                    eom = True
            End Select
        End If
        
    Dim areaCheck As Boolean
        areaCheck = False
        
        '�������݊�������
        If thisBK.Worksheets("���C��").Cells(26, 16).Value = thisBK.Worksheets("���C��").Cells(26, 16).Value Then
            areaCheck = True
        End If
        
        '�����������Ă��邩
        If eom And areaCheck Then
        Else
            '�����ĂȂ�
            MsgBox "�쐬�̏����������Ă��Ȃ��悤�ł��B" & vbCrLf & _
                   "�쐬�����������A�S�̍쐬����������܂ł͎��s�ł��܂���B", vbExclamation
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
    
    '***** ���s
    
    '�t�@�C�������邩�������� ==========================
    
    Call BookCreate(thisBK, toDay)
    
    '�i�[���� ==========================================
    
    Dim R As Long
    Dim area As Variant
    Dim areaRng As Range
    Dim areaDic As Object
    Set areaDic = CreateObject("Scripting.Dictionary")

        R = AreaSh.Cells(AreaSh.Rows.Count, 1).End(xlUp).Row
    
        AreaSh.Activate
    
    Set areaRng = AreaSh.Range(AreaSh.Cells(2, 1), AreaSh.Cells(R, 1))
    
        '�v�シ��������Ɋi�[����
        For Each area In areaRng
            areaDic.Add area.Value, area.Offset(0, 1).Value
        Next area
    
    Set areaRng = Nothing
    Set AreaSh = Nothing
    
    '�u�b�N�����ɊJ���A�����擾���Ă��� ============
    
    '�t�@�C���i���o�[
    Dim num As Integer
        num = searchmonth(toDay)
    
    '�t�@�C���p�X
    Dim areaPth As String
    
    '�t�@�C���擾��
    Dim areaCol As Long
        areaCol = areaBookGetCol(toDay)
        
    '�t�@�C��sh
    Dim AreaSh As Worksheet
                
        Set AreaSh = Nothing
    
    '��N����p�z��
    Dim lastRevenue() As Currency '8���ށ@������+���v
    ReDim lastRevenue(7)
        
        
    '***** �{���� *****
    
        '�t�@�C�����J���ĕϐ��Ɋi�[
        Set BK = Workbooks.Open(thisBK.path & "\\" & Year(toDay) & ".xlsx")
        
        '���t�͂����ĂȂ�����������
        If BK.Worksheets("�x��").Cells(3, 4).Value = "" Then BK.Worksheets("�x��").Cells(3, 4).Value = Year(toDay) & "/1/1"
        
    
    '�t�@�C����1���J����
    For Each area In areaDic
    
        '****** �t�@�C���ɃV�[�g�i�[�A�Ȃ���΃R�s�[����
          On Error Resume Next
    
            Set AreaSh = BK.Worksheets(area)
        
          On Error GoTo 0
            
            '�Ȃ��Ƃ�
            If AreaSh Is Nothing Then
                BK.Worksheets("").Visible = True
                BK.Worksheets("").Copy after:=BK.Worksheets(BK.Sheets.Count)
                Set AreaSh = BK.Worksheets(BK.Sheets.Count)
                AreaSh.name = area
                AreaSh.Cells(3, 2).Value = area
                BK.Worksheets("").Visible = False
            End If
    
        
        '���N�t�@�C���p�X
        areaPth = thisBK.path & "\" & areaDic.Item(area) & "\" & Year(toDay) & "" & area & num & ".xlsx"
        '���N�t�@�C�����J���ĕϐ��i�[
        Set areaBK = Workbooks.Open(areaPth)
        
        '==================== �f�[�^�]�L ====================
        
        Call outputAreaData(areaBK, areaCol, AreaSh, toDay)
            
            '�m�i�V
            Set AreaSh = Nothing
        
        '==================== ���ԍ\����쐬 ====================
        
        Call salesRatioMonth(BK, area, areaBK, areaCol, toDay)
                
            '���N�t�@�C���A�ۑ����Ȃ��ŕ���
            areaBK.Close False
            Set areaBK = Nothing
        
        
        '***** ���N�̃u�b�N���J��
        
            '��N�t�@�C���p�X
            areaPth = thisBK.path & "\" & areaDic.Item(area) & "\" & Year(toDay) - 1 & "" & area & num & ".xlsx"
            
            On Error Resume Next
            
                '��N�t�@�C�����J���ĕϐ��i�[
                Set areaBK = Workbooks.Open(areaPth)
                
            On Error GoTo 0
            
            '���N�Ȃ��Ƃ�
            If areaBK Is Nothing Then
                GoTo noStore '����
            End If

        '==================== ��N����擾 ====================
        
        '���ޖ����v�̎擾�A���Z
        lastRevenue = areaRevenueGet(areaBK, area, areaCol, lastRevenue)
        
            '���N�t�@�C���A�ۑ����Ȃ��ŕ���
            areaBK.Close False
            Set areaBK = Nothing
noStore:
    Next area '��
        
        Set areaDic = Nothing
        
        
    '***** �]�L���� *****
        
    '================== �x�����v�v�Z ====================
    
    Call allAreaCalc(BK, toDay)
    
    '================== �x����Όv�Z ====================
    
    Call allStoreYoYCalc(BK, toDay, lastRevenue)
    
    '================= �N�ԍ\����]�L ===================
    
    Call yearRatio(BK, toDay)
    
    '***** �I���̏���
    
        BK.Worksheets("�x��").Activate
        
        '�t�@�C����ۑ����ďI��
        BK.Close True
        Set BK = Nothing
        
        thisBK.Worksheets("���C��").Activate
        Set thisBK = Nothing

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "�������������܂����B", vbInformation

End Sub

'�V�K�u�b�N�쐬
Private Sub BookCreate(ByRef thisBK As Workbook, toDay As Date)

Application.ScreenUpdating = False

    '�p�X�i�[�p
    Dim pth As String
        pth = thisBK.path
                
    Dim yer As Integer
        yer = Year(toDay)
                    
    '�t�H���_���Ƀt�@�C�������邩�T��
    If Dir(pth & "\\" & yer & ".xlsx") = "" Then
        
        '�t�H���_�̗L�����Ď�
        On Error GoTo createErr
        
        '�Ȃ���΁A���{�t�@�C�����R�s�[���A�u�b�N�Ƃ��ĐV�K�쐬 -> �R�s�[��path , �V�K�t�@�C��path
        FileCopy pth & "\" & "���{.xlsx", pth & "\\" & yer & ".xlsx"
        
        On Error GoTo 0
    End If
    
    '��������
    Exit Sub
    
'���G���[�Ώ�
'�t�@�C���A�t�H���_��������Ȃ������� ================================
createErr:

        MsgBox "�u���{.xlsx�v�܂��́u�t�H���_�v������̏ꏊ�Ɍ�����܂���ł����B" & vbCrLf & _
               "�u���{.xlsx�v��u�t�H���_�v�̏ꏊ�A���O���m�F���Ă��珈������蒼���Ă��������B", vbCritical
        
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        
        End  '�v���O�������f
        
End Sub

'�u�b�N����̎擾���Ԃ�
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
    
    '�߂�l
    areaBookGetCol = c

End Function

'�t�@�C���̏����t�@�C���ɓ]�L
Private Sub outputAreaData(ByRef areaBK As Workbook, ByRef areaCol As Long, ByRef sh As Worksheet, toDay As Date)

    'areaBK �t�@�C���AareaCol �擾��Ash �V�[�g�Amon �쐬��
       
    '�o�͗�i�[
    Dim Col As Integer
        Col = Month(toDay) + 3
        
    Dim areaR As Long
        areaR = 140      '�s����
        
    Dim R As Long
        R = 5
        
    Dim i As Integer '0-7 ������+���v�A8����
    Dim j As Integer '���ޓ��A9����
    
    
    '�t�@�C���A�V�[�g�i�[
    Dim areaws As Worksheet
    Set areaws = areaBK.Worksheets(sh.name)
    
        '�f�[�^�]�L
        For i = 0 To 7
            For j = 0 To 8
            
                sh.Cells(R, Col).Value = areaws.Cells(areaR, areaCol).Value
                
                If j = 0 Then
                    R = R + 2
                Else
                    R = R + 1
                End If
                
                areaR = areaR + 1
            
            Next j '���̍���
        Next i '���̕���
        
        '����������
        sh.Cells(94, Col).Value = Date
        
    Set areaws = Nothing

End Sub

'���ԍ\����쐬
Private Sub salesRatioMonth(ByRef BK As Workbook, ByRef area As Variant, ByRef areaBK As Workbook, ByRef areaCol As Long, ByRef toDay As Date)

    'bk �t�@�C���Aarea ���Aareabk �t�@�C���Aareacol �擾��Atoday �쐬��

    Dim MonSh As Worksheet
    Set MonSh = BK.Worksheets("���ԍ\����")
    
    '�t�@�C��sh�A�ŏI�s
    Dim R As Long
        R = MonSh.Cells(MonSh.Rows.Count, 2).End(xlUp).Row
        
    '���v��A�ŏI�s
    Dim sr As Long
        sr = MonSh.Cells(MonSh.Rows.Count, 7).End(xlUp).Row
        
    '******** �挎�����𐸍�
    
        '���t���������ł͂Ȃ�������
        If MonSh.Cells(1, 5).Value <> toDay Then
            '���t����������
            MonSh.Cells(1, 5).Value = toDay
        
            '�͈̓Z������|����     range("A3:Br") �F�x�X��
            MonSh.Range(MonSh.Cells(3, 1), MonSh.Cells(R + 1, 2)).Value = ""
            
            '���v��                 range("G6:Gsr")
            MonSh.Range(MonSh.Cells(6, 7), MonSh.Cells(sr + 1, 7)).Value = ""

        End If
    
    '******** �����A�����@�������ݏ���
                
    '�t�@�C��sh�A�������݊J�n�s�̍Ď擾
        R = MonSh.Cells(MonSh.Rows.Count, 2).End(xlUp).Row + 1
       sr = MonSh.Cells(MonSh.Rows.Count, 7).End(xlUp).Row + 1

    '���擾
    Dim areaName As String
        areaName = area
        
        '���v��ɖ������
        MonSh.Cells(sr, 7).Value = areaName

    Dim storeName As String
    Dim Ws As Worksheet
        
        
        '�t�@�C���̃V�[�g���P����
        For Each Ws In areaBK.Worksheets
        
            storeName = Ws.name
            
            '�x�X�V�[�g�Ȃ�΁A���v���擾�]�L
            If storeName <> areaName And storeName <> "�x�X" Then
                With MonSh
                    .Cells(R, 1).Value = areaName '��
                    .Cells(R, 2).Value = storeName '�x�X��
                    .Cells(R, 3).Value = Ws.Cells(203, areaCol).Value '������z
                End With
                R = R + 1
            End If
        Next Ws
        
    Set MonSh = Nothing
        
End Sub

'��N�̍��ڍ��v���擾�A�v�Z
Private Function areaRevenueGet(ByRef areaBK As Workbook, ByRef area As Variant, ByRef areaCol As Long, revenue() As Currency) As Currency()

    'areaBK �t�@�C���Aarea ���AareaCol �擾��Arevenue ��N���v���Z�z��

    Dim areaws As Worksheet
    Set areaws = areaBK.Worksheets(area)
    
    Dim i As Integer
    Dim R As Long
        R = 140     '�s
    
        For i = 0 To UBound(revenue)
            '���Z
            revenue(i) = revenue(i) + areaws.Cells(R, areaCol).Value
            R = R + 9
        Next i
    
    Set areaws = Nothing
    
    '�߂�
    areaRevenueGet = revenue()

End Function

'�x���V�[�g�̌v�Z
Private Sub allAreaCalc(ByRef BK As Workbook, ByRef toDay As Date)
    
    'BK �t�@�C���AtaDay �쐬��
        
    '�擾�A�o�͗�
    Dim Col As Long
        Col = Month(toDay) + 3

    '���v�擾�A���Z
    Dim cateCalc(7, 3) As Currency '8����4����
    
    '���ڍs
    Dim R As Long
        
    Dim i As Integer
    Dim j As Integer
    
    Dim Ws As Worksheet
    Dim shName As String
    
    ' ********************* ���z�擾�A���Z
    
        '�t�@�C���A�V�[�g�S��
        For Each Ws In BK.Worksheets
        
            shName = Ws.name
            R = 5 '����
            
            '�V�[�g�Ȃ�΁A
            If shName <> "�x��" And shName <> "" And shName <> "���ԍ\����" And shName <> "�N�ԍ\����" Then
            
                '���ނ���
                For i = 0 To UBound(cateCalc, 1)
                    '���ڂ���
                    For j = 0 To UBound(cateCalc, 2)
                        
                        cateCalc(i, j) = cateCalc(i, j) + Ws.Cells(R, Col).Value
                        
                        '�s�̉��Z
                        Select Case j
                            Case 0, 2       '����A�q���̎擾��
                                R = R + 3
                            Case 1, 3       '�_���A�e���z�̎擾��
                                R = R + 2
                        End Select
                    
                    Next j '���̍���
                Next i '���̕���
            End If
        Next Ws '���̃V�[�g
        
    ' ********************* �x���V�[�g�ɏo��
    
    Set Ws = BK.Worksheets("�x��")
    
        R = 5
        
        For i = 0 To UBound(cateCalc, 1)
            For j = 0 To UBound(cateCalc, 2)
                
                Ws.Cells(R, Col).Value = cateCalc(i, j)
                
                    Select Case j
                        Case 0, 2       '����A�q���̏o�͎�
                            R = R + 3
                        Case 1, 3       '�_���A�e���z�̏o�͎�
                            R = R + 2
                    End Select
            Next j
        Next i
        
    Set Ws = Nothing
        
End Sub

'�x���̍�Δ���v�Z����
Private Sub allStoreYoYCalc(ByRef BK As Workbook, ByRef toDay As Date, lastRevenue() As Currency)

    'bk �t�@�C���Atoday �쐬���Alastrevenue ��N�����ޔ���

    Dim Ws As Worksheet
    Set Ws = BK.Worksheets("�x��")
    
    Dim Col As Long
        Col = Month(toDay) + 3  '�o�͑ΏہFC��ȍ~
        
    Dim i As Integer
    Dim R As Long
        R = 5 '�v�Z���ڍs   �o�͍�΍s = 7
    
    Dim val As Currency

        For i = 0 To UBound(lastRevenue)
                        
            '0���Z�Ȃ�΃[���ŏ�������
            If lastRevenue(i) = 0 Then
                Ws.Cells(R + 2, Col).Value = 0
            Else
                '���N���̔���擾
                val = Ws.Cells(R, Col).Value
                Ws.Cells(R + 2, Col).Value = val / lastRevenue(i)
            End If
            
            R = R + 10
            
        Next i '���̕���
    
        '����������
        Ws.Cells(85, Col).Value = Date
    
    Set Ws = Nothing
    
End Sub

'�t�@�C���N�ԍ\���䏈��
Private Sub yearRatio(ByRef BK As Workbook, ByRef toDay As Date)

    Dim MonSh As Worksheet
    Dim YearSh As Worksheet
    
    Set MonSh = BK.Worksheets("���ԍ\����")
    Set YearSh = BK.Worksheets("�N�ԍ\����")

    Dim Col As Long
        Col = Month(toDay) + 2  '�o�͑ΏہFB��ȍ~
        
        '�\����擾�̂��ߍČv�Z
        MonSh.Calculate
        
        MonSh.Activate
    
    '****** ���ԃV�[�g����x�X���A�x���\����擾
       
    '���ꕨ
    Dim storeDic As Object
    Set storeDic = CreateObject("Scripting.Dictionary")
    
    '���ԃV�[�g�ŏI�s
    Dim R As Long
        R = MonSh.Cells(MonSh.Rows.Count, 2).End(xlUp).Row
        
    Dim store As Variant
    Dim storeRng As Range
    
        Set storeRng = MonSh.Range(MonSh.Cells(3, 2), MonSh.Cells(R, 2))
           
            For Each store In storeRng
                '�l�̎擾
                storeDic.Add store.Value, store.Offset(0, 3).Value
                
            Next store
        
        Set storeRng = Nothing

            
    '****** �N�ԃV�[�g�ɏo��
    
        YearSh.Activate
    
    '�V�x�X�`�F�b�N
    Dim newStore As Boolean
        newStore = False
        
    Dim rng As Variant
        
    '�N�ԃV�[�g�ŏI�s
        R = YearSh.Cells(YearSh.Rows.Count, 2).End(xlUp).Row
        
        '�x�X�͈�
        Set storeRng = YearSh.Range(YearSh.Cells(3, 2), YearSh.Cells(R, 2))
   
        
            '�擾�x�X�����ɉ�
            For Each store In storeDic
            
                '�͈͂Ɏx�X�����邩����
                If WorksheetFunction.CountIf(YearSh.Columns(2), store) > 0 Then
                               
                    'B��A�x�X���Ńt�B���^�[��������
                    YearSh.Columns(2).AutoFilter Field:=1, Criteria1:=store
                    
                        For Each rng In storeRng.SpecialCells(xlCellTypeVisible)
                            '�x�X�̍\����o��
                            If rng.Value = store Then
                                YearSh.Cells(rng.Row, Col).Value = storeDic(store)
                                Exit For '���̎x�X
                            End If
                        Next rng
                
                '�Ȃ���΁A�V�x�X����
                Else
                    R = R + 1
                    YearSh.Cells(R, 2).Value = store                 '�x�X��
                    YearSh.Cells(R, Col).Value = storeDic(store) '�\����
                    newStore = True '�t���O����
                End If
                
            Next store
            
        '�t�B���^�[����
        If YearSh.AutoFilterMode = True Then
            YearSh.AutoFilterMode = False
        End If
         
        Set storeRng = Nothing
        
            
    '****** �V�x�X���͂����������A�\�̕��בւ����s��
        
        If newStore Then
        
            '���בւ��̂��߂ɍČv�Z
            YearSh.Calculate
            
            '���i�[
            Dim storeArr() As String
            Dim Ws As Worksheet
            Dim shName As String
            Dim i As Integer
            
                ReDim storeArr(BK.Worksheets.Count - 5)
                i = 0
                
                For Each Ws In BK.Worksheets
                
                    shName = Ws.name
                    '��v�V�[�g����Ȃ����
                    If shName <> "�x��" And shName <> "" And shName <> "���ԍ\����" And shName <> "�N�ԍ\����" Then
                        storeArr(i) = shName
                        i = i + 1
                    End If
                Next Ws
            
                    
            '�ŏI�s�Ď擾
            R = YearSh.Cells(YearSh.Rows.Count, 2).End(xlUp).Row
            
            '���בւ��͈�
            Set storeRng = YearSh.Range(YearSh.Cells(3, 1), YearSh.Cells(R, Col))
            
                '���בւ���
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
