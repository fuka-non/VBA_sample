Option Explicit
'���C������
Sub main()

'��ʍX�V�I�t
Application.ScreenUpdating = False
'Excel�֐��A�����v�Z�I�t
Application.Calculation = xlCalculationManual


    Dim thisBK As Workbook
    
    Dim tws As Worksheet
    Dim thisS As Worksheet
    Dim data_ws As Worksheet
    Dim store_ws As Worksheet
        
    Set thisBK = ThisWorkbook
    Set tws = thisBK.Worksheets("���C��")
    Set thisS = thisBK.Worksheets("�x�X�R�[�h")
    
    
    Dim R As Long
        
    '���C���u�b�N�V�[�g�̃t�B���^����Ɖ���
        Call searchFilter(thisBK)

    
    '������ **************************************************************************
            
    '���C���V�[�g�̕K�v���ɓ��͂����邩���� ==============================
        
        If tws.Cells(2, 1).Value = "" Then
            MsgBox "�쐬�x�������͂���Ă��܂���B���͂��Ă�����s���Ă��������B", vbCritical
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
        If tws.Cells(9, 6).Value = "" Then
            MsgBox "�쐬���t�����͂���Ă��܂���B���͂��Ă�����s���Ă��������B", vbCritical
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
    '�f�[�^�V�[�g�ɕs�����Ȃ������� ==============================
    
    '���o�V�[�g������p
    Dim data_name() As Variant
        data_name = Array("����", "�e��")
    
    'Each�p
    Dim area As Variant
    Dim store As Variant
    Dim Ws As Worksheet
    Dim dt As Variant
        
    Dim shCheck(3) As Boolean
    Dim s As Integer
        s = 0
    
        For Each dt In data_name
            For Each Ws In thisBK.Worksheets
                '�V�[�g�����݂�����t���O�𗧂Ă�
                If Ws.name = dt Then
                    shCheck(s) = True
                    s = s + 1
                    Exit For
                End If
            Next Ws
        Next dt
                
        '�V�[�g����
        For Each dt In shCheck
            If dt = False Then
                MsgBox "�����p�̃f�[�^�V�[�g���K�v��������Ă��܂���B" & vbCrLf & "�m�F���Ă�����s���Ă��������B", vbCritical
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                Exit Sub
            End If
        Next dt
        

    '�o�͎x���擾���� ==========================================
    
    tws.Activate
    
        '���C���V�[�g�A���͎x���ŏI�s���i�[
        R = tws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    '���͎x���͈͂��i�[
    Dim areaRng As Range
    Set areaRng = tws.Range(tws.Cells(2, 1), tws.Cells(R, 1)) 'A2�ȍ~
    
     'Dictionary�I�u�W�F�N�g�̐錾
    Dim areaDic As Object
    Set areaDic = CreateObject("Scripting.Dictionary")
    
    
    '�������t
    Dim toDay As Date
    
            '�쐬���t�i�[
            toDay = tws.Cells(9, 6).Value
    
    
    '�����t�@�C�����l�̊i�[�@�����Fdate
    Dim num As Integer
        num = search_month(toDay)

    
        'Dic�ɏ������ގx��������
        For Each area In areaRng
            If area.Value <> "" Then
            '�������A�v�f�̒ǉ�     .Add �L�[,�l
             areaDic.Add area.Value, area.Offset(0, 1).Value    '�x����,�R�[�h+�x����
            End If
        Next area
    
    
    '�����������ݎx���̔��� =======================================
    
    '�x���͈͂̐��l�����p
    Dim rc() As Integer
        rc() = areaCheck(tws, toDay)
    
    '�����ςݎx�����������ޔ͈�
    Set areaRng = tws.Range(tws.Cells(3, 6), tws.Cells(6, 8))
        
        
    '�{���� **************************************************************************
        
    '=========== �擾�������͎x�������ɁA�������J�n���� ===========
     
    '�V�K����p
    Dim newcheck As Boolean  'book�p
    Dim newStore As Boolean  '�V�x�X�p
    
        newcheck = False
        newStore = False
    
    '�x���u�b�N�p
    Dim areaBK As Workbook
    Dim areaInfo(2) As String
    
    '�����o����
    Dim c As Long
    
    '�x�X�p�z��
    Dim stores() As String
    
            
    '�G���[�`�F�b�N
    Dim errcheck As Boolean
    
    Dim filePth As String
    
        R = 0
        

    '�o�͂���x�����P����for�@area = dic�̃L�[
    For Each area In areaDic
    
        '�����t�@�C������      �����AItem(area):�R�[�h+�x�������� , area:�x���� , �t�@�C���i���o�[
        newcheck = BookCreate(areaDic.Item(area), area, num, toDay)
        
        filePth = thisBK.path & "\" & areaDic.Item(area) & "\" & Year(toDay) & "_" & area & num & ".xlsx"
                
        '�x���t�@�C�����J���ĕϐ��i�[
        Set areaBK = Workbooks.Open(filePth)

'�G���[�߂�
newCreateAreaSheet:

                '�V�Kbook�������Ƃ��A�x���V�[�g���ƃZ�����A�Y���x�����ɕύX����
                If newcheck Then
                    areaBK.Worksheets("�x��").name = area
                    areaBK.Worksheets(area).Cells(3, 2).Value = area
                    '�R�s�[�p�ЂȌ^�V�[�g��\��
                    areaBK.Worksheets("�x�X").Visible = True
                End If
        
        '���C���u�b�N�̃��C���V�[�g��active�����Ă���
        tws.Activate

        '�V�K�x��book���������̂܂ܕ���ꂽ���̊Ď�
        On Error GoTo nothingArea
        
        '�������ݗ�𔻒肵�āA����i�[�@�@�����F�������ݓ��t ,�x���u�b�N
        c = out_col(toDay, areaBK.Worksheets(area))
        
        On Error GoTo 0
        
        
        '�x�X�R�[�h�V�[�g���A�N�e�B�u������
        thisS.Activate
        
        '�����x���Ńt�B���^   �x���� D��
        thisS.Cells(1, 1).AutoFilter Field:=4, Criteria1:=area
        
        '�x���x�X���i�[
        stores() = storeNames(thisS)
        
        
        '========================== �x�X���̃f�[�^����������ł�������
        
        
        '�x�X�z����P����
        For Each store In stores
        
'�G���[�߂�
newCreateStoreSheet:

            '�V�Kbook/�V�K�V�x�X�A���Ȃ�΁A�x�X�V�[�g�̍쐬,���̑�����
            If newcheck Or newStore Then
                'areaBk.Activate
                areaBK.Worksheets("�x�X").Copy after:=areaBK.Worksheets(areaBK.Sheets.Count)
                Set store_ws = areaBK.Worksheets(areaBK.Sheets.Count)
                store_ws.name = store
                store_ws.Cells(3, 2).Value = store
                thisBK.Activate
            Else
               '�G���[�L���b�`  try-catch
               On Error GoTo storeWsError
               
                '�x���u�b�N�A�]�L�p�V�[�g�̊i�[
                Set store_ws = areaBK.Worksheets(store)
                    '�����F�����r���ŐV�x�X������������̓G���[�ɂȂ邽�߁A�Ώ���
               
               '�G���[�Ď����s
               On Error GoTo 0
                
            End If
            
            
            '******* �f�[�^�p�z����񂵂āA�x�X�̊e���l��]�L
            For Each dt In data_name
                
                '���C���u�b�N�A���o�p�V�[�g�̊i�[
                Set data_ws = thisBK.Worksheets(dt)
                
                '�e���V�[�g���𔻒�
                If InStr(data_ws.name, "�e��") > 0 Then
                    '�e���V�[�g�F�x�X�Ńt�B���^   �x�X�� A��
                    data_ws.Cells(1, 1).AutoFilter Field:=1, Criteria1:=store
                Else
                    '����/�x�X��΃V�[�g�F�x�X�Ńt�B���^   �x�X�� C��
                    data_ws.Cells(1, 1).AutoFilter Field:=3, Criteria1:=store
                End If
            
                '�]�L����  �����F�e�V�[�g�Q�ƁA�������ݗ�
                errcheck = outData(store_ws, data_ws, c)
                
                '��񂪂Ȃ��Ƃ��A�����𒆒f����
                If errcheck Then
                    '�x��book�͕ۑ���������
                    Application.DisplayAlerts = False
                    areaBK.Close False
                    Application.DisplayAlerts = True
                    Application.ScreenUpdating = True
                    Application.Calculation = xlCalculationAutomatic
                    Exit Sub
                End If
                
                Set data_ws = Nothing
                
                
            Next dt '���̕��ރf�[�^�V�[�g��
            
            '�������̋L�q
            store_ws.Cells(230, c).Value = Date
            
            Set store_ws = Nothing
                            
        Next store  'stores��forEach�@���̎x�X�V�[�g��
        
        
        '=================== �x���V�[�g���v�̌v�Z���� ===================
        
        
        '�x��ws�A�e���ލ��v�̌v�Z
        Call areaCalc(areaBK, area, c)
                
        
        '����΂̓���
        '��Nbook�̋L�q�Ȃ�΁A��΂͕s�v�B���N�x�̏����ł���΁A��Ώ������Ăяo��
        If Year(toDay) = Year(Date) Then
            Call lastYearCalc(areaBK, area, c)
        ElseIf toDay = Year(Date) - 1 & "12/31" And Month(Date) = 1 Then
            '1�����A12�����̏���
            Call lastYearCalc(areaBK, area, c)
        End If
        
                                
        '�V�Kbook or �V�x�X�������́A�ЂȌ^�V�[�g���\����
        If newcheck Or newStore Then areaBK.Worksheets("�x�X").Visible = False
        
        '�V�[�g�\���͎x���V�[�g��擪��
        areaBK.Worksheets(area).Activate
        
            'areaBk�̏����擾���Ă���
            areaInfo(0) = area
            areaInfo(1) = areaBK.name
            areaInfo(2) = areaBK.path

        '�x���u�b�N�����
        areaBK.Close savechanges:=True
        
        Erase stores
        
        
        '���M�pbook�̍쐬 =========================
        
        If Year(toDay) = Year(Date) Then
            Call areaBookCopy(areaInfo, toDay)
        ElseIf toDay = Year(Date) - 1 & "12/31" And Month(Date) = 1 Then
            '1�����A12�����̏���
            Call areaBookCopy(areaInfo, toDay)
        End If
       
        Set areaBK = Nothing
        Erase areaInfo
        
        '�����x���̏������� =====================
        areaRng(rc(0), rc(1)) = area
        tws.Cells(2 + R, 1).Value = ""
        
            rc(1) = rc(1) + 1
                R = R + 1
                
            If rc(1) = 4 Then
                rc(1) = 1
                rc(0) = rc(0) + 1
            End If
    
    Next area  'areaDic ���̎x����
    
        Erase rc
        Erase data_name
        Set areaRng = Nothing
    
    '�I���̏����@*********************************************************
    
    '���C���u�b�N�V�[�g�̃t�B���^����Ɖ���
    Call searchFilter(thisBK)
    
    tws.Activate
    
    '�쐬���̓���
    tws.Cells(2, 6).Value = toDay
    
    '�x���\��́A�S�x���������I��������w����t�N���A
    If tws.Cells(26, 17).Value = tws.Cells(27, 17).Value Then tws.Cells(9, 6).Value = ""
                          
                          
    Set areaDic = Nothing
    Set thisBK = Nothing
    Set tws = Nothing
    Set thisS = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    
    MsgBox "�������������܂����B", vbInformation
    
    '�{�������I��
    Exit Sub

'���G���[�Ώ�
'�V�x�X���A�G���[���� ========================================
storeWsError:
                '�V�K�x�X�������r���ɂł����Ƃ�
                newStore = True
                areaBK.Worksheets("�x�X").Visible = True
                err.Clear
                
                     Application.ScreenUpdating = False
                
                '���C�������ɖ߂�
                Resume newCreateStoreSheet
                

'book�Ɏx���V�[�g���Ȃ������Ƃ��A�G���[���� ========================================
nothingArea:
                newcheck = True
                err.Clear
                
                    Application.ScreenUpdating = False
                
                Resume newCreateAreaSheet
End Sub

'�������ݎx���̃A�h���X��Ԃ�
Private Function areaCheck(ByRef tws As Worksheet, toDay As Date) As Integer()

    
    '�z��v�f�p
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
    
     '�������ݎx�������
     If tws.Cells(2, 6).Value <> toDay Then
        '���t���������
        
        '�����x���͑S�ċ󔒂�
            outRng.Value = ""
        
     Else
        '���t���ꏏ��������
        '���ɏ����ς݂̎x�����c�����߁A���͉\range�͈͂��m���߂�
        For Each v In outRng
                                     'i �c�@j ���@[i][j] 4,3�܂�
            If v.Value <> "" Then
                '�L���ς݂ł����
                j = j + 1
                If j = 4 Then
                    j = 1
                    i = i + 1
                End If
            Else
                '�󔒂���������
                rc(0) = i
                rc(1) = j
                Exit For '���[�v������
            End If
        Next v
     End If
    
    '�߂�l
    areaCheck = rc()
    
    Set outRng = Nothing

End Function

'�V�K�Ō��u�b�N�쐬
Function BookCreate(area As Variant, name As Variant, num As Integer, toDay As Date) As Boolean

Application.ScreenUpdating = False

    '�p�X�i�[�p
    Dim pth As String
        pth = ThisWorkbook.path
        
    Dim check As Boolean
        check = False
        
    Dim yer As Integer
        
        yer = Year(toDay)
        
            
    '�x���t�H���_���Ɏx���t�@�C�������邩�T��
    If Dir(pth & "\" & area & "\" & yer & "_" & name & num & ".xlsx") = "" Then
        
        '�t�H���_�̗L�����Ď�
        On Error GoTo createErr
        
        '�Ȃ���΁A���{�t�@�C�����R�s�[���A�x���u�b�N�Ƃ��ĐV�K�쐬 -> �R�s�[��path , �V�K�t�@�C��path
        FileCopy pth & "\" & "���{.xlsx", pth & "\" & area & "\" & yer & "_" & name & num & ".xlsx"
        check = True
        
        On Error GoTo 0
    End If
    
    '�����Ԃ�
    BookCreate = check
    Exit Function
    
'���G���[�Ώ�
'�x���t�H���_��������Ȃ������� ================================
createErr:

        MsgBox "�u���{.xlsx�v�܂��́u" & area & "�t�H���_�v������̏ꏊ�Ɍ�����܂���ł����B" & vbCrLf & _
               "�u���{.xlsx�v��u" & area & "�t�H���_�v�̏ꏊ�A���O���m�F���Ă��珈������蒼���Ă��������B", vbCritical
        
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        
        End  '�v���O�������f
        
End Function

'�������ݗ��Ԃ��֐�
Private Function out_col(out_date As Date, ByRef aws As Worksheet) As Long

    '�������ݓ��t�@out_date
    '�x���V�[�g�@aws
            
    '�������ݗpbook�A�V�[�g���A�N�e�B�u��
    aws.Activate
    
    '���Z�����󔒂̂Ƃ��A���t����͂��鏈��
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
  
        '���͌����肵�A���t�͈͂��i�[
        Select Case Month(out_date)
            Case Month(aws.Cells(3, 4).Value)
                Set dayRng = aws.Range(Cells(4, 4), Cells(4, 8))
            Case Month(aws.Cells(3, 9).Value)
                Set dayRng = aws.Range(Cells(4, 9), Cells(4, 13))
            Case Month(aws.Cells(3, 14).Value)
                Set dayRng = aws.Range(Cells(4, 14), Cells(4, 18))
        End Select
        
        
    '�������ݓ��t�̔���
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
        
        '���͊��ԓ��t�̔���
        For Each d In dayRng
            '�l����v�����Z���̗��Ԃ�
            If d.Value = out_d Then
                    Col = d.Column
                    Exit For
            End If
        Next d
        
        On Error GoTo 0
        
        '���t�w��Ɍ�肪����Ƃ�
        If Col = 0 Then
        
            MsgBox "�o�͗�̎擾�ŃG���[���������܂����B" & vbCrLf & _
                   "���蓮�p���ɍ쐬�������T�̗������t�œ��͂��A�Ď��s���Ă��������B" & vbCrLf & _
                   "�����𒆒f���܂��B", vbCritical
            
            '�x��book��ۑ���������
            Application.DisplayAlerts = False
            aws.Parent.Close False
            Application.DisplayAlerts = True
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        
            End  '�v���O�����S�I���i���f�j
        
        End If
                
  
    '�o�͗��Ԃ�
    out_col = Col
    Exit Function

'�G���[�Ώ�
'range�Q�Ƃ̎擾�R�� ==================================================
dayRangeErr:
            
            MsgBox "�o�͌��̔���ŃG���[���������܂����B" & vbCrLf & _
                   "���蓮�p���ɍ쐬�������T�̗������t�œ��͂��A�Ď��s���Ă��������B" & vbCrLf & _
                   "�����𒆒f���܂��B", vbCritical
            
            '�x��book��ۑ���������
            Application.DisplayAlerts = False
            aws.Parent.Close False
            Application.DisplayAlerts = True
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        
            End  '�v���O�����S�I���i���f�j
    
End Function

'�t�@�C�����l����
Function search_month(mydate As Date) As Integer
        
    Dim m As Integer
    Dim this_m As Integer

        m = Month(mydate)
    
    '�Y���t�@�C�����l����
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

'�x���͈͂̎x�X���z���Ԃ��֐�
Private Function storeNames(ByRef Ws As Worksheet) As String()
    
    'ws �x�X�R�[�hsheet
    
    Dim R As Long
    Dim rng As Range
    Dim var As Variant
    
    Dim stores() As String
    Dim i As Integer

        
        '�͈͂́A�ŏI�s���i�[
        R = Ws.Cells(Rows.Count, 5).End(xlUp).Row

        '�x�X���͈͂��i�[
        Set rng = Ws.Range(Ws.Cells(2, 5), Ws.Cells(R, 5)) 'E�J�n�s:E�ŏI�s
    
    '�v�f��������
    ReDim stores(rng.SpecialCells(xlCellTypeVisible).Count - 1)
    
    i = 0
    
    '�͈͂��񂵂Ďx�X����z��Ɋi�[     ���Z���݂̂ŉ񂷎w��
    For Each var In rng.SpecialCells(xlCellTypeVisible)
    
        stores(i) = var.Value
                
        i = i + 1
        
    Next var
    
    storeNames = stores()
    
    Set rng = Nothing

End Function

'�V�[�g�f�[�^�]�L����
Function outData(ByRef store_ws As Worksheet, ByRef data_ws As Worksheet, ByRef outCol As Long) As Boolean

    'store_ws �������ݗp�x�Xws , data_ws ���ows , outCol �����o����
        
    Dim c As Long '�ŏI��
    Dim R As Long '�ŏI�s
    
    Dim i As Integer
    Dim y As Long
    
    Dim outRng As Range
    Dim rng As Variant
    
    Dim str As String
        
            
    data_ws.Activate
    
    '�ŏI�s�̊i�[
    R = data_ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    i = 1
    
    
    If InStr(data_ws.name, "�e��") > 0 Then
    
    '===================== �e���V�[�g���̏��� =====================
    
    'range�Œ��o�V�[�g���͈͊i�[
    
        '�ŏI��̊i�[
        c = data_ws.Cells(1, Columns.Count).End(xlToLeft).Column
        '�e���V�[�g�A���l�݂̂̊i�[
        Set outRng = data_ws.Range(data_ws.Cells(2, 2), data_ws.Cells(R, c)) 'range("B2:cr")
        
        Dim val As String
        Dim j As Integer
            j = 0
        
        For Each rng In outRng.SpecialCells(xlCellTypeVisible)
                                      
                        
            '�x�X���Ȃ�������
            If VarType(rng.Value) = vbString Then
                MsgBox "�u" & store_ws.name & "�v��[" & data_ws.name & "]�f�[�^��������܂���ł����B" & vbCrLf & _
                       "�m�F��A���̎x���̏��������蒼���Ă��������B", vbCritical
                outData = True
                Exit Function
            End If
            
            '****** �f�[�^��������
            
            If i = 1 Then
                '�����o���s�̔���
                val = data_ws.Cells(1, 2 + j).Value
                str = Left(val, InStr(val, "_") - 1)
                
                Select Case str
                            Case "�H�i"
                                y = 66
                            Case Else
                                '�]���ȗ�A�s������������ continue
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
            j = j + 1 '���̍��ڐ���

        Next rng
        
    Else
    
    '===================== �݌v�V�[�g���̏��� =====================
    
    'range�Œ��o�V�[�g���͈͊i�[

        '�ŏI��̊i�[
        c = data_ws.Cells(2, Columns.Count).End(xlToLeft).Column
        '����/��΃V�[�g�Ȃ�΁A���ޔ���̂��߂�A����܂߂�
        Set outRng = data_ws.Range(data_ws.Cells(2, 1), data_ws.Cells(R, c)) 'range("A2:cr")
                

        For Each rng In outRng.SpecialCells(xlCellTypeVisible)
        
            'i��1 - 10 �����[�e����
            Select Case i
            
                Case 1
                
                    str = rng.Value
                    
                    Select Case str
                        Case "�H�i"
                            y = 59
                        Case Else
                        '�t�B���^�[�������炸�A�f�[�^���Ȃ��Ƃ�
                            MsgBox "�u" & store_ws.name & "�v��[" & data_ws.name & "]�f�[�^��������܂���ł����B" & vbCrLf & _
                            "�m�F��A���̎x���̏��������蒼���Ă��������B", vbCritical
                            outData = True
                            Exit Function
                    End Select  'case 1����case�I���
                
                Case 2, 3
                    '�������Ȃ�
                
                Case 4, 6, 7, 8, 9, 10
                    '�f�[�^�o��
                        store_ws.Cells(y, outCol).Value = rng.Value
                        y = y + 1
                Case 5
                    '��΂̏o��
                        store_ws.Cells(y, outCol).Value = rng.Value / 100
                        y = y + 1
            End Select
            
            'i�̏�����
            If i = 10 Then
                i = 1
            Else
                i = i + 1
            End If
            
        Next rng
        
    End If
        
    '���Ȃ��I��
    outData = False
    
End Function

'�x���V�[�g���ތv�Z
Sub areaCalc(ByRef areaBK As Workbook, ByRef area As Variant, ByRef outCol As Long)

    'areaBk �x���u�b�N�Q�ƁAarea �����x�����Q�ƁAoutCol �������ݗ�Q��
    
'Excel�֐��Čv�Z
Application.Calculate
     
    Dim Ws As Worksheet
    Dim str As String
    
    Dim i As Integer
    Dim j As Integer
    Dim y As Long '�s���Z�p
    
    Dim val As String
    
    
    '���l������񎟌��z��
    Dim cate_scores(22, 3) As Currency
    
    
        '�x���u�b�N�̑S�V�[�g���P����
        For Each Ws In areaBK.Worksheets
                    
            str = Ws.name  '�V�[�g�l�[���i�[
              y = 5        '�擾�J�n�s
            
            '�x�����A�x�X�ЂȌ^�V�[�g�ȊO�Ȃ�
            If str <> area And str <> "�x�X" Then
            
                
            '����0-21��]  2,2,3,2
                
                '�e���ނ̒l�����Z����       (�񎟌��z��Ȃ̂�,����1)
                For i = 0 To UBound(cate_scores, 1)
                                                 
                                
                    '���Z�ΏۃZ�� 4�@����A�_���A�q���A�e���z
                    For j = 0 To 3
                    
                        val = Ws.Cells(y, outCol).Value
                        
                        '�e���ނ̍��v�����Z����
                        If val <> "" Then
                            cate_scores(i, j) = cate_scores(i, j) + Ws.Cells(y, outCol).Value
                        End If
                        
                        If j <> 2 Then
                           y = y + 2
                        Else
                           y = y + 3
                        End If
                        
                    Next j '���̃Z����
                                   
                Next i '���̕��ނ�
            End If  '�V�[�g����
        Next Ws  '���̃V�[�g��
        
        '================ �x���V�[�g�ɏ����o��
       
    '�x���V�[�g�i�[
    Set Ws = areaBK.Worksheets(area)
        
        y = 5 '�����o���s
        
        For i = 0 To UBound(cate_scores, 1)
        
                
                '����A�e���z�����ɏ����o��
                For j = 0 To 3
                        
                    Ws.Cells(y, outCol).Value = cate_scores(i, j)
                        
                    If j <> 2 Then
                        y = y + 2
                    Else
                        y = y + 3
                    End If
                        
                Next j '���̃Z����

        Next i '���̕��ނ�
        
        '�������̋L�q
        Ws.Cells(221, outCol).Value = Date
        
Set Ws = Nothing
                
End Sub

'���M�p�̎x���u�b�N���쐬����
Private Sub areaBookCopy(areaInfo() As String, ByRef out_date As Date)

    'areaInfo(0) �x�����AareaInfo(1) �x��book���AareaInfo �x��book�p�X�AoutDay �쐬��
    
    '���M�p���t�𔻒肷�� ==================================
                
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
            '���t�𕶎���Ƃ��Ċi�[
            dateStr = Year(out_date) & "." & Month(out_date) & "." & out_d
        
    
    '�t�@�C���R�s�[���� ====================================
    
    Dim sendFileName As String
        '���M�p�t�@�C�������i�[     '�x����
        sendFileName = "�y" & areaInfo(0) & "�z�񍐏�" & dateStr & ".xlsx"
    
        '�쐬�����x���t�@�C�����R�s�[���A���M�p�t�@�C���Ƃ��ĐV�K�쐬 -> �����쐬�����x���t�@�C��path�� , ���M�t�@�C��path���i��΃p�X�j
        FileCopy areaInfo(2) & "\" & areaInfo(1), areaInfo(2) & "\" & sendFileName
        
            'areaInfo(2) : path�@�A�@areaInfo(1) : �x���u�b�N��

End Sub

'�t�B���^�[����Ɖ���
Private Sub searchFilter(ByRef thisBK As Workbook)

    Dim Ws As Worksheet

    For Each Ws In thisBK.Worksheets
        If Ws.AutoFilterMode = True Then
            Ws.AutoFilterMode = False  '�t�B���^����Ă������
        End If
    Next Ws '���V�[�g

End Sub

'���book���J���āA���N�x�̍�΂��v�Z����
Private Sub lastYearCalc(ByRef areaBK As Workbook, ByRef area As Variant, ByRef outCol As Long)

    'areaBK ���N�x��book�Q�ƁAarea �x�����AoutCol �����o����

    Dim Ws As Worksheet
    Dim yer As Integer
    Dim nameStr As String
    Dim fileName As String
    
        fileName = areaBK.name
        
        '�N�x�̐؂�o��     yyyy
        yer = CInt(Left(fileName, InStr(fileName, "_") - 1)) - 1
        
        '���O�̐؂�o��     name.xlsx
        nameStr = Right(fileName, Len(fileName) - InStr(fileName, "_"))
        
    
    Dim filePth As String
        
        filePth = areaBK.path & "\" & yer & "_" & nameStr
        
    Dim lastBK As Workbook
        
        '�O��book�̗L���Ď�
        On Error GoTo nothingLastBook
        
        '��Ηp�t�@�C�����J���ĕϐ��i�[
        Set lastBK = Workbooks.Open(filePth)
        
        On Error GoTo 0
    
    Dim str As String
    Dim i As Integer
    Dim y As Long '�s���Z
    
    'Dictionary �e�x�X�̔����������
    Dim storeDic As Object
    Set storeDic = CreateObject("Scripting.Dictionary")
    
    '��������z�� �x��ws = 25���� , �x�Xws = 1����
    Dim revenues() As Currency
    
    
    '============== �O����book����A�e�V�[�g�̔���z���擾���� ==============
            
    '�V�[�g���Ƃɔ���i�[����
    For Each Ws In lastBK.Worksheets
    
        
        str = Ws.name
        
        '�ЂȌ^,�x���V�[�g�ȊO�Ȃ�΁A�x�X���v�̂�revenue�Ɏ擾
        If str <> "�x�X" And str <> area Then
             
             '1����
             ReDim revenues(0)
   
             y = 212 '�x�X���v�s
                
                    revenues(0) = Ws.Cells(y, outCol).Value
                
        '�x���V�[�g�Ȃ�΁A�S����revenue�Ɏ擾
        ElseIf str = area Then
               
                '24����
                ReDim revenues(23)
               
                y = 5 '��ʐH�i�v�s����
                
               '���ލ��A24��]
                For i = 0 To UBound(revenues)
                    revenues(i) = Ws.Cells(y, outCol).Value
                    y = y + 9
                Next i  '���̍��ڂ�
        End If
        
        '�����Ɋi�[
        storeDic.Add str, revenues

    Next Ws  'lastBk.���̃V�[�g��
    
    
    '��Ηpbook��ۑ���������
    Application.DisplayAlerts = False
    lastBK.Close False
    Application.DisplayAlerts = True

    '���book�͗m�i�V
    Set lastBK = Nothing
    
    
    '======================== ����book�Ɋe�V�[�g�̍�Δ���v�Z���� ========================
    
    Dim sd As Variant
    Dim sale As Currency
    
    '�x�X�������P����
    For Each sd In storeDic
        
        '�O���ŕX�����x�X���������Ƃ��p�� try-catch
        On Error GoTo closedStore
        
            '�x�X�V�[�gset
            Set Ws = areaBK.Worksheets(sd)
        
        On Error GoTo 0
        
        str = Ws.name
        
        '���I�z��̏�����
        Erase revenues
        
        'item��z��ɓ��꒼��
        revenues = storeDic(sd)
        
        '�x���V�[�g�̂Ƃ�
        If str = area Then
        
            '��ʐH�i����s����
            y = 5
        
            '�e���ڂƌv�Z
            For i = 0 To UBound(revenues)
            
                If y = 212 Then Ws.Cells(y, outCol).Calculate
            
                '0�ŏ��Z�͂ł��Ȃ��̂ŁA�z��v�f�𔻒肵�Ă���
                If revenues(i) = 0 Then
                    '0���Z�ł���΁A0%�œ��͂���
                    Ws.Cells(y + 1, outCol).Value = 0
                Else
                    sale = Ws.Cells(y, outCol).Value '��������̎擾
                    Ws.Cells(y + 1, outCol).Value = sale / revenues(i) '��������/�O������
                End If
                
                y = y + 9

            Next i '���̍���
            
        '�x�X�V�[�g�̎�
        ElseIf str <> area And str <> "�x�X" Then
        
            '����s
            y = 212

            '�x�X���v�̂݌v�Z
                '0�ŏ��Z�͂ł��Ȃ��̂ŁA�z��v�f�𔻒肵�Ă���
                If revenues(0) = 0 Then
                    '0���Z�ł���΁A0%�œ��͂���
                    Ws.Cells(y + 1, outCol).Value = 0
                Else
                    sale = Ws.Cells(y, outCol).Value '��������̎擾
                    Ws.Cells(y + 1, outCol).Value = sale / revenues(0) '��������/�O������
                End If
        End If
closeNext:
    Next sd '���̎x�X��

Set Ws = Nothing
Set storeDic = Nothing
Erase revenues

Exit Sub '�G���[�Ȃ��A�������𐳏�I��

'���G���[�Ώ�
'�O������book��������Ȃ������Ƃ� ===================================================
nothingLastBook:

        '���N����̐V�x��book���Ȃ�
        err.Clear
        Application.ScreenUpdating = False
        
        Exit Sub  '��Ώ����𔲂���

'�O���ɂ͎x�X���������̂ɁA����Ȃ������Ƃ� ===========================================
closedStore:

        err.Clear
        Application.ScreenUpdating = False
        GoTo closeNext

End Sub

