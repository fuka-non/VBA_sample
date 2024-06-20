Option Explicit
'�ߋ��̃f�[�^���܂Ƃ߂č�鏈��
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
        
    Set tws = thisBK.Worksheets("�܂Ƃߗp")
    Set thisS = thisBK.Worksheets("�x�X�R�[�h")


    '�f�[�^�t�H���_�[��I�����Ă��炤 ============================================
    Dim folderPath As String
    
      With Application.FileDialog(msoFileDialogFolderPicker)
       '
        .InitialFileName = thisBK.path
        .Title = "�f�[�^�t�@�C�����������t�H���_�[��I��"
       
       '�L�����Z�����͏����𔲂���
       If .Show = 0 Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
       End If
       '�t�H���_�p�X�i�[
       folderPath = .SelectedItems(1)
      End With
         
            
    '�܂Ƃߗp�V�[�g�̕K�v���ɓ��͂����邩���� ==============================
        
        If tws.Cells(2, 1).Value = "" Then
            MsgBox "�쐬�����͂���Ă��܂���B���͂��Ă�����s���Ă��������B", vbCritical
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
        
    '���͂��ꂽ�]�L�̊i�[ ========================================================
    
    Dim r As Long
        '�܂Ƃߗp�V�[�g�A���͍ŏI�s���i�[
        r = tws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    '���͔͈͂��i�[
    Dim areaRng As Range
    Set areaRng = tws.Range(tws.Cells(2, 1), tws.Cells(r, 1)) 'A2�ȍ~
    
    Dim area As Variant
    
     'Dictionary�I�u�W�F�N�g�̐錾
    Dim areaDic As Object
    Set areaDic = CreateObject("Scripting.Dictionary")
            
        'Dic�ɏ������ނ�����
        For Each area In areaRng
            If area.Value <> "" Then
            '�������A�v�f�̒ǉ�     .Add �L�[,�l
             areaDic.Add area.Value, area.Offset(0, 1).Value    '��,�R�[�h+��
            End If
        Next area
    
    '�����ς݂��������ޔ͈�
    Set areaRng = tws.Range(tws.Cells(3, 6), tws.Cells(6, 8))
    Dim rng As Variant
    
                
    '�t�H���_�[���t�@�C�����i�[ ============================
    
    Dim dateSplit() As String
    Dim tmp As String
    Dim toDay As Date
    
    Dim fso, file, files
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set files = fso.GetFolder(folderPath).files
    
    Dim num As Integer

            
    '=========== �擾�������͂ɏ����J�n ===========
     
    '�V�K����p
    Dim newcheck As Boolean  'book�p
    Dim newStore As Boolean  '
    
        newcheck = False
        
    '�����o����
    Dim c As Long
    
    '�x�X�p����
    Dim stores As Object
    Set stores = CreateObject("Scripting.Dictionary")
    
    Dim store As Variant
    Dim storeDate As Date
    
    '���o�V�[�g������p
    Dim data_name() As Variant
        data_name = Array("����", "")
    Dim dt As Variant
            
    '�G���[�`�F�b�N
    Dim errcheck As Boolean
    
    'area�t�@�C���p�X
    Dim areaPth As String
    
    Dim areaFileName As String
        
        r = 0
    
            
    '�o�͂�����P����
    For Each area In areaDic
    
           
        '�����Ńt�B���^   �� D��
        thisS.Cells(1, 1).AutoFilter Field:=4, Criteria1:=area
        thisS.Activate
    
        '�x�X���i�[
        Set stores = storeDatas(thisS)

                
            '====================== �t�H���_�����P���� ======================
            
            For Each file In files

                '********* �f�[�^�t�@�C�����J��
                                
                '�f�[�^�t�@�C�����J���ĕϐ��i�[
                Set dataBK = Workbooks.Open(file)
                
                '�f�[�^�t�@�C�����V�[�g�̑̍ق𐮂���
                Call fixUpData(dataBK)
                
                '�t�@�C��������]�L���t���擾����
                dateSplit = Split(dataBK.name, ".")
                tmp = dateSplit(0) & "/" & dateSplit(1) & "/" & dateSplit(2)
                toDay = CDate(tmp)
                               
                               
                '********* �t�@�C�����J��
                
                     '�쐬��������
                    num = search_month(toDay)
                
                    '�����̃t�@�C������
                    newcheck = bookCreate(areaDic.Item(area), area, num, toDay)
                    
                    '�t�@�C�����̊i�[
                    areaFileName = Year(toDay) & "_" & area & num & ".xlsx"
                    
                        '�t�@�C�����J����Ă��邩����
                        If areaBKcheck(areaFileName) = False Then
                            
                            '�J���Ă���t�@�C�������ɂ���΁A�ۑ����ĂƂ���
                            If Not (areaBK Is Nothing) Then areaBK.Close True
                            
                            '�J����Ă��Ȃ���΃t�@�C�����J��
                            areaPth = thisBK.path & "\" & areaDic.Item(area) & "\" & Year(toDay) & "_" & area & num & ".xlsx"
                            '�t�@�C�����J���ĕϐ��i�[
                            Set areaBK = Workbooks.Open(areaPth)
                            
                        End If
               
'�G���[�߂�
newCreateAreaSheet:

                '�V�Kbook�������Ƃ��A�V�[�g���ƃZ�����A�Y�����ɕύX����
                If newcheck Then
                    areaBK.Worksheets("").name = area
                    areaBK.Worksheets(area).Cells(3, 2).Value = area
                    '�R�s�[�p�ЂȌ^�V�[�g��\��
                    areaBK.Worksheets("�x�X").Visible = True
                End If
        

        '�V�Kbook���������̂܂ܕ���ꂽ���̊Ď�
        On Error GoTo nothingArea2
        
            '�������ݗ�𔻒肵�āA����i�[�@�@�����F�������ݓ��t ,�u�b�N
            c = lastOut_col(toDay, areaBK.Worksheets(area))
        
        On Error GoTo 0
        
        
        '=============== �x�X���̃f�[�^����������ł������� ===============
               
        '�i�[�x�X���P����
        For Each store In stores
            
            '�x�X�̉c�ƊJ�n�����i�[
            storeDate = stores(store)
            newStore = False
        
'�G���[�߂�
newCreateStoreSheet:

            '�V�Kbook/�V�X�A���A���A�f�[�^�N�Z�����x�X�c�Ɠ����Â���΁A�x�X�V�[�g�̍쐬,�x�X������
            If (newcheck Or newStore) And (storeDate < toDay) Then
                areaBK.Worksheets("�x�X").Copy after:=areaBK.Worksheets(areaBK.Sheets.Count)
                Set store_ws = areaBK.Worksheets(areaBK.Sheets.Count)
                store_ws.name = store
                store_ws.Cells(3, 2).Value = store
                
            '�V�Kbook�ł͂Ȃ��A�f�[�^����
            ElseIf storeDate < toDay Then
               '�G���[�L���b�`  try-catch
               On Error GoTo storeWsError2
               
                    '�u�b�N�A�]�L�p�V�[�g�̊i�[
                    Set store_ws = areaBK.Worksheets(store)
                    '�����F�����r���ŐV�X������������̓G���[�ɂȂ邽�߁A�Ώ���
               
               '�G���[�Ď����s
               On Error GoTo 0
                                   
            '�V�x�X�A���̃f�[�^�N�Z�����_�ł͉c�Ƃ��Ă��Ȃ�
            Else
                GoTo nextStore  '���̎x�X��
            End If
               
            
            '******* �f�[�^�p�z����񂵂āA�x�X�̊e���l��]�L
            For Each dt In data_name
                
                '�f�[�^�u�b�N�A���o�p�V�[�g�̊i�[
                Set data_ws = dataBK.Worksheets(dt)
                
                '�V�[�g���𔻒�
                If InStr(data_ws.name, "") > 0 Then
                    '�V�[�g�F�x�X�Ńt�B���^   �x�X�� A��
                    data_ws.Cells(1, 1).AutoFilter Field:=1, Criteria1:=store
                Else
                    '���ރV�[�g�F�x�X�Ńt�B���^   �x�X�� C��
                    data_ws.Cells(1, 1).AutoFilter Field:=3, Criteria1:=store
                End If
            
                    '�]�L����  �����F�e�V�[�g�Q�ƁA�������ݗ�
                    errcheck = outData(store_ws, data_ws, c)
                
                '��񂪂Ȃ��Ƃ��A�����𒆒f����
                If errcheck Then
                    'book�͕ۑ���������
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
                
nextStore:
        Next store  '���̎x�X�V�[�g��
        
        '=================== �V�[�g���v�̌v�Z����(�f�[�^������) ===================
               
        'ws�A�e���ލ��v�̌v�Z
        Call areaCalc(areaBK, area, c)

        '******************* �f�[�^�t�@�C���؂�ւ�
        
            '�f�[�^�t�@�C�������
            Application.DisplayAlerts = False
            dataBK.Close False
            Application.DisplayAlerts = True
                                
            '��񃊃Z�b�g
            Erase dateSplit
            Set dataBK = Nothing
            
        Next file '���̃f�[�^�t�@�C��
        
        '=================== �]�L���� ===================
        
        '�ЂȌ^�V�[�g���\����
        If areaBK.Worksheets("�x�X").Visible = True Then areaBK.Worksheets("�x�X").Visible = False
        
        '�V�[�g�\���̓V�[�g��擪��
        areaBK.Worksheets(area).Activate
        
        '�u�b�N��ۑ����ĕ���
        areaBK.Close savechanges:=True
        
        Set stores = Nothing
        Set areaBK = Nothing
              
            tws.Activate

        '�����̏������� =====================
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
        
    Next area '����
    
    
    '�I���̏����@*********************************************************
    
    thisS.AutoFilterMode = False  '�t�B���^����
    
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
    
    MsgBox "�������������܂����B", vbInformation

    '�{�������I��
    Exit Sub

'���G���[�Ώ�
'�V�X���A�G���[���� ========================================
storeWsError2:
                '�V�K�x�X�������r���ɂł����Ƃ�
                newStore = True
                areaBK.Worksheets("�x�X").Visible = True
                err.Clear
                
                     Application.ScreenUpdating = False
                
                '�G���[�Ď����s
                On Error GoTo 0
                
                '�{�����ɖ߂�
                Resume newCreateStoreSheet
                

'book�ɃV�[�g���Ȃ������Ƃ��A�G���[���� ========================================
nothingArea2:
                newcheck = True
                err.Clear
                
                    Application.ScreenUpdating = False
        
                On Error GoTo 0
                
                Resume newCreateAreaSheet
End Sub

'�͈͂̎x�X���A�J�X����Ԃ��֐�
Private Function storeDatas(ByRef ws As Worksheet) As Object
    
    'ws �x�X�R�[�hsheet
    
    Dim r As Long
    Dim rng As Range
    Dim var As Variant
    
    Dim stores As Object
    Set stores = CreateObject("Scripting.Dictionary")

        
        '�͈͂́A�ŏI�s���i�[
        r = ws.Cells(Rows.Count, 5).End(xlUp).Row

        '�x�X���͈͂��i�[
        Set rng = ws.Range(ws.Cells(2, 5), ws.Cells(r, 5)) 'E�J�n�s:E�ŏI�s
    
        
    '�͈͂��񂵂Ďx�X����z��Ɋi�[     ���Z���݂̂ŉ񂷎w��
    For Each var In rng.SpecialCells(xlCellTypeVisible)
    
        stores.Add var.Value, var.Offset(0, 1).Value
                       
    Next var
    
    '�߂�l
    Set storeDatas = stores
    
    Set rng = Nothing

End Function

'�u�b�N���J����Ă��邩�`�F�b�N
Private Function areaBKcheck(areaFileName As String) As Boolean

    ' �J���Ă��邷�ׂẴu�b�N�𑖍�
    Dim wb As Workbook
    For Each wb In Workbooks
        ' �u�b�N������v������True��Ԃ�
        If wb.name = areaFileName Then
            areaBKcheck = True
            Exit Function
        End If
    Next wb
    
End Function

'�������ރf�[�^�̑̍ق𐮂���
Private Sub fixUpData(ByRef dataBK As Workbook)

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


    Dim ws As Worksheet
    Dim str As String
    
    Dim cate As Variant
    Dim cateArr As Variant
                
    Dim lastCol As Long
    Dim j As Long
    
    
    '============================== �V�[�g�̑̍ق����낦��
    
    '�f�[�^�u�b�N�̃V�[�g���ׂĂ��P��������
    For Each ws In dataBK.Worksheets
    
        str = ws.name

        '********************** �V�[�g�́A��폜������
        If InStr(str, "") > 0 Then

            '�Z����������
            ws.UsedRange.UnMerge

            '1�s�ځ`12�s�ڂ܂ō폜
            ws.Range("1:12").Delete shift:=xlUp


            '�ŏI�񂪉ςȂ̂ŁA�V�[�g���ɍŏI�s���i�[����
            lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

                '�ŏI�񂩂�t���[�v������
                For j = lastCol To 1 Step -1

                    'j��Z���̕�������i�[
                    str = ws.Cells(1, j).Value

                '������̐���
                Select Case str <> ""
                    '�܂܂�Ă�����Ȃɂ����Ȃ�
                    Case str Like "�x�X?", str Like "*��", str Like "*��"

                    '�����ȊO�͗�폜
                    Case Else
                        ws.Columns(j).Delete

                End Select

                '�󔒗񎞂̏���
                If str = "" Then ws.Columns(j).Delete

                Next j '���̗��
                        
        End If
            
            ws.Columns("A").AutoFit
            
    Next ws '���̃V�[�g��
    
    
    '============================== ���ރV�[�g�����
                
    Dim r As Long
    Dim c As Long
    
    Dim lastR As Long
    Dim startR As Long
    Dim i As Integer
               
    Dim bun As Worksheet
    Set bun = dataBK.Sheets.Add(dataBK.Worksheets(1))
        
        bun.name = "����"
        cateArr = Array("�啪��", "������")
        
        lastR = 0
       startR = 1
            i = 1
            j = 0

        '��/�����ރV�[�g����l�������Ă���
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
        
        
        '�񕝒���
        bun.Columns.AutoFit
        
        Set ws = Nothing

End Sub

'�������ݗ��Ԃ��֐�
Private Function lastOut_col(out_date As Date, ByRef aws As Worksheet) As Long

    '�������ݓ��t�@out_date
    '�V�[�g�@aws
            
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
        
        '���͊��ԓ��t�̔���
        For Each d In dayRng
            '�l����v�����Z���̗��Ԃ�
            If d.Value = out_d Then
                    col = d.Column
                    Exit For
            End If
        Next d
        
        On Error GoTo 0
        
        '���t�w��Ɍ�肪����Ƃ�
        If col = 0 Then
        
            MsgBox "�o�͗�̎擾�ŃG���[���������܂����B" & vbCrLf & _
                   "�����𒆒f���܂��B", vbCritical
            
            'book��ۑ���������
            Application.DisplayAlerts = False
            aws.Parent.Close False
            Application.DisplayAlerts = True
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        
            End  '�v���O�����S�I���i���f�j
        
        End If
                
  
    '�o�͗��Ԃ�
    lastOut_col = col
    Exit Function

'�G���[�Ώ�
'range�Q�Ƃ̎擾�R�� ==================================================
dayRangeErr:
            
            MsgBox "�o�͌��̔���ŃG���[���������܂����B" & vbCrLf & _
                   "�����𒆒f���܂��B", vbCritical
            
            'book��ۑ���������
            Application.DisplayAlerts = False
            aws.Parent.Close False
            Application.DisplayAlerts = True
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        
            End  '�v���O�����S�I���i���f�j
    
End Function


