Option Explicit
'�������ރf�[�^�����C���u�b�N�Ɏ擾����
Sub dataInput()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


    Dim thisBK As Workbook
    Set thisBK = ThisWorkbook
    Dim ws As Worksheet
    Dim str As String
    
    Dim cate As Variant
    Dim cateArr As Variant
        cateArr = Array("����", "�e��������")
        
    '============================== ���C���u�b�N�Ƀf�[�^�V�[�g���c���Ă���Ȃ��������
       
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
    
    '============================== �f�[�^�t�@�C���̎擾 ==============================
    
        ChDir thisBK.path
    
    Dim openFile As String
        
    '�f�[�^�t�@�C����I��ł��炤
    openFile = Application.GetOpenFilename("Microsoft Excel�u�b�N,*.xlsx")
    
        If openFile = "False" Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
    
    Dim dataBK As Workbook
    Set dataBK = Workbooks.Open(openFile)
        
    Dim lastCol As Long
    Dim j As Long
    
    
    '============================== �V�[�g�̑̍ق����낦��
    
    '�f�[�^�u�b�N�̃V�[�g���ׂĂ��P��������
    For Each ws In dataBK.Worksheets
    
        str = ws.name

        '********************** �e���V�[�g�́A��폜������
        If InStr(str, "�e��") > 0 Then

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
                    Case str Like "�X", str Like "*�e����", str Like "*�e����"

                    '�����ȊO�͗�폜
                    Case Else
                        ws.Columns(j).Delete

                End Select

                '�󔒗�A�s�v�񂾂������̏���
                If str = "" Or InStr(str, "�p") > 0 Then ws.Columns(j).Delete

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
        
        
    '============================== ���C���u�b�N�Ƀf�[�^�������Ă���
    
    For Each ws In dataBK.Worksheets
    
        str = ws.name
    
        If str = "����" Or InStr(str, "�e��") > 0 Then
            '�f�[�^�V�[�g�Ȃ�΁A���C���u�b�N�ɃV�[�g�R�s�[
            dataBK.Worksheets(str).Copy after:=thisBK.Worksheets(thisBK.Sheets.Count)
        End If
    Next ws
    
        '�f�[�^�t�@�C�������
        Application.DisplayAlerts = False
        dataBK.Close False
        Application.DisplayAlerts = True
    
    thisBK.Worksheets("���C��").Activate
    
    Set dataBK = Nothing
    Set thisBK = Nothing
    Set ws = Nothing
    Set bun = Nothing
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "�f�[�^�̎擾���������܂����B", vbInformation

End Sub

'���C���u�b�N�̃f�[�^�V�[�g���폜����
Sub sheetDelete()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ans As VbMsgBoxResult
    
        ans = MsgBox("���̃t�@�C���Ɏ�荞�܂ꂽ�f�[�^�V�[�g���폜���܂��B" & vbCrLf & "��낵���ł����H", vbYesNo + vbExclamation, "�V�[�g�̍폜")
    
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
        cateArr = Array("����", "�e��")
        
    '============================== ���C���u�b�N�̃f�[�^�V�[�g����������
       
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
