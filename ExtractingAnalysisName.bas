Attribute VB_Name = "Module1"

Option Explicit

Sub sample6()
    Dim file As String                      '�J����Analysis�̃t�@�C�����O
    Dim i As Long: i = 1                '�Z���Ԓn
    Dim filePath As String             '�t�@�C���̃p�X
    Dim opnAls As Workbook       '�J����Analysis�u�b�N
    Dim WS As Worksheet             '���[�N�V�[�g
    Dim find_name As String        'Sheet1��T���ϐ�
    Dim flg As Boolean                 'Sheet1�����邩�ǂ����̃t���O
    Dim fileNum As Long             'Analysis�̃t�@�C���ԍ�
    Dim newSheet As Worksheet '�V�u�b�N
    Dim newSheetName As String  '�V�u�b�N�̖��O
    Dim rptName As String            '���|�[�g���[��
    Dim techKey() As String
    Dim prmptName() As String
    Dim alsWs As Worksheet
    Dim j As Long
    Dim lstNum As Long
    
    
    'Application.ScreenUpdating = False
    filePath = ThisWorkbook.Worksheets("Sheet1").Range("E3").Value & "\" 'E3�Ńt�@�C���p�X���擾
    file = Dir(filePath & "*.xlsx") '�p�X�z���g���q�u.xlsx�v�̍ŏ��̃t�@�C�������擾
    find_name = "Sheet2"
    flg = False
    
       '-----------------�t�H���_���̃u�b�N�J��Ԃ��J������-----------------
    
    Do While file <> ""
    
        '-----------------���̃u�b�N��Sheet2�����邩�ǂ���-----------------
        
            For Each WS In ThisWorkbook.Worksheets
            'ws�̖��O��find_name("Sheet2")�Ɠ�����������
                If WS.name = find_name Then flg = True 'flg��true��
            Next WS
            
        '-----------------Sheet2�������---------------------------------
        fileNum = Left(file, 4)     '�t�@�C�����̓��S�����擾�A�܂�Ǘ��ԍ�
            If flg = True Then
                Dim p, s, d1, d2 As Long
                s = ThisWorkbook.Worksheets("Sheet2").Cells(1, 1).End(xlDown).row
                d1 = ThisWorkbook.Worksheets("Sheet2").Cells(1, 1).Value
                If d1 = fileNum Then
                     GoTo Label1  '�J�����ɔ��
                End If
            
                For p = 2 To s
                d2 = ThisWorkbook.Worksheets("Sheet2").Cells(p, 1).Value
                If d2 = fileNum Then
                    GoTo Label1 '�J�����ɔ��
                End If
                Next p
            End If

        flg = False
        Set opnAls = Workbooks.Open(filePath & file) 'Analysis�Y���u�b�N���J��opnAls�ɓ����
        'Set opnAls = Workbooks.Open(filePath & "0914_�׽�ݔ̔����чD���ǉ�_? �F�_�l_�̔��Ǘ� .xlsx") 'Analysis�Y���u�b�N���J��opnAls�ɓ����
        file = opnAls.name          'Analysis�Y���u�b�N(opnAls)�̃t�@�C�������擾
'        fileNum = Left(file, 4)     '�t�@�C�����̓��S�����擾

         
                'rptName = opnAls.Worksheets("Sheet1").Cells(1, 1).Value '�Y���t�@�C����A1�̒l���擾
                Dim lResult As Variant
                lResult = Application.Run("SAPGetSourceInfo", "DS_1", "DataSourceName")
                Dim Worksheet As Worksheet, flag As Boolean
                flag = False
                                
                Dim del As String
                del = "_" ' �A���_�[�X�R�A��T��
                Dim a As String
                a = InStrRev(file, del)  '�Ō��"_"�̈ʒu���擾
                Dim b As String
                b = Len(file)    'file���̒������擾
                Dim c As String
                c = b - a  'file������Ō�̃A���X�R�̈ʒu������
                
                Debug.Print InStrRev(file, del)
                Dim fileLastNm As String
                fileLastNm = Right(file, c) 'right�ōŌ�̃A���X�R�̌����擾
                
                For Each Worksheet In ThisWorkbook.Worksheets
                    If Worksheet.name = "Sheet2" Then flag = True
                Next Worksheet
                If flag = False Then
                    Set newSheet = ThisWorkbook.Worksheets.Add '�V�V�[�g�쐬��newBook�ɑ��
                    newSheet.name = "Sheet2" '�V�V�[�g�̖��O���擾
                    newSheetName = newSheet.name
                End If
                    
                Set newSheet = ThisWorkbook.Worksheets(1)

                If newSheet.Cells(i, 2).Value = "" And newSheet.Cells(i, 1).Value = "" And newSheet.Cells(i, 3).Value = "" Then  '�V�K�쐬����newBook�̃V�[�g�P��A1�̒l��null��������
                    newSheet.Cells(i, 2).Value = lResult 'rptName��V�u�b�N��A1�ɏo��
                    newSheet.Cells(i, 1).Value = Format(fileNum, "0000")
                    newSheet.Cells(i, 3).Value = fileLastNm
                Else
                    i = newSheet.Cells(1, 1).End(xlDown).row + 1
                    Debug.Print i
                    newSheet.Cells(i, 2).Value = lResult 'rptName��V�u�b�N��A1�ɏo��
                    newSheet.Cells(i, 1).Value = Format(fileNum, "0000")
                    newSheet.Cells(i, 3).Value = fileLastNm
                End If

                i = i + 1 '�C���N�������g���āA��s���ɏo�͂���B
'            On Error Resume Next
            Set alsWs = opnAls.Worksheets(1)
'
'          'newSheet�̍Ō�̍s���擾
           lstNum = getLastRow(newSheet)
            
           Debug.Print lstNum
           
'            Dim flagflag As Boolean
'            flagflag = False
'            'Sheet2��A�񂪃t�@�C�����ɂ�������
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
'               Debug.Print fileNum & "���܂ޔz��͑��݂���"
'           Else
'                Debug.Print fileNum & "���܂ޔz��͑��݂��܂���"
'           End If
           
'           Dim varResult
'           varResult = Filter(testArray, fileNum)
'
'           If UBound(varResult) <> -1 Then
'                Debug.Print file & "�͔z����ɑ��݂���"
'           Else
'                Debug.Print file & "�͔z����ɑ��݂��Ȃ�"
'           End If
'��������
'           Debug.Print UBound(testArray) '6
'           Debug.Print LBound(testArray)   '1
'           'newBook.Worksheets("Sheet1").Range("C1", Cells(LBound(testArray, 1), UBound(testArray, 1))).Value = testArray
'           Dim k As Long
'           Dim l As Long
'           For k = 3 To lstNum - 1
'               newBook.Worksheets("Sheet1").Cells(1, k).Value = testArray(l)
'           Next k
'�����܂ł���Ă���r��


        DoEvents
        Workbooks(file).Close SaveChanges:=False  '�����グ��Analysis�u�b�N��ۑ����Ȃ��ŕ���
Label1:
        file = Dir
        ThisWorkbook.Save
    Loop
    
    'Call find_NA
    
    'A1���L�[�ɂ��ď����ɂ���
    Call Range("A:C").Sort( _
    Key1:=Range("A1"), _
    Order1:=xlAscending)
    
    MsgBox "�I���ł�����"
    'Application.ScreenUpdating = True
End Sub

'���[�u���b�N�̍ŏI�s���擾���郁�\�b�h
Function getLastRow(WS As Worksheet, Optional CheckCol As Long = 1) As Long
    getLastRow = WS.Cells(1, CheckCol).End(xlDown).row
End Function



Public Function find_NA()
    Dim filePath As String
    Dim fileNum As String
    Dim newSheet As Worksheet
    Dim lstNum As Long
    Dim i As Long
    filePath = ThisWorkbook.Worksheets("Sheet1").Range("E3").Value & "\" 'E3�Ńt�@�C���p�X���擾
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
