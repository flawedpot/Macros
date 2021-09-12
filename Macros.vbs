Option Explicit

'���ׂẴV�[�g���g�嗦100%�ɂ���A1�Z����I������
Sub AlignAll()

    Dim zoomRate As String      '�g�嗦
    Dim i As Integer
    
    zoomRate = InputBox("�g�嗦����͂��Ă�������", "�g�嗦����", 100)
    
    '�g�嗦���͂����͂��ꂽ�ꍇ
    If zoomRate <> "" Then
    
        '���ׂẴ��[�N�V�[�g�ɑ΂��Ď��{
        For i = 1 To Worksheets.Count
            With ActiveWindow
                .Zoom = zoomRate    '�g�嗦��ݒ�
                .ScrollRow = 1      '��ԏ�ɃX�N���[��
                .ScrollColumn = 1   '��ԍ��ɃX�N���[��
            End With
            Worksheets(i).Range("A1").Select    'A1�Z���I��
        Next i
        '�擪�V�[�g���A�N�e�B�x�[�g
        Worksheets(1).Activate
        
    Else
    
        '�����I��
        Exit Sub
        
    End If
    
End Sub

'�I�𒆂̃Z���̉��ɍs��ǉ�����
Sub AddRow()

    Rows(ActiveCell.Row + 1).Insert Shift:=xlDown
    
End Sub

'�l�\��t�����s��
'(�ʓr�}�N���̃V���[�g�J�b�g�L�[�ɖ{�}�N����ݒ肷�邱��)
Sub PasteValue()

    Selection.PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
        
End Sub

'�I��͈͂��s��������
Sub MergeRow()

    Call MergeAreas(0)
    
End Sub

'�I��͈͂�񌋍�����
Sub MergeColumn()

    Call MergeAreas(1)

End Sub

'������0�̂Ƃ��A�I��͈͂��s��������
'������1�̂Ƃ��A�I��͈͂�񌋍�����
Sub MergeAreas(ByVal mode As Integer)

    '�ϐ��錾
    Dim startRow As Integer     '�I��͈͂̍���Z���̍s
    Dim startColumn As Integer  '�I��͈͂̍���Z���̗�
    Dim endRow As Integer       '�I��͈͂̉E���Z���̍s
    Dim endColumn As Integer    '�I��͈͂̉E���Z���̗�
    Dim i, j As Integer
    
    '���ׂĂ̑I��͈͂ɑ΂��Ď��s
    For i = 1 To Selection.Areas.Count
    
        '�ϐ��̐ݒ�
        startRow = Selection.Areas(i).Row
        startColumn = Selection.Areas(i).Column
        endRow = startRow + Selection.Areas(i).Rows.Count - 1
        endColumn = startColumn + Selection.Areas(i).Columns.Count - 1
        
        '�I��͈͂��ォ�珇�ɍs����
        If mode = 0 Then
            For j = startRow To endRow
                Range(Cells(j, startColumn), Cells(j, endColumn)).Merge
            Next j
            
        '�I��͈͂������珇�ɗ񌋍�
        Else
            For j = startColumn To endColumn
                Range(Cells(startRow, j), Cells(endRow, j)).Merge
            Next j
        End If
        
    Next i

End Sub

'����������argStr�������͈�argRange�̍Ōォ�琔���ĉ��Ԗڂɂ��邩�����߁A
'�Y������s��argCol��ڂɂ��镶����Ԃ�
Function VLOOKUPREV(ByVal argRange1 As Range, _
                    ByVal argRange2 As Range, _
                    ByVal argCol As Integer) As String

    '�ϐ��錾
    Dim ret As String   '�߂�l
    Dim i As Integer
    
    '�ϐ�������
    ret = ""
    
    '����1�̍s�܂��͗�2�ȏ�̂Ƃ��G���[��Ԃ�
    If argRange1.Rows.Count > 1 Or _
       argRange1.Columns.Count > 1 Then
        
        ret = CVErr(xlErrValue)
        
    Else
        '����2�͈̔͂̍ŉE�񂩂����1�ƈ�v����Z����T��
        For i = 1 To argRange2.Rows.Count
        
            '��v�����s�ɂ������3�̗�ɂ��镶�����߂�l�ɐݒ�
            If argRange1.Value = argRange2.Cells(i, argRange2.Columns.Count).Value Then
                ret = argRange2.Cells(i, argCol).Value
                Exit For
            
            '��v���Ȃ��܂ܒT���I��������G���[��Ԃ�
            ElseIf i = argRange2.Rows.Count Then
                ret = CVErr(xlErrValue)
            End If
            
        Next i
        
    End If
    
    '�߂�l��ݒ�
    VLOOKUPREV = ret
    
End Function