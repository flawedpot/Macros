Option Explicit

'すべてのシートを拡大率100%にしてA1セルを選択する
Sub AlignAll()

    Dim zoomRate As String      '拡大率
    Dim i As Integer
    
    zoomRate = InputBox("拡大率を入力してください", "拡大率入力", 100)
    
    '拡大率入力が入力された場合
    If zoomRate <> "" Then
    
        'すべてのワークシートに対して実施
        For i = 1 To Worksheets.Count
            With ActiveWindow
                .Zoom = zoomRate    '拡大率を設定
                .ScrollRow = 1      '一番上にスクロール
                .ScrollColumn = 1   '一番左にスクロール
            End With
            Worksheets(i).Range("A1").Select    'A1セル選択
        Next i
        '先頭シートをアクティベート
        Worksheets(1).Activate
        
    Else
    
        '処理終了
        Exit Sub
        
    End If
    
End Sub

'選択中のセルの下に行を追加する
Sub AddRow()

    Rows(ActiveCell.Row + 1).Insert Shift:=xlDown
    
End Sub

'値貼り付けを行う
'(別途マクロのショートカットキーに本マクロを設定すること)
Sub PasteValue()

    Selection.PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
        
End Sub

'選択範囲を行結合する
Sub MergeRow()

    Call MergeAreas(0)
    
End Sub

'選択範囲を列結合する
Sub MergeColumn()

    Call MergeAreas(1)

End Sub

'引数が0のとき、選択範囲を行結合する
'引数が1のとき、選択範囲を列結合する
Sub MergeAreas(ByVal mode As Integer)

    '変数宣言
    Dim startRow As Integer     '選択範囲の左上セルの行
    Dim startColumn As Integer  '選択範囲の左上セルの列
    Dim endRow As Integer       '選択範囲の右下セルの行
    Dim endColumn As Integer    '選択範囲の右下セルの列
    Dim i, j As Integer
    
    'すべての選択範囲に対して実行
    For i = 1 To Selection.Areas.Count
    
        '変数の設定
        startRow = Selection.Areas(i).Row
        startColumn = Selection.Areas(i).Column
        endRow = startRow + Selection.Areas(i).Rows.Count - 1
        endColumn = startColumn + Selection.Areas(i).Columns.Count - 1
        
        '選択範囲を上から順に行結合
        If mode = 0 Then
            For j = startRow To endRow
                Range(Cells(j, startColumn), Cells(j, endColumn)).Merge
            Next j
            
        '選択範囲を左から順に列結合
        Else
            For j = startColumn To endColumn
                Range(Cells(startRow, j), Cells(endRow, j)).Merge
            Next j
        End If
        
    Next i

End Sub

'検索文字列argStrが検索範囲argRangeの最後から数えて何番目にあるかを求め、
'該当する行のargCol列目にある文字を返す
Function VLOOKUPREV(ByVal argRange1 As Range, _
                    ByVal argRange2 As Range, _
                    ByVal argCol As Integer) As String

    '変数宣言
    Dim ret As String   '戻り値
    Dim i As Integer
    
    '変数初期化
    ret = ""
    
    '引数1の行または列が2以上のときエラーを返す
    If argRange1.Rows.Count > 1 Or _
       argRange1.Columns.Count > 1 Then
        
        ret = CVErr(xlErrValue)
        
    Else
        '引数2の範囲の最右列から引数1と一致するセルを探索
        For i = 1 To argRange2.Rows.Count
        
            '一致した行にある引数3の列にある文字列を戻り値に設定
            If argRange1.Value = argRange2.Cells(i, argRange2.Columns.Count).Value Then
                ret = argRange2.Cells(i, argCol).Value
                Exit For
            
            '一致しないまま探索終了したらエラーを返す
            ElseIf i = argRange2.Rows.Count Then
                ret = CVErr(xlErrValue)
            End If
            
        Next i
        
    End If
    
    '戻り値を設定
    VLOOKUPREV = ret
    
End Function