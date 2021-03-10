Option Explicit

'#VBA100本ノック 12本目
'A1から始まる表範囲のC列に金額が入っています｡
'しかし､ところどころに結合されたセルがあります｡
'セル結合を解除し、入っている金額を整数で均等に割り振ってください。（2枚目画像）
'端数処理方法は任意とします｡
'※結合セルには正の整数しか入っていません｡

Sub vbaknocks_012()

    Dim ws As Worksheet: Set ws = Worksheets("セル結合")
    Dim rng As Range
    Dim area As Range
    Dim val As Integer
    Dim residue As Integer
    
    For Each rng In ws.Range("A1").CurrentRegion
        If rng.MergeCells Then
            Set area = rng.MergeArea    '結合セル範囲を記憶
            val = rng.Value    '結合セルの値を記憶
            rng.MergeCells = False  '結合解除
            area.Value = Int(val / area.Rows.Count)  '値を結合セル数で割って各セルに代入
            residue = val - (rng.Value * area.Rows.Count)  '割り振り前後の値の差を計算
            area.Resize(residue).Value = rng.Value + 1  '値の前後差を各セルに割り振る
        End If
        Set area = Nothing
    Next
    
    Set ws = Nothing
End Sub

'解答例
'セル結合の判定は､
'MergeArea.Count > 1
'これでも判定できます｡
'結合範囲のセル数が > 1 なら結合されていることになります。
'セル結合を解除するには､
'Range.MergeCells = False
'Range.UnMerge
'上はプロパティ､下はメソッドです｡
'シートやValue等､省略できるものは全て省略しました｡
'
'Sub VBA100_12_01()
'    Dim rng As Range
'    Dim i As Long, v As Long, m As Long
'    For i = 2 To Cells(Rows.Count, 3).End(xlUp).Row
'        Set rng = Cells(i, 3).MergeArea
'        If rng.Count > 1 Then
'            rng.UnMerge
'            v = rng(1)
'            rng = Int(rng(1) / rng.Count)
'            m = v - (rng(1) * rng.Count)
'            rng.Resize(m) = rng(1) + 1
'        End If
'    Next
'End Sub
'
'
'割り振りはいろいろな方法があります｡
'四捨五入すると面倒なので､切り捨ててから差分を割り振る方がマイナスにならなくて簡単だと思います｡
'上では、差分数のResizeした範囲に+1しています。
'今回は特に補足もないので省略せずに書き直したVBAだけ掲載しました｡
'
'
'訂正です｡
'm=０、つまり割り切れてあまりが無い場合の判定が漏れていました。
'
'
'Sub VBA100_12_01()
'    Dim rng As Range
'    Dim i As Long, v As Long, m As Long
'    For i = 2 To Cells(Rows.Count, 3).End(xlUp).Row
'        Set rng = Cells(i, 3).MergeArea
'        If rng.Count > 1 Then
'            rng.UnMerge
'            v = rng(1)
'            rng = Int(rng(1) / rng.Count)
'            m = v - (rng(1) * rng.Count)
'            If m <> 0 Then rng.Resize(m) = rng(1) + 1
'        End If
'    Next
'End Sub
'
'
'補足
'Sub VBA100_12_02()
'    Dim target As Range
'    Set target = ActiveSheet.Range("A1").CurrentRegion
'    Set target = Intersect(target, target.Offset(1, 2))
'
'    Dim rng As Range
'    Dim i As Long, v As Long, m As Long
'    For Each rng In target
'        If rng.MergeCells Then
'            Set rng = rng.MergeArea
'            rng.UnMerge
'            v = rng.Item(1)
'            rng.Value = Int(rng.Item(1) / rng.Count)
'            m = v - (rng.Item(1) * rng.Count)
'            If m <> 0 Then rng.Resize(m) = rng.Item(1) + 1
'        End If
'    Next
'End Sub
'
'最初のVBAとやっていることは同じです｡
'ForをFor Eachに変更して､あとはプロパティ等を省略せずに少し丁寧に書き直しただけの違いです｡
'
