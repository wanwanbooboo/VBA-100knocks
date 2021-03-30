Option Explicit
'#VBA100本ノック 15本目
'「2020年04月」から「2021年03月」のシートがあります。
'しかし､シートの順番が狂ってしまっています｡
'「2020年04月」から昇順に並べ替えてください。
'・上記以外のシートは存在しません。
'・シート名は「yyyy年mm月」これで統一されています。
Sub vbaknocks_015()
    Dim i As Integer
    Dim num As Integer: num = Worksheets.Count
    
    If num = 1 Then
        MsgBox "並び替える必要はありません"
        Exit Sub
    End If
    
    Do While num > 0
        For i = 1 To num - 1
            If Worksheets(i).name > Worksheets(i + 1).name Then
                Worksheets(i).Move After:=Worksheets(i + 1)
            End If
        Next i
        num = num - 1
    Loop
End Sub

'解答例
'方法は大きく分けて3通り
'・シート名が年月限定として、4月から順に並べる
'・シート名の大小比較して直接シートを並べ替える
'・シート名を新規シートに出力してソートし、その順に並べる
'今回の問題の条件限定なら最初の方法が最も簡単でしょう｡
'
'Sub VBA100_15_01()
'    Const startYM As Date = #4/1/2020#
'    Dim i As Long
'    On Error Resume Next '万一のシート抜け対応
'    For i = 1 To 12
'        Sheets(Format(DateAdd("m", i - 1, startYM), "yyyy年mm月")).Move After:=Sheets(Sheets.Count)
'    Next
'    Sheets(1).Select
'End Sub
'
'
'2 番目の方法はいわゆるバブルソートをシートで行えばよいでしょう｡
'3 番目の方法はエクセルの機能を活用した方法になります｡
'3 番目と考え方は同じですが､配列にシート名を入れて配列を並べ替える方法もあります｡
'ただし､VBAでは標準で配列の並べ替えがサポートされていません｡
'
'Sub VBA100_15_02()
'    Dim i As Long
'    Dim j As Long
'    For i = 12 To 1 Step -1
'        For j = 1 To i - 1
'            If Sheets(j).name > Sheets(j + 1).name Then
'                Sheets(j).Move After:=Sheets(j + 1)
'            End If
'        Next
'    Next
'    Sheets(1).Select
'End Sub
'
'
'最新のMicrosoft 365ならシートのSort関数が使えます。
'これが使えるようになると､VBAでの配列ソートが簡単におこなえるようになりますね｡
