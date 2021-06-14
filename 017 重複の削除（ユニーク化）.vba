Option Explicit
'#VBA100本ノック 17本目
'画像1のように部・課・氏名の「社員」シートがあります。
'このデータを基に、画像2のように部・課マスタを作成してください。
'※部・課でユニーク化するという事ことです。
'シート「部・課マスタ」は存在している前提で構いません。
'※マスタなのでコード順にしてください｡


Sub vbaknocks_017()


    Dim myDic As Object
    Set myDic = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    Dim j As Integer
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Set ws1 = Worksheets("社員")
    Set ws2 = Worksheets("部・課マスタ")
    Dim starttime As Double
    Dim endtime As Double
    Dim arry() As Variant
    Dim n As Integer
    
    n = 0
    starttime = Timer
  
    With ws1
        For i = 1 To .Cells(Rows.Count, 1).End(xlUp).Row
            If Not myDic.exists(.Cells(i, 4).Value) Then
                myDic.Add .Cells(i, 4).Value, .Cells(i, 6).Value
                n = n + 1
                ReDim Preserve arry(1 To 4, 1 To n)
                    For j = 1 To 4
                        arry(j, n) = .Cells(i, j + 2).Value
                    Next j
            End If
        Next i
    End With
    
    With ws2
        arry = WorksheetFunction.Transpose(arry)
        .Range(.Cells(1, 1), .Cells(12, 4)).Value = arry
        .Range("A1").Sort key1:=.Range("B1"), order1:=xlAscending, Header:=xlYes
    End With
    
    endtime = Timer

    MsgBox "Process time: " & endtime - starttime

    Set myDic = Nothing
    Set ws1 = Nothing
    Set ws2 = Nothing
End Sub

'ユニーク化する方法は沢山あります｡
'・関数+（オートフィルター/1行ずつ抽出）
'・並べ替えて上下比較
'・Dictionaryを使う
'・フィルターオプションの設定
'・重複の削除
'・ピボットテーブル
'・Power Query
'・UNIQUE関数
'色々あますが､まずはフィルターオプションの設定から｡
'
'Sub VBA100_17_01()
'    Dim ws社員 As Worksheet
'    Dim ws部課 As Worksheet
'    Set ws社員 = Worksheets("社員")
'    Set ws部課 = Worksheets("部・課マスタ")
'
'    ws部課.Cells.Clear
'    ws社員.Columns("C:F").AdvancedFilter Action:=xlFilterCopy, _
'                                         CopyToRange:=ws部課.Range("A1"), _
'                                         Unique:=True
'
'    With ws部課
'        .Range("A1").CurrentRegion.Sort key1:=.Range("A1"), order1:=xlAscending, _
'                                        key2:=.Range("B1"), order2:=xlAscending, _
'                                        Header:=xlYes
'    End With
'End Sub
'
'
'フィルターオプションの設定は､あくまでユニーク化にも使えるということであって､
'何十万件から重複データを消すというような場合はお勧めしません｡
'次にユニーク化と言ったらDictionaryが思い浮かんだ人も多いのではないでしょうか｡
'Dictionaryは用途が広く､使い慣れると何かと便利です｡
'
'Sub VBA100_17_02()
'    Dim ws社員 As Worksheet
'    Dim ws部課 As Worksheet
'    Set ws社員 = Worksheets("社員")
'    Set ws部課 = Worksheets("部・課マスタ")
'
'    Dim dic As Object
'    Set dic = CreateObject("Scripting.Dictionary")
'
'    Dim i As Long, tmp As String
'    With ws社員
'        For i = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
'            tmp = .Cells(i, 3) & vbTab & .Cells(i, 4)
'            If Not dic.exists(tmp) Then
'                dic.Add tmp, .Cells(i, 3).Resize(, 4).Value
'            End If
'        Next
'    End With
'
'    ws部課.Range("A1").CurrentRegion.Offset(1).ClearContents
'    Dim j As Long, v As Variant
'    j = 2
'    For Each v In dic.items
'        ws部課.Cells(j, 1).Resize(, 4).Value = v
'        j = j + 1
'    Next
'
'    With ws部課
'        .Range("A1").CurrentRegion.Sort key1:=.Range("A1"), order1:=xlAscending, _
'                                        key2:=.Range("B1"), order2:=xlAscending, _
'                                        Header:=xlYes
'    End With
'End Sub
'
'
'その他､関数 フィルター､並べ替えてから上下比較のVBAサンプルを記事補足に掲載しました｡
'方法は沢山あるので､いろいろ挑戦してみると面白いと思います｡
