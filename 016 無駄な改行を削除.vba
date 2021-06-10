Option Explicit
'#VBA100本ノック 16本目
'セル内改行はAlt Enterですね｡
'引数の文字列から無駄な改行（LF）を削除して返すFunctionを作成してください。
'※CRLFはLFに変換する｡
'■無駄な改行とは
'・文字列の前後の改行
'・2連続の改行
'サンプル:改行(\n)
'\n無駄な\n\n改行を\n\n\n削除\n\n
'↓
'無駄な\n改行を\n削除

Function vbaknocks_016(ByVal str As String) As String
            
    str = Replace(str, vbCrLf, vbLf)
    
    Do Until InStr(str, vbLf & vbLf) = 0
        str = Replace(str, vbLf & vbLf, vbLf)
    Loop
    
    If Left(str, 1) = vbLf Then
        str = Right(str, Len(str) - 1)
    End If
    
     If Right(str, 1) = vbLf Then
        str = Left(str, Len(str) - 1)
    End If
    
    vbaknocks_016 = str
End Function


'#VBA100本ノック 16本目 解答
'文字列操作としては基本になります｡
'・Do...Loop（再帰含む）
'・正規表現
'・Split+Join(TEXTJOIN含む)
'・１文字ずつ処理
'・その他
'大きくわけるとこのようになると思います｡
'まずは基本と言いますか、最も単純な方法のDo…Loopから。
'
'Function VBA100_16_01(ByVal arg As String) As String
'    arg = Replace(arg, vbCrLf, vbLf)
'
'    Dim oLen As Long
'    Do
'        oLen = Len(arg)
'        arg = Replace(arg, vbLf & vbLf, vbLf)
'    Loop Until Len(arg) = oLen
'
'    If Left(arg, 1) = vbLf Then arg = Mid(arg, 2)
'    If Right(arg, 1) = vbLf Then arg = Left(arg, Len(arg) - 1)
'    VBA100_16_01 = arg
'End Function
'
'
'文字列操作では必ず出てくる正規表現｡
'なかなか使いこなすのは大変ですが､使えるととても便利です｡
'すこしずつ慣れていければ良いと思います｡
'Split Join､1文字ずつ処理､その他､これらのサンプルVBAは記事補足に掲載しました｡
'
'Function VBA100_16_02(ByVal arg As String) As String
'    With CreateObject("VBScript.RegExp")
'        .Global = True
'        .Pattern = "^\n+|\n+$|\n+(?=\n)"
'        VBA100_16_02 = .Replace(Replace(arg, vbCrLf, vbLf), "")
'    End With
'End Function
'
'
'補足
'Do...Loop（再帰含む）
'Do...Loopの終了条件は、置換する文字が無くなるまでになります。
'先のVBAでは､Len関数でReplace前後の文字列長に変化がなくなるまでにしています｡
'もっと単純に､置換対象の文字がなくなるまでという判定もできます｡
'
'    Do While InStr(arg, vbLf & vbLf) > 0
'        arg = Replace(arg, vbLf & vbLf, vbLf)
'    Loop
'
'vbLf & vbLf
'これが2回出てきてしまうので､この場合は事前に変数に入れておきたいところです｡
'
'終了条件が明確なDo...Loopなので、これは結構簡単に再帰に書き換えができます。
'（ただし、この内容では効率があまりよろしくないです。）
'再帰についてはいずれお題を出す予定です｡
'
'正規表現
'パターンの書き方はいろいろありますので､先のパターンはあくまで1例です｡
'CRLFの置換は､VBAのReplace一発なので､これを使っています｡
'
'VBAで正規表現を利用する（RegExp）｜VBA技術解説
'正規表現は複雑なパターンマッチングとテキストの検索置換するためのツールです、VBAで正規表現を使う場合はRegExpオブジェクトを使用します、RegExpは、VBScriptに正規表現として用意されているオブジェクトです。目次 メタ文字 正規表現 正規表現RegExpの使い方 RegExpオブジェクト RegExp…
'
'Split Join(TEXTJOIN含む)
'Split関数でLFで分割して配列化してから処理するものです｡
'この場合､LFLFとつづいている場合は､配列内に空の要素が出来てしまうので､これの対処が必要になります｡
'
'Function VBA100_16_04(ByVal arg As String) As String
'    If arg = "" Then Exit Function
'    Dim i As Long, v As Variant
'    Dim ary() As String
'    ReDim ary(1 To Len(arg))
'    For Each v In Split(Replace(arg, vbCrLf, vbLf), vbLf)
'        i = i + 1
'        If v <> "" Then ary(i) = v & vbLf
'    Next
'    VBA100_16_04 = Join(ary, "")
'    VBA100_16_04 = Left(VBA100_16_04, Len(VBA100_16_04) - 1)
'End Function
'
'
'Excel2016以降なら､シートのTEXTJOIN関数が使えます｡
'これなら､空の要素を無視してくれるので､とても簡単に済みます｡
'
'Function VBA100_16_05(ByVal arg As String) As String
'    VBA100_16_05 = WorksheetFunction.TextJoin(vbLf, True, IIf(arg = "", "", Split(Replace(arg, vbCrLf, vbLf), vbLf)))
'End Function
'
'1 文字ずつ処理
'とにかく､自力で1文字ずつLFを処理していこうというものです｡
'
'Function VBA100_16_03(ByVal arg As String) As String
'    arg = Replace(arg, vbCrLf, vbLf)
'
'    Dim rtn As String
'    Dim i As Long
'    Dim flgLF As Boolean
'    Dim flgOut As Boolean
'
'    flgLF = False
'    For i = 1 To Len(arg)
'        flgOut = True
'        Select Case True
'            Case flgLF
'                flgLF = False
'            Case Mid(arg, i, 1) = vbLf
'                flgOut = False
'            Case Else
'                flgLF = True
'        End Select
'        If flgOut Then rtn = rtn & Mid(arg, i, 1)
'    Next
'
'    If Right(rtn, 1) = vbLf Then rtn = Left(rtn, Len(rtn) - 1)
'    VBA100_16_03 = rtn
'End Function
'
'この方法の場合は､細部の書き方は人により千差万別になると思います｡
'興味があれば読み解いてみてください｡
'
'その他
'今回のようなお題としては単純な処理は､方法が多種多用に存在します｡
'私の想像の及ばない方法もあるかもしれません｡
'以下は､シートのTrim関数を使った､ちょっとトリッキーな方法になります｡
'
'Function VBA100_16_06(ByVal arg As String) As String
'    arg = Replace(arg, vbCrLf, vbLf)
'    arg = Replace(arg, " ", Chr(1))
'    arg = Replace(arg, "　", Chr(2))
'
'    arg = Replace(arg, vbLf, " ")
'    arg = WorksheetFunction.Trim(arg)
'    arg = Replace(arg, " ", vbLf)
'
'    arg = Replace(arg, Chr(1), " ")
'    arg = Replace(arg, Chr(2), "　")
'    VBA100_16_06 = arg
'End Function
'
'Chr(1)とChr(2)を犠牲にして、半角空白と全角空白を一旦退避してしてから、
'LFを空白置換して､Trimで前後の空白と間の無駄な空白を取り除いています｡
'その後で、空白をLFに戻して、Chr(1)とChr(2)に変更しておいた半角・全角空白を元に戻しています。
'
'Chr(1)とChr(2)は、単に通常の入力では入らない文字コードなら何でも良いので、分かり易いコードを選んだという事です。
