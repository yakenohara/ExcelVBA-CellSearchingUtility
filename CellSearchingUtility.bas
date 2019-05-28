Attribute VB_Name = "Module1"
'
' 指定シートに一致するセルが存在するかどうかを返す
'
' findThis: 検索するセル
' inThisSheetName: 検索シート名
' lookAtPart: True で部分一致指定, false で完全一致指定
'
' Parameters of Range.Find method
'
' | Argment         | Constant    | Description                    |
' | --------------- | ----------- | ------------------------------ |
' | What            | -           | 検索するデータを指定(必須)     |
' | After           | -           | 検索を開始するセルを指定       |
' | LookIn          | xlFormulas  | 検索対象を数式に指定           |
' |                 | xlValues    | 検索対象を値に指定             |
' |                 | xlComents   | 検索対象をコメント文に指定     |
' | LookAt          | xlPart      | 一部が一致するセルを検索       |
' |                 | xlWhole     | 全部が一致するセルを検索       |
' | SearchOrder     | xlByRows    | 検索方向を列で指定             |
' |                 | xlByColumns | 検索方向を行で指定             |
' | SearchDirection | xlNext      | 順方向で検索(デフォルトの設定) |
' |                 | xlPrevious  | 逆方向で検索                   |
' | MatchCase       | True        | 大文字と小文字を区別           |
' |                 | False       | 区別しない(デフォルトの設定)   |
' | MatchByte       | True        | 半角と全角を区別する           |
' |                 | False       | 区別しない(デフォルトの設定)   |
'
' Parameters of Range.Find method
'
' | Constant   | Number | Display | Desctiption                                                                              |
' | ---------- | -----: | ------- | ---------------------------------------------------------------------------------------- |
' | xlErrDiv0  |   2007 | #DIV/0! | 0割り                                                                                    |
' | XlErrNA    |   2042 | #N/A    | 計算や処理の対象となるデータがない、または正当な結果が得られない                         |
' | xlErrName  |   2029 | #NAME?  | Excelの関数では利用できない名前(存在しない関数名等)が使用されている                      |
' | XlErrNull  |   2000 | #NULL!  | 半角空白文字の参照演算子で指定した2つのセル範囲に、共通部分がない(`=SUM(A1:A3 C1:C3)`等) |
' | XlErrNum   |   2036 | #NUM!   | 使用できる範囲外の数値を指定したか、それが原因で関数の解が見つからない                   |
' | XlErrRef   |   2023 | #REF!   | 数式内で無効なセルが参照されている                                                       |
' | XlErrValue |   2015 | #VALUE! | 関数の引数の形式が間違っている(数値を指定すべきところに文字列を指定等)                   |
'
Public Function isDefinedInThisSheet(ByVal findThis As Variant, ByVal inThisSheetName As Variant, Optional ByVal lookAtPart As Variant = True) As Variant
    
    Dim ret As Variant
    Dim sheetWasFound As Boolean 'シートが見つかったかどうか
    Dim cellWasFound As Boolean '見つかったかどうか
    Dim lookAtParam As Variant 'Range.Find method の LookAt parameter 用設定値
    
    'Range.Find method の LookAt parameter 用設定値の決定
    If lookAtPart Then '完全一致指定の場合
        lookAtParam = xlPart
    
    Else '完全一致指定でない場合
        lookAtParam = xlWhole
    
    End If
    
    
    'デフォルトで`見つからなかった`を設定
    sheetWasFound = False
    cellWasFound = False
    
    'シート網羅ループ
    For Each sht In Worksheets
        
        If sht.Name = inThisSheetName Then '指定シートが見つかった場合
        
            sheetWasFound = True
        
            '完全一致で検索
            Set foundobj = sht.UsedRange.Find( _
                What:=findThis, _
                LookAt:=lookAtParam _
            )
            
            If Not (foundobj Is Nothing) Then '見つかった場合
                cellWasFound = True
            
            End If
            
            Exit For 'break
        
        End If
        
    Next sht
    
    If Not sheetWasFound Then 'シートが見つからなかった場合
        ret = CVErr(xlErrNA) '#N/Aを返す
    
    Else
        ret = cellWasFound 'セルが見つからなかったかどうかを返す
    
    End If
    
    isDefinedInThisSheet = ret
    
End Function
