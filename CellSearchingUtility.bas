Attribute VB_Name = "CellSearchingUtility"
'<License>------------------------------------------------------------
'
' Copyright (c) 2019 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
' 指定範囲内を検索して最初に見つかったセルを返す
'
' ## Parameters
'
'  - keyWord
'     セル検索キーワード
'
'  - fromThisRange
'     検索対象範囲
'
'  - lookAtPart (Optional. TRUE as default)
'     Cell検索キーワード `keyWord` を部分一致で検索する場合に `TRUE`
'     完全一致で検索する場合は `FALSE` を指定する
'
' ## Returns
'
'  最初に見つかったセルの Range Object
'  セルが見つからなかった場合は `#N/A` を返却する
'
Public Function matchedCellInRange(ByVal keyWord As Variant, ByVal fromThisRange As Range, Optional ByVal lookAtPart As Boolean = True) As Variant

    Dim ret As Variant '返却値
    Dim lookAtParam As Variant 'Range.Find method の LookAt parameter 用設定値
    Dim rangeBrokenKeyword As Variant
    
    'Range.Find method の LookAt parameter 用設定値の決定
    If lookAtPart Then '部分一致指定の場合
        lookAtParam = xlPart
    
    Else '完全一致指定の場合
        lookAtParam = xlWhole
    End If
    
    
    'note
    '
    ' 日付型もしくは通貨型の書式設定をしたセルの Range Objectの .Value が返す値の型は、
    ' セルに設定されている値によって、以下のように複雑に変化する。
    ' 動作予測しにくいので、Boolean/Double/String/Error/Empty 型しか返さない .Value2 で取得する。
    '
    ' 日付型
    '   -> Date 型 か、String 型(1900年1月1日~9999年12月31日(※1)の範囲外の日付)
    ' 通貨型
    '   -> Currency 型 か、
    '      .Value にアクセスしただけで Exception(
    '        セルの値に -922,337,203,685,477 ～ 922,337,203,685,477(※2) の範囲外の値が設定されていた場合に発生する
    '      )※
    ' で値を取得する。
    '
    ' ※1
    ' https://support.office.com/ja-jp/article/excel-%e3%81%ae%e4%bb%95%e6%a7%98%e3%81%a8%e5%88%b6%e9%99%90-1672b34d-7043-467e-8e27-269d656771c3?ui=ja-JP&rs=ja-JP&ad=JP
    '
    ' ※2
    ' https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/data-type-summary
    '
    'Range.Value2 method 相当の操作で 検索キーワードを取得
    rangeBrokenKeyword = getValue2(keyWord)
    
    If IsError(rangeBrokenKeyword) Then
        'note Range.Find method の What:= に指定すると Exception が発生するので、ハジく
        ret = CVErr(xlErrValue) '#VALUE! を返す
        
    Else
    
        ' Option Parameter settings of Range.Find method
        '
        '| Parameter       | Meaning                                       |
        '| --------------- | --------------------------------------------- |
        '| After           | セル範囲の先頭が検査1発目のセルとなるように、 |
        '|                 | 検索開始位置をセル範囲最後にする              |
        '| LookIn          | 検索対象を数式に指定                          |
        '| LookAt          | 完全一致 / 部分一致 (引数設定による)          |
        '| SearchOrder     | 検索方向を行で指定                            |
        '| SearchDirection | 順方向で検索                                  |
        '| MatchCase       | 大文字と小文字を区別しない                    |
        '| MatchByte       | 半角と全角を区別しない                        |
        '| SearchFormat    | 書式で検索しない                              |
        '
        Set searchResult = fromThisRange.Find( _
            What:=keyWord, _
            After:=fromThisRange.Item(fromThisRange.Count), _
            LookIn:=xlValues, _
            LookAt:=lookAtParam, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            MatchByte:=False, _
            SearchFormat:=False _
        )
        
        If Not searchResult Is Nothing Then '見つかったとき
            Set ret = searchResult
            
        Else '見つからなかった時
            ret = CVErr(xlErrNA) '#N/Aを返却
        
        End If
    
    End If
    
    If IsObject(ret) Then
        Set matchedCellInRange = ret
    Else
        matchedCellInRange = ret
    End If
    
End Function

'
' ThisWorkbook 内の指定シートからセルを検索して、
' 最初に見つかったセルを返す
'
' ## Parameters
'
'  - keyWord
'     セル検索キーワード
'
'  - inThisSheet
'     検索対象シート
'     数値型で指定した場合はシート番号(1 based)
'     文字列型で指定した場合はシート名として扱われる
'
'  - lookAtPart (Optional. TRUE as default)
'     Cell検索キーワード `keyWord` を部分一致で検索する場合に `TRUE`
'     完全一致で検索する場合は `FALSE` を指定する
'
' ## Returns
'
'  最初に見つかったセルの Range Object
'  エラー時は以下を返却する
'
'  - #N/A
'     セルが見つからなかった場合
'
'  - #NUM!
'     検索対象シートが存在しない場合
'
'  - #VALUE!
'     検索対象シートの指定引数 `inThisSheet` に
'     文字列型でも数値型でもない型で値が指定されている
'
Public Function matchedCellInSheet(ByVal keyWord As Variant, ByVal inThisSheet As Variant, Optional ByVal lookAtPart As Boolean = True) As Variant
    
    Dim ret As Variant '返却値
    Dim rangeBrokenKeyword As Variant
    Dim rangeBrokenSheetName As Variant '検索対象シートを指定する為の引数
    Dim searchFromThisSheet As Variant  '検索対象シート
    Dim lookAtParam As Variant 'Range.Find method の LookAt parameter 用設定値
    
    'Range.Value2 method 相当の操作で 検索キーワードを取得
    rangeBrokenKeyword = getValue2(keyWord)
    
    If IsError(rangeBrokenKeyword) Then
        'note Range.Find method の What:= に指定すると Exception が発生するので、ハジく
        ret = CVErr(xlErrValue) '#VALUE! を返す
        
    Else
    
        'Range.Value2 method 相当の操作で シート名を取得
        rangeBrokenSheetName = getValue2(inThisSheet)
        
        '検索対象シートの設定
        Select Case (TypeName(rangeBrokenSheetName))
            
            Case "String" 'シート名指定の場合
            
                Set searchFromThisSheet = getSheetObjFromString(rangeBrokenSheetName)
                
                If searchFromThisSheet Is Nothing Then
                    ret = CVErr(xlErrNum) '#NUM! を返す
                End If
            
            Case "Double" 'Index No(1 based) 指定の場合
            
                'ワークシート数チェック
                If (rangeBrokenSheetName <= ThisWorkbook.Worksheets.Count) Then '存在するワークシート数の範囲内の場合
                    Set searchFromThisSheet = ThisWorkbook.Worksheets(rangeBrokenSheetName)
                
                Else '存在するワークシート数の範囲外の場合
                    ret = CVErr(xlErrNum) '#NUM! を返す
                    
                End If
            
            Case Else '不明型の場合
                ret = CVErr(xlErrValue) '#VALUE! を返す
                
        End Select
        
        If Not (IsError(ret)) Then
            
            'Range.Find method の LookAt parameter 用設定値の決定
            If lookAtPart Then '部分一致指定の場合
                lookAtParam = xlPart
            Else '完全一致指定の場合
                lookAtParam = xlWhole
            End If
            
            ' Option Parameter settings of Range.Find method
            '
            '| Parameter       | Meaning                                       |
            '| --------------- | --------------------------------------------- |
            '| After           | セル範囲の先頭が検査1発目のセルとなるように、 |
            '|                 | 検索開始位置をセル範囲最後にする              |
            '| LookIn          | 検索対象を数式に指定                          |
            '| LookAt          | 完全一致 / 部分一致 (引数設定による)          |
            '| SearchOrder     | 検索方向を行で指定                            |
            '| SearchDirection | 順方向で検索                                  |
            '| MatchCase       | 大文字と小文字を区別しない                    |
            '| MatchByte       | 半角と全角を区別しない                        |
            '| SearchFormat    | 書式で検索しない                              |
            '
            Set fromThisRange = searchFromThisSheet.UsedRange
            Set foundobj = fromThisRange.Find( _
                What:=rangeBrokenKeyword, _
                After:=fromThisRange.Item(fromThisRange.Count), _
                LookIn:=xlValues, _
                LookAt:=lookAtParam, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, _
                MatchCase:=False, _
                MatchByte:=False, _
                SearchFormat:=False _
            )
            
            If (foundobj Is Nothing) Then '見つからなかった場合
                ret = CVErr(xlErrNA) '#N/Aを返却
            Else '見つかった場合
                Set ret = foundobj '見つかったCellの Range Objectを返却
                
            End If
            
        End If
    End If
    
    If IsObject(ret) Then
        Set matchedCellInSheet = ret
    Else
        matchedCellInSheet = ret
    End If
    
End Function

'
' 指定引数が セル参照(Range Object) の場合は .Value2 でセル内の値を、
' そうでない場合は プリミティブ型が指定されたと判断して、
' .Value2 が取りうる値のタイプ
' Double/String/Boolean/Error/Empty
' のいづれかに Cast して返す
'
Private Function getValue2(ByVal variant_unkown As Variant) As Variant

    Dim ret As Variant
    
    If (TypeName(variant_unkown) = "Range") Then
        ret = variant_unkown.Value2
    
    Else
        ret = getValue2FromPrimitive(variant_unkown)
        
    End If
    
    getValue2 = ret
    
End Function

'
' セル内の値が取りうるタイプ
' Double/Currency/Date/String/Boolean/Error/Empty
' が、.Value2 で値を取得することで、タイプが
' Double/String/Boolean/Error/Empty (CurrencyとDate型がDoubleにキャストされる)
' のいづれかに Cast されるように、
' プリミティブな値を格納する変数が取りうるタイプ(Decimal, Long, LongLong等)を
' Double/String/Boolean/Error/Empty
' のいづれかに Cast して返す
'
' キャスト不可能なタイプ(Object等)の場合は #VALUE! を返す
'
Private Function getValue2FromPrimitive(ByVal variant_primitive As Variant) As Variant

    Dim ret As Variant '返却値
    
    On Error GoTo EXCEPTION_CAST 'CDbl() で Exception 時に Go
    
    '
    ' Case statement 内の※
    '  VBA の1データ型(組み込みのデータ型, Intrinsic data type)だが、プリミティブ型とはみなさない
    '
    Select Case TypeName(variant_primitive)
        Case "Boolean"
            ret = variant_primitive 'Boolean のまま格納
        
        Case "Byte"
            ret = CDbl(variant_primitive)
        
        'Case "Collection" ->対応しない※
        
        Case "Currency"
            ret = CDbl(variant_primitive)
        
        Case "Date"
            ret = CDbl(variant_primitive) 'シリアル値を Double として取得
            
        Case "Decimal"
            ret = CDbl(variant_primitive)
        
        'Case "Dictionary" ->対応しない※
        
        Case "Double"
            ret = variant_primitive 'そのまま格納
        
        Case "Integer"
            ret = CDbl(variant_primitive)
        
        Case "Long"
            ret = CDbl(variant_primitive)
        
        Case "LongLong"
            ret = CDbl(variant_primitive)
        
        'Case "LongPtr" ->対応しない※
        
        'Case "Object" ->対応しない※
        
        Case "Single"
            ret = CDbl(variant_primitive)
        
        Case "String" '(可変長文字列、固定長文字列どちらでも)
            ret = variant_primitive 'そのまま格納
            
        Case Else
            
            If _
            ( _
                (IsError(variant_primitive)) Or _
                (IsEmpty(variant_primitive)) _
            ) Then
                'Error かEmptyの場合はそのまま返す
                ret = variant_primitive
            Else
                ret = CVErr(xlErrValue) '#VALUE! を返す
            End If
            
        
    End Select
    
    getValue2FromPrimitive = ret
    Exit Function
    
EXCEPTION_CAST: 'CDbl() で Exception発生

    ret = CVErr(xlErrValue) '#VALUE! を返す
    getValue2FromPrimitive = ret
    Exit Function
    
End Function

'
' ThisWorkbook から Sheet Object を シート名を使って取得する
' シートが存在しない場合は、Nothing を返す
'
Private Function getSheetObjFromString(ByVal sheetName As String) As Variant
    
    Dim ret As Variant
    
    On Error GoTo NOT_FOUND
    Set getSheetObjFromString = ThisWorkbook.Worksheets(sheetName)
    Exit Function
    
NOT_FOUND: ' シートが存在しない場合
    Set getSheetObjFromString = Nothing
    Exit Function
    
End Function
