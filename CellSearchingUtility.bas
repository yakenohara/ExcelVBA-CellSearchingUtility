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
    Dim rangeBrokenSheetName As Variant '検索対象シートを指定する為の引数
    Dim searchFromThisSheet As Variant  '検索対象シート
    Dim lookAtParam As Variant 'Range.Find method の LookAt parameter 用設定値
    
    '検索対象シート `inThisSheet` の指定が Range Object の場合は、
    'inThisSheet.value でシートを検索する
    If (TypeName(inThisSheet)) = "Range" Then 'セル範囲指定の場合(1つだけのセル選択の場合もここで処理する)
        rangeBrokenSheetName = inThisSheet.Item(1).Value '1つめのセル内の値
    Else
        If IsObject(inThisSheet) Then
            Set rangeBrokenSheetName = inThisSheet
        Else
            rangeBrokenSheetName = inThisSheet
        End If
    End If
    
    
    '検索対象シートの設定
    Select Case (TypeName(rangeBrokenSheetName))
        
        Case "String" 'シート名指定の場合
        
            Set searchFromThisSheet = getSheetObjFromString(rangeBrokenSheetName)
            
            If searchFromThisSheet Is Nothing Then
                ret = CVErr(xlErrNum) '#NUM! を返す
            End If
        
        Case "Byte", "Integer", "Long", "Single", "Double" 'Index No(1 based) 指定の場合
        
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
        
        If (foundobj Is Nothing) Then '見つからなかった場合
            ret = CVErr(xlErrNA) '#N/Aを返却
        Else '見つかった場合
            Set ret = foundobj '見つかったCellの Range Objectを返却
            
        End If
        
    End If
    
    If IsObject(ret) Then
        Set matchedCellInSheet = ret
    Else
        matchedCellInSheet = ret
    End If
    
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
