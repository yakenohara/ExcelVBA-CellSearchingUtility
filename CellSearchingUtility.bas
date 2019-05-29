Attribute VB_Name = "CellSearchingUtility"
'
' 指定シートに一致するセルが存在するかどうかを返す
'
' findThis: 検索キーワード
' inThisSheetName: 検索シート名
' lookAtPart: True で部分一致指定, false で完全一致指定
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

'
'指定文字列を検索してヒットしたセルの位置を返す
'
' findThis: 検索キーワード
' searchRange: 検索対象範囲
' getRow: True ヒットしたセルの位置の行数指定, Falseで列指定
' lookAtPart: True で部分一致指定, false で完全一致指定
'
Public Function findCellAndGetPosition(ByVal findThis As Variant, ByVal searchRange As Range, ByVal getRow As Variant, Optional ByVal lookAtPart As Variant = True) As Variant

    Dim ret As Variant
    Dim lookAtParam As Variant 'Range.Find method の LookAt parameter 用設定値
    
    'Range.Find method の LookAt parameter 用設定値の決定
    If lookAtPart Then '完全一致指定の場合
        lookAtParam = xlPart
    
    Else '完全一致指定でない場合
        lookAtParam = xlWhole
    
    End If
    
    '検索実行
    Set searchResult = searchRange.Find( _
        What:=findThis, _
        LookAt:=lookAtParam _
    )
    'xlPart:部分一致有効
    'After:先頭から開始するように、最終セルを指定

    If Not searchResult Is Nothing Then '見つかったとき
        
        If getRow Then '行位置取得指定のとき
            ret = searchResult.Row
            
        Else '列位置取得指定のとき
            ret = searchResult.Column
        
        End If
        
        
    Else '見つからなかった時
        ret = CVErr(xlErrNA) '#N/Aを返却
    
    End If
    
    findCellAndGetPosition = ret
    
End Function

