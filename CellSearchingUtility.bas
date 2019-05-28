Attribute VB_Name = "Module1"
'
' 指定シートに一致するセルが存在するかどうかを返す
'
' findThis: 検索するセル
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
