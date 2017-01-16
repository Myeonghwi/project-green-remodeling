Attribute VB_Name = "ScrapeMdl"
Sub IsTrueRange()

    Dim arrFactor
    
    arrFactor = Array("Repla_Insulation", "Repla_Window", "Repla_Shading", "Repla_Lighting")

        For i = LBound(arrFactor) To UBound(arrFactor)
        
            If Range(arrFactor(i)).Offset(0, 1) = True Then

                ScrapeList

            Else

                'List(1, i) = False

            End If
        Next

End Sub

Sub CountList()
    
    Dim arrFactor
    
    arrFactor = Array("Repla_Insulation", "Repla_Window", "Repla_Shading", "Repla_Lighting")
    
    rowCount = 0
    colCount = 0
    
        For i = LBound(arrFactor) To UBound(arrFactor)
            
            Set rngRow_ = Range(Range(arrFactor(i)).Offset(2, 0), Range(arrFactor(i)).End(xlDown))
            
            For Each rngRow In rngRow_  '각 요소의 열의 합 구하기
                
                rowCount = rowCount + 1
                
            Next

        Next
        
        Set rngCol_ = Range(Range(arrFactor(0)).Offset(1, 0), Range(arrFactor(0)).Offset(1, 0).End(xlToRight))
        
        For Each rngCol In rngCol_  '행의 합 구하기
                    
            colCount = colCount + 1
            
        Next
        
        Range("Repla_rowCount").Value = rowCount
        Range("Repla_colCount").Value = colCount
        
End Sub

Sub ScrapeList()
    
    Dim arrFactor
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    arrFactor = Array("Repla_Insulation", "Repla_Window", "Repla_Shading", "Repla_Lighting")
    
        CountList
        
        ReDim lst(rowCount, colCount)
        
        j = 0
        k = 0
        l = 0
        
        For i = LBound(arrFactor) To UBound(arrFactor)
        
            Set rngRow_ = Range(Range(arrFactor(i)).Offset(2, 0), Range(arrFactor(i)).End(xlDown))
            Set rngCol_ = Range(Range(arrFactor(0)).Offset(1, 0), Range(arrFactor(0)).Offset(1, 0).End(xlToRight))
            
            For Each rngRow In rngRow_
            
                For Each rngCol In rngCol_
            
                    lst(j, k) = Range(arrFactor(i)).Offset(l + 2, k).Value
                    
                    k = k + 1
                Next
                k = 0
                j = j + 1
                l = l + 1
            Next
            l = 0
        Next

End Sub
