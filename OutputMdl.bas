Attribute VB_Name = "OutputMdl"

'Sub ParseXML()
'
'    Dim xmlDoc As MSXML2.DOMDocument60
'    Dim xmlNode As MSXML2.IXMLDOMNode
'
'    Dim coolingString As String
'    Dim heatingString As String
'
'    Set xmlDoc = New MSXML2.DOMDocument60
'
'    xmlDoc.async = False
'    xmlDoc.Load (ThisWorkbook.Path & OUTPUT_PATH & "\TestCase3Table.xml")
'
'
'    Set coolingSet = xmlDoc.getElementsByTagName("Zonecoolingsummarymonthly")
'    Set heatingSet = xmlDoc.getElementsByTagName("Zoneheatingsummarymonthly")
'
'        'Zone1~7까지(나중에 Zone을 하나로 통일해야함)
'        For i = 1 To 7
'            For Each NodeLevel1 In coolingSet
'                For Each NodeLevel2 In NodeLevel1.getElementsByTagName("for")
'                    If NodeLevel2.Text = "1FLOOR:ZONE" & i Then
'                        'Debug.Print NodeLevel1.getElementsByTagName("CustomMonthlyReport")(0).Text
'                        coolingString = NodeLevel1.getElementsByTagName("CustomMonthlyReport")(0).Text
'                        'For Each NodeLevel3 In NodeLevel1.getElementsByTagName("CustomMonthlyReport")
'                            'Debug.Print NodeLevel3.Text
'                            'NodeLevel3.Text
'                        'Next
'                    End If
'                Next
'            Next
'
'            For Each NodeLevel1 In heatingSet
'                For Each NodeLevel2 In NodeLevel1.getElementsByTagName("for")
'                    If NodeLevel2.Text = "1FLOOR:ZONE" & i Then
'                        heatingString = NodeLevel1.getElementsByTagName("CustomMonthlyReport")(0).Text
'                    End If
'                Next
'            Next
'
'            Call SplitText(coolingString, heatingString)
'
'        Next
'
'        For i = 1 To 12
'            Range("ElecConsumption").Cells(1, i) = montlyCoolingEnergy(i)
'            Range("GasConsumption").Cells(1, i) = montlyHeatingEnergy(i)
'        Next
'
'End Sub

Sub ParseXML()

    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode

    Dim coolingString As String
    Dim heatingString As String
    Dim lighthingElecString As String
    Dim equipElecString As String

    Set xmlDoc = New MSXML2.DOMDocument60

    xmlDoc.async = False
    xmlDoc.Load (ThisWorkbook.Path & OUTPUT_PATH & "\TestCase3Table.xml")


    Set heatingSet = xmlDoc.getElementsByTagName("Zoneheatingsummarymonthly") 'heating energy
    Set coolingSet = xmlDoc.getElementsByTagName("Zonecoolingsummarymonthly") 'cooling energy
    Set elecSet = xmlDoc.getElementsByTagName("Zoneelectricsummarymonthly") 'equipment electricity energy

        For Each NodeLevel1 In coolingSet
            For Each NodeLevel2 In NodeLevel1.getElementsByTagName("for")
                If NodeLevel2.Text = "1FLOOR:ZONE" & 7 Then
                    'Debug.Print NodeLevel1.getElementsByTagName("CustomMonthlyReport")(0).Text
                    coolingString = NodeLevel1.getElementsByTagName("CustomMonthlyReport")(0).Text
                    'For Each NodeLevel3 In NodeLevel1.getElementsByTagName("CustomMonthlyReport")
                        'Debug.Print NodeLevel3.Text
                        'NodeLevel3.Text
                    'Next
                End If
            Next
        Next

        For Each NodeLevel1 In heatingSet
            For Each NodeLevel2 In NodeLevel1.getElementsByTagName("for")
                If NodeLevel2.Text = "1FLOOR:ZONE" & 7 Then
                    heatingString = NodeLevel1.getElementsByTagName("CustomMonthlyReport")(0).Text
                End If
            Next
        Next
        
        For Each NodeLevel1 In elecSet
            For Each NodeLevel2 In NodeLevel1.getElementsByTagName("for")
                If NodeLevel2.Text = "1FLOOR:ZONE" & 7 Then
                    lighthingElecString = NodeLevel1.getElementsByTagName("CustomMonthlyReport")(0).Text
                End If
            Next
        Next
        
        For Each NodeLevel1 In elecSet
            For Each NodeLevel2 In NodeLevel1.getElementsByTagName("for")
                If NodeLevel2.Text = "1FLOOR:ZONE" & 7 Then
                    equipElecString = NodeLevel1.getElementsByTagName("CustomMonthlyReport")(3).Text
                End If
            Next
        Next

        Call SplitText(coolingString, heatingString, lighthingElecString, equipElecString)

        For i = 1 To 12
            Range("ElecConsumption").Cells(1, i) = montlyElecConsumption(i)
            Range("GasConsumption").Cells(1, i) = montlyGasConsumption(i)
        Next

End Sub


Sub SplitText(coolingString As String, heatingString As String, lighthingElecString As String, equipElecString As String)
    
    Dim coolingTemp() As String
    Dim heatingTemp() As String
    Dim lighthingTemp() As String
    Dim equipTemp() As String
        
        coolingTemp = Split(coolingString, " ")
        heatingTemp = Split(heatingString, " ")
        lighthingTemp = Split(lighthingElecString, " ")
        equipTemp = Split(equipElecString, " ")
        
        For i = 1 To 12
            montlyElecConsumption(i) = (montlyElecConsumption(i) + CDbl(coolingTemp(i)) + CDbl(lighthingTemp(i)) + CDbl(equipTemp(i))) / 3600000 'J to kWh
            montlyGasConsumption(i) = (montlyGasConsumption(i) + CDbl(heatingTemp(i))) / 40000000 'J to m3
        Next
        
End Sub

Sub OutputResult(Generation As Integer)

        Range("Result_No").Cells(Generation + 1, 1) = Generation

        For i = 1 To 12
        
            Range("Result_Elec").Cells(Generation + 1, i + 1) = montlyElecConsumption(i)
            Range("Result_Gas").Cells(Generation + 1, i + 1) = montlyGasConsumption(i)
            
        Next
        
End Sub

Sub OutputOptResult()

End Sub
