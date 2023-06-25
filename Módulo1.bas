Attribute VB_Name = "Módulo1"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub apisofbot()

    Dim objRequest, json, lstDespesas As Object
    Dim token, strUrl, strResponse, strVars As String
    Dim key, category_api As Variant
    Dim row As Long
    
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    
    token = Range("I12")
    category_api = Cells(1, 1).Value
    row = 2


    Dim ArrayValues As ArrayList
    Set ArrayValues = New ArrayList
    Dim qtd_preenchido As Integer
    qtd_preenchido = 0
    
    If category_api = "despesas" Then
        
        qtd = Range("A1:A15").Rows.Count
        
    ElseIf category_api = "unidades" Then
        
        qtd = Range("A1:A4").Rows.Count
        
    ElseIf category_api = "orgaos" Then
    
        qtd = Range("A1:A5").Rows.Count
        
    ElseIf category_api = "liquidacoes" Then
    
        qtd = Range("A1:A4").Rows.Count
    
    End If

    For i = 2 To qtd

        cel_1 = Cells(i, 1)
        cel_2 = Cells(i, 2)
        
        If cel_2 <> "" Then
            ArrayValues.Add cel_1 & "=" & cel_2
            strVars = strVars & ArrayValues(qtd_preenchido) & "&"
            qtd_preenchido = qtd_preenchido + 1
            
            
        End If
        
        Next i
        
    Debug.Print strVars
    'Debug.Print ArrayValues(1)
    'Debug.Print qtd_preenchido
    

    strUrl = "https://gatewayapi.prodam.sp.gov.br:443/financas/orcamento/sof/v3.0.1/" & category_api & "?" & strVars
        
    Debug.Print strUrl
    
    With objRequest
        .Open "GET", strUrl, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Authorization", "Bearer " & token
        .Send
        
        If .Status <> 200 Then
            MsgBox "Erro: " & .ResponseText
            Exit Sub
        End If
        
        While .readyState <> 4
            DoEvents
        Wend
        
        strResponse = .ResponseText
        Set json = JsonConverter.ParseJson(strResponse)
        Set lstDespesas = json("lstDespesas")
        
        For Each key In lstDespesas
            Cells(row, 4) = key
            Cells(row, 5) = lstDespesas(key)
            row = row + 1
            
            'Debug.Print lstDespesas(key)
        Next key
        
    End With
    
    'Ajustar formatação
    Range("D2:E17").Select
    ActiveWindow.SmallScroll Down:=-6
    Selection.NumberFormat = "General"
    
    
End Sub

