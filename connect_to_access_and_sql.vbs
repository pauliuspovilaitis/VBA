Public Sub intitialise()
    Dim dbfullname As String
    dbfullname = ThisWorkbook.Worksheets("settings").Range("B2").Value

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" _
       & dbfullname & ";"
    
    Set rs = CreateObject("ADODB.Recordset")
    Set by_call_rs = CreateObject("ADODB.Recordset")
End Sub

 
Sub Retrieve(from_time As Date, till_time As Date, from_date As Date, till_date As Date)

    Call intitialise
    
     Dim from_string As String
     Dim till_string As String
     
     from_string = Format(from_date, "YYYY-MM-DD") & " " & Format(from_time, "hh:mm:ss")
     till_string = Format(till_date, "YYYY-MM-DD") & " " & Format(till_time, "hh:mm:ss")

     Dim per_CALL_results_i As Long
     per_CALL_results_i = 1
     
      
    '---------------------------------MAIN QUERY-----------------~~-
     Dim by_call_query As String
     by_call_query = "SELECT * FROM DRL_CallData WHERE (StartSystem BETWEEN #" & from_string & "# And #" & till_string & "#) "

     
     With by_call_rs
           Set .ActiveConnection = cn
               .Source = by_call_query
               .Open , , 3, 3
     End With
    
     Dim total_ANS_CALL_count As Long
     total_ANS_CALL_count = by_call_rs.RecordCount
     
     Dim calls_offered As Long
     Dim calls_handled As Long
     Dim avg_talk_time As Double
     Dim avg_acw_time As Double
     Dim sl As Double
       
     calls_offered = 0
     calls_handled = 0
     avg_talk_time = 0
     avg_acw_time = 0
     sl = 0
   
     Dim acd_1 As Long
     Dim acd_2 As Long
     Dim acd_3 As Long
     Dim acd_4 As Long
     
     acd_1 = 0
     acd_2 = 0
     acd_3 = 0
     acd_4 = 0
   
   
        Do While Not by_call_rs.EOF
        
        On Error Resume Next
        
            'calls offered
            calls_offered = calls_offered + 1
            
            'calls_handled
            If by_call_rs.Fields("Status") = "Answered" Then
              calls_handled = calls_handled + 1
            End If
            
            'sum of talk time
          
            avg_talk_time = avg_talk_time + Val(by_call_rs.Fields("TalkTime"))

            
            'sum of acw time
            avg_acw_time = avg_acw_time + Val(by_call_rs.Fields("AfterCallWorkTime"))
 
            'acd time 1 [0;10]
            If by_call_rs.Fields("Status") = "Answered" Then
                If Second(by_call_rs.Fields("StartofWaiting") - by_call_rs.Fields("StartofHandling")) <= 10 Then
                  acd_1 = acd_1 + 1
                End If
            End If
            
            'acd time 2 (10;20]
            If by_call_rs.Fields("Status") = "Answered" Then
                If Second(by_call_rs.Fields("StartofWaiting") - by_call_rs.Fields("StartofHandling")) > 10 And Second(by_call_rs.Fields("StartofWaiting") - by_call_rs.Fields("StartofHandling")) <= 20 Then
                  acd_2 = acd_2 + 1
                End If
            End If
            
           'acd time 3 (20;30]
            If by_call_rs.Fields("Status") = "Answered" Then
                If Second(by_call_rs.Fields("StartofWaiting") - by_call_rs.Fields("StartofHandling")) > 20 And Second(by_call_rs.Fields("StartofWaiting") - by_call_rs.Fields("StartofHandling")) <= 30 Then
                  acd_3 = acd_3 + 1
                End If
            End If
            
            'acd time 4 (30;40]
            If by_call_rs.Fields("Status") = "Answered" Then
                If Second(by_call_rs.Fields("StartofWaiting") - by_call_rs.Fields("StartofHandling")) > 30 And Second(by_call_rs.Fields("StartofWaiting") - by_call_rs.Fields("StartofHandling")) <= 40 Then
                  acd_4 = acd_4 + 1
                End If
            End If
                               
             by_call_rs.movenext
             per_CALL_results_i = per_CALL_results_i + 1
        Loop
           
    by_call_rs.Close
                   
 'projektavimas
 
    Dim per_row As Long
    per_row = 2
    
    Do While ThisWorkbook.Worksheets("data").Cells(per_row, 1).Value <> ""
        per_row = per_row + 1
    Loop
                   
    With ThisWorkbook.Worksheets("data")
    
        .Cells(per_row, 1).Value = from_date
        .Cells(per_row, 1).NumberFormat = "mm/dd/yy"
        
        
        .Cells(per_row, 2).Value = from_time + "12:30:00 AM"
        .Cells(per_row, 2).NumberFormat = "hh:mm"
        
        
        .Cells(per_row, 3).Value = "TCSDATA"
        .Cells(per_row, 4).Value = "identity_string"
        .Cells(per_row, 5).Value = "acd_Group4"
        
        .Cells(per_row, 6).Value = calls_offered
        .Cells(per_row, 7).Value = calls_handled
        
        
         .Cells(per_row, 6).NumberFormat = "@"
         .Cells(per_row, 6).Value = Right(CStr("0000000000") & CStr(.Cells(per_row, 6).Value), 8)

         .Cells(per_row, 7).NumberFormat = "@"
         .Cells(per_row, 7).Value = Right(CStr("0000000000") & CStr(.Cells(per_row, 7).Value), 8)

        
        
        If avg_talk_time <> 0 Then
          .Cells(per_row, 8).Value = avg_talk_time / calls_handled
        Else
            .Cells(per_row, 8).Value = 0
        End If
        
        
        If avg_acw_time <> 0 Then
          .Cells(per_row, 9).Value = avg_acw_time / calls_handled
        Else
         .Cells(per_row, 9).Value = 0
        End If
        
        
        
        Dim avg_acd As Long
        Dim avg_acw As Long
        
        If calls_handled <> 0 Then
            avg_acd = avg_talk_time / calls_handled
            Else
            avg_acd = 0
        End If
        
        If calls_handled <> 0 Then
            avg_acw = avg_acw_time / calls_handled
            Else
            avg_acw = 0
        End If
        
        'ASA
        
        .Cells(per_row, 10).Value = avg_acd + avg_acw
      
        'SL
        If calls_handled <> 0 Then
           .Cells(per_row, 11).Value = (acd_1 + acd_2) / calls_handled * 100
        Else
           .Cells(per_row, 11).Value = 0
        End If
                    
        'avg pos staff default
        
        .Cells(per_row, 12).NumberFormat = "@"
        .Cells(per_row, 12).Value = "0000000.00"
        
        .Cells(per_row, 8).Value = Round(.Cells(per_row, 8).Value, 2)
        .Cells(per_row, 9).Value = Round(.Cells(per_row, 9).Value, 2)
        .Cells(per_row, 10).Value = Round(.Cells(per_row, 10).Value, 1)
        .Cells(per_row, 11).Value = Round(.Cells(per_row, 11).Value, 2)
     
    End With
                   
End Sub
