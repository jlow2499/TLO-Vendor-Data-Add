Private Sub DelayMs(ms As Long)
    Debug.Print TimeValue(Now)
    Application.Wait (Now + (ms * 0.00000001))
    Debug.Print TimeValue(Now)
End Sub


Sub CLEAR()
Set aRange = Sheets("TLO BOT").Range("A5.ZZ50000")
aRange.ClearContents
End Sub


Sub TLO_Add()


Dim CurrentHost As Object
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession
Dim irow As Long
irow = 5
  file = Range("A" & irow).Value

  
  Do
  If Range("A" & irow).Value = "" Then
    Application.StatusBar = "Credit AR Add Complete"
    MsgBox "Add Complete"
    Exit Sub
    End If
  

        If CurrentHost.GetText(0, 22, 52) = "ENTER SELECTION (.,FILE#,/,STATUS,-nnnnn,Tn,/R,HELP)" Then
        CurrentHost.Output Range("A" & irow).Value & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("A" & irow).Value & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "26" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "26" & ChrW$(13)
        End If
        
    If Range("EN" & irow) = "" Then
        
        Call DelayMs(100)
        
        Else
        
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 22) = "SELECTION (/A,/T,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        End If
        
        Call DelayMs(200)

        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("DW" & irow).Value & " " & Range("DY" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("DW" & irow).Value & " " & Range("DY" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
        If Range("EB" & irow).Value = "" Then
        
            Call DelayMs(100)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "/11" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/11" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
                CurrentHost.Output "335" & ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
            Else
                Call DelayMs(600)
                CurrentHost.Output "335" & ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Left(Range("EN" & irow).Value, 3) & "-" & Mid(Range("EN" & irow).Value, 4, 3) & "-" & Right(Range("EN" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Left(Range("EN" & irow).Value, 3) & "-" & Mid(Range("EN" & irow).Value, 4, 3) & "-" & Right(Range("EN" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EO" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EO" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "//" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "//" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output ChrW$(13)
            End If
            
        Else
        
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EB" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EB" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "/8" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/8" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EC" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EC" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("ED" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("ED" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
                If Len(Range("EE" & irow).Value) = 3 Then
                CurrentHost.Output "00" & Range("EE" & irow).Value
                ElseIf Len(Range("EE" & irow).Value) = 4 Then
                CurrentHost.Output "0" & Range("EE" & irow).Value
                Else
                CurrentHost.Output Range("EE" & irow).Value
                End If
                Call DelayMs(200)
                CurrentHost.Output ChrW$(13)
            Else
                Call DelayMs(600)
                If Len(Range("EE" & irow).Value) = 3 Then
                CurrentHost.Output "00" & Range("EE" & irow).Value
                ElseIf Len(Range("EE" & irow).Value) = 4 Then
                CurrentHost.Output "0" & Range("EE" & irow).Value
                Else
                CurrentHost.Output Range("EE" & irow).Value
                End If
                Call DelayMs(200)
                CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
               
            If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
            CurrentHost.Output "335" & ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "335" & ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Left(Range("EN" & irow).Value, 3) & "-" & Mid(Range("EN" & irow).Value, 4, 3) & "-" & Right(Range("EN" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Left(Range("EN" & irow).Value, 3) & "-" & Mid(Range("EN" & irow).Value, 4, 3) & "-" & Right(Range("EN" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EO" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EO" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "//" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "//" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output ChrW$(13)
            End If
            
        End If
        
        
                
    End If
    
    
    
    ' part two
    
    If Range("EP" & irow).Value = "" Then
        
        Call DelayMs(100)
        
        Else
               
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 22) = "SELECTION (/A,/T,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        End If
        
        Call DelayMs(200)

        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("DW" & irow).Value & " " & Range("DY" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("DW" & irow).Value & " " & Range("DY" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
        If Range("EH" & irow).Value = "" Then
        
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "/11" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/11" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
                CurrentHost.Output "335" & ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
            Else
                Call DelayMs(600)
                CurrentHost.Output "335" & ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Left(Range("EP" & irow).Value, 3) & "-" & Mid(Range("EP" & irow).Value, 4, 3) & "-" & Right(Range("EP" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Left(Range("EP" & irow).Value, 3) & "-" & Mid(Range("EP" & irow).Value, 4, 3) & "-" & Right(Range("EP" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EQ" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EQ" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "//" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "//" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output ChrW$(13)
            End If
            
        Else
        
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EH" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EH" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "/8" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/8" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EI" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EI" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EJ" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EJ" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
                If Len(Range("EK" & irow).Value) = 3 Then
                CurrentHost.Output "00" & Range("EK" & irow).Value
                ElseIf Len(Range("EK" & irow).Value) = 4 Then
                CurrentHost.Output "0" & Range("EK" & irow).Value
                Else
                CurrentHost.Output Range("EK" & irow).Value
                End If
                Call DelayMs(200)
                CurrentHost.Output ChrW$(13)
            Else
                Call DelayMs(600)
                If Len(Range("EK" & irow).Value) = 3 Then
                CurrentHost.Output "00" & Range("EK" & irow).Value
                ElseIf Len(Range("EK" & irow).Value) = 4 Then
                CurrentHost.Output "0" & Range("EK" & irow).Value
                Else
                CurrentHost.Output Range("EK" & irow).Value
                End If
                Call DelayMs(200)
                CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
               
            If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
            CurrentHost.Output "335" & ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "335" & ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Left(Range("EP" & irow).Value, 3) & "-" & Mid(Range("EP" & irow).Value, 4, 3) & "-" & Right(Range("EP" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Left(Range("EP" & irow).Value, 3) & "-" & Mid(Range("EP" & irow).Value, 4, 3) & "-" & Right(Range("EP" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("EQ" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("EQ" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "//" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "//" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output ChrW$(13)
            End If
            
        End If
      
    End If
    
    ' phone 3, yes i am a sloppy programmer, but it works and is done quickly
    
    If Range("ER" & irow).Value = "" Then
        Call DelayMs(100)
    Else
    
        If CurrentHost.GetText(0, 22, 22) = "SELECTION (/A,/T,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        End If
        
        Call DelayMs(200)

        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("DW" & irow).Value & " " & Range("DY" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("DW" & irow).Value & " " & Range("DY" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "/11" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/11" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
                CurrentHost.Output "335" & ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
            Else
                Call DelayMs(600)
                CurrentHost.Output "335" & ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
                Call DelayMs(100)
                CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Left(Range("ER" & irow).Value, 3) & "-" & Mid(Range("ER" & irow).Value, 4, 3) & "-" & Right(Range("ER" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Left(Range("ER" & irow).Value, 3) & "-" & Mid(Range("ER" & irow).Value, 4, 3) & "-" & Right(Range("ER" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Range("ES" & irow).Value & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Range("ES" & irow).Value & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "//" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "//" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output ChrW$(13)
            End If

    End If
    
   'POE round 1!!!
    
    If Range("ET" & irow).Value = "" Then
        Call DelayMs(100)
    Else
    
        If CurrentHost.GetText(0, 22, 22) = "SELECTION (/A,/T,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        End If
        
        Call DelayMs(200)

        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("ET" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("ET" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
        Call DelayMs(200)

        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("EU" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("EU" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("EV" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("EV" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("EW" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("EW" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
                If Len(Range("EX" & irow).Value) = 3 Then
                CurrentHost.Output "00" & Range("EX" & irow).Value
                ElseIf Len(Range("EX" & irow).Value) = 4 Then
                CurrentHost.Output "0" & Range("EX" & irow).Value
                Else
                CurrentHost.Output Range("EX" & irow).Value
                End If
                Call DelayMs(200)
                CurrentHost.Output ChrW$(13)
            Else
                Call DelayMs(600)
                If Len(Range("EX" & irow).Value) = 3 Then
                CurrentHost.Output "00" & Range("EX" & irow).Value
                ElseIf Len(Range("EX" & irow).Value) = 4 Then
                CurrentHost.Output "0" & Range("EX" & irow).Value
                Else
                CurrentHost.Output Range("EX" & irow).Value
                End If
                Call DelayMs(200)
                CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
               
            If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
            CurrentHost.Output "POE" & ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "POE" & ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
         If Range("EY" & irow).Value = "" Then
            Call DelayMs(100)
         Else
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Left(Range("EY" & irow).Value, 3) & "-" & Mid(Range("EY" & irow).Value, 4, 3) & "-" & Right(Range("EY" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Left(Range("EY" & irow).Value, 3) & "-" & Mid(Range("EY" & irow).Value, 4, 3) & "-" & Right(Range("EY" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
         End If
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "//" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "//" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output ChrW$(13)
            End If
    
    
    End If
    
    ' POE add round 2... FIGHT!!!
    
    If Range("FB" & irow).Value = "" Then
        Call DelayMs(100)
    Else
    
        If CurrentHost.GetText(0, 22, 22) = "SELECTION (/A,/T,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        Call DelayMs(200)
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output "TLO" & ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        End If
        
        Call DelayMs(200)

        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("FB" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("FB" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
        Call DelayMs(200)

        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("FC" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("FC" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("FD" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("FD" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output Range("FE" & irow).Value
        Call DelayMs(100)
        CurrentHost.Output ChrW$(13)
        Else
        Call DelayMs(600)
        CurrentHost.Output Range("FE" & irow).Value
        Call DelayMs(200)
        CurrentHost.Output ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
                If Len(Range("FF" & irow).Value) = 3 Then
                CurrentHost.Output "00" & Range("FF" & irow).Value
                ElseIf Len(Range("FF" & irow).Value) = 4 Then
                CurrentHost.Output "0" & Range("FF" & irow).Value
                Else
                CurrentHost.Output Range("FF" & irow).Value
                End If
                Call DelayMs(200)
                CurrentHost.Output ChrW$(13)
            Else
                Call DelayMs(600)
                If Len(Range("FF" & irow).Value) = 3 Then
                CurrentHost.Output "00" & Range("FF" & irow).Value
                ElseIf Len(Range("FF" & irow).Value) = 4 Then
                CurrentHost.Output "0" & Range("FF" & irow).Value
                Else
                CurrentHost.Output Range("FF" & irow).Value
                End If
                Call DelayMs(200)
                CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
               
            If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
            CurrentHost.Output "POE" & ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "POE" & ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            Call DelayMs(100)
            CurrentHost.Output ChrW$(13)
            
            End If
            
            Call DelayMs(200)
            
         If Range("FG" & irow).Value = "" Then
            Call DelayMs(100)
         Else
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output Left(Range("FG" & irow).Value, 3) & "-" & Mid(Range("FG" & irow).Value, 4, 3) & "-" & Right(Range("FG" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output Left(Range("FG" & irow).Value, 3) & "-" & Mid(Range("FG" & irow).Value, 4, 3) & "-" & Right(Range("FG" & irow).Value, 4)
            Call DelayMs(400)
            CurrentHost.Output ChrW$(13)
            End If
            
            Call DelayMs(200)
            
         End If
            
            If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
            CurrentHost.Output "//" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "//" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
            
            Call DelayMs(200)
            
            If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
            CurrentHost.Output ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output ChrW$(13)
            End If
    
    End If
    
        If CurrentHost.GetText(0, 22, 24) = "SELECTION (n,/A,/T,/H,/)" Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
        
        If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
            CurrentHost.Output "/" & ChrW$(13)
            Else
            Call DelayMs(600)
            CurrentHost.Output "/" & ChrW$(13)
            End If
    
    irow = irow + 1
    Loop

End Sub
