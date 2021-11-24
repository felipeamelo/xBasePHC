**//+-----------------------------------------------------------------------------------//
**//|Funcao....: code128()
**//|Autor.....: Felipe A. Melo
**//|Data......: 14/05/2020
**//|Descricao.: Função que codifica valor para formato EAN128 para ser interpretado
**//|Observação: 
**//+-----------------------------------------------------------------------------------//
Function code128(chaine)
Local i,lVldOk

Private checksum
Private mini
Private dummy
Private tableB  = .T.
Private code128 = ""

lVldOk = .T.

If Len(chaine) > 0
   **//Check for valid characters
   For i = 1 To Len(chaine)
      cVarTmp = SubStr(chaine,i,1)

      Do Case
      	Case Asc(cVarTmp) == 32
      		lVldOk = .F.
      		Exit
      	Case Asc(cVarTmp) >= Asc(Chr(126)) .And. Asc(cVarTmp) <= Asc(Chr(203))
      		lVldOk = .F.
       		Exit
      EndCase
   Next i

   tableB = .T.
   code128 = ""

   If lVldOk
      ** //i  devient l index sur la chaine / i  become the string index
      i = 1 
      Do While i  <= Len(chaine)
         If tableB
            **//See if interesting to switch to table C
            **//yes for 4 digits at start or end, else if 6 digits
            mini = IIf(i = 1 Or i + 3 = Len(chaine), 4, 6)
            
            **//Teste valor mini
            fTestNum(chaine)

            **//Choice of table C
            If mini < 0  
               If i = 1  
                  **//Starting with table C
                  code128 = Chr(205)
               Else 
                  **//Switch to table C
                  code128 = code128 + Chr(199)
               EndIf
               tableB = .F.
            Else
               If i = 1 
                  **//Starting with table B
                  code128 = Chr(204) 
               EndIf
            EndIf
         EndIf
         
         If !tableB 
            **//We are on table C, try to process 2 digits
            mini = 2
            
            **//Teste valor mini
            fTestNum(chaine)

            If mini < 0 
               **//OK for 2 digits, process it
               dummy = Val(Mid (chaine , i , 2))
               dummy = IIf(dummy < 95, dummy + 32, dummy + 100)
               code128 = code128 + Chr(dummy)
               i = i + 2
            Else 
               **//We haven t 2 digits, switch to table B
               code128 = code128 + Chr(200)
               tableB = .T.
            EndIf
         EndIf
        
         If tableB
            **//Process 1 digit with table B
            code128 = code128 + SubStr(chaine,i,1)
            i = i + 1
         EndIf
      EndDo

      **//Calculation of the checksum
      For i = 1 To Len(code128)
          dummy = Asc(SubStr(code128,i,1))
          dummy = IIf(dummy < 127, dummy - 32, dummy - 100)
          If i = 1
             checksum = dummy 
          EndIf
          checksum =(checksum + (i - 1) * dummy ) % 103
      Next i
      
      **//Calculation of the checksum ASCII code
      checksum = IIf(checksum < 95, checksum + 32, checksum + 100)
      
      **//Add the checksum and the STOP
      code128 = code128 + Chr(checksum) + Chr(206)
   EndIf
EndIf

Return(code128)

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fTestNum()
**//|Autor.....: Felipe A. Melo
**//|Data......: 14/05/2020
**//|Descricao.: 
**//|Observação: 
**//+-----------------------------------------------------------------------------------//
Function fTestNum(chaine)
**//if the mini% characters from i% are numeric, then mini%=0
Local i 
i = 0
mini = mini - 1
If i + mini <= Len(chaine) Then
   Do While mini >= 0
      If Asc(SubStr(chaine, i + mini, 1)) < 48 Or Asc(SubStr(chaine, i + mini, 1)) > 57
         Exit
      EndIf
      mini = mini - 1
   EndDo
EndIf

Return
