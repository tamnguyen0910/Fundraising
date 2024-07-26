Attribute VB_Name = "Module1"
' Fundraising from consolidation to communication


Sub orders_aggregation()

'this sub helps aggregate multiple sheets in different excel files with same structure into a new workbook
'and save this new workbook in the same path of this workbook and name it "Order Aggregation - ddmmyy"
'with ddmmyy the date doing this aggregation
'This helps to save a draf version of raw data befor any manupulation

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fd As FileDialog
    Dim fileName As Variant
    Dim destWb As Workbook
    Dim destSheet As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range
    Dim destRange As Range
    Dim currentDate As String

    ' Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = True

    ' Show the dialog box. If the .Show method returns True, the
    ' user picked at least one file. If the .Show method returns
    ' False, the user clicked Cancel.
    If fd.Show = True Then
        ' Create a new workbook for aggregated data
        currentDate = Format(Date, "ddmmyy")
        Set destWb = Workbooks.Add
        Set destSheet = destWb.Sheets(1)
        destSheet.Name = "Aggregated Data"
        
        ' Initialize the destination row
        lastRow = 1

        ' Loop through each file selected
        For Each fileName In fd.SelectedItems
            ' Open the workbook
            Set wb = Workbooks.Open(fileName)

            ' Loop through all the sheets in the workbook
            For Each ws In wb.Sheets
                ' Copy the data from the source sheet to the destination sheet
                Set sourceRange = ws.UsedRange
                If lastRow = 1 Then
                    ' Copy the entire range including headers
                    Set destRange = destSheet.Cells(lastRow, 1)
                    sourceRange.Copy Destination:=destRange
                    lastRow = lastRow + sourceRange.Rows.Count
                Else
                    ' Copy the range excluding headers
                    Set destRange = destSheet.Cells(lastRow, 1)
                    sourceRange.Offset(1, 0).Resize(sourceRange.Rows.Count - 1).Copy Destination:=destRange
                    lastRow = lastRow + sourceRange.Rows.Count - 1
                End If
            Next ws

            ' Close the workbook
            wb.Close SaveChanges:=False
        Next fileName

        ' Save the new workbook
        destWb.SaveAs ThisWorkbook.Path & "\Order Aggregations - " & currentDate & ".xlsx"
        destWb.Close SaveChanges:=True
    End If

    Set fd = Nothing
End Sub

Sub ImporterColonnes()

'this sub helps import the orders form aggregation file created above to this workbook
'then check if the data is with no error
'then extract usefull information, range it in a specified order of this workbook, so the macro works even when
'sources files changes
'and perform a test on estimated amounts


Dim Fichier, Wbk1 As Workbook, Wbk2 As Workbook
Dim Colonnes(), col As Integer, Resultat As Variant
Dim NbLignes As Integer, Lig As Integer
Dim F As Variant
Dim i, j, k, l, m As Integer
Dim b()


    Sheets("Collecte").Activate

    Set Wbk2 = ThisWorkbook

    Colonnes = Array("ISIN", "Mouvement", "Fonds", "Nombre de parts", "Montant", "Devise", "Date VL", "Date reglement", "Donneur d'ordre", "Investisseur", "Description", "Ordre en Montant / Parts")  'colonnes qui nous intéressent

  Fichier = Application.GetOpenFilename("Fichiers Excels, *.xls*") 'Sélection du fichier

  If Fichier <> False Then 'Si l'utilisateur choisit le fichier aggregation

    Set Wbk1 = Workbooks.Open(Fichier) 'ouverture du fichier
    With Wbk1.Sheets("Aggregated Data")


            ReDim F(1 To Cells(Rows.Count, 1).End(xlUp).Row, UBound(Colonnes) - LBound(Colonnes) + 1)


            For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
               

                For j = 1 To UBound(Colonnes) - LBound(Colonnes)
                
                    F(i, j) = Cells(i, j)
                  

                    b = ExtraitCol(F, j)
                
                    
                    For k = 5 To 27 'this suppose the number of orders aggregated do not exceed 27, can change according to needs
                        If Test(Wbk2.Sheets("Collecte").Cells(6, k), b) Then
                        
                            Wbk2.Sheets("Collecte").Cells(i, k).Offset(5, 0) = b(i)
                       
                       
                        End If
                    Next k
            
                Next j

            Next i
    
            For l = 2 To Cells(Rows.Count, 1).End(xlUp).Row + 7
                 Wbk2.Sheets("Collecte").Cells(l, 14).Offset(5, 0) = Wbk1.Sheets("Aggregated Data").Cells(l, 12)
            Next l
            
    
        Dim derlig As Integer
        derlig = Range("E" & Rows.Count).End(xlUp).Row + 5
        With ThisWorkbook.Worksheets("Collecte")
            
            '.Range("C7:C" & derlig).Formula = "=IF(S7<>"""";VLOOKUP(S7,DB!A:B,1,FALSE);"""")"
            '.Range("Q7:Q" & derlig).Formula = "=IF(X7<>"""",G7*Y7,"""")"
            .Range("Q7:Q" & derlig).Formula = "=H7"
            .Range("P7:P" & derlig).Formula = "=IF(S7<>"""",IF(ABS((H7/Q7)-1)<0.5%,""OK"",""NON""),"""")"
            '.Range("S7:S" & derlig).Formula = "=IF(C7<>"""",VLOOKUP(C7,DB!A:B,2,FALSE),"""")"
            .Range("V7:V" & derlig).Formula = "=IF(S7<>"""",VLOOKUP(S7,DB!B:O,14,FALSE),"""")"
            '.Range("V7:V" & derlig).Formula = "=I7"
            .Range("Z7:Z" & derlig).Formula = "=IF(S7<>"""",WORKDAY(TODAY(),-1),"""")"
            .Range("AA7:AA" & derlig).Formula = "=IF(S7<>"""",VLOOKUP(S7,DB!B:C,2,FALSE),"""")"
            '.Range("AA7:AA" & derlig).Formula = "=F7"
            .Range("X7:X" & derlig).Formula = "=J7"
            
            For m = 7 To derlig
                If .Range("N" & m) = "" Then
                    .Range("N" & m).Value = "Parts"
                End If
            Next m
            
        End With
    End With
    Wbk1.Close

  Else: MsgBox "Vous n'avez pas sélectionné de fichier."
  
  
  End If


Set Wbk1 = Nothing
Set Wbk2 = Nothing
End Sub

Function ExtraitCol(a, col)
' extracts a specific column from a 2D array and returns it as a 1D array.

  ReDim b(LBound(a) To UBound(a))
  Dim i As Integer
  For i = LBound(a) To UBound(a): b(i) = a(i, col): Next i
  ExtraitCol = b
End Function

Function Test(strIn As String, b()) As Boolean

    Test = Not (IsError(Application.Match(strIn, b, 0)))
    
End Function

Sub histo()
'after extracting useful data, save it in another sheet name "Histo_Orders"
lastligne = 1
lastRow = 7


'**Attention  à défiltrer avant toute tentative d'écriture sur l'historique**
ThisWorkbook.Sheets("Histo_Orders").Rows("1:1").AutoFilter
'*******'

Sheets("Collecte").Activate

lastligne = ThisWorkbook.Sheets("Histo_Orders").Range("A" & Rows.Count).End(xlUp).Row + 1

Do While ThisWorkbook.Sheets("Collecte").Cells(4, lastRow).Text <> ""

    lastRow = lastRow + 1
    
Loop

lastRow = lastRow - 1

With ThisWorkbook

    For j = 7 To lastRow
        
        .Sheets("Histo_Orders").Range("A" & lastligne).Value = .Sheets("Fax").Range("d" & j).Value
        .Sheets("Histo_Orders").Range("b" & lastligne).Value = .Sheets("Fax").Range("e" & j).Value
        .Sheets("Histo_Orders").Range("c" & lastligne).Value = .Sheets("Fax").Range("f" & j).Value
        .Sheets("Histo_Orders").Range("d" & lastligne).Value = .Sheets("Fax").Range("g" & j).Value
        .Sheets("Histo_Orders").Range("e" & lastligne).Value = .Sheets("Fax").Range("h" & j).Value
        .Sheets("Histo_Orders").Range("f" & lastligne).Value = .Sheets("Fax").Range("i" & j).Value
        .Sheets("Histo_Orders").Range("g" & lastligne).Value = .Sheets("Fax").Range("j" & j).Value
        .Sheets("Histo_Orders").Range("h" & lastligne).Value = .Sheets("Fax").Range("k" & j).Value
        .Sheets("Histo_Orders").Range("i" & lastligne).Value = .Sheets("Fax").Range("l" & j).Value
        .Sheets("Histo_Orders").Range("j" & lastligne).Value = .Sheets("Fax").Range("m" & j).Value
        .Sheets("Histo_Orders").Range("k" & lastligne).Value = .Sheets("Fax").Range("n" & j).Value
        .Sheets("Histo_Orders").Range("l" & lastligne).Value = .Sheets("Fax").Range("o" & j).Value
        
        lastligne = lastligne + 1
    Next j
    
End With


End Sub

Sub Importer_mails()

'in an other pre-arranged name "Mail", we prepare the cut off annoucement mail
'first pull out all orders related to a cut-off date spécified in cell D5 and the fund specified in cell D3
'than based on type of sending (to validation or final file) specified in cell D7 to apply a list of receivers appropriately
'than perform conversion calculation and total calculation in terms of fund and share class


    Dim ws_mm As Worksheet
    Dim histofax As Worksheet
    Dim max_histo As Integer
    Dim freq As String
    
    Dim k As Integer
    Dim p As Integer
    Dim q As Integer
    
    Dim last_sub As Integer
    Dim last_red As Integer
    
    Dim temp_s As Integer
    Dim temp_r As Integer
    
    Set ws_mm = Sheets("MAIL")
    Set histofax = Sheets("Histo_Orders")
    freq = ws_mm.Range("D3").Text
    
  
    '**Attention à réinitialiser les filtres actifs de l'historique avant d'importer les valeurs**
    ThisWorkbook.Sheets("Histo_Orders").Rows("1:1").AutoFilter
    '*******'
    
    'determine la derniere ligne
    max_histo = MaxIDColCells(2, 1, histofax)
    
    'nettoyage
    histofax.Range("A10000:H" & CStr(10000 + max_histo)).ClearContents
    

    'EXTRACTION DES LIGNES QUI NOUS INTERESSENT (CUT OFF ET FREQUENCE)
    k = 10000
    
    For j = 2 To max_histo
    
        If histofax.Range("H" & CStr(j)).Value = ws_mm.Range("D5").Value And histofax.Range("G" & CStr(j)).Text = freq And histofax.Range("M" & CStr(j)).Value <> "ANNULE" Then
            
            'ISIN
            histofax.Range("A" & CStr(j)).Copy _
                Destination:=histofax.Range("A" & CStr(k))
                
            'Fonds
            histofax.Range("J" & CStr(j)).Copy _
                Destination:=histofax.Range("B" & CStr(k))
                
            'nbpart/estimation
            histofax.Range("F" & CStr(j)).Copy _
                Destination:=histofax.Range("C" & CStr(k))
                
            'currency
            histofax.Range("K" & CStr(j)).Copy _
                Destination:=histofax.Range("D" & CStr(k))
                
            'VL du
            histofax.Range("D" & CStr(j)).Copy _
                Destination:=histofax.Range("E" & CStr(k))
                
            'Reglement
            histofax.Range("E" & CStr(j)).Copy _
                Destination:=histofax.Range("F" & CStr(k))
                
            'mvt
            histofax.Range("C" & CStr(j)).Copy _
                Destination:=histofax.Range("G" & CStr(k))
            
                
            k = k + 1
        End If
    
    
    Next j
    
    k = k - 10001
    p = 0
    q = 0
    
    
    
    For l = 0 To k
    
        temp_s = 13
        temp_r = 32
    
        last_sub = MaxIDColCells(14, 3, ws_mm)
        last_red = MaxIDColCells(33, 3, ws_mm)
        
        'IMPORTS DES SOUSCRIPTIONS SANS DOUBLONS
        If histofax.Range("G" & CStr(10000 + l)).Text = "Subscription" Then
        
            For S = 14 To last_sub
        
                If ws_mm.Range("C" & CStr(S)).Text = histofax.Range("A" & CStr(10000 + l)).Text Then
                
                    ws_mm.Range("F" & CStr(S)).Value = ws_mm.Range("F" & CStr(S)).Value + histofax.Range("C" & CStr(10000 + l)).Value
            
                    Exit For
                
                Else
                
                temp_s = temp_s + 1
            
                End If
            
            Next S
            
            If temp_s = last_sub Then
                
                histofax.Range("A" & CStr(10000 + l) & ":B" & CStr(10000 + l)).Copy
                    ws_mm.Range("C" & CStr(last_sub + 1)).PasteSpecial xlPasteValues
                    
                histofax.Range("C" & CStr(10000 + l) & ":F" & CStr(10000 + l)).Copy
                    ws_mm.Range("F" & CStr(last_sub + 1)).PasteSpecial xlPasteValues
                    
                Call ReNameFund(ws_mm, 14 + p, 4)
                p = p + 1
            
            End If
        
        Else
        'IMPORT DES RACHATS
        
            For S = 33 To last_red
        
                If ws_mm.Range("C" & CStr(S)).Text = histofax.Range("A" & CStr(10000 + l)).Text Then
                
                    ws_mm.Range("E" & CStr(S)).Value = ws_mm.Range("E" & CStr(S)).Value + histofax.Range("C" & CStr(10000 + l)).Value
            
                    Exit For
                
                Else
                
                temp_r = temp_r + 1
            
                End If
            
            Next S
            
            If temp_r = last_red Then
                
                histofax.Range("A" & CStr(10000 + l) & ":C" & CStr(10000 + l)).Copy
                    ws_mm.Range("C" & CStr(last_red + 1)).PasteSpecial xlPasteValues
                    
                histofax.Range("D" & CStr(10000 + l) & ":F" & CStr(10000 + l)).Copy
                    ws_mm.Range("G" & CStr(last_red + 1)).PasteSpecial xlPasteValues
                    
                Call ReNameFund(ws_mm, 33 + q, 4)
                q = q + 1
            
            End If
        
        
        End If
    
    Next l
    
    'On n'oublie pas de faire le ménage après soi
    histofax.Range("A10000:Z11000").ClearContents
    
    
End Sub

Sub Macro_Mail()

'now we do the email : formating with tableau and color !

    Application.DisplayAlerts = False
    Dim ws As Worksheet
    Set ws = Sheets("MAIL")
    ws.Range("P14:P25").ClearContents
    ws.Range("T14:T25").ClearContents
    ws.Range("P33:P48").ClearContents
    ws.Range("T33:T48").ClearContents
    Dim dif As String
    Dim plage As Range
    
    
    If ws.Range("D7").Value = "" Then
        
        MsgBox "Merci de renseigner la liste de diffusion"
        Exit Sub
        
    Else
        
        dif = ws.Range("D7").Value
    
    End If
    
    
    
    If ws.Range("C14").Value = "" And ws.Range("C33").Value = "" Then
    
        Call EMailPasDeCollecte_Macro(dif)
        
    Else
        Call EMail_Macro(dif)
        
    End If
    Application.DisplayAlerts = True


End Sub

Sub EMailPasDeCollecte_Macro(dif As String)
    Dim ws As Worksheet
    Dim iNbRows As Integer, iNbCols As Integer, i As Integer, j As Integer, iCellLength As Integer, X As Integer, iNbValTabs As Integer, iRepere As Integer, iCptRows As Integer
    Dim strHTML As String, sTo As String, sCc As String
    Dim Mail As Object
    Application.DisplayAlerts = True
    Set ws = Sheets("MAIL")
    sTo = ""
    sCc = ""
    With CreateObject("Outlook.Application")
        Set Mail = .CreateItem(0)
        With Mail
            .Subject = ws.Range("D11").Text & "Cut-off du " & ws.Range("D5").Value
            strHTML = ""
            strHTML = strHTML & "<HEAD>"
            strHTML = strHTML & "<style type='text/css'><!-- table, th, td {border: 1px solid black;border-collapse: collapse;font-family:arial;}table{width: 900px;}td{padding: 5px;}.classun{width: 100px;}.classdeux{width: 300px;} --></style>"
            strHTML = strHTML & "</HEAD>"
            strHTML = strHTML & "<BODY>"
            strHTML = strHTML & "Bonjour,<BR><BR>Pas de <b>Souscription</b> sur "
            strHTML = strHTML & ws.Range("D11").Value & ", date de cut-off du "
            strHTML = strHTML & ws.Range("D5").Value & " en trade date du "
            strHTML = strHTML & ws.Range("G11").Value & ". <BR> "
            strHTML = strHTML & " <BR>Pas de <b>Rachat</b> sur "
            strHTML = strHTML & ws.Range("D30").Value & ", date de cut-off du "
            strHTML = strHTML & ws.Range("D5").Value & " en trade date du "
            strHTML = strHTML & ws.Range("G30").Value & ". <BR>"
            strHTML = strHTML & "<BR>Cordialement,<BR><BR>SIP Reporting & Business Support. <BR>"
            strHTML = strHTML & "</BODY>"
            strHTML = strHTML & ""
            .HTMLBody = strHTML
            If dif = "Compliance niveau 1" Then
                iNbRows = WorksheetFunction.Max(MaxIDColCells(2, 9, Sheets("DB")), MaxIDColCells(2, 10, Sheets("DB")))
                sTo = sTo & Sheets("DB").Range("I2").Value
                sCc = sCc & Sheets("DB").Range("J2").Value
                For i = 3 To iNbRows
                    sTo = sTo & ";" & Sheets("DB").Range("I" & i).Value
                    sCc = sCc & ";" & Sheets("DB").Range("J" & i).Value
                Next i
                .To = sTo
                .CC = sCc
            ElseIf dif = "Compliance niveau 2" Then
                iNbRows = WorksheetFunction.Max(MaxIDColCells(2, 11, Sheets("DB")), MaxIDColCells(2, 12, Sheets("DB")))
                sTo = sTo & Sheets("DB").Range("K2").Value
                sCc = sCc & Sheets("DB").Range("L2").Value
                For i = 3 To iNbRows
                    sTo = sTo & ";" & Sheets("DB").Range("K" & i).Value
                    sCc = sCc & ";" & Sheets("DB").Range("L" & i).Value
                Next i
                .To = sTo
                .CC = sCc
            Else
                iNbRows = WorksheetFunction.Max(MaxIDColCells(2, 13, Sheets("DB")), MaxIDColCells(2, 14, Sheets("DB")))
                sTo = sTo & Sheets("DB").Range("M2").Value
                sCc = sCc & Sheets("DB").Range("N2").Value
                For i = 3 To iNbRows
                    sTo = sTo & ";" & Sheets("DB").Range("M" & i).Value
                    sCc = sCc & ";" & Sheets("DB").Range("N" & i).Value
                Next i
                .To = sTo
                .CC = sCc
            End If
            .Display
        End With
        Set Mail = Nothing
    End With
    Application.DisplayAlerts = True
End Sub

Sub EMail_Macro(dif As String)
'
    Dim ws As Worksheet
    Dim iNbRows As Integer, iNbCols As Integer, i As Integer, j As Integer, iCellLength As Integer, X As Integer, iNbValTabs As Integer, iRepere As Integer, iCptRows As Integer
    Dim strHTML As String, sTo As String, sCc As String, estim As String, estim1 As String
    Dim Mail As Object
    Dim temp As String
    
    Dim max_sub As Integer
    Dim max_red As Integer
    
    Application.DisplayAlerts = True
    
    Set ws = Sheets("MAIL")
    
    max_sub = MaxIDColCells(14, 3, ws)
    max_red = MaxIDColCells(33, 3, ws)
    
    'Check if we need to open a fx history file
    Dim need_fx As Boolean
    need_fx = False
    Dim isin As String
    Dim indice As Integer
    Dim trouve As Boolean
    Dim end_tab_fx As Long
    Dim k As Long
    Dim wbfx As Workbook
    Dim wbm As Worksheet
    Dim wbd As Worksheet
    
    
    Set wbm = ThisWorkbook.Sheets("Mail")
    Set wbd = ThisWorkbook.Sheets("DB")
    
    'SOUSCRIPTIONS
    i = 14
    While i < max_sub + 1 And wbm.Cells(i, 3) <> ""
        'isin = wbm.Cells(i, 3).Value
        'check if ccy part and ccy fund are the same
        If wbm.Cells(i, 7).Value <> wbm.Cells(i, 15).Value Then
            'check if there alrealdy is an opened fx historic file
            If need_fx = False Then
                MsgBox "Une conversion de devise est requise. Veuillez sélectionner le dernier historique de taux de change généré par Datahub."
                need_fx = True
                'Select a fx file for the conversion, choose the last one generated by datahub
                ' Get the fx historic file's path
                Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
                Application.FileDialog(msoFileDialogOpen).title = "Selectionner le dernier export fx"
                Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\aenprdneoxam\Batch\OUT\EXPORT_FX"
                iChoice = Application.FileDialog(msoFileDialogOpen).Show
                If iChoice <> 0 Then
                
                    ' Open the ReferentialDb
                    sDbPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
                    Set wbfx = Workbooks.Open(sDbPath)
                    end_tab_fx = Sheets(1).Range("A2").End(xlDown).Row
                    
                    'si le fichier est au format csv, appliquer le code si dessous
                    'Columns("A:A").Select
                    'Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                    'TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                    'Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                    ':=Array(1, 1), TrailingMinusNumbers:=True
                    
                    wbfx.Sheets(1).Columns("A:W").Sort key1:=Range("B1"), order1:=xlDescending, Header:=xlYes
                Else
                    MsgBox "Vous n'avez pas sélectionné de fichier. Veuillez recommencer."
                    Exit Sub
                End If
            End If
            If need_fx = True Then
                trouve = False
                k = 2
                    
                While trouve = False And k < end_tab_fx + 1
                    base_curr = Left(wbfx.Sheets(1).Cells(k, 1).Value, 3)
                    quote_curr = Mid(wbfx.Sheets(1).Cells(k, 1).Value, 4, 3)
                    If CDate(wbfx.Sheets(1).Cells(k, 2).Value) = wbm.Range("Q" & i).Value And base_curr = wbm.Cells(i, 7).Value And quote_curr = wbm.Cells(i, 15).Value Then
                        wbm.Range("P" & i) = wbfx.Sheets(1).Cells(k, 3).Value
                        wbm.Range("R" & i) = wbm.Range("P" & i).Value * wbm.Range("F" & i).Value
                        trouve = True
                    End If
                    k = k + 1
                Wend
            End If
        Else
            wbm.Range("P" & i) = 1
            wbm.Range("R" & i) = wbm.Range("P" & i).Value * wbm.Range("F" & i).Value
        End If
        i = i + 1
    Wend
    
    'RACHATS
    i = 33
    While i < max_red + 1 And wbm.Cells(i, 3) <> ""
        If wbm.Cells(i, 7).Value <> wbm.Cells(i, 15).Value Then
            If need_fx = False Then
                MsgBox "Une conversion de devise est requise. Veuillez sélectionner le dernier historique de taux de change généré par Datahub."
                need_fx = True
                'Select a fx file for the conversion, choose the last one generated by datahub
                ' Get the fx historic file's path
                Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
                Application.FileDialog(msoFileDialogOpen).title = "Selectionner le dernier export fx"
                Application.FileDialog(msoFileDialogOpen).InitialFileName = "\\aenprdneoxam\Batch\OUT\EXPORT_FX"
                iChoice = Application.FileDialog(msoFileDialogOpen).Show
                If iChoice <> 0 Then
                
                    ' Open the ReferentialDb
                    sDbPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
                    Set wbfx = Workbooks.Open(sDbPath)
                    end_tab_fx = Sheets(1).Range("A2").End(xlDown).Row
                    
                    'si le fichier est au format csv, appliquer le code si dessous
                    'Columns("A:A").Select
                    'Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                    'TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                    'Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                    ':=Array(1, 1), TrailingMinusNumbers:=True
                    
                    wbfx.Sheets(1).Columns("A:W").Sort key1:=Range("B1"), order1:=xlDescending, Header:=xlYes
                Else
                    MsgBox "Vous n'avez pas sélectionné de fichier. Veuillez recommencer."
                    Exit Sub
                End If
            End If
            
            If need_fx = True Then
                trouve = False
                k = 2
                        
                While trouve = False And k < end_tab_fx + 1
                    base_curr = Left(wbfx.Sheets(1).Cells(k, 1).Value, 3)
                    quote_curr = Mid(wbfx.Sheets(1).Cells(k, 1).Value, 4, 3)
                    If CDate(wbfx.Sheets(1).Cells(k, 2).Value) = wbm.Range("Q" & i).Value And base_curr = wbm.Cells(i, 7).Value And quote_curr = wbm.Cells(i, 15).Value Then
                        wbm.Range("P" & i) = wbfx.Sheets(1).Cells(k, 3).Value
                        wbm.Range("R" & i) = wbm.Range("P" & i).Value * wbm.Range("F" & i).Value
                        trouve = True
                    End If
                    k = k + 1
                Wend
            End If
        Else
            wbm.Range("P" & i) = 1
            wbm.Range("R" & i) = wbm.Range("P" & i).Value * wbm.Range("F" & i).Value
        End If
        i = i + 1
    Wend
    
    '''close the fx file without saving it
    'wbfx.Close False
    
    'tri les tableaux des souscriptions et rachats en fond et calcule le montant equivalent par fond
    'Call calc_somme_fond
    
    'writing email
    sTo = ""
    sCc = ""
    With CreateObject("Outlook.Application")
        Set Mail = .CreateItem(0)
        With Mail
            .Subject = ws.Range("D11").Text & " Cut-off" & " du " & wbm.Range("D5").Value
            strHTML = ""
            strHTML = strHTML & "<HEAD>"
            strHTML = strHTML & "<style type='text/css'><!-- table, th, td {border: 1px solid black;border-collapse: collapse;font-family:arial;}table{width: 900px;}td{padding: 5px;}.classun{width: 100px; text-align:center;}.classdeux{width: 300px; text-align:center;} --></style>"
            strHTML = strHTML & "</HEAD>"
            strHTML = strHTML & "<BODY>"
            strHTML = strHTML & "Bonjour,<BR><BR>-" & vbTab & "Ci-dessous les <b>Souscriptions</b> "
            strHTML = strHTML & wbm.Range("D11").Value & ", date de cut-off du "
            strHTML = strHTML & wbm.Range("D5").Value & " en trade date du "
            strHTML = strHTML & wbm.Range("G11").Value & " : <BR><BR> "
            
            If max_sub > 13 Then
            
                'TABLEAUX DES SOUSCRIPTIONS
                For i = 14 To max_sub
                    strHTML = strHTML & "<TABLE cellpadding='0' cellspacing='0'><B>"
                    strHTML = strHTML & "<TR halign='middle'nowrap>"
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 3) & "</FONT></TD>" 'ISIN
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classdeux'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 4) & "</FONT></TD>" 'Fonds
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 5) & "</FONT></TD>" 'nb parts
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 6) & "</FONT></TD>" 'estimation
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 7) & "</FONT></TD>" 'ccy
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 8) & "</FONT></TD>" 'vl du
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 9) & "</FONT></TD>" 'reglement
                    strHTML = strHTML & "</TR>"
                    
                    strHTML = strHTML & "<TR halign='middle'nowrap>"
                    
                    For j = 3 To 9
                        If j = 6 Or j = 5 Then
                            estim = ""
                            estim1 = ""
                            strHTML = strHTML & "<TD bgcolor='#FFFFFF'align='center'class='classun'><FONT SIZE=2><U>"
                            temp = wbm.Cells(i, j).Value
                            temp_integer = Fix(temp)
                            
                            If j = 5 Then
                                temp_digi = Round(wbm.Cells(i, j) - temp_integer, 3)
                                If Len(temp_digi) = 3 Then
                                    temp_digi = temp_digi & "00"
                                ElseIf Len(temp_digi) = 4 Then
                                    temp_digi = temp_digi & "0"
                                End If
                            Else
                                temp_digi = Round(wbm.Cells(i, j) - temp_integer, 2)
                                If Len(temp_digi) = 3 Then
                                    temp_digi = temp_digi & "0"
                                End If
                            End If
                            
                            
                            iCellLength = Len(temp_integer) Mod 3
                            If iCellLength <> 0 Then
                                estim = estim & Left(temp_integer, iCellLength)
                            End If
                            X = iCellLength + 1
                            
                            
                            While X < Len(temp_integer)
                                estim = estim & " " & Mid(temp_integer, X, 3)
                                X = X + 3
                            Wend

                            If InStr(estim, ",") > 0 Then
                            
                            Else
                                estim = Replace(estim, ".", ",")
                            End If
                            
                            If temp_digi <> 0 Then
                                estim = estim & temp_digi
                            End If
                            
                            estim1 = Replace(estim, " ,", ",")
                            
                            
                            If Not IsError(Replace(estim1, ".", ",")) Then
                                estim1 = Replace(estim1, "0,", ",")
                            Else
                                estim1 = Replace(estim1, "0.", ",")
                            End If
                            
                            strHTML = strHTML & estim1 & "</FONT></TD>"
    
                        Else
                        
                            strHTML = strHTML & "<TD bgcolor='#FFFFFF'align='center'class='classun'><FONT SIZE=2><U>" & wbm.Cells(i, j).Value & "</U></FONT></TD>"
                        
                        End If
                    Next j
                    
                    
                    '''tableaux des totaux souscriptions
                    If wbm.Range("T" & i) <> "" Then
                        strHTML = strHTML & "</B></TABLE><BR>"
                        strHTML = strHTML & "<TABLE cellpadding='0' cellspacing='0'><B>"
                        strHTML = strHTML & "<TR halign='middle'nowrap>"
                        'strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2></FONT></TD>"
                        strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2></FONT></TD>"
                        strHTML = strHTML & "<TD  width='27'bgcolor='#FFFFFF'align='left'class='classdeux'><FONT COLOR ='BLACK' SIZE=2><FONT face='arial'><I>Total des souscriptions en " & Range("O" & i) & " équivalent<I></FONT></TD>"
                        strHTML = strHTML & "<TD width='10' bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2></FONT></TD>"
                        Range("R" & i) = Round(Range("R" & i), 0)
                        strHTML = strHTML & "<TD bgcolor='#FFFFFF'align='center'><FONT SIZE=2><FONT face='arial'><I><U>"
                        iCellLength = Len(Range("T" & i)) Mod 3
                        If iCellLength <> 0 Then
                            strHTML = strHTML & Left(Range("T" & i), iCellLength) & " "
                        End If
                        X = iCellLength + 1
                        While X < Len(Range("T" & i))
                            strHTML = strHTML & Mid(Range("T" & i), X, 3) & " "
                        X = X + 3
                        Wend
                        strHTML = strHTML & "</FONT></TD>"
                        strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2><FONT face='arial'><I><U>" & Range("O" & i) & "<U><I></FONT></TD>"
                        strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2><FONT face='arial'><I><U>" & Range("H" & i) & "<U><I></FONT></TD>"
                        strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2><FONT face='arial'><I><U>" & Range("I" & i) & "<U><I></FONT></TD>"
                    
                    End If
                    
                    strHTML = strHTML & "</TR>"
                    strHTML = strHTML & "</B></TABLE><BR>"
                    
                    
                Next i
                
            Else
            
            strHTML = strHTML & "Pas de Souscriptions.<BR>"
            
            End If
            
            
            strHTML = strHTML & " <BR>-" & vbTab & "Ci-dessous les <b>Rachats</b>  "
            strHTML = strHTML & wbm.Range("D30").Value & ", date de cut-off du "
            strHTML = strHTML & wbm.Range("D5").Value & " en trade date du "
            strHTML = strHTML & wbm.Range("G30").Value & " : <BR><BR>"
            
            If max_red > 32 Then
            
                'TABLEAU DES RACHATS
                For i = 33 To max_red
                    strHTML = strHTML & "<TABLE cellpadding='0' cellspacing='0'><B>"
                    strHTML = strHTML & "<TR halign='middle'nowrap>"
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 3) & "</FONT></TD>" 'ISIN
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classdeux'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 4) & "</FONT></TD>"
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 5) & "</FONT></TD>"
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 6) & "</FONT></TD>"
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 7) & "</FONT></TD>"
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 8) & "</FONT></TD>"
                    strHTML = strHTML & "<TD bgcolor='#006A8D'align='left'class='classun'><FONT COLOR ='WHITE' SIZE=2>" & wbm.Cells(13, 9) & "</FONT></TD>"
                    strHTML = strHTML & "</TR>"
                    
                    strHTML = strHTML & "<TR halign='middle'nowrap>"
                    
                    For j = 3 To 9
                        'If j = 5 Then
                            
                            'strHTML = strHTML & "<TD bgcolor='#FFFFFF'align='center'class='classun'><FONT SIZE=2><U>" & "- " & Cells(i, j) & "</U></FONT></TD>"
                        
                        If j = 6 Or j = 5 Then
                            strHTML = strHTML & "<TD bgcolor='#FFFFFF'align='center'class='classun'><FONT SIZE=2><U>"
                            estim = ""
                            estim1 = ""
                         
                            temp = wbm.Cells(i, j).Value
                            temp_integer = Fix(temp)
                            
                            If j = 5 Then
                                temp_digi = Round(temp - temp_integer, 3)
                            
                                If Len(temp_digi) = 3 Then
                                    temp_digi = temp_digi & "00"
                                ElseIf Len(temp_digi) = 4 Then
                                    temp_digi = temp_digi & "0"
                                End If
                            
                            Else
                                temp_digi = Round(temp - temp_integer, 2)
                                If Len(temp_digi) = 3 Then
                                    temp_digi = temp_digi & "0"
                                End If
                            End If

                            iCellLength = Len(temp_integer) Mod 3
                            If iCellLength <> 0 Then
                                estim = estim & Left(temp_integer, iCellLength) & " "
                            End If
                            X = iCellLength + 1
                            While X < Len(temp_integer)
                                estim = estim & Mid(temp_integer, X, 3) & " "
                                X = X + 3
                            Wend

                            If InStr(estim, ",") > 0 Then

                            Else

                                estim = Replace(estim, ".", ",")

                            End If
                            estim1 = Replace(estim, " ,", ",")
                            Debug.Print estim1 & " " & temp_digi
                            If temp_digi <> 0 Then
                                estim1 = estim1 & temp_digi
                            End If
                            
                            If Not IsError(Replace(estim1, ".", ",")) Then
                                estim1 = Replace(estim1, " 0,", ",")
                            Else
                                estim1 = Replace(estim1, " 0.", ",")
                            End If
                            
                            
                           
                            strHTML = strHTML & "- " & estim1 & "</FONT></TD>"
    
                        Else
                        
                            strHTML = strHTML & "<TD bgcolor='#FFFFFF'align='center'class='classun'><FONT SIZE=2><U>" & wbm.Cells(i, j) & "</U></FONT></TD>"
                        
                        End If
                    Next j
                    
                    '''tableaux des totaux rachats
                    If wbm.Range("T" & i) <> "" Then
                        strHTML = strHTML & "</B></TABLE><BR>"
                        strHTML = strHTML & "<TABLE cellpadding='0' cellspacing='0'><B>"
                        strHTML = strHTML & "<TR halign='middle'nowrap>"
                        'strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2></FONT></TD>"
                        strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2></FONT></TD>"
                        strHTML = strHTML & "<TD  width='27'bgcolor='#FFFFFF'align='left'class='classdeux'><FONT COLOR ='BLACK' SIZE=2><FONT face='arial'><I>Total des rachats en " & Range("O" & i) & " équivalent<I></FONT></TD>"
                        strHTML = strHTML & "<TD bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2></FONT></TD>"
                        Range("R" & i) = Round(Range("R" & i), 0)
                        strHTML = strHTML & "<TD bgcolor='#FFFFFF'align='center'><FONT SIZE=2><FONT face='arial'><I><U>"
                        iCellLength = Len(Range("T" & i)) Mod 3
                        If iCellLength <> 0 Then
                            strHTML = strHTML & Left(Range("T" & i), iCellLength) & " "
                        End If
                        X = iCellLength + 1
                        While X < Len(Range("T" & i))
                            strHTML = strHTML & Mid(Range("T" & i), X, 3) & " "
                        X = X + 3
                        Wend
                        strHTML = strHTML & "</FONT></TD>"
                        strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2><FONT face='arial'><I><U>" & Range("O" & i) & "<U><I></FONT></TD>"
                        strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2><FONT face='arial'><I><U>" & Range("H" & i) & "<U><I></FONT></TD>"
                        strHTML = strHTML & "<TD width='11%'bgcolor='#FFFFFF'align='left'class='classun'><FONT COLOR ='BLACK' SIZE=2><FONT face='arial'><I><U>" & Range("I" & i) & "<U><I></FONT></TD>"
                    
                    End If
                    
                    strHTML = strHTML & "</TR>"
                    strHTML = strHTML & "</B></TABLE><BR>"
                    
                Next i
                
            Else
            
            strHTML = strHTML & "Pas de Rachats.<BR>"
            
            End If
            
            strHTML = strHTML & "<BR>Cordialement,<BR><BR>SIP Reporting & Business Support. <BR>"
            strHTML = strHTML & "</BODY>"
            strHTML = strHTML & ""
            .HTMLBody = strHTML
            If dif = "Compliance niveau 1" Then
                iNbRows = WorksheetFunction.Max(MaxIDColCells(2, 9, wbd), MaxIDColCells(2, 10, wbd))
                sTo = sTo & wbd.Range("I2").Value
                sCc = sCc & wbd.Range("J2").Value
                For i = 3 To iNbRows
                    sTo = sTo & ";" & wbd.Range("I" & i).Value
                    sCc = sCc & ";" & wbd.Range("J" & i).Value
                Next i
                .To = sTo
                .CC = sCc
            ElseIf dif = "Compliance niveau 2" Then
                iNbRows = WorksheetFunction.Max(MaxIDColCells(2, 11, wbd), MaxIDColCells(2, 12, wbd))
                sTo = sTo & wbd.Range("K2").Value
                sCc = sCc & wbd.Range("L2").Value
                For i = 3 To iNbRows
                    sTo = sTo & ";" & wbd.Range("K" & i).Value
                    sCc = sCc & ";" & wbd.Range("L" & i).Value
                Next i
                .To = sTo
                .CC = sCc
            Else
                iNbRows = WorksheetFunction.Max(MaxIDColCells(2, 13, wbd), MaxIDColCells(2, 14, wbd))
                sTo = sTo & wbd.Range("M2").Value
                sCc = sCc & wbd.Range("N2").Value
                For i = 3 To iNbRows
                    sTo = sTo & ";" & wbd.Range("M" & i).Value
                    sCc = sCc & ";" & wbd.Range("N" & i).Value
                Next i
                .To = sTo
                .CC = sCc
            End If
            .Display
        End With
        Set Mail = Nothing
    End With
    Application.DisplayAlerts = True
End Sub


' ----------------------------------------
' Function    : MaxIDColCells
' Description : Get the max ID of non empty cells of a column "iOptionalColID" starting from "iStartingLine"
'               If iOptionalColID is not set, use the first column ID
' ----------------------------------------
Function MaxIDColCells(Optional iStartingLine As Integer = 1, Optional iOptionalColID As Integer = 1, Optional wsOptional As Worksheet, Optional wbOptional As Workbook) As Integer
    Dim iNbCells As Integer
    Dim sColLetter As String
    Dim ws As Worksheet
    Dim wb As Workbook
    ' If wbOptional is not set, use the active workbook
    If wbOptional Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = wbOptional
    End If
    ' If wsOptional is not set, use the active sheet
    If wsOptional Is Nothing Then
        Set ws = wb.ActiveSheet
    Else
        Set ws = wsOptional
    End If
    iNbCells = iStartingLine
    ' Convert "iOptionalColID" into its letter
    sColLetter = ConvertToLetter(iOptionalColID)
    While Not ws.Range(sColLetter & iNbCells).Value Like ""
        iNbCells = iNbCells + 1
    Wend
    MaxIDColCells = iNbCells - 1
End Function

' ----------------------------------------
' Function    : ConvertToLetter
' Description : Convert a column ID "iCol" into its letter
' ----------------------------------------
Function ConvertToLetter(iCol As Integer) As String
    ConvertToLetter = Split(Cells(1, iCol).Address, "$")(1)
End Function

Function ReNameFund(sheet As Worksheet, i As Integer, j As Integer)
    
    Dim final As String
    Dim fund As String
    Dim fundArray As Variant
    Dim F As Variant
    
    fundArray = Array("B CHF", "B EUR", "B USD", "B GBP", "B JPY", "D CHF", "D EUR", "D USD", "I C CHF", "I C EUR", "I C USD", "I D CHF", "I D EUR", "I D USD", "M C USD", "B1 CHF", "B1 EUR", "B1 GBP", "B1 JPY", "B1 USD", "B2 CHF", "B2 EUR", "B2 GBP", "B2 JPY", "B2 USD", "B3 CHF", "B3 EUR", "B3 GBP", "B3 JPY", "B3 USD", "M C EUR", "MC CHF", "D2 CHF")
    
    final = "TAM Project -" & sheet.Range("D11").Text & "-"
    
    fund = sheet.Cells(i, j).Text
    
    For Each F In fundArray
    
        If InStr(fund, F) > 0 Then
        
            final = final & F
            
            If InStr(final, "USD") > 0 Then
            
            Else
            
                final = final & " H"
            
            End If
            
            sheet.Cells(i, j).Value = final
            
        End If
    
    Next F
    
End Function

' After all, lets not forget to reinitiate the workbook
Sub Reset()

    Dim ws As Worksheet
    
    Set ws = Sheets("MAIL")
    
    ws.Range("D3:D7").Value = ""
    ws.Range("C14:D25").Value = ""
    ws.Range("F14:I25").Value = ""
    ws.Range("C33:E48").Value = ""
    ws.Range("G33:I48").Value = ""
    ws.Range("P14:P25").Value = ""
    ws.Range("R14:R25").Value = ""
    ws.Range("T14:T25").Value = ""
    ws.Range("P33:P48").Value = ""
    ws.Range("R33:R48").Value = ""
    ws.Range("T33:T48").Value = ""

    Sheets("Collecte").Range("C7:AA" & Application.WorksheetFunction.CountA(Range("S:S")) + 5).Value = ""

End Sub


