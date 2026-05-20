Attribute VB_Name = "Module3"
Option Explicit

' --- ZMIENNE GLOBALNE (KOMUNIKACJA Z FORMULARZEM) ---
Public SzukaneKlucze As Collection
Public SciezkiSerwer As Collection
Public SzukanaDataRaw As String
Public CzySzukacOutlook As Boolean
Public CzySzukacSerwer As Boolean

' --- FLAGI SESJI (LOGIKA: NADPISZ/POMIÑ WSZYSTKIE) ---
Dim G_NadpiszWszystkie As Boolean
Dim G_PominWszystkie As Boolean

' Zmienne daty
Dim DataOd As Date
Dim DataDo As Date
Dim UzytoDaty As Boolean

' --- KONFIGURACJA (ODCZYT Z REJESTRU) ---
Private Const APP_CFG As String = "OutlookWI_Search"
Private Const SEC_CFG As String = "Konfiguracja"

Private Function PobierzKonfiguracje(sciezka As String, klucz As String, domyslna As String) As String
    PobierzKonfiguracje = GetSetting(APP_CFG, SEC_CFG, klucz, domyslna)
End Function

Private Function ProgDomyslny(k As String, d As String) As String
    ProgDomyslny = PobierzKonfiguracje("", k, d)
End Function

Private Function PobierzSlownikKluczy() As Variant
    Dim raw As String
    raw = PobierzKonfiguracje("", "SlownikKluczy", "umowa|podwykonaw|aneks|zlecenie|kontrakt|draft|wniosek|wdr|faktura|protokol|zamowienie|sprawozdanie|pozwolenie|decyzja")
    PobierzSlownikKluczy = Split(raw, "|")
End Function

Private Function PobierzMinRozmiarZalacznika() As Long
    PobierzMinRozmiarZalacznika = CLng(PobierzKonfiguracje("", "MinRozmiarZalacznika", "51200"))
End Function

Private Function PobierzSciezkeBazowa() As String
    PobierzSciezkeBazowa = PobierzKonfiguracje("", "SciezkaBazowa", "D:\DANE\")
End Function

Private Function CzyZapisywacMetadane() As Boolean
    CzyZapisywacMetadane = CBool(PobierzKonfiguracje("", "ZapisujMetadane", "True"))
End Function

' =========================================================
' 1. G£ÓWNA PROCEDURA WYSZUKIWANIA
' =========================================================
Sub ZaawansowaneWyszukiwanieWI()
    Dim folderZapisu As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim docelowyFolder As String
    
    ' RESET STANU
    Set SzukaneKlucze = New Collection
    Set SciezkiSerwer = New Collection
    SzukanaDataRaw = "": CzySzukacOutlook = False: CzySzukacSerwer = False
    UzytoDaty = False: G_NadpiszWszystkie = False: G_PominWszystkie = False

    ' Pokazanie okna formularza
    frmSzukajWI.Show
    
    If SzukaneKlucze.Count = 0 And SzukanaDataRaw = "" Then Exit Sub
    
    UzytoDaty = AnalizujDate(SzukanaDataRaw, DataOd, DataDo)
    
    docelowyFolder = PobierzSciezkeZapisu()
    If docelowyFolder = "" Then
        MsgBox "Anulowano zapisywanie plików.", vbInformation, "Przerwano"
        Exit Sub
    End If
    
    Dim nazwaFolderu As String
    If SzukaneKlucze.Count > 0 Then nazwaFolderu = CleanFileName(CStr(SzukaneKlucze(1))) Else nazwaFolderu = "Wyniki_" & Format(Now, "HHmm")
    
    folderZapisu = docelowyFolder & nazwaFolderu & "_wyniki\"
    If Not fso.FolderExists(folderZapisu) Then fso.CreateFolder folderZapisu

    On Error GoTo ErrorHandler
    
    Dim licznikOutlook As Integer: licznikOutlook = 0
    Dim licznikSerwer As Integer: licznikSerwer = 0
    
    ' --- ETAP A: OUTLOOK ---
    If CzySzukacOutlook Then
        ' [NOWE] Inicjalizacja paska przed startem pêtli
        frmSzukajWI.AktualizujPostep 0, "Przygotowanie skanowania Outlooka..."
        
        Dim ns As Outlook.NameSpace: Set ns = Application.GetNamespace("MAPI")
        Dim items As Outlook.items, mail As Object, atmt As Outlook.Attachment
        Dim i As Long, filtrSQL As String, warunkiTekst As String, warunekDaty As String
        Dim klucz As Variant, destFile As String
        
        For Each klucz In SzukaneKlucze
            Dim warunekPojedynczy As String
            warunekPojedynczy = BuildOutlookSQLForField(CStr(klucz))
            If warunkiTekst <> "" Then warunkiTekst = warunkiTekst & " AND " & warunekPojedynczy Else warunkiTekst = warunekPojedynczy
        Next klucz
        
        If UzytoDaty Then
            If DataOd > 0 Then warunekDaty = """urn:schemas:httpmail:datereceived"" >= '" & Format(DataOd, "yyyy-mm-dd 00:00") & "'"
            If DataDo > 0 Then
                Dim wDo As String: wDo = """urn:schemas:httpmail:datereceived"" <= '" & Format(DataDo, "yyyy-mm-dd 23:59") & "'"
                If warunekDaty <> "" Then warunekDaty = "(" & warunekDaty & " AND " & wDo & ")" Else warunekDaty = wDo
            End If
        End If
        
        filtrSQL = "@SQL=" & IIf(warunkiTekst <> "" And warunekDaty <> "", "(" & warunkiTekst & ") AND " & warunekDaty, warunkiTekst & warunekDaty)
        Set items = ns.GetDefaultFolder(olFolderInbox).items.Restrict(filtrSQL)
        
        ' [NOWE] Obliczamy ca³kowit¹ iloœæ maili do paska
        Dim totalItems As Long
        totalItems = items.Count
        
        For i = totalItems To 1 Step -1
            ' [NOWE] Odœwie¿anie paska co 5 maili
            If i Mod 5 = 0 Then
                frmSzukajWI.AktualizujPostep (totalItems - i) / totalItems, "Przeszukiwanie Outlooka..."
            End If
            
            If TypeOf items.item(i) Is mailItem Then
                Set mail = items.item(i)
                For Each atmt In mail.Attachments
                    If atmt.Size > 51200 And Not CzyToObraz(atmt.fileName) Then
                        destFile = folderZapisu & "Outlook_" & Format(mail.ReceivedTime, "yyyy-mm-dd_") & CleanFileName(atmt.fileName)
                        If ZarzadzajKonfliktem(destFile, fso) Then
                            atmt.SaveAsFile destFile
                            If CzyZapisywacMetadane Then ZapiszMetadane destFile, mail, fso
                            licznikOutlook = licznikOutlook + 1
                        End If
                    End If
                Next atmt
            End If
        Next i
    End If

    ' --- ETAP B: SERWER ---
    If CzySzukacSerwer Then
        Dim sciezkaItem As Variant
        ' [NOWE] Zmienne do œledzenia postêpu folderów
        Dim folderIndex As Integer: folderIndex = 1
        Dim totalFolders As Integer: totalFolders = SciezkiSerwer.Count
        
        For Each sciezkaItem In SciezkiSerwer
            ' [NOWE] Odœwie¿anie paska przy zmianie folderu
            frmSzukajWI.AktualizujPostep (folderIndex / totalFolders), "Skanowanie folderu: " & fso.GetFileName(CStr(sciezkaItem))
            
            If fso.FolderExists(CStr(sciezkaItem)) Then
                licznikSerwer = licznikSerwer + SzukajRekurencyjnie(CStr(sciezkaItem), folderZapisu, fso)
            End If
            folderIndex = folderIndex + 1
        Next sciezkaItem
    End If
    
    ' [NOWE] Wymuszenie 100% po zakoñczeniu
    frmSzukajWI.AktualizujPostep 1, "Zakoñczono!"
    ' [NOWE] Zamkniêcie formularza z paskiem
    Unload frmSzukajWI

    MsgBox "Eksport zakoñczony!" & vbCrLf & "Outlook: " & licznikOutlook & " | Serwer: " & licznikSerwer, vbInformation, "WI Status"
    Call Shell("explorer.exe " & folderZapisu, vbNormalFocus)
    Exit Sub

ErrorHandler:
    Unload frmSzukajWI ' [NOWE] Zabezpieczenie przed zablokowaniem formularza przy b³êdzie
    MsgBox "B³¹d: " & Err.Description, vbCritical
End Sub

' =========================================================
' FUNKCJE POMOCNICZE
' =========================================================

Private Function ZarzadzajKonfliktem(ByVal sciezka As String, fso As Object) As Boolean
    If Not fso.FileExists(sciezka) Then ZarzadzajKonfliktem = True: Exit Function
    If G_NadpiszWszystkie Then ZarzadzajKonfliktem = True: Exit Function
    If G_PominWszystkie Then ZarzadzajKonfliktem = False: Exit Function
    
    Dim pyt As VbMsgBoxResult
    pyt = MsgBox("Plik ju¿ istnieje: " & vbCrLf & Mid(sciezka, InStrRev(sciezka, "\") + 1) & vbCrLf & vbCrLf & _
                 "[TAK] - Nadpisz ten jeden" & vbCrLf & _
                 "[NIE] - Pomiñ ten jeden" & vbCrLf & _
                 "[ANULUJ] - OPCJA DLA WSZYSTKICH", vbYesNoCancel + vbQuestion, "Konflikt plików")
                 
    If pyt = vbYes Then ZarzadzajKonfliktem = True
    If pyt = vbNo Then ZarzadzajKonfliktem = False
    If pyt = vbCancel Then
        Dim zb As VbMsgBoxResult
        zb = MsgBox("Czy nadpisaæ WSZYSTKIE duble?" & vbCrLf & "[TAK] = Wszystkie" & vbCrLf & "[NIE] = Pomiñ wszystkie", vbYesNo + vbExclamation, "Decyzja zbiorcza")
        If zb = vbYes Then G_NadpiszWszystkie = True: ZarzadzajKonfliktem = True Else G_PominWszystkie = True: ZarzadzajKonfliktem = False
    End If
End Function

Private Function BuildOutlookSQLForField(ByVal searchExpr As String) As String
    Dim tokens() As String, sql As String, currentOp As String, i As Integer, w As String, termSQL As String
    tokens = Split(searchExpr, " ")
    currentOp = "AND"
    sql = ""
    
    For i = 0 To UBound(tokens)
        w = UCase(tokens(i))
        If w = "AND" Or w = "OR" Then
            currentOp = w
        ElseIf w = "NOT" Then
            currentOp = "AND NOT"
        Else
            Dim term As String
            term = Replace(Replace(tokens(i), "'", "''"), "*", "%")
            If InStr(term, "%") = 0 Then term = "%" & term & "%"
            
            termSQL = "(""urn:schemas:httpmail:textdescription"" LIKE '" & term & "' " & _
                      "OR ""urn:schemas:httpmail:subject"" LIKE '" & term & "' " & _
                      "OR ""urn:schemas:httpmail:fromname"" LIKE '" & term & "')"
            
            If sql = "" Then
                sql = IIf(currentOp = "AND NOT", "NOT " & termSQL, termSQL)
            Else
                sql = sql & " " & currentOp & " " & termSQL
            End If
            currentOp = "AND"
        End If
    Next i
    BuildOutlookSQLForField = "(" & sql & ")"
End Function

Function AnalizujDate(ByVal txt As String, ByRef dOd As Date, ByRef dDo As Date) As Boolean
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim m As Object: txt = Trim(txt)
    If txt = "" Then Exit Function
    regEx.Global = True: regEx.Pattern = "\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4}"
    Set m = regEx.Execute(txt)
    If m.Count >= 2 Then
        dOd = CDate(m(0).Value): dDo = CDate(m(1).Value)
    ElseIf m.Count = 1 Then
        dOd = CDate(m(0).Value): dDo = dOd
        If InStr(txt, ">") > 0 Then dDo = 0 Else If InStr(txt, "<") > 0 Then dDo = dOd: dOd = 0
    End If
    AnalizujDate = (m.Count > 0)
End Function

Function SzukajTylkoWGlownymFolderze(ByVal sciezka As String, ByVal cel As String, fso As Object) As Integer
    On Error Resume Next
    Dim plik As Object, ile As Integer, folderPath As String
    Dim match As Boolean, klucz As Variant, dest As String
    ile = 0
    folderPath = IIf(Right(sciezka, 1) <> "\", sciezka & "\", sciezka)

    For Each plik In fso.GetFolder(folderPath).Files
        match = True
        For Each klucz In SzukaneKlucze
            If InStr(1, plik.Name, CStr(klucz), vbTextCompare) = 0 Then
                match = False
                Exit For
            End If
        Next klucz

        If match And Not CzyToObraz(plik.Name) Then
            dest = cel & "Serwer_" & plik.Name
            If ZarzadzajKonfliktem(dest, fso) Then
                plik.Copy dest, True
                ile = ile + 1
            End If
        End If
    Next plik
    SzukajTylkoWGlownymFolderze = ile
End Function
Function SzukajRekurencyjnie(ByVal sciezka As String, ByVal cel As String, fso As Object) As Integer
    On Error Resume Next
    Dim ile As Integer, folderPath As String
    Dim plik As Object, podFolder As Object
    Dim match As Boolean, klucz As Variant, dest As String
    
    ile = 0
    folderPath = IIf(Right(sciezka, 1) <> "\", sciezka & "\", sciezka)
    
    If Not fso.FolderExists(folderPath) Then
        SzukajRekurencyjnie = 0
        Exit Function
    End If
    
    For Each plik In fso.GetFolder(folderPath).Files
        match = True
        For Each klucz In SzukaneKlucze
            If InStr(1, plik.Name, CStr(klucz), vbTextCompare) = 0 Then
                match = False
                Exit For
            End If
        Next klucz
        
        If match And Not CzyToObraz(plik.Name) Then
            dest = cel & "Serwer_" & plik.Name
            If ZarzadzajKonfliktem(dest, fso) Then
                plik.Copy dest, True
                ile = ile + 1
            End If
        End If
    Next plik
    
    For Each podFolder In fso.GetFolder(folderPath).SubFolders
        ile = ile + SzukajRekurencyjnie(podFolder.Path, cel, fso)
    Next podFolder
    
    SzukajRekurencyjnie = ile
End Function

Private Function CzyToObraz(ByVal nazwaPliku As String) As Boolean
    Dim ext As String
    ext = LCase(Mid(nazwaPliku, InStrRev(nazwaPliku, ".")))
    CzyToObraz = (ext = ".png" Or ext = ".jpg" Or ext = ".jpeg" Or _
                  ext = ".gif" Or ext = ".bmp" Or ext = ".tiff" Or _
                  ext = ".tif" Or ext = ".webp" Or ext = ".svg" Or _
                  ext = ".ico" Or ext = ".emz" Or ext = ".wmz")
End Function
Private Sub ZapiszMetadane(ByVal plikDocelowy As String, mail As Object, fso As Object)
    On Error Resume Next
    Dim metaPath As String
    metaPath = Left(plikDocelowy, InStrRev(plikDocelowy, ".") - 1) & "".meta""
    
    Dim meta As String
    meta = ""From: "" & mail.SenderName & "" <"" & mail.SenderEmailAddress & "">"" & vbCrLf & _
           ""To: "" & mail.To & vbCrLf & _
           ""CC: "" & mail.CC & vbCrLf & _
           ""Subject: "" & mail.Subject & vbCrLf & _
           ""Received: "" & Format(mail.ReceivedTime, ""yyyy-mm-dd HH:mm:ss"") & vbCrLf & _
           ""OriginalFile: "" & fso.GetFileName(plikDocelowy)
    
    Dim ts As Object
    Set ts = fso.CreateTextFile(metaPath, True, False)
    ts.Write meta
    ts.Close
End Sub
Function CleanFileName(n As String) As String
    Dim v As Variant
    CleanFileName = n
    For Each v In Array("/", "\", ":", "*", "?", Chr(34), "<", ">", "|")
        CleanFileName = Replace(CleanFileName, CStr(v), "_")
    Next v
End Function

Private Function OcjenReguly(subjekt As String, nadawca As String, nazwaPliku As String) As String
    Dim i As Integer, warunek As String, cel As String, pattern As String
    Dim reguly As String, regula As Variant
    OcjenReguly = """"
    
    For i = 1 To 20
        warunek = GetSetting(APP_CFG, ""Reguly\" & i, ""Warunek"", """")
        If warunek = """" Then Exit For
        
        pattern = GetSetting(APP_CFG, ""Reguly\" & i, ""Wzorzec"", """")
        cel = GetSetting(APP_CFG, ""Reguly\" & i, ""Cel"", """")
        If cel = """" Then GoTo NastêpnaRegula
        
        Select Case warunek
            Case ""Temat""
                If InStr(1, subjekt, pattern, vbTextCompare) > 0 Then OcjenReguly = cel
            Case ""Nadawca""
                If InStr(1, nadawca, pattern, vbTextCompare) > 0 Then OcjenReguly = cel
            Case ""Zalacznik""
                If InStr(1, nazwaPliku, pattern, vbTextCompare) > 0 Then OcjenReguly = cel
        End Select
        
        If OcjenReguly <> """" Then Exit Function
NastêpnaRegula:
    Next i
End Function
Function PobierzSciezkeZapisu() As String
    Dim OstatniaSciezka As String, odpowiedz As VbMsgBoxResult
    Dim objShell As Object, objFolder As Object, nowaSciezka As String
    
    OstatniaSciezka = VBA.GetSetting("OutlookWI_Search", "Ustawienia", "OstatniKatalogZapisu", "")
    
    If OstatniaSciezka <> "" Then
        odpowiedz = MsgBox("Wyniki s¹ gotowe do skopiowania." & vbCrLf & vbCrLf & _
                           "Ostatnie miejsce zapisu to:" & vbCrLf & OstatniaSciezka & vbCrLf & vbCrLf & _
                           "Czy ZAPISAÆ JAK POPRZEDNIO?", vbYesNoCancel + vbQuestion, "Kreator zapisu")
        If odpowiedz = vbYes Then
            PobierzSciezkeZapisu = OstatniaSciezka
            Exit Function
        End If
    End If
    
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "Wybierz nowy folder na wyniki:", 0, 0)
    
    If Not objFolder Is Nothing Then
        nowaSciezka = objFolder.Self.path
        If Right(nowaSciezka, 1) <> "\" Then nowaSciezka = nowaSciezka & "\"
        VBA.SaveSetting "OutlookWI_Search", "Ustawienia", "OstatniKatalogZapisu", nowaSciezka
        PobierzSciezkeZapisu = nowaSciezka
    Else
        PobierzSciezkeZapisu = ""
    End If
End Function

Sub WyszukajPlikiZAnimacja()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folderPath As String
    folderPath = "D:\DANE" ' Zmieñ na swój folder docelowy
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder nie istnieje!", vbCritical
        Exit Sub
    End If
    
    ' Zbieramy informacje o plikach
    Dim oFolder As Object
    Set oFolder = fso.GetFolder(folderPath)
    
    Dim totalFiles As Long
    totalFiles = oFolder.Files.Count ' Wa¿ne: to sprawdza tylko g³ówny katalog, nie podfoldery
    
    If totalFiles = 0 Then
        MsgBox "Folder jest pusty.", vbInformation
        Exit Sub
    End If
    
    ' Pokazujemy formularz w trybie nieblokuj¹cym
    frmSzukajWI.Show vbModeless
    
    Dim oFile As Object
    Dim currentIndex As Long
    currentIndex = 0
    
    ' Rozpoczynamy pêtlê po plikach
    For Each oFile In oFolder.Files
        currentIndex = currentIndex + 1
        
        ' --- AKTUALIZACJA PASKA ---
        If currentIndex Mod 10 = 0 Then
            frmSzukajWI.AktualizujPostep currentIndex / totalFiles, "Przetwarzanie plikow..."
        End If
    Next oFile
    
    Unload frmSzukajWI
End Sub
