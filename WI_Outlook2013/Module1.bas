Attribute VB_Name = "Module1"
Public Sub ProcesujIZapiszMail(ByVal oMail As Outlook.mailItem)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim nrProjektu As String
    Dim folderProjektu As String
    Dim atmt As Outlook.Attachment
    Dim sciezkaZapisu As String
    
    ' 1. Ekstrakcja numeru projektu z tematu (szukamy 5 cyfr)
    nrProjektu = SzukajNumeruProjektu(oMail.Subject)
    
    If nrProjektu <> "" Then
        ' 2. ZnajdŸ folder projektu na dysku D:
        ' U¿ywamy funkcji, któr¹ napisaliœmy wczeœniej
        folderProjektu = FindFolderByPrefix("D:\DANE\", nrProjektu, fso)
        
        If folderProjektu <> "" Then
            sciezkaZapisu = folderProjektu & "\poczta_automatyczna\"
            If Not fso.FolderExists(sciezkaZapisu) Then fso.CreateFolder sciezkaZapisu
            
            ' 3. Zapisz za³¹czniki, jeœli pasuj¹ do s³ów kluczowych
            For Each atmt In oMail.Attachments
                ' Wykorzystujemy Twoj¹ funkcjê CzyToDokumentUmowny
                If CzyToDokumentUmowny(atmt.fileName) Then
                    atmt.SaveAsFile sciezkaZapisu & Format(oMail.ReceivedTime, "yyyy-mm-dd_HHmm_") & atmt.fileName
                End If
            Next atmt
        End If
    End If
End Sub

' Funkcja pomocnicza do wyci¹gania numeru z tematu
Private Function SzukajNumeruProjektu(ByVal temat As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\d{5}" ' Szuka dok³adnie 5 cyfr obok siebie
    
    If regEx.Test(temat) Then
        SzukajNumeruProjektu = regEx.Execute(temat)(0).Value
    Else
        SzukajNumeruProjektu = ""
    End If
End Function

