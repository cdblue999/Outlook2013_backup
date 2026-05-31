' Copyright (C) 2026 ZMS
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.
Attribute VB_Name = "Module2"
Option Explicit

Private Const APP_NAME As String = "OutlookProjectArchiver"
Private Const BASE_SEARCH_PATH As String = "D:\DANE\"

Sub EksportujUmowyDraft_Pro()
    Dim nrProjektu As String, domyslnyNr As String
    Dim fso As Object
    Dim outlookApp As Outlook.Application
    Dim ns As Outlook.NameSpace
    Dim items As Outlook.items
    Dim mail As Object
    Dim atmt As Outlook.Attachment
    Dim projectFolderPath As String, targetPath As String, safeFileName As String
    Dim i As Long, licznik As Integer
    
    ' 1. Pobranie ostatnio u�ywanego numeru (User Friendly)
    domyslnyNr = GetSetting(APP_NAME, "Settings", "LastProject", "")
    nrProjektu = InputBox("Podaj 5-cyfrowy numer projektu:", "Eksport Um�w DRAFT", domyslnyNr)
    
    If nrProjektu = "" Then Exit Sub
    If Not nrProjektu Like "#####" Then
        MsgBox "B��d: Numer projektu musi mie� dok�adnie 5 cyfr.", vbCritical, "B��d wej�cia"
        Exit Sub
    End If
    
    SaveSetting APP_NAME, "Settings", "LastProject", nrProjektu
    
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 2. Sprawdzenie dost�pno�ci zasob�w
    If Not fso.FolderExists(BASE_SEARCH_PATH) Then
        MsgBox "B��d: Dysk D: lub folder " & BASE_SEARCH_PATH & " jest od��czony.", vbCritical
        Exit Sub
    End If
    
    ' 3. Inteligentne szukanie folderu (np. 05400 - Zadanie X)
    projectFolderPath = FindFolderByPrefix(BASE_SEARCH_PATH, nrProjektu, fso)
    
    If projectFolderPath = "" Then
        projectFolderPath = BASE_SEARCH_PATH & nrProjektu
        If Not fso.FolderExists(projectFolderPath) Then fso.CreateFolder projectFolderPath
    End If
    
    targetPath = projectFolderPath & "\umowa_draft\"
    If Not fso.FolderExists(targetPath) Then fso.CreateFolder targetPath
    
    ' 4. Procesowanie Outlook (MAPI)
    Set outlookApp = New Outlook.Application
    Set ns = outlookApp.GetNamespace("MAPI")
    
    ' Filtr SQL - poprawiony format daty na ISO (bezpieczniejszy)
    ' Szuka maili od 1 lutego 2026
    Dim filter As String
    filter = "@SQL=(""urn:schemas:httpmail:subject"" LIKE '%" & nrProjektu & "%') AND " & _
             """urn:schemas:httpmail:datereceived"" >= '2026-02-01 00:00'"
             
    Set items = ns.GetDefaultFolder(olFolderInbox).items.Restrict(filter)
    licznik = 0
    
    For i = items.Count To 1 Step -1
        If TypeOf items.item(i) Is mailItem Then
            Set mail = items.item(i)
            
            For Each atmt In mail.Attachments
                If CzyToDokumentUmowny(atmt.fileName) Then
                    ' Nazewnictwo: RRRR-MM-DD_NazwaPliku (u�atwia sortowanie w D:\DANE)
                    safeFileName = Format(mail.ReceivedTime, "yyyy-mm-dd") & "_" & CleanFileName(atmt.fileName)
                    
                    If Not fso.FileExists(targetPath & safeFileName) Then
                        atmt.SaveAsFile targetPath & safeFileName
                        licznik = licznik + 1
                    End If
                End If
            Next atmt
        End If
    Next i
    
    MsgBox "Eksport zako�czony sukcesem!" & vbCrLf & _
           "Projekt: " & nrProjektu & vbCrLf & _
           "Nowych plik�w: " & licznik, vbInformation, "Sukces WI"
    Exit Sub

ErrorHandler:
    MsgBox "B��d krytyczny: " & Err.Description, vbCritical
End Sub

' --- FUNKCJA SZUKAJ�CA FOLDERU PO NUMERZE ---
Public Function FindFolderByPrefix(root As String, prefix As String, fso As Object) As String
    Dim subFolder As Object
    On Error Resume Next ' Na wypadek braku dost�pu do jakiego� folderu
    For Each subFolder In fso.GetFolder(root).SubFolders
        If Left(subFolder.Name, Len(prefix)) = prefix Then
            FindFolderByPrefix = subFolder.path
            Exit Function
        End If
    Next subFolder
    FindFolderByPrefix = ""
End Function

' --- FILTR DOKUMENT�W (Dodany klucz 'draft' i 'wniosek') ---
Public Function CzyToDokumentUmowny(fname As String) As Boolean
    Dim ext As String, n As String
    Dim klucze As Variant, k As Variant
    n = LCase(fname)
    If InStrRev(n, ".") = 0 Then Exit Function
    ext = Mid(n, InStrRev(n, "."))
    
    ' Tylko dokumenty edytowalne i PDF
    If Not (ext Like ".pdf" Or ext Like ".doc*" Or ext Like ".xls*") Then Exit Function
    
    klucze = Array("umowa", "podwykonaw", "aneks", "zlecenie", "kontrakt", "draft", "wniosek", "wdr")
    For Each k In klucze
        If InStr(n, k) > 0 Then
            CzyToDokumentUmowny = True
            Exit Function
        End If
    Next k
End Function

' --- CZYSZCZENIE NAZWY PLIKU ---
Private Function CleanFileName(fname As String) As String
    Dim v As Variant
    CleanFileName = fname
    For Each v In Array("/", "\", ":", "*", "?", """", "<", ">", "|")
        CleanFileName = Replace(CleanFileName, v, "_")
    Next v
End Function

