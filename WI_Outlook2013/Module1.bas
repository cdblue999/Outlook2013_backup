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
        ' 2. Znajd� folder projektu na dysku D:
        ' U�ywamy funkcji, kt�r� napisali�my wcze�niej
        folderProjektu = FindFolderByPrefix("D:\DANE\", nrProjektu, fso)
        
        If folderProjektu <> "" Then
            sciezkaZapisu = folderProjektu & "\poczta_automatyczna\"
            If Not fso.FolderExists(sciezkaZapisu) Then fso.CreateFolder sciezkaZapisu
            
            ' 3. Zapisz za��czniki, je�li pasuj� do s��w kluczowych
            For Each atmt In oMail.Attachments
                ' Wykorzystujemy Twoj� funkcj� CzyToDokumentUmowny
                If CzyToDokumentUmowny(atmt.fileName) Then
                    atmt.SaveAsFile sciezkaZapisu & Format(oMail.ReceivedTime, "yyyy-mm-dd_HHmm_") & atmt.fileName
                End If
            Next atmt
        End If
    End If
End Sub

' Funkcja pomocnicza do wyci�gania numeru z tematu
Private Function SzukajNumeruProjektu(ByVal temat As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\d{5}" ' Szuka dok�adnie 5 cyfr obok siebie
    
    If regEx.Test(temat) Then
        SzukajNumeruProjektu = regEx.Execute(temat)(0).Value
    Else
        SzukajNumeruProjektu = ""
    End If
End Function

