VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSzukajWI 
   Caption         =   "Wyszukiwarka Korespondencji WI"
   ClientHeight    =   8640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8150
   OleObjectBlob   =   "frmSzukajWI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSzukajWI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const APP_NAME As String = "OutlookWI_Search"

' --- DEKLARACJA API DO OTWIERANIA STRON WWW ---
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If
' ----------------------------------------------

Private Sub UserForm_Initialize()
    ' Odczyt pól tekstowych z rejestru
    Me.txtNumer.Value = GetSetting(APP_NAME, "LastSearch", "Numer", "")
    Me.txtNadawca.Value = GetSetting(APP_NAME, "LastSearch", "Nadawca", "")
    Me.txtData.Value = GetSetting(APP_NAME, "LastSearch", "Data", "")
    Me.txtDotyczy.Value = GetSetting(APP_NAME, "LastSearch", "Dotyczy", "")
    Me.txtKomentarz.Value = GetSetting(APP_NAME, "LastSearch", "Komentarz", "")
    
    ' --- NOWOŒÆ: Checkboxy domyœlnie zaznaczone na True ---
    Me.chkSzukajOutlook.Value = CBool(GetSetting(APP_NAME, "LastSearch", "ChkOutlook", "True"))
    Me.chkSzukajSerwer.Value = CBool(GetSetting(APP_NAME, "LastSearch", "ChkSerwer", "True"))
    
    ' --- Odczyt listy œcie¿ek ---
    Dim zapisaneSciezki As String
    Dim tablicaSciezek() As String
    Dim s As Variant
    
    zapisaneSciezki = GetSetting(APP_NAME, "LastSearch", "SciezkiList", "")
    
    If zapisaneSciezki <> "" Then
        tablicaSciezek = Split(zapisaneSciezki, "|")
        Me.lstSciezki.Clear
        For Each s In tablicaSciezek
            If Trim(CStr(s)) <> "" Then Me.lstSciezki.AddItem CStr(s)
        Next s
    End If
End Sub
Private Sub btnSzukaj_Click()
    ' 1. Walidacja
    If Trim(Me.txtNumer.Value & Me.txtNadawca.Value & Me.txtDotyczy.Value) = "" Then
        MsgBox "Podaj przynajmniej jedno kryterium.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Pakowanie œcie¿ek z listy do jednego ci¹gu tekstowego
    Dim i As Integer
    Dim doZapisu As String
    Set SciezkiSerwer = New Collection
    
    doZapisu = ""
    For i = 0 To Me.lstSciezki.ListCount - 1
        SciezkiSerwer.Add Me.lstSciezki.List(i)
        doZapisu = doZapisu & Me.lstSciezki.List(i) & IIf(i < Me.lstSciezki.ListCount - 1, "|", "")
    Next i
    
    ' 3. ZAPISYWANIE WSZYSTKIEGO DO REJESTRU
    SaveSetting APP_NAME, "LastSearch", "Numer", Me.txtNumer.Value
    SaveSetting APP_NAME, "LastSearch", "Nadawca", Me.txtNadawca.Value
    SaveSetting APP_NAME, "LastSearch", "Data", Me.txtData.Value
    SaveSetting APP_NAME, "LastSearch", "Dotyczy", Me.txtDotyczy.Value
    SaveSetting APP_NAME, "LastSearch", "Komentarz", Me.txtKomentarz.Value
    SaveSetting APP_NAME, "LastSearch", "SciezkiList", doZapisu
    
    ' --- NOWOŒÆ: Zapisywanie stanu checkboxów ---
    SaveSetting APP_NAME, "LastSearch", "ChkOutlook", CStr(Me.chkSzukajOutlook.Value)
    SaveSetting APP_NAME, "LastSearch", "ChkSerwer", CStr(Me.chkSzukajSerwer.Value)
    
    ' 4. Przygotowanie kluczy dla modu³u g³ównego
    Set SzukaneKlucze = New Collection
    If CleanInput(Me.txtNumer.Value) <> "" Then SzukaneKlucze.Add CleanInput(Me.txtNumer.Value)
    If CleanInput(Me.txtNadawca.Value) <> "" Then SzukaneKlucze.Add CleanInput(Me.txtNadawca.Value)
    If CleanInput(Me.txtDotyczy.Value) <> "" Then SzukaneKlucze.Add CleanInput(Me.txtDotyczy.Value)
    If CleanInput(Me.txtKomentarz.Value) <> "" Then SzukaneKlucze.Add CleanInput(Me.txtKomentarz.Value)
    
    CzySzukacOutlook = Me.chkSzukajOutlook.Value
    CzySzukacSerwer = Me.chkSzukajSerwer.Value
    SzukanaDataRaw = CleanInput(Me.txtData.Value)
    
    Me.Hide
End Sub

Private Sub btnAnuluj_Click()
    Set SzukaneKlucze = New Collection
    Set SciezkiSerwer = New Collection
    SzukanaDataRaw = ""
    Me.Hide
End Sub

' -------------------------------------------------------------------------
' POPRAWIONY KOD OBS£UGI LISTY ŒCIE¯EK DLA OUTLOOKA
' -------------------------------------------------------------------------
Private Sub btnAddPath_Click()
    Dim objShell As Object
    Dim objFolder As Object
    
    Set objShell = CreateObject("Shell.Application")
    ' Wywo³anie natywnego okna wyboru folderu systemu Windows
    Set objFolder = objShell.BrowseForFolder(0, "Wybierz folder do przeszukiwania na serwerze:", 0, 0)
    
    If Not objFolder Is Nothing Then
        Me.lstSciezki.AddItem objFolder.Self.path
    End If
End Sub

Private Sub btnRemovePath_Click()
    Dim i As Integer
    ' Pêtla od ty³u, ¿eby nie zgubiæ indeksów przy usuwaniu
    For i = Me.lstSciezki.ListCount - 1 To 0 Step -1
        If Me.lstSciezki.Selected(i) Then
            Me.lstSciezki.RemoveItem i
        End If
    Next i
End Sub
' -------------------------------------------------------------------------

Private Function CleanInput(val As String) As String
    Dim temp As String
    temp = Trim(val)
    temp = Replace(temp, vbCr, " ")
    temp = Replace(temp, vbLf, " ")
    Do While InStr(temp, "  ") > 0
        temp = Replace(temp, "  ", " ")
    Loop
    If Replace(Replace(temp, "*", ""), "?", "") = "" Then temp = ""
    CleanInput = temp
End Function

' -------------------------------------------------------------------------
' FUNKCJA ANIMUJ¥CA PASEK POSTÊPU
' -------------------------------------------------------------------------
Public Sub AktualizujPostep(procent As Single, tekst As String)
    On Error Resume Next
    
    ' Jeœli formularz jest z jakiegoœ powodu ukryty, upewnij siê, ¿e jest widoczny bez blokowania kodu
    If Me.Visible = False Then Me.Show vbModeless
    
    ' Maksymalna szerokoœæ to szerokoœæ etykiety pe³ni¹cej rolê t³a
    Dim maxWidth As Single
    maxWidth = Me.lblPrgBg.Width
    
    ' Zabezpieczenie przed wartoœciami z kosmosu
    If procent < 0 Then procent = 0
    If procent > 1 Then procent = 1
    
    ' Zmiana szerokoœci kolorowego paska
    Me.lblPrgBar.Width = maxWidth * procent
    
    ' Aktualizacja tekstu (opcjonalnie)
    Me.lblPrgText.Caption = tekst & " (" & Format(procent, "0%") & ")"
    
    ' BARDZO WA¯NE: Wymusza odœwie¿enie grafiki okna. Bez tego pasek siê "zawiesi" do koñca pêtli.
    DoEvents
End Sub

Private Sub WI_logo_Click()
    Dim url As String
    url = "https://www.wi.wroc.pl/"
    
    ShellExecute 0, "open", url, "", "", 1
End Sub
