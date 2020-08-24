Attribute VB_Name = "fkApp"
Option Explicit

'
' Beschreibung:
' Liesst einen Stringwert (Text) aus einer INI-Datei und gibt diesen als Rueckgabewert zurueck.
'
' Parameter:
' lpApplicationName    Namen des Abschnittes in der INI-Datei.
'                      "VBNullString"-Zeichen liefert die Namen aller Sektionen.
'
' lpKeyName            Schluessel des zu lesenden Eintrags.
'                      "VBNullString"-Zeichen + ein existierender Abschnitt "lpApplicationName", liefert alle im Abschnitt gespeicherten Schluessel.
'
' nDefault             Vorgabe-/Standardwert, sofern kein Eintrag passender Eintrag vorhanden ist.
'
' lpReturnedString     Puffer, der den Rueckgabewert enthaelt (muss mit ausreichend Leerstellen gefuellt sein)
'
' nSize                Groesse des Puffers in Bytes.
'
' lpFileName           Pfadangabe der INI-Datei
'
' Rueckgabewert:
' Bei erfolgreichem Aufruf enthaelt der Rueckgabewert die Laenge des gelesenen Strings (Textes), andernfalls wird "0" zurueckgegeben.
'
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'
' Beschreibung:
' Diese Funktion schreibt einen Wert in eine INI-Datei oder loescht bestimmte Abschnitt der INI-Datei. Die INI-Datei wird automatisch erstellt, sollte sie noch nicht existieren..
'
' Parameter:
' lpApplicationName  Name des Abschnittes in der INI-Datei.
'                    "VBNullString"-Zeichen loescht alle Abschnitte der INI-Datei!
'
' lpKeyName          Schluessel des Eintrags, dessen Wert gespeichert werden soll.
'                    "VBNullString" und ein existierender Abschnitt "lpApplicationName", wird der gesamte Abschnitt aus der Datei geloescht.
'
' lpString           Wert, der im Abschnitt unter dem angegebenen Schluessel gespeichert werden soll. uebergibt man hier ein "VBNullString"-Zeichen so wird der Schluessel aus dem Abschnitt der INI-Datei geloescht.
'                    "VBNullString"-Zeichen => der Schluessel wird aus dem Abschnitt der INI-Datei geloescht.
'
' lpFileName         Pfadangabe der INI-Datei. Existiert diese Datei nicht, wird sie automatisch erstellt.
'
' Rueckgabewert:
' War der Funktionsaufruf erfolgreich, so ist der Rueckgabewert 1, andernfalls wird 0 als Rueckgabewert zurueckgeliefert.
'
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const FORMAT_MMYYYY = "mm.yyyy"
Public Const FORMAT_01 = "01."
Public Const FORMAT_DDMMYYYY = "dd.mm.yyyy"
Public Const FORMAT_DDMMYY = "dd.mm.yy"
Public Const FORMAT_FLOAT0 = "###,##0"
Public Const FORMAT_FLOAT2 = "###,##0.00"
Public Const FORMAT_FLOAT3 = "###,##0.000"
Public Const FORMAT_FLOAT5 = "###,##0.00000"

'################################################################################
'
Public Sub testWW()

Dim windgeschw             As Double
Dim windrichtung           As Double
Dim eigengeschw            As Double
Dim rwk                    As Double
Dim differenz_wind_und_rwk As Double
Dim sin_diff_wind_rwk      As Double
Dim sin_mal_wind_geschw    As Double
Dim luv_wink               As Double

    windgeschw = 12.964
    windrichtung = 200
    rwk = 63
    eigengeschw = 140
    
    windgeschw = 33
    windrichtung = 96
    rwk = 140
    eigengeschw = 90

    If (windrichtung > rwk) Then
    
        differenz_wind_und_rwk = windrichtung - rwk
        
    Else
    
        differenz_wind_und_rwk = rwk - windrichtung
        
    End If

    
    sin_diff_wind_rwk = Sin(differenz_wind_und_rwk)
    
    
    sin_mal_wind_geschw = sin_diff_wind_rwk * windgeschw
    
    luv_wink = sin_mal_wind_geschw / eigengeschw
   
    Call wl("windgeschw             =>" & windgeschw & "<")
    Call wl("windrichtung           =>" & windrichtung & "<")
    Call wl("eigengeschw            =>" & eigengeschw & "<")
    Call wl("rwk                    =>" & rwk & "<")
    Call wl("differenz_wind_und_rwk =>" & differenz_wind_und_rwk & "<")
    Call wl("sin_diff_wind_rwk      =>" & sin_diff_wind_rwk & "<")
    Call wl("sin_mal_wind_geschw    =>" & sin_mal_wind_geschw & "<")
    Call wl("luv_wink               =>" & luv_wink & "<  " & ArcSin(luv_wink))
         
    'Sin(Luv) = ( Sin( Differenz zwischen Windrichtung und Rwk ) * Windgeschwindigkeit) / Eigengeschwindigkeit
    
End Sub

'################################################################################
'
' Erstellt einen String mit mindestens der angegebenen Laenge.
' Ist "pString" laenger als "pLaenge", ist das Ergebnis entsprechend der Laenge von pString.
' Ist "pString" kuerzer als "pLaenge", wird mit Leerzeichen aufgefuellt
' Das Ergebnis beginnt mit "pString" und es wird nach links hin aufgefuellt.
'
' ? ">" & getFeldLinksMin( "1234567890123", 10 ) & "<" = >1234567890123<
' ? ">" & getFeldLinksMin( "",             -10 ) & "<" = ><
' ? ">" & getFeldLinksMin( "",              10 ) & "<" = >          <
'
' PARAMETER: pString        = die Ausgangszeichenfolge
' PARAMETER: pLaenge        = die Mindestlaenge an Zeichen des Ergebnisses
'
' RETURN : einen String mit mindestens angegebenen Laenge, welcher mit pString beginnt
'
Public Function getFeldLinksMin(pString As String, pMinBreite As Integer) As String

    '
    ' Pruefung: Stringlaenge kleiner Mindesbreite ?
    ' Ist der Eingabestring kuerzer als die Sollbreite, muss das Ergebnis bis
    ' zur definierten Mindestbreite aufgefuellt werden.
    '
    ' Ist der Eingabestring laenger als die Sollbreite, wird dem Aufrufer der
    ' Eingabestring zurueckgegeben.
    '
    If (Len(pString) < pMinBreite) Then

        getFeldLinksMin = pString & String(pMinBreite - Len(pString), " ")

    Else

        getFeldLinksMin = pString

    End If

End Function

'################################################################################
'
' Erstellt einen String mit mindestens der angegebenen Laenge und rechter Ausrichtung von pString.
' Ist "pString" laenger als "pLaenge", ist das Ergebnis entsprechend der Länge von pString.
' Ist "pString" kuerzer als "pLaenge", wird mit Leerzeichen aufgefuellt
'
' ? ">" & getFeldRechtsMin( "12345", 10 ) & "<"            =>     12345<
' ? ">" & getFeldRechtsMin( "123456789012345", 10 ) & "<"  =>123456789012345<
'
' PARAMETER: pString        = die Ausgangszeichenfolge
' PARAMETER: pLaenge        = die Mindestlaenge an Zeichen des Ergebnisses
'
' RETURN : einen String mit mindestens angegebenen Laenge, welcher auf pString endet
'
Public Function getFeldRechtsMin(pString As String, pLaenge As Integer) As String

    If (pLaenge > Len(pString)) Then

        getFeldRechtsMin = String(pLaenge - Len(pString), " ") & pString

    Else

        getFeldRechtsMin = pString

    End If

End Function

'################################################################################
'
Public Function getFeldRechtsMinInteger(ByVal pIntegerZahl As Integer, pLaenge As Integer) As String

Dim str_integer As String

    str_integer = "" & pIntegerZahl
    
    If (pLaenge > Len(str_integer)) Then

        getFeldRechtsMinInteger = String(pLaenge - Len(str_integer), " ") & str_integer

    Else

        getFeldRechtsMinInteger = str_integer

    End If

End Function

'################################################################################
'
Public Function getMaxInteger(ByVal pZahl1 As Integer, ByVal pZahl2 As Integer) As Integer

    If (pZahl1 > pZahl2) Then
    
        getMaxInteger = pZahl1
    
    Else
    
        getMaxInteger = pZahl2
        
    End If
    
End Function

'################################################################################
'
' Liefert die Anzahl Zeichen der Zahl fuer die String-Formatierung zurueck.
'
' <        10 = 1 Zeichen
' <       100 = 2 Zeichen
' <     1.000 = 3 Zeichen
' <    10.000 = 4 Zeichen
' <   100.000 = 6 Zeichen
' < 1.000.000 = 7 Zeichen
' sonst       8 Zeichen
'
'
Public Function getAnzahlStellen(ByVal pZahl As Integer) As Integer

    If (pZahl < 10) Then
        
        getAnzahlStellen = 1
    
    ElseIf (pZahl < 100) Then
        
        getAnzahlStellen = 2
    
    ElseIf (pZahl < 1000) Then
        
        getAnzahlStellen = 3
    
    ElseIf (pZahl < 10000) Then
        
        getAnzahlStellen = 4
    
    ElseIf (pZahl < 100000) Then
        
        getAnzahlStellen = 5
    
    ElseIf (pZahl < 1000000) Then
        
        getAnzahlStellen = 6
    
    ElseIf (pZahl < 10000000) Then
        
        getAnzahlStellen = 7
    
    ElseIf (pZahl < 100000000) Then
        
        getAnzahlStellen = 8
    
    ElseIf (pZahl < 1000000000) Then
        
        getAnzahlStellen = 9
    
    Else
        
        getAnzahlStellen = 10
    
    End If

End Function

'################################################################################
'
Public Function getZahlX(pEingabe As String, pVorgabe As Double) As Double

On Error GoTo errGetZahl

'Print ( 2.34 = CDbl( "2.34" ) ), CDbl( "2.34" )  = Falsch 234
'Print ( 2.34 = CDbl( "2,34" ) ), CDbl( "2,34" )  = Wahr   2,34

Dim fkt_ergebnis         As Double
Dim pos_suchtext         As Integer
Dim such_trennzeichen    As String
Dim dezimal_trennzeichen As String

    fkt_ergebnis = pVorgabe

    '
    ' Pruefung: Eingabestring gesetzt ?
    '
    ' Ist die Eingabe nicht gesetzt (getrimmt ein Leerstring), bekommt der Aufrufer
    ' den Wert aus "pVorgabe" zurueck.
    '
    If (Trim(pEingabe) <> "") Then

        '
        ' Ermittlung des Dezimal-Trennzeichens
        '
        ' Es muss ermittelt werden, welche Konvertierung den Double-Wert 2.34 ergibt.
        ' Die Konvertierung mittels der Funktion CDbl ist Laenderabhaengig, d.h. die
        ' Funktion CDbl nimmt einmal das Komma als Dezimaltrennzeichen, ein anderes
        ' mal den Punkt.
        '
        ' Print ( 2.34 = CDbl( "2.34" ) ), CDbl( "2.34" )  = Falsch 234
        '
        ' Print ( 2.34 = CDbl( "2,34" ) ), CDbl( "2,34" )  = Wahr   2,34
        '
        If (2.34 = CDbl("2.34")) Then

            such_trennzeichen = ","

            dezimal_trennzeichen = "."

        ElseIf (2.34 = CDbl("2,34")) Then

            such_trennzeichen = "."

            dezimal_trennzeichen = ","

        End If

        '
        ' Ermittlung der Position des ermittelten Dezimaltrennzeichens
        '
        pos_suchtext = InStr(pEingabe, such_trennzeichen)

        '
        ' Pruefung: Ist in der Eingabe ein Tausendertrennzeichen vorhanden?
        '
        ' Ist in der Eingabe ein Tausendertrennzeichen vorhanden, muss die
        ' Eingabe auf das korrekte Tausendertrennzeichen geaendert werden.
        '
        If (pos_suchtext > 0) Then

            pEingabe = Left(pEingabe, pos_suchtext - 1) & dezimal_trennzeichen & Mid(pEingabe, pos_suchtext + 1)

        End If

        '
        ' Eingabe in einen Doublewert umwandeln
        '
        fkt_ergebnis = CDbl("0" & Trim(pEingabe))

    End If

EndFunktion:

    On Error Resume Next

    DoEvents

    getZahlX = fkt_ergebnis

    Exit Function

errGetZahl:

    fkt_ergebnis = pVorgabe

    Resume EndFunktion

End Function

'################################################################################
'
Public Function getInteger(pBetrag As String, pVorgabe As Integer) As String

On Error GoTo errGetInteger

    getInteger = CInt(pBetrag)

    Exit Function

errGetInteger:

    getInteger = pVorgabe

    Exit Function

End Function

'################################################################################
'
Public Function getDouble(pBetrag As String, pVorgabe As Integer) As Double

On Error GoTo errGetDouble

    getDouble = getZahlX(pBetrag, 0#)

    Exit Function

errGetDouble:

    getDouble = pVorgabe

    Exit Function

End Function

'################################################################################
'
Public Function get5nk(pBetrag As Double) As Double

    get5nk = CDbl(Fix(pBetrag * 100000#) * 0.00001)

End Function

'################################################################################
'
' getStringStundenMinuten(   45 ) = 00:45
' getStringStundenMinuten(   60 ) = 01:00
' getStringStundenMinuten(   75 ) = 01:15
'
' getStringStundenMinuten(  234 ) = 03:54
' getStringStundenMinuten( -234 ) = 00:00
'
Public Function getStringStundenMinuten(ByVal pMinuten As Long) As String

On Error Resume Next

    If (pMinuten > 0) Then
    
        Dim anzahl_stunden As Long
        
        anzahl_stunden = Fix((pMinuten \ 60))

        getStringStundenMinuten = IIf(anzahl_stunden < 10, "0", "") & anzahl_stunden & ":" & Right("00" & (pMinuten - (anzahl_stunden * 60)), 2)
        
    Else
    
        getStringStundenMinuten = "00:00"
        
    End If

End Function

'################################################################################
'
' Inverse Sin
'
Function ArcSin(X As Double) As Double

    ArcSin = Atn(X / Sqr(-X * X + 1))
    
End Function

'################################################################################
'
' Inverse Cosin
'
Function ArcCos(X As Double) As Double

    ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    
End Function

'################################################################################
'
' PARAMETER: pIniDateiName  = Pfad und Name der Ini-Datei
' PARAMETER: pSection       = Sektion der Ini-Datei
' PARAMETER: pSchluesselName = der Schluessel aus der Sektion
' PARAMETER: pVorgabeWert   = der Vorgabewert, wenn der Schluessel nicht vorhanden ist
'
' RETURN : den Wert des angegebenen Schluessels, oder den Vorgabewert wenn der Schluessel nicht existiert
'
Public Function readIniDateiText(pIniDateiName As String, pSection As String, pSchluesselName As String, pVorgabeWert As String) As String

Dim return_string As String

    return_string = String(255, Chr(0))

    readIniDateiText = Left(return_string, GetPrivateProfileString(pSection, ByVal pSchluesselName, pVorgabeWert, return_string, Len(return_string), pIniDateiName))

End Function

'################################################################################
'
' PARAMETER: pSection       = Sektion der Ini-Datei
' PARAMETER: pSchluesselName = der Schluessel aus der Sektion
' PARAMETER: pVorgabeWert   = der Vorgabewert, wenn der Schluessel nicht vorhanden ist
'
' RETURN : den Wert des angegebenen Schluessels, oder den Vorgabewert wenn der Schluessel nicht existiert
'
Public Function readIniText(pSection As String, pSchluesselName As String, pVorgabeWert As String) As String

    readIniText = readIniDateiText(getAnwIniDateiName(), pSection, pSchluesselName, pVorgabeWert)

End Function

'################################################################################
'
' Liest einen boolschen Wert aus der INI-Datei.
'
' PARAMETER: pSection       = Sektion der Ini-Datei
' PARAMETER: pSchluesselName = der Schluessel aus der Sektion
' PARAMETER: pVorgabeWert   = der boolsche Vorgabewert, wenn der Schluessel nicht vorhanden ist
'
' RETURN : TRUE wenn der Ini-Wert "true, wahr, 1" ist, sonst false
'
Public Function readIniBoolean(pSection As String, pSchluesselName As String, pVorgabeWert As Boolean) As String

Dim wert_aus_ini_datei As String

    '
    'HINWEIS: In VB sind die boolschen Werte eingedeutsch worden!
    ' ? true  = Wahr
    ' ? false = Falsch
    '
    ' Der Schluessel wird aus der INI-Datei gelesen und in Kleinbuchstaben gewandelt.
    '
    wert_aus_ini_datei = LCase(readIniDateiText(getAnwIniDateiName(), pSection, pSchluesselName, "" & pVorgabeWert))

    '
    ' Der so gelesene INI-Wert wird gegen die Strings "true", "wahr" und "1" geprueft.
    ' Ist der INI-Wert gleich einem der Pruefwoerter wird das Ergebnis auf TRUE gestellt.
    ' Ist der INI-Wert ungleich dieser Woerter, wird das Funktionsergebnis auf FALSE gestellt.
    '
    If ((wert_aus_ini_datei = "true") Or (wert_aus_ini_datei = "wahr") Or (wert_aus_ini_datei = "1")) Then

        readIniBoolean = True

    Else

        readIniBoolean = False

    End If

End Function

'################################################################################
'
' Liest einen Integer Wert aus der INI-Datei.
'
' PARAMETER: pSection       = Sektion der Ini-Datei
' PARAMETER: pSchluesselName = der Schluessel aus der Sektion
' PARAMETER: pVorgabeWert   = der Vorgabewert, wenn der Schluessel nicht vorhanden ist
'
' RETURN : den aus der Ini gelesenen Wert, oder den Vorgabewert
'
Public Function readIniInteger(pSection As String, pSchluesselName As String, pVorgabeWert As Integer) As String

On Error GoTo errReadIniInteger

    readIniInteger = CInt(readIniDateiText(getAnwIniDateiName(), pSection, pSchluesselName, "" & pVorgabeWert))

    Exit Function

errReadIniInteger:

    readIniInteger = pVorgabeWert

    Exit Function

End Function

'################################################################################
'
' PARAMETER: pIniDateiName  = Pfad und Name der Ini-Datei
' PARAMETER: pSection       = Sektion der Ini-Datei
' PARAMETER: pSchluesselName = der Schluessel aus der Sektion
' PARAMETER: pValue         = der zu schreibende Wert
'
' RETURN : 1 bei erfolgreicher Speicherung, sonst 0
'
Public Function writeIniDateiText(pIniDateiName As String, pSection As String, pSchluesselName As String, pValue As String) As Long

    writeIniDateiText = WritePrivateProfileString(pSection, pSchluesselName, pValue, pIniDateiName)

End Function

'################################################################################
'
' PARAMETER: pSection       = Sektion der Ini-Datei
' PARAMETER: pSchluesselName = der Schluessel aus der Sektion
' PARAMETER: pValue         = der zu schreibende Wert
'
' RETURN : 1 bei erfolgreicher Speicherung, sonst 0
'
Public Function writeIniText(pSection As String, pSchluesselName As String, pValue As String) As Long

    writeIniText = WritePrivateProfileString(pSection, pSchluesselName, pValue, getAnwIniDateiName())

End Function

'################################################################################
'
' PARAMETER: pSection       = Sektion der Ini-Datei
' PARAMETER: pSchluesselName = der Schluessel aus der Sektion
' PARAMETER: pValue         = der zu schreibende Wert
'
' RETURN : 1 bei erfolgreicher Speicherung, sonst 0
'
Public Function writeIniBoolean(pSection As String, pSchluesselName As String, pValue As Boolean) As Long

    '
    'HINWEIS: In VB sind die boolschen Werte eingedeutsch worden!
    ' ? true  = Wahr
    ' ? false = Falsch
    '
    writeIniBoolean = writeIniText(pSection, pSchluesselName, IIf(pValue, "true", "false"))

End Function

'################################################################################
'
' RETURN : den Pfad und Dateinamen auf die INI-Datei der Anwendung
'
Public Function getAnwIniDateiName() As String

    If (Right(App.Path, 1) = "\") Then

        getAnwIniDateiName = App.Path & "datei_name.ini"

    Else

        getAnwIniDateiName = App.Path & "\datei_name.ini"

    End If

End Function

'################################################################################
'
Public Sub wl(pString As String)

    Debug.Print pString

End Sub

