Attribute VB_Name = "modHelp"
Option Explicit
' Module    : modHelp
' Purpose   : User-facing help content in 13 languages.
'             Currently covers the Autodesk Apprentice Server installation
'             warning: short MsgBox body + full README written to sandbox.
' Reference : Appendix A.13, A.61
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Documentation accuracy + header completion):
'   Added missing Requires section (U() from modLang, PRJ_PutClipboard
'   from modClipboard, PRJ_EnsureSandbox / PRJ_SandboxPath from
'   modSandbox, PRJ_FSO from modCOM).
'   Completed the truncated v9.13.5 changelog entry (was left as a
'   bare line with no content).
'   Moved Notes section to the end of the header per Style B
'   convention; extended it to document the line-continuation
'   policy (this module contains no line continuations).
'   No functional changes.
'
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.19.3 (Header sync):
'   Version bumped to match PRJ_VERSION. Footer added.
'   No code changes.
'
' SIZE NOTE: This module exceeds the 800 LOC guideline (section 1.10)
' and is documented as a justified exception.
'
' Justification:
'   1. Single responsibility: multi-language help text content.
'   2. Content is prose, not logic. Splitting into 13 per-language
'      modules would trigger over-modularization (section 1.1).
'   3. Analogous to external help files (e.g. .hlp or .chm) in .NET.
'   4. Adding a new language = adding one Private Function.
'   5. Navigation within the module is by language code, not by order.
'
' This is a deliberate architectural choice, not debt.
'
' v9.13.5 (ARCH-26 - §4.3 / §4.4 / §4.5):
'   HELP_WriteApprenticeREADME uses FSO Unicode:=True (UTF-16LE with
'   BOM) instead of bare Print # (ANSI codepage). Prior code wrote
'   in CP1252 on EN Windows, corrupting JP/KO/ZH/RU/PL/CS content.
'   Notepad, browsers, and Notepad++ all handle UTF-16LE correctly
'   on every locale. The behavior is documented inline in the
'   HELP_WriteApprenticeREADME procedure.
'
' Requires  : U() (modLang).
'             PRJ_PutClipboard (modClipboard).
'             PRJ_EnsureSandbox, PRJ_SandboxPath (modSandbox).
'             PRJ_FSO (modCOM).
'
' Notes     : Kept separate from modLang so the localization engine stays a flat
'             key->string table without multi-line vbCrLf-laden prose.
'             Clipboard service (PRJ_PutClipboard) lives in modClipboard.
'             Two public entry points:
'               HELP_ApprenticeShortMsg(sLang)      MsgBox body (4-step summary)
'               HELP_WriteApprenticeREADME(sLang)   write full guide to sandbox,
'                                                   return path ("" on failure)
'             sLang is the 2-or-3 letter code from LANG_ActiveCode():
'               en, de, fr, it, es, pt, ru, pl, cs, ja, ko, zh, zht
'             Unknown codes fall back to EN.
'             Non-ASCII characters are wrapped in U("...\uXXXX...") to survive
'             any editor save-encoding. U() lives in modLang.
'             Log files stay English regardless of UI language (A.66); this
'             module only affects user-facing dialogs and README files.
'             No line-continuation characters in this module.

' Installer and download constants (single-line, no continuations)
Private Const APP_REG_PATH_2027 As String = "C:\Program Files\Autodesk\Inventor Apprentice Server 2027\Bin\ApprenticeRegSvr.exe"
Private Const APP_REG_PATH_2026 As String = "C:\Program Files\Autodesk\Inventor Apprentice Server 2026\Bin\ApprenticeRegSvr.exe"
Private Const APP_DOWNLOAD_URL As String = "https://www.autodesk.com/support/technical/article/caas/sfdcarticles/sfdcarticles/Inventor-Apprentice-Server.html"

' Public entry points

' HELP_ApprenticeShortMsg - MsgBox body. Fits a standard vbInformation dialog.
Public Function HELP_ApprenticeShortMsg(ByVal sLang As String) As String
    Select Case LCase$(sLang)
        Case "de":  HELP_ApprenticeShortMsg = ApprenticeShort_DE()
        Case "fr":  HELP_ApprenticeShortMsg = ApprenticeShort_FR()
        Case "it":  HELP_ApprenticeShortMsg = ApprenticeShort_IT()
        Case "es":  HELP_ApprenticeShortMsg = ApprenticeShort_ES()
        Case "pt":  HELP_ApprenticeShortMsg = ApprenticeShort_PT()
        Case "ru":  HELP_ApprenticeShortMsg = ApprenticeShort_RU()
        Case "pl":  HELP_ApprenticeShortMsg = ApprenticeShort_PL()
        Case "cs":  HELP_ApprenticeShortMsg = ApprenticeShort_CS()
        Case "ja":  HELP_ApprenticeShortMsg = ApprenticeShort_JA()
        Case "ko":  HELP_ApprenticeShortMsg = ApprenticeShort_KO()
        Case "zh":  HELP_ApprenticeShortMsg = ApprenticeShort_ZH()
        Case "zht": HELP_ApprenticeShortMsg = ApprenticeShort_ZHT()
        Case Else:  HELP_ApprenticeShortMsg = ApprenticeShort_EN()
    End Select
End Function

' HELP_ApprenticeREADMEBody - full guide as a single string.
Public Function HELP_ApprenticeREADMEBody(ByVal sLang As String) As String
    Select Case LCase$(sLang)
        Case "de":  HELP_ApprenticeREADMEBody = ApprenticeReadme_DE()
        Case "fr":  HELP_ApprenticeREADMEBody = ApprenticeReadme_FR()
        Case "it":  HELP_ApprenticeREADMEBody = ApprenticeReadme_IT()
        Case "es":  HELP_ApprenticeREADMEBody = ApprenticeReadme_ES()
        Case "pt":  HELP_ApprenticeREADMEBody = ApprenticeReadme_PT()
        Case "ru":  HELP_ApprenticeREADMEBody = ApprenticeReadme_RU()
        Case "pl":  HELP_ApprenticeREADMEBody = ApprenticeReadme_PL()
        Case "cs":  HELP_ApprenticeREADMEBody = ApprenticeReadme_CS()
        Case "ja":  HELP_ApprenticeREADMEBody = ApprenticeReadme_JA()
        Case "ko":  HELP_ApprenticeREADMEBody = ApprenticeReadme_KO()
        Case "zh":  HELP_ApprenticeREADMEBody = ApprenticeReadme_ZH()
        Case "zht": HELP_ApprenticeREADMEBody = ApprenticeReadme_ZHT()
        Case Else:  HELP_ApprenticeREADMEBody = ApprenticeReadme_EN()
    End Select
End Function

' HELP_WriteApprenticeREADME - write full guide to
'   %TEMP%\VBADR\README_Apprentice_<LANG>.txt
'   Returns file path on success, "" on any failure. Never raises.
'
' v9.13.5 (ARCH-26): uses FSO Unicode:=True (UTF-16LE with BOM) instead
'   of bare Print # (ANSI codepage). Prior code wrote in CP1252 on EN
'   Windows, corrupting JP/KO/ZH/RU/PL/CS content. Notepad, browsers,
'   and Notepad++ all handle UTF-16LE correctly on every locale.
Public Function HELP_WriteApprenticeREADME(ByVal sLang As String) As String
    Dim sBody As String
    Dim sPath As String
    Dim f As Object

    HELP_WriteApprenticeREADME = ""

    On Error GoTo Fail

    PRJ_EnsureSandbox
    sBody = HELP_ApprenticeREADMEBody(sLang)
    sPath = PRJ_SandboxPath() & "\README_Apprentice_" & UCase$(sLang) & ".txt"

    Set f = PRJ_FSO().CreateTextFile(sPath, True, True)
    If f Is Nothing Then Exit Function
    f.Write sBody
    f.Close
    Set f = Nothing

    HELP_WriteApprenticeREADME = sPath
    Exit Function

Fail:
    On Error Resume Next
    If Not f Is Nothing Then f.Close
    Set f = Nothing
    HELP_WriteApprenticeREADME = ""
    On Error GoTo 0
End Function

' Short MsgBox bodies - 13 languages

Private Function ApprenticeShort_EN() As String
    Dim s As String
    s = "Apprentice Server is not registered on this machine." & vbCrLf & vbCrLf
    s = s & "Inventor 2026 and later require a separate Apprentice Server installation." & vbCrLf
    s = s & "Without it, batch processing uses the full Inventor engine (10-50x slower)." & vbCrLf & vbCrLf
    s = s & "Quick install:" & vbCrLf
    s = s & "1. Download from: " & APP_DOWNLOAD_URL & vbCrLf
    s = s & "2. Run the installer (standard user; admin rights only for setup)." & vbCrLf
    s = s & "3. Register from a normal Command Prompt:" & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & "4. Restart Inventor." & vbCrLf & vbCrLf
    s = s & "See the full guide at:"
    ApprenticeShort_EN = s
End Function

Private Function ApprenticeShort_DE() As String
    Dim s As String
    s = U("Apprentice Server ist auf diesem Rechner nicht registriert.") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 und h\u00f6her erfordern eine separate Apprentice-Server-Installation.") & vbCrLf
    s = s & U("Ohne sie verwendet die Stapelverarbeitung die vollst\u00e4ndige Inventor-Engine (10-50x langsamer).") & vbCrLf & vbCrLf
    s = s & U("Schnellinstallation:") & vbCrLf
    s = s & U("1. Herunterladen von: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. Installationsprogramm ausf\u00fchren (Standardbenutzer; Admin-Rechte nur f\u00fcr Setup).") & vbCrLf
    s = s & U("3. Registrieren \u00fcber eine normale Eingabeaufforderung:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Inventor neu starten.") & vbCrLf & vbCrLf
    s = s & U("Vollst\u00e4ndige Anleitung unter:")
    ApprenticeShort_DE = s
End Function

Private Function ApprenticeShort_FR() As String
    Dim s As String
    s = U("Apprentice Server n'est pas enregistr\u00e9 sur cette machine.") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 et versions ult\u00e9rieures n\u00e9cessitent une installation s\u00e9par\u00e9e d'Apprentice Server.") & vbCrLf
    s = s & U("Sans cela, le traitement par lots utilise le moteur Inventor complet (10 \u00e0 50x plus lent).") & vbCrLf & vbCrLf
    s = s & U("Installation rapide :") & vbCrLf
    s = s & U("1. T\u00e9l\u00e9charger depuis : ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. Ex\u00e9cuter l'installateur (utilisateur standard ; droits admin uniquement pour l'installation).") & vbCrLf
    s = s & U("3. Enregistrer depuis une invite de commandes normale :") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Red\u00e9marrer Inventor.") & vbCrLf & vbCrLf
    s = s & U("Voir le guide complet ici :")
    ApprenticeShort_FR = s
End Function

Private Function ApprenticeShort_IT() As String
    Dim s As String
    s = U("Apprentice Server non \u00e8 registrato su questo computer.") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 e versioni successive richiedono un'installazione separata di Apprentice Server.") & vbCrLf
    s = s & U("Senza di esso, l'elaborazione batch usa il motore Inventor completo (10-50 volte pi\u00f9 lento).") & vbCrLf & vbCrLf
    s = s & U("Installazione rapida:") & vbCrLf
    s = s & U("1. Scarica da: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. Esegui l'installer (utente standard; privilegi admin solo per il setup).") & vbCrLf
    s = s & U("3. Registra da un normale Prompt dei comandi:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Riavvia Inventor.") & vbCrLf & vbCrLf
    s = s & U("Vedi la guida completa qui:")
    ApprenticeShort_IT = s
End Function

Private Function ApprenticeShort_ES() As String
    Dim s As String
    s = U("Apprentice Server no est\u00e1 registrado en este equipo.") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 y versiones posteriores requieren una instalaci\u00f3n separada de Apprentice Server.") & vbCrLf
    s = s & U("Sin \u00e9l, el procesamiento por lotes usa el motor completo de Inventor (10-50x m\u00e1s lento).") & vbCrLf & vbCrLf
    s = s & U("Instalaci\u00f3n r\u00e1pida:") & vbCrLf
    s = s & U("1. Descargue desde: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. Ejecute el instalador (usuario est\u00e1ndar; privilegios de admin solo para la instalaci\u00f3n).") & vbCrLf
    s = s & U("3. Registre desde un S\u00edmbolo del sistema normal:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Reinicie Inventor.") & vbCrLf & vbCrLf
    s = s & U("Consulte la gu\u00eda completa en:")
    ApprenticeShort_ES = s
End Function

Private Function ApprenticeShort_PT() As String
    Dim s As String
    s = U("O Apprentice Server n\u00e3o est\u00e1 registrado nesta m\u00e1quina.") & vbCrLf & vbCrLf
    s = s & U("O Inventor 2026 e vers\u00f5es posteriores exigem uma instala\u00e7\u00e3o separada do Apprentice Server.") & vbCrLf
    s = s & U("Sem ele, o processamento em lote usa o motor Inventor completo (10-50x mais lento).") & vbCrLf & vbCrLf
    s = s & U("Instala\u00e7\u00e3o r\u00e1pida:") & vbCrLf
    s = s & U("1. Baixe de: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. Execute o instalador (usu\u00e1rio padr\u00e3o; privil\u00e9gios de admin apenas para o setup).") & vbCrLf
    s = s & U("3. Registre a partir de um Prompt de Comando normal:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Reinicie o Inventor.") & vbCrLf & vbCrLf
    s = s & U("Consulte o guia completo em:")
    ApprenticeShort_PT = s
End Function

Private Function ApprenticeShort_RU() As String
    Dim s As String
    s = U("\u041f\u0440\u0438\u043b\u043e\u0436\u0435\u043d\u0438\u0435 Apprentice Server \u043d\u0435 \u0437\u0430\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0438\u0440\u043e\u0432\u0430\u043d\u043e \u043d\u0430 \u044d\u0442\u043e\u043c \u043a\u043e\u043c\u043f\u044c\u044e\u0442\u0435\u0440\u0435.") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 \u0438 \u0431\u043e\u043b\u0435\u0435 \u043f\u043e\u0437\u0434\u043d\u0438\u0435 \u0432\u0435\u0440\u0441\u0438\u0438 \u0442\u0440\u0435\u0431\u0443\u044e\u0442 \u043e\u0442\u0434\u0435\u043b\u044c\u043d\u043e\u0439 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438 Apprentice Server.") & vbCrLf
    s = s & U("\u0411\u0435\u0437 \u043d\u0435\u0433\u043e \u043f\u0430\u043a\u0435\u0442\u043d\u0430\u044f \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0430 \u0438\u0441\u043f\u043e\u043b\u044c\u0437\u0443\u0435\u0442 \u043f\u043e\u043b\u043d\u044b\u0439 \u0434\u0432\u0438\u0436\u043e\u043a Inventor (\u0432 10-50 \u0440\u0430\u0437 \u043c\u0435\u0434\u043b\u0435\u043d\u043d\u0435\u0435).") & vbCrLf & vbCrLf
    s = s & U("\u0411\u044b\u0441\u0442\u0440\u0430\u044f \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430:") & vbCrLf
    s = s & U("1. \u0421\u043a\u0430\u0447\u0430\u0442\u044c \u0441: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. \u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u044c \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a (\u043e\u0431\u044b\u0447\u043d\u044b\u0439 \u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u044c; \u043f\u0440\u0430\u0432\u0430 \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0430\u0442\u043e\u0440\u0430 \u0442\u043e\u043b\u044c\u043a\u043e \u0434\u043b\u044f \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438).") & vbCrLf
    s = s & U("3. \u0417\u0430\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0438\u0440\u043e\u0432\u0430\u0442\u044c \u0438\u0437 \u043e\u0431\u044b\u0447\u043d\u043e\u0439 \u043a\u043e\u043c\u0430\u043d\u0434\u043d\u043e\u0439 \u0441\u0442\u0440\u043e\u043a\u0438:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. \u041f\u0435\u0440\u0435\u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u044c Inventor.") & vbCrLf & vbCrLf
    s = s & U("\u041f\u043e\u043b\u043d\u043e\u0435 \u0440\u0443\u043a\u043e\u0432\u043e\u0434\u0441\u0442\u0432\u043e:")
    ApprenticeShort_RU = s
End Function

Private Function ApprenticeShort_PL() As String
    Dim s As String
    s = U("Apprentice Server nie jest zarejestrowany na tym komputerze.") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 i nowsze wymagaj\u0105 oddzielnej instalacji Apprentice Server.") & vbCrLf
    s = s & U("Bez niego przetwarzanie wsadowe korzysta z pe\u0142nego silnika Inventor (10-50x wolniej).") & vbCrLf & vbCrLf
    s = s & U("Szybka instalacja:") & vbCrLf
    s = s & U("1. Pobierz z: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. Uruchom instalator (zwyk\u0142y u\u017cytkownik; uprawnienia administratora tylko do instalacji).") & vbCrLf
    s = s & U("3. Zarejestruj z normalnego wiersza polecenia:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Uruchom ponownie Inventor.") & vbCrLf & vbCrLf
    s = s & U("Pe\u0142ny przewodnik:")
    ApprenticeShort_PL = s
End Function

Private Function ApprenticeShort_CS() As String
    Dim s As String
    s = U("Apprentice Server nen\u00ed na tomto po\u010d\u00edta\u010di zaregistrov\u00e1n.") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 a nov\u011bj\u0161\u00ed vy\u017eaduj\u00ed samostatnou instalaci Apprentice Server.") & vbCrLf
    s = s & U("Bez n\u011bj d\u00e1vkov\u00e9 zpracov\u00e1n\u00ed pou\u017e\u00edv\u00e1 pln\u00fd engine Inventor (10-50x pomalej\u0161\u00ed).") & vbCrLf & vbCrLf
    s = s & U("Rychl\u00e1 instalace:") & vbCrLf
    s = s & U("1. St\u00e1hn\u011bte z: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. Spus\u0165te instal\u00e1tor (b\u011b\u017en\u00fd u\u017eivatel; pr\u00e1va spr\u00e1vce pouze pro instalaci).") & vbCrLf
    s = s & U("3. Zaregistrujte z b\u011b\u017en\u00e9ho p\u0159\u00edkazov\u00e9ho \u0159\u00e1dku:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Restartujte Inventor.") & vbCrLf & vbCrLf
    s = s & U("\u00dapln\u00fd pr\u016fvodce:")
    ApprenticeShort_CS = s
End Function

Private Function ApprenticeShort_JA() As String
    Dim s As String
    s = U("\u3053\u306e\u30de\u30b7\u30f3\u306b Apprentice Server \u304c\u767b\u9332\u3055\u308c\u3066\u3044\u307e\u305b\u3093\u3002") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 \u4ee5\u964d\u3067\u306f\u3001Apprentice Server \u3092\u500b\u5225\u306b\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u3059\u308b\u5fc5\u8981\u304c\u3042\u308a\u307e\u3059\u3002") & vbCrLf
    s = s & U("\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u3057\u306a\u3044\u5834\u5408\u3001\u30d0\u30c3\u30c1\u51e6\u7406\u306f\u5b8c\u5168\u306a Inventor \u30a8\u30f3\u30b8\u30f3\u3092\u4f7f\u7528\u3057\u307e\u3059 (10\u301c50 \u500d\u9045\u3044)\u3002") & vbCrLf & vbCrLf
    s = s & U("\u30af\u30a4\u30c3\u30af \u30a4\u30f3\u30b9\u30c8\u30fc\u30eb:") & vbCrLf
    s = s & U("1. \u30c0\u30a6\u30f3\u30ed\u30fc\u30c9: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. \u30a4\u30f3\u30b9\u30c8\u30fc\u30e9\u30fc\u3092\u5b9f\u884c (\u6a19\u6e96\u30e6\u30fc\u30b6\u30fc\u3001\u7ba1\u7406\u8005\u6a29\u9650\u306f\u30bb\u30c3\u30c8\u30a2\u30c3\u30d7\u306e\u307f)\u3002") & vbCrLf
    s = s & U("3. \u901a\u5e38\u306e\u30b3\u30de\u30f3\u30c9 \u30d7\u30ed\u30f3\u30d7\u30c8\u304b\u3089\u767b\u9332:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Inventor \u3092\u518d\u8d77\u52d5\u3057\u307e\u3059\u3002") & vbCrLf & vbCrLf
    s = s & U("\u5b8c\u5168\u306a\u30ac\u30a4\u30c9:")
    ApprenticeShort_JA = s
End Function

Private Function ApprenticeShort_KO() As String
    Dim s As String
    s = U("\uc774 \ucef4\ud4e8\ud130\uc5d0 Apprentice Server\uac00 \ub4f1\ub85d\ub418\uc5b4 \uc788\uc9c0 \uc54a\uc2b5\ub2c8\ub2e4.") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 \uc774\uc0c1\uc5d0\uc11c\ub294 Apprentice Server\ub97c \ubcc4\ub3c4\ub85c \uc124\uce58\ud574\uc57c \ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("\uc124\uce58\ud558\uc9c0 \uc54a\uc73c\uba74 \ubc30\uce58 \ucc98\ub9ac\ub294 \uc804\uccb4 Inventor \uc5d4\uc9c4\uc744 \uc0ac\uc6a9\ud569\ub2c8\ub2e4 (10-50\ubc30 \ub290\ub9bc).") & vbCrLf & vbCrLf
    s = s & U("\ube60\ub978 \uc124\uce58:") & vbCrLf
    s = s & U("1. \ub2e4\uc74c\uc5d0\uc11c \ub2e4\uc6b4\ub85c\ub4dc: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. \uc124\uce58 \uad00\ub9ac\uc790 \uc2e4\ud589 (\ud45c\uc900 \uc0ac\uc6a9\uc790, \uad00\ub9ac\uc790 \uad8c\ud55c\uc740 \uc124\uce58\uc5d0\ub9cc \ud544\uc694).") & vbCrLf
    s = s & U("3. \uc77c\ubc18 \uba85\ub839 \ud504\ub86c\ud504\ud2b8\uc5d0\uc11c \ub4f1\ub85d:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. Inventor\ub97c \ub2e4\uc2dc \uc2dc\uc791\ud569\ub2c8\ub2e4.") & vbCrLf & vbCrLf
    s = s & U("\uc804\uccb4 \uac00\uc774\ub4dc:")
    ApprenticeShort_KO = s
End Function

Private Function ApprenticeShort_ZH() As String
    Dim s As String
    s = U("\u6b64\u8ba1\u7b97\u673a\u4e0a\u672a\u6ce8\u518c Apprentice Server\u3002") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 \u53ca\u66f4\u9ad8\u7248\u672c\u9700\u8981\u5355\u72ec\u5b89\u88c5 Apprentice Server\u3002") & vbCrLf
    s = s & U("\u5982\u679c\u4e0d\u5b89\u88c5\uff0c\u6279\u5904\u7406\u5c06\u4f7f\u7528\u5b8c\u6574\u7684 Inventor \u5f15\u64ce\uff08\u901f\u5ea6\u6162 10-50 \u500d\uff09\u3002") & vbCrLf & vbCrLf
    s = s & U("\u5feb\u901f\u5b89\u88c5:") & vbCrLf
    s = s & U("1. \u4ece\u4ee5\u4e0b\u5730\u5740\u4e0b\u8f7d: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. \u8fd0\u884c\u5b89\u88c5\u7a0b\u5e8f\uff08\u6807\u51c6\u7528\u6237\uff1b\u4ec5\u5b89\u88c5\u65f6\u9700\u8981\u7ba1\u7406\u5458\u6743\u9650\uff09\u3002") & vbCrLf
    s = s & U("3. \u4ece\u666e\u901a\u547d\u4ee4\u63d0\u793a\u7b26\u6ce8\u518c:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. \u91cd\u65b0\u542f\u52a8 Inventor\u3002") & vbCrLf & vbCrLf
    s = s & U("\u5b8c\u6574\u6307\u5357:")
    ApprenticeShort_ZH = s
End Function

Private Function ApprenticeShort_ZHT() As String
    Dim s As String
    s = U("\u6b64\u96fb\u8166\u4e0a\u672a\u767b\u9304 Apprentice Server\u3002") & vbCrLf & vbCrLf
    s = s & U("Inventor 2026 \u53ca\u66f4\u65b0\u7248\u672c\u9700\u8981\u53e6\u5916\u5b89\u88dd Apprentice Server\u3002") & vbCrLf
    s = s & U("\u82e5\u672a\u5b89\u88dd\uff0c\u6279\u6b21\u8655\u7406\u5c07\u4f7f\u7528\u5b8c\u6574\u7684 Inventor \u5f15\u64ce\uff08\u901f\u5ea6\u6162 10-50 \u500d\uff09\u3002") & vbCrLf & vbCrLf
    s = s & U("\u5feb\u901f\u5b89\u88dd:") & vbCrLf
    s = s & U("1. \u5f9e\u4ee5\u4e0b\u4f4d\u7f6e\u4e0b\u8f09: ") & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("2. \u57f7\u884c\u5b89\u88dd\u7a0b\u5f0f\uff08\u6a19\u6e96\u4f7f\u7528\u8005\uff1b\u50c5\u5b89\u88dd\u6642\u9700\u8981\u7cfb\u7d71\u7ba1\u7406\u54e1\u6b0a\u9650\uff09\u3002") & vbCrLf
    s = s & U("3. \u5f9e\u4e00\u822c\u547d\u4ee4\u63d0\u793a\u5b57\u5143\u767b\u9304:") & vbCrLf
    s = s & "     """ & APP_REG_PATH_2027 & """ /install" & vbCrLf
    s = s & U("4. \u91cd\u65b0\u555f\u52d5 Inventor\u3002") & vbCrLf & vbCrLf
    s = s & U("\u5b8c\u6574\u6307\u5357:")
    ApprenticeShort_ZHT = s
End Function

' Full README bodies - 13 languages

Private Function ApprenticeReadme_EN() As String
    Dim s As String
    s = "Batch PDF Publisher - Apprentice Server Installation Guide" & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & "WHY IS THIS NEEDED?" & vbCrLf
    s = s & "Autodesk Inventor 2026 and later no longer ship Apprentice Server" & vbCrLf
    s = s & "as part of the main installer. It is now a separate download that" & vbCrLf
    s = s & "must be installed and registered manually. Without it, Batch PDF" & vbCrLf
    s = s & "Publisher falls back to the full Inventor engine, which is 10 to 50" & vbCrLf
    s = s & "times slower when reading sheet counts from IDW files." & vbCrLf & vbCrLf
    s = s & "BEFORE YOU START" & vbCrLf
    s = s & "- Install all pending Windows updates and reboot." & vbCrLf
    s = s & "- You need administrator rights for the installer only." & vbCrLf
    s = s & "- Free disk space: at least 2.2 GB (1.0 GB on system drive)." & vbCrLf
    s = s & "- Close all Autodesk applications (Inventor, Vault, Design Assistant)." & vbCrLf & vbCrLf
    s = s & "STEP 1 - DOWNLOAD" & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & "Choose the version matching your Inventor:" & vbCrLf
    s = s & "- Inventor 2027  ->  Apprentice Server 2027" & vbCrLf
    s = s & "- Inventor 2026  ->  Apprentice Server 2026" & vbCrLf & vbCrLf
    s = s & "STEP 2 - INSTALL" & vbCrLf
    s = s & "1. Double-click the downloaded .sfx.exe file." & vbCrLf
    s = s & "2. Follow the installer prompts." & vbCrLf
    s = s & "3. Click Finish when done." & vbCrLf
    s = s & "4. Reboot if prompted." & vbCrLf & vbCrLf
    s = s & "STEP 3 - REGISTER (REQUIRED)" & vbCrLf
    s = s & "Open Command Prompt (cmd.exe) AS A NORMAL USER, not administrator," & vbCrLf
    s = s & "and run:" & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & "If you installed Apprentice Server 2026, replace 2027 with 2026 in" & vbCrLf
    s = s & "the path above." & vbCrLf & vbCrLf
    s = s & "STEP 4 - VERIFY" & vbCrLf
    s = s & "Start Design Assistant 2027. On the Help menu, choose About" & vbCrLf
    s = s & "Design Assistant. The About box should show:" & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & "STEP 5 - RESTART INVENTOR AND RE-TEST" & vbCrLf
    s = s & "1. Close Inventor completely." & vbCrLf
    s = s & "2. Start Inventor." & vbCrLf
    s = s & "3. Run Batch PDF Publisher (Run_BatchPDFPublisher)." & vbCrLf
    s = s & "4. The Apprentice warning dialog should no longer appear." & vbCrLf
    s = s & "5. Check the log:" & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & "   There should be no line 'WARN [PRJ_GetApprentice]'." & vbCrLf & vbCrLf
    s = s & "TROUBLESHOOTING" & vbCrLf
    s = s & "Q: The warning dialog still appears after installing." & vbCrLf
    s = s & "A: You skipped Step 3 (register). Apprentice Server requires" & vbCrLf
    s = s & "   explicit COM registration via ApprenticeRegSvr.exe /install." & vbCrLf & vbCrLf
    s = s & "Q: Registration command reports 'file not found'." & vbCrLf
    s = s & "A: The version number in the path is wrong. Check the actual" & vbCrLf
    s = s & "   installation folder under C:\Program Files\Autodesk\." & vbCrLf & vbCrLf
    s = s & "Q: I cannot install Apprentice Server (no admin rights)." & vbCrLf
    s = s & "A: Batch PDF Publisher still works but is 10-50x slower." & vbCrLf
    s = s & "   Ask your IT department to install it." & vbCrLf & vbCrLf
    s = s & "Q: How do I suppress the warning without installing?" & vbCrLf
    s = s & "A: Delete this registry key and restart Inventor:" & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_EN = s
End Function

Private Function ApprenticeReadme_DE() As String
    Dim s As String
    s = U("Batch PDF Publisher - Installationsanleitung f\u00fcr Apprentice Server") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("WARUM IST DAS ERFORDERLICH?") & vbCrLf
    s = s & U("Autodesk Inventor 2026 und h\u00f6her liefern Apprentice Server nicht mehr") & vbCrLf
    s = s & U("im Hauptinstaller mit. Es ist jetzt ein separater Download, der manuell") & vbCrLf
    s = s & U("installiert und registriert werden muss. Ohne ihn verwendet Batch PDF") & vbCrLf
    s = s & U("Publisher die vollst\u00e4ndige Inventor-Engine, die beim Lesen der Blattanzahl") & vbCrLf
    s = s & U("aus IDW-Dateien 10- bis 50-mal langsamer ist.") & vbCrLf & vbCrLf
    s = s & U("BEVOR SIE BEGINNEN") & vbCrLf
    s = s & U("- Installieren Sie alle ausstehenden Windows-Updates und starten Sie neu.") & vbCrLf
    s = s & U("- Administratorrechte sind nur f\u00fcr das Installationsprogramm erforderlich.") & vbCrLf
    s = s & U("- Freier Speicherplatz: mindestens 2,2 GB (1,0 GB auf Systemlaufwerk).") & vbCrLf
    s = s & U("- Schlie\u00dfen Sie alle Autodesk-Anwendungen (Inventor, Vault, Design Assistant).") & vbCrLf & vbCrLf
    s = s & U("SCHRITT 1 - HERUNTERLADEN") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("W\u00e4hlen Sie die zu Ihrer Inventor-Version passende Version:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("SCHRITT 2 - INSTALLIEREN") & vbCrLf
    s = s & U("1. Doppelklicken Sie auf die heruntergeladene .sfx.exe-Datei.") & vbCrLf
    s = s & U("2. Folgen Sie den Anweisungen des Installationsprogramms.") & vbCrLf
    s = s & U("3. Klicken Sie abschlie\u00dfend auf Fertig stellen.") & vbCrLf
    s = s & U("4. Starten Sie das System bei Aufforderung neu.") & vbCrLf & vbCrLf
    s = s & U("SCHRITT 3 - REGISTRIEREN (ERFORDERLICH)") & vbCrLf
    s = s & U("\u00d6ffnen Sie die Eingabeaufforderung (cmd.exe) ALS NORMALER BENUTZER,") & vbCrLf
    s = s & U("nicht als Administrator, und f\u00fchren Sie aus:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Wenn Sie Apprentice Server 2026 installiert haben, ersetzen Sie 2027") & vbCrLf
    s = s & U("durch 2026 im obigen Pfad.") & vbCrLf & vbCrLf
    s = s & U("SCHRITT 4 - \u00dcBERPR\u00dcFEN") & vbCrLf
    s = s & U("Starten Sie Design Assistant 2027. W\u00e4hlen Sie im Men\u00fc Hilfe die Option") & vbCrLf
    s = s & U("\u00dcber Design Assistant. Das Info-Fenster sollte anzeigen:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("SCHRITT 5 - INVENTOR NEU STARTEN UND ERNEUT TESTEN") & vbCrLf
    s = s & U("1. Schlie\u00dfen Sie Inventor vollst\u00e4ndig.") & vbCrLf
    s = s & U("2. Starten Sie Inventor.") & vbCrLf
    s = s & U("3. F\u00fchren Sie Batch PDF Publisher aus (Run_BatchPDFPublisher).") & vbCrLf
    s = s & U("4. Der Apprentice-Warnungsdialog sollte nicht mehr erscheinen.") & vbCrLf
    s = s & U("5. Pr\u00fcfen Sie das Protokoll:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   Es sollte keine Zeile 'WARN [PRJ_GetApprentice]' geben.") & vbCrLf & vbCrLf
    s = s & U("FEHLERBEHEBUNG") & vbCrLf
    s = s & U("F: Der Warnungsdialog erscheint auch nach der Installation.") & vbCrLf
    s = s & U("A: Sie haben Schritt 3 (\u00dcberpr\u00fcfung) \u00fcbersprungen. Apprentice Server") & vbCrLf
    s = s & U("   ben\u00f6tigt eine explizite COM-Registrierung.") & vbCrLf & vbCrLf
    s = s & U("F: Der Registrierungsbefehl meldet 'Datei nicht gefunden'.") & vbCrLf
    s = s & U("A: Die Versionsnummer im Pfad ist falsch. Pr\u00fcfen Sie den tats\u00e4chlichen") & vbCrLf
    s = s & U("   Ordner unter C:\Program Files\Autodesk\.") & vbCrLf & vbCrLf
    s = s & U("F: Ich kann Apprentice Server nicht installieren (keine Admin-Rechte).") & vbCrLf
    s = s & U("A: Batch PDF Publisher funktioniert weiterhin, ist aber 10-50x langsamer.") & vbCrLf
    s = s & U("   Bitten Sie Ihre IT-Abteilung um Installation.") & vbCrLf & vbCrLf
    s = s & U("F: Wie unterdr\u00fccke ich die Warnung ohne Installation?") & vbCrLf
    s = s & U("A: L\u00f6schen Sie diesen Registrierungsschl\u00fcssel und starten Sie Inventor neu:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_DE = s
End Function

Private Function ApprenticeReadme_FR() As String
    Dim s As String
    s = U("Batch PDF Publisher - Guide d'installation d'Apprentice Server") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("POURQUOI EST-CE N\u00c9CESSAIRE ?") & vbCrLf
    s = s & U("Autodesk Inventor 2026 et versions ult\u00e9rieures ne fournissent plus") & vbCrLf
    s = s & U("Apprentice Server avec l'installateur principal. Il s'agit d\u00e9sormais d'un") & vbCrLf
    s = s & U("t\u00e9l\u00e9chargement s\u00e9par\u00e9 \u00e0 installer et enregistrer manuellement. Sans lui,") & vbCrLf
    s = s & U("Batch PDF Publisher utilise le moteur Inventor complet, qui est 10 \u00e0 50") & vbCrLf
    s = s & U("fois plus lent lors de la lecture du nombre de feuilles des fichiers IDW.") & vbCrLf & vbCrLf
    s = s & U("AVANT DE COMMENCER") & vbCrLf
    s = s & U("- Installez toutes les mises \u00e0 jour Windows en attente et red\u00e9marrez.") & vbCrLf
    s = s & U("- Les droits administrateur ne sont n\u00e9cessaires que pour l'installation.") & vbCrLf
    s = s & U("- Espace disque libre : au moins 2,2 Go (1,0 Go sur le lecteur syst\u00e8me).") & vbCrLf
    s = s & U("- Fermez toutes les applications Autodesk (Inventor, Vault, Design Assistant).") & vbCrLf & vbCrLf
    s = s & U("\u00c9TAPE 1 - T\u00c9L\u00c9CHARGEMENT") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("Choisissez la version correspondant \u00e0 votre Inventor :") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("\u00c9TAPE 2 - INSTALLATION") & vbCrLf
    s = s & U("1. Double-cliquez sur le fichier .sfx.exe t\u00e9l\u00e9charg\u00e9.") & vbCrLf
    s = s & U("2. Suivez les invites de l'installateur.") & vbCrLf
    s = s & U("3. Cliquez sur Terminer lorsque c'est fait.") & vbCrLf
    s = s & U("4. Red\u00e9marrez si demand\u00e9.") & vbCrLf & vbCrLf
    s = s & U("\u00c9TAPE 3 - ENREGISTREMENT (OBLIGATOIRE)") & vbCrLf
    s = s & U("Ouvrez l'invite de commandes (cmd.exe) EN TANT QU'UTILISATEUR NORMAL,") & vbCrLf
    s = s & U("pas en tant qu'administrateur, et ex\u00e9cutez :") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Si vous avez install\u00e9 Apprentice Server 2026, remplacez 2027 par 2026") & vbCrLf
    s = s & U("dans le chemin ci-dessus.") & vbCrLf & vbCrLf
    s = s & U("\u00c9TAPE 4 - V\u00c9RIFICATION") & vbCrLf
    s = s & U("Lancez Design Assistant 2027. Dans le menu Aide, s\u00e9lectionnez") & vbCrLf
    s = s & U("\u00c0 propos de Design Assistant. La bo\u00eete doit afficher :") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("\u00c9TAPE 5 - RED\u00c9MARRER INVENTOR ET RETESTER") & vbCrLf
    s = s & U("1. Fermez compl\u00e8tement Inventor.") & vbCrLf
    s = s & U("2. D\u00e9marrez Inventor.") & vbCrLf
    s = s & U("3. Ex\u00e9cutez Batch PDF Publisher (Run_BatchPDFPublisher).") & vbCrLf
    s = s & U("4. La bo\u00eete de dialogue d'avertissement ne devrait plus appara\u00eetre.") & vbCrLf
    s = s & U("5. V\u00e9rifiez le journal :") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   Il ne devrait pas y avoir de ligne 'WARN [PRJ_GetApprentice]'.") & vbCrLf & vbCrLf
    s = s & U("D\u00c9PANNAGE") & vbCrLf
    s = s & U("Q : L'avertissement appara\u00eet encore apr\u00e8s l'installation.") & vbCrLf
    s = s & U("R : Vous avez saut\u00e9 l'\u00e9tape 3 (enregistrement). Apprentice Server") & vbCrLf
    s = s & U("   n\u00e9cessite un enregistrement COM explicite.") & vbCrLf & vbCrLf
    s = s & U("Q : La commande d'enregistrement indique 'fichier introuvable'.") & vbCrLf
    s = s & U("R : Le num\u00e9ro de version dans le chemin est incorrect. V\u00e9rifiez le") & vbCrLf
    s = s & U("   dossier r\u00e9el sous C:\Program Files\Autodesk\.") & vbCrLf & vbCrLf
    s = s & U("Q : Je ne peux pas installer Apprentice Server (pas de droits admin).") & vbCrLf
    s = s & U("R : Batch PDF Publisher fonctionne toujours mais est 10-50x plus lent.") & vbCrLf
    s = s & U("   Demandez \u00e0 votre service informatique de l'installer.") & vbCrLf & vbCrLf
    s = s & U("Q : Comment supprimer l'avertissement sans installer ?") & vbCrLf
    s = s & U("R : Supprimez cette cl\u00e9 de registre et red\u00e9marrez Inventor :") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_FR = s
End Function

Private Function ApprenticeReadme_IT() As String
    Dim s As String
    s = U("Batch PDF Publisher - Guida all'installazione di Apprentice Server") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("PERCH\u00c9 \u00c8 NECESSARIO?") & vbCrLf
    s = s & U("Autodesk Inventor 2026 e versioni successive non includono pi\u00f9 Apprentice") & vbCrLf
    s = s & U("Server nel programma di installazione principale. Ora \u00e8 un download") & vbCrLf
    s = s & U("separato che deve essere installato e registrato manualmente. Senza di") & vbCrLf
    s = s & U("esso, Batch PDF Publisher usa il motore Inventor completo, che \u00e8 10-50") & vbCrLf
    s = s & U("volte pi\u00f9 lento durante la lettura del numero di fogli dai file IDW.") & vbCrLf & vbCrLf
    s = s & U("PRIMA DI INIZIARE") & vbCrLf
    s = s & U("- Installa tutti gli aggiornamenti di Windows in sospeso e riavvia.") & vbCrLf
    s = s & U("- Sono necessari privilegi di amministratore solo per il programma di installazione.") & vbCrLf
    s = s & U("- Spazio su disco libero: almeno 2,2 GB (1,0 GB sull'unit\u00e0 di sistema).") & vbCrLf
    s = s & U("- Chiudi tutte le applicazioni Autodesk (Inventor, Vault, Design Assistant).") & vbCrLf & vbCrLf
    s = s & U("FASE 1 - DOWNLOAD") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("Scegli la versione corrispondente al tuo Inventor:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("FASE 2 - INSTALLAZIONE") & vbCrLf
    s = s & U("1. Fai doppio clic sul file .sfx.exe scaricato.") & vbCrLf
    s = s & U("2. Segui le istruzioni del programma di installazione.") & vbCrLf
    s = s & U("3. Fai clic su Fine al termine.") & vbCrLf
    s = s & U("4. Riavvia se richiesto.") & vbCrLf & vbCrLf
    s = s & U("FASE 3 - REGISTRAZIONE (OBBLIGATORIA)") & vbCrLf
    s = s & U("Apri il Prompt dei comandi (cmd.exe) COME UTENTE NORMALE, non come") & vbCrLf
    s = s & U("amministratore, ed esegui:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Se hai installato Apprentice Server 2026, sostituisci 2027 con 2026") & vbCrLf
    s = s & U("nel percorso sopra.") & vbCrLf & vbCrLf
    s = s & U("FASE 4 - VERIFICA") & vbCrLf
    s = s & U("Avvia Design Assistant 2027. Nel menu ?, seleziona Informazioni su") & vbCrLf
    s = s & U("Design Assistant. La finestra dovrebbe mostrare:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("FASE 5 - RIAVVIA INVENTOR E RIPROVA") & vbCrLf
    s = s & U("1. Chiudi completamente Inventor.") & vbCrLf
    s = s & U("2. Avvia Inventor.") & vbCrLf
    s = s & U("3. Esegui Batch PDF Publisher (Run_BatchPDFPublisher).") & vbCrLf
    s = s & U("4. La finestra di avviso Apprentice non dovrebbe pi\u00f9 apparire.") & vbCrLf
    s = s & U("5. Controlla il registro:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   Non dovrebbe esserci la riga 'WARN [PRJ_GetApprentice]'.") & vbCrLf & vbCrLf
    s = s & U("RISOLUZIONE DEI PROBLEMI") & vbCrLf
    s = s & U("D: La finestra di avviso appare ancora dopo l'installazione.") & vbCrLf
    s = s & U("R: Hai saltato la Fase 3 (registrazione). Apprentice Server richiede") & vbCrLf
    s = s & U("   una registrazione COM esplicita.") & vbCrLf & vbCrLf
    s = s & U("D: Il comando di registrazione segnala 'file non trovato'.") & vbCrLf
    s = s & U("R: Il numero di versione nel percorso \u00e8 errato. Controlla la cartella") & vbCrLf
    s = s & U("   effettiva sotto C:\Program Files\Autodesk\.") & vbCrLf & vbCrLf
    s = s & U("D: Non posso installare Apprentice Server (nessun diritto di amministratore).") & vbCrLf
    s = s & U("R: Batch PDF Publisher funziona comunque ma \u00e8 10-50x pi\u00f9 lento.") & vbCrLf
    s = s & U("   Chiedi al tuo reparto IT di installarlo.") & vbCrLf & vbCrLf
    s = s & U("D: Come faccio a sopprimere l'avviso senza installare?") & vbCrLf
    s = s & U("R: Elimina questa chiave di registro e riavvia Inventor:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_IT = s
End Function

Private Function ApprenticeReadme_ES() As String
    Dim s As String
    s = U("Batch PDF Publisher - Gu\u00eda de instalaci\u00f3n de Apprentice Server") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("\u00bfPOR QU\u00c9 ES NECESARIO?") & vbCrLf
    s = s & U("Autodesk Inventor 2026 y versiones posteriores ya no incluyen Apprentice") & vbCrLf
    s = s & U("Server en el instalador principal. Ahora es una descarga separada que") & vbCrLf
    s = s & U("debe instalarse y registrarse manualmente. Sin \u00e9l, Batch PDF Publisher") & vbCrLf
    s = s & U("usa el motor completo de Inventor, que es 10 a 50 veces m\u00e1s lento al") & vbCrLf
    s = s & U("leer el n\u00famero de hojas de los archivos IDW.") & vbCrLf & vbCrLf
    s = s & U("ANTES DE COMENZAR") & vbCrLf
    s = s & U("- Instale todas las actualizaciones de Windows pendientes y reinicie.") & vbCrLf
    s = s & U("- Solo necesita privilegios de administrador para el instalador.") & vbCrLf
    s = s & U("- Espacio libre en disco: al menos 2,2 GB (1,0 GB en la unidad del sistema).") & vbCrLf
    s = s & U("- Cierre todas las aplicaciones de Autodesk (Inventor, Vault, Design Assistant).") & vbCrLf & vbCrLf
    s = s & U("PASO 1 - DESCARGAR") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("Elija la versi\u00f3n que coincida con su Inventor:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("PASO 2 - INSTALAR") & vbCrLf
    s = s & U("1. Haga doble clic en el archivo .sfx.exe descargado.") & vbCrLf
    s = s & U("2. Siga las indicaciones del instalador.") & vbCrLf
    s = s & U("3. Haga clic en Finalizar cuando termine.") & vbCrLf
    s = s & U("4. Reinicie si se le solicita.") & vbCrLf & vbCrLf
    s = s & U("PASO 3 - REGISTRAR (OBLIGATORIO)") & vbCrLf
    s = s & U("Abra el S\u00edmbolo del sistema (cmd.exe) COMO USUARIO NORMAL, no como") & vbCrLf
    s = s & U("administrador, y ejecute:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Si instal\u00f3 Apprentice Server 2026, reemplace 2027 por 2026 en la") & vbCrLf
    s = s & U("ruta anterior.") & vbCrLf & vbCrLf
    s = s & U("PASO 4 - VERIFICAR") & vbCrLf
    s = s & U("Inicie Design Assistant 2027. En el men\u00fa Ayuda, seleccione Acerca de") & vbCrLf
    s = s & U("Design Assistant. El cuadro debe mostrar:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("PASO 5 - REINICIAR INVENTOR Y VOLVER A PROBAR") & vbCrLf
    s = s & U("1. Cierre Inventor completamente.") & vbCrLf
    s = s & U("2. Inicie Inventor.") & vbCrLf
    s = s & U("3. Ejecute Batch PDF Publisher (Run_BatchPDFPublisher).") & vbCrLf
    s = s & U("4. El cuadro de advertencia de Apprentice ya no deber\u00eda aparecer.") & vbCrLf
    s = s & U("5. Compruebe el registro:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   No deber\u00eda haber la l\u00ednea 'WARN [PRJ_GetApprentice]'.") & vbCrLf & vbCrLf
    s = s & U("SOLUCI\u00d3N DE PROBLEMAS") & vbCrLf
    s = s & U("P: El aviso sigue apareciendo despu\u00e9s de la instalaci\u00f3n.") & vbCrLf
    s = s & U("R: Ha omitido el Paso 3 (registro). Apprentice Server requiere un") & vbCrLf
    s = s & U("   registro COM expl\u00edcito.") & vbCrLf & vbCrLf
    s = s & U("P: El comando de registro informa 'archivo no encontrado'.") & vbCrLf
    s = s & U("R: El n\u00famero de versi\u00f3n en la ruta es incorrecto. Compruebe la carpeta") & vbCrLf
    s = s & U("   real en C:\Program Files\Autodesk\.") & vbCrLf & vbCrLf
    s = s & U("P: No puedo instalar Apprentice Server (sin derechos de administrador).") & vbCrLf
    s = s & U("R: Batch PDF Publisher sigue funcionando pero es 10-50x m\u00e1s lento.") & vbCrLf
    s = s & U("   Pida a su departamento de TI que lo instale.") & vbCrLf & vbCrLf
    s = s & U("P: \u00bfC\u00f3mo suprimo el aviso sin instalarlo?") & vbCrLf
    s = s & U("R: Elimine esta clave del registro y reinicie Inventor:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_ES = s
End Function

Private Function ApprenticeReadme_PT() As String
    Dim s As String
    s = U("Batch PDF Publisher - Guia de instala\u00e7\u00e3o do Apprentice Server") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("POR QUE ISSO \u00c9 NECESS\u00c1RIO?") & vbCrLf
    s = s & U("O Autodesk Inventor 2026 e vers\u00f5es posteriores n\u00e3o incluem mais o") & vbCrLf
    s = s & U("Apprentice Server no instalador principal. Agora \u00e9 um download separado") & vbCrLf
    s = s & U("que deve ser instalado e registrado manualmente. Sem ele, o Batch PDF") & vbCrLf
    s = s & U("Publisher usa o motor Inventor completo, que \u00e9 10 a 50 vezes mais lento") & vbCrLf
    s = s & U("ao ler a contagem de folhas de arquivos IDW.") & vbCrLf & vbCrLf
    s = s & U("ANTES DE COME\u00c7AR") & vbCrLf
    s = s & U("- Instale todas as atualiza\u00e7\u00f5es do Windows pendentes e reinicie.") & vbCrLf
    s = s & U("- Voc\u00ea precisa de privil\u00e9gios de administrador somente para o instalador.") & vbCrLf
    s = s & U("- Espa\u00e7o livre em disco: pelo menos 2,2 GB (1,0 GB na unidade do sistema).") & vbCrLf
    s = s & U("- Feche todos os aplicativos Autodesk (Inventor, Vault, Design Assistant).") & vbCrLf & vbCrLf
    s = s & U("PASSO 1 - DOWNLOAD") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("Escolha a vers\u00e3o correspondente ao seu Inventor:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("PASSO 2 - INSTALAR") & vbCrLf
    s = s & U("1. Clique duas vezes no arquivo .sfx.exe baixado.") & vbCrLf
    s = s & U("2. Siga as instru\u00e7\u00f5es do instalador.") & vbCrLf
    s = s & U("3. Clique em Concluir quando terminar.") & vbCrLf
    s = s & U("4. Reinicie se solicitado.") & vbCrLf & vbCrLf
    s = s & U("PASSO 3 - REGISTRAR (OBRIGAT\u00d3RIO)") & vbCrLf
    s = s & U("Abra o Prompt de Comando (cmd.exe) COMO USU\u00c1RIO NORMAL, n\u00e3o como") & vbCrLf
    s = s & U("administrador, e execute:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Se voc\u00ea instalou o Apprentice Server 2026, substitua 2027 por 2026 no") & vbCrLf
    s = s & U("caminho acima.") & vbCrLf & vbCrLf
    s = s & U("PASSO 4 - VERIFICAR") & vbCrLf
    s = s & U("Inicie o Design Assistant 2027. No menu Ajuda, selecione Sobre o") & vbCrLf
    s = s & U("Design Assistant. A caixa deve mostrar:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("PASSO 5 - REINICIAR O INVENTOR E TESTAR NOVAMENTE") & vbCrLf
    s = s & U("1. Feche completamente o Inventor.") & vbCrLf
    s = s & U("2. Inicie o Inventor.") & vbCrLf
    s = s & U("3. Execute o Batch PDF Publisher (Run_BatchPDFPublisher).") & vbCrLf
    s = s & U("4. A caixa de aviso do Apprentice n\u00e3o deve mais aparecer.") & vbCrLf
    s = s & U("5. Verifique o log:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   N\u00e3o deve haver a linha 'WARN [PRJ_GetApprentice]'.") & vbCrLf & vbCrLf
    s = s & U("SOLU\u00c7\u00c3O DE PROBLEMAS") & vbCrLf
    s = s & U("P: O aviso ainda aparece ap\u00f3s a instala\u00e7\u00e3o.") & vbCrLf
    s = s & U("R: Voc\u00ea pulou o Passo 3 (registro). O Apprentice Server exige registro") & vbCrLf
    s = s & U("   COM expl\u00edcito.") & vbCrLf & vbCrLf
    s = s & U("P: O comando de registro informa 'arquivo n\u00e3o encontrado'.") & vbCrLf
    s = s & U("R: O n\u00famero da vers\u00e3o no caminho est\u00e1 errado. Verifique a pasta real") & vbCrLf
    s = s & U("   em C:\Program Files\Autodesk\.") & vbCrLf & vbCrLf
    s = s & U("P: N\u00e3o consigo instalar o Apprentice Server (sem direitos de admin).") & vbCrLf
    s = s & U("R: O Batch PDF Publisher ainda funciona, mas \u00e9 10-50x mais lento.") & vbCrLf
    s = s & U("   Pe\u00e7a ao seu departamento de TI para instalar.") & vbCrLf & vbCrLf
    s = s & U("P: Como suprimo o aviso sem instalar?") & vbCrLf
    s = s & U("R: Exclua esta chave do registro e reinicie o Inventor:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_PT = s
End Function

Private Function ApprenticeReadme_RU() As String
    Dim s As String
    s = U("Batch PDF Publisher - \u0420\u0443\u043a\u043e\u0432\u043e\u0434\u0441\u0442\u0432\u043e \u043f\u043e \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0435 Apprentice Server") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("\u0417\u0410\u0427\u0415\u041c \u042d\u0422\u041e \u041d\u0423\u0416\u041d\u041e?") & vbCrLf
    s = s & U("Autodesk Inventor 2026 \u0438 \u0431\u043e\u043b\u0435\u0435 \u043f\u043e\u0437\u0434\u043d\u0438\u0435 \u0432\u0435\u0440\u0441\u0438\u0438 \u0431\u043e\u043b\u044c\u0448\u0435 \u043d\u0435 \u0432\u043a\u043b\u044e\u0447\u0430\u044e\u0442") & vbCrLf
    s = s & U("Apprentice Server \u0432 \u043e\u0441\u043d\u043e\u0432\u043d\u043e\u0439 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a. \u0422\u0435\u043f\u0435\u0440\u044c \u044d\u0442\u043e \u043e\u0442\u0434\u0435\u043b\u044c\u043d\u0430\u044f") & vbCrLf
    s = s & U("\u0437\u0430\u0433\u0440\u0443\u0437\u043a\u0430, \u043a\u043e\u0442\u043e\u0440\u0443\u044e \u043d\u0443\u0436\u043d\u043e \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c \u0438 \u0437\u0430\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0438\u0440\u043e\u0432\u0430\u0442\u044c \u0432\u0440\u0443\u0447\u043d\u0443\u044e.") & vbCrLf
    s = s & U("\u0411\u0435\u0437 \u043d\u0435\u0433\u043e Batch PDF Publisher \u0438\u0441\u043f\u043e\u043b\u044c\u0437\u0443\u0435\u0442 \u043f\u043e\u043b\u043d\u044b\u0439 \u0434\u0432\u0438\u0436\u043e\u043a Inventor,") & vbCrLf
    s = s & U("\u043a\u043e\u0442\u043e\u0440\u044b\u0439 \u0432 10-50 \u0440\u0430\u0437 \u043c\u0435\u0434\u043b\u0435\u043d\u043d\u0435\u0435 \u043f\u0440\u0438 \u0447\u0442\u0435\u043d\u0438\u0438 \u043a\u043e\u043b\u0438\u0447\u0435\u0441\u0442\u0432\u0430 \u043b\u0438\u0441\u0442\u043e\u0432 \u0438\u0437 IDW.") & vbCrLf & vbCrLf
    s = s & U("\u041f\u0415\u0420\u0415\u0414 \u041d\u0410\u0427\u0410\u041b\u041e\u041c") & vbCrLf
    s = s & U("- \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u0435 \u0432\u0441\u0435 \u043e\u0436\u0438\u0434\u0430\u044e\u0449\u0438\u0435 \u043e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u0438\u044f Windows \u0438 \u043f\u0435\u0440\u0435\u0437\u0430\u0433\u0440\u0443\u0437\u0438\u0442\u0435 \u0441\u0438\u0441\u0442\u0435\u043c\u0443.") & vbCrLf
    s = s & U("- \u041f\u0440\u0430\u0432\u0430 \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0430\u0442\u043e\u0440\u0430 \u043d\u0443\u0436\u043d\u044b \u0442\u043e\u043b\u044c\u043a\u043e \u0434\u043b\u044f \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a\u0430.") & vbCrLf
    s = s & U("- \u0421\u0432\u043e\u0431\u043e\u0434\u043d\u043e\u0435 \u043c\u0435\u0441\u0442\u043e \u043d\u0430 \u0434\u0438\u0441\u043a\u0435: \u043d\u0435 \u043c\u0435\u043d\u0435\u0435 2,2 \u0413\u0411 (1,0 \u0413\u0411 \u043d\u0430 \u0441\u0438\u0441\u0442\u0435\u043c\u043d\u043e\u043c \u0434\u0438\u0441\u043a\u0435).") & vbCrLf
    s = s & U("- \u0417\u0430\u043a\u0440\u043e\u0439\u0442\u0435 \u0432\u0441\u0435 \u043f\u0440\u0438\u043b\u043e\u0436\u0435\u043d\u0438\u044f Autodesk (Inventor, Vault, Design Assistant).") & vbCrLf & vbCrLf
    s = s & U("\u0428\u0410\u0413 1 - \u0417\u0410\u0413\u0420\u0423\u0417\u041a\u0410") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u0432\u0435\u0440\u0441\u0438\u044e, \u0441\u043e\u043e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0443\u044e\u0449\u0443\u044e \u0432\u0430\u0448\u0435\u043c\u0443 Inventor:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("\u0428\u0410\u0413 2 - \u0423\u0421\u0422\u0410\u041d\u041e\u0412\u041a\u0410") & vbCrLf
    s = s & U("1. \u0414\u0432\u0430\u0436\u0434\u044b \u0449\u0451\u043b\u043a\u043d\u0438\u0442\u0435 \u043f\u043e \u0441\u043a\u0430\u0447\u0430\u043d\u043d\u043e\u043c\u0443 \u0444\u0430\u0439\u043b\u0443 .sfx.exe.") & vbCrLf
    s = s & U("2. \u0421\u043b\u0435\u0434\u0443\u0439\u0442\u0435 \u0438\u043d\u0441\u0442\u0440\u0443\u043a\u0446\u0438\u044f\u043c \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a\u0430.") & vbCrLf
    s = s & U("3. \u041f\u043e \u0437\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u0438\u0438 \u043d\u0430\u0436\u043c\u0438\u0442\u0435 \u0413\u043e\u0442\u043e\u0432\u043e.") & vbCrLf
    s = s & U("4. \u041f\u0440\u0438 \u043d\u0435\u043e\u0431\u0445\u043e\u0434\u0438\u043c\u043e\u0441\u0442\u0438 \u043f\u0435\u0440\u0435\u0437\u0430\u0433\u0440\u0443\u0437\u0438\u0442\u0435 \u0441\u0438\u0441\u0442\u0435\u043c\u0443.") & vbCrLf & vbCrLf
    s = s & U("\u0428\u0410\u0413 3 - \u0420\u0415\u0413\u0418\u0421\u0422\u0420\u0410\u0426\u0418\u042f (\u041e\u0411\u042f\u0417\u0410\u0422\u0415\u041b\u042c\u041d\u041e)") & vbCrLf
    s = s & U("\u041e\u0442\u043a\u0440\u043e\u0439\u0442\u0435 \u043a\u043e\u043c\u0430\u043d\u0434\u043d\u0443\u044e \u0441\u0442\u0440\u043e\u043a\u0443 (cmd.exe) \u041a\u0410\u041a \u041e\u0411\u042b\u0427\u041d\u042b\u0419") & vbCrLf
    s = s & U("\u041f\u041e\u041b\u042c\u0417\u041e\u0412\u0410\u0422\u0415\u041b\u042c, \u043d\u0435 \u043a\u0430\u043a \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0430\u0442\u043e\u0440, \u0438 \u0432\u044b\u043f\u043e\u043b\u043d\u0438\u0442\u0435:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("\u0415\u0441\u043b\u0438 \u0432\u044b \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u043b\u0438 Apprentice Server 2026, \u0437\u0430\u043c\u0435\u043d\u0438\u0442\u0435 2027") & vbCrLf
    s = s & U("\u043d\u0430 2026 \u0432 \u043f\u0443\u0442\u0438 \u0432\u044b\u0448\u0435.") & vbCrLf & vbCrLf
    s = s & U("\u0428\u0410\u0413 4 - \u041f\u0420\u041e\u0412\u0415\u0420\u041a\u0410") & vbCrLf
    s = s & U("\u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 Design Assistant 2027. \u0412 \u043c\u0435\u043d\u044e \u0421\u043f\u0440\u0430\u0432\u043a\u0430 \u0432\u044b\u0431\u0435\u0440\u0438\u0442\u0435") & vbCrLf
    s = s & U("\u041e \u043f\u0440\u043e\u0433\u0440\u0430\u043c\u043c\u0435 Design Assistant. \u0412 \u043e\u043a\u043d\u0435 \u0434\u043e\u043b\u0436\u043d\u043e \u0431\u044b\u0442\u044c:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("\u0428\u0410\u0413 5 - \u041f\u0415\u0420\u0415\u0417\u0410\u041f\u0423\u0421\u041a INVENTOR \u0418 \u041f\u041e\u0412\u0422\u041e\u0420\u041d\u042b\u0419 \u0422\u0415\u0421\u0422") & vbCrLf
    s = s & U("1. \u041f\u043e\u043b\u043d\u043e\u0441\u0442\u044c\u044e \u0437\u0430\u043a\u0440\u043e\u0439\u0442\u0435 Inventor.") & vbCrLf
    s = s & U("2. \u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 Inventor.") & vbCrLf
    s = s & U("3. \u0417\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 Batch PDF Publisher (Run_BatchPDFPublisher).") & vbCrLf
    s = s & U("4. \u0414\u0438\u0430\u043b\u043e\u0433 \u043f\u0440\u0435\u0434\u0443\u043f\u0440\u0435\u0436\u0434\u0435\u043d\u0438\u044f Apprentice \u0431\u043e\u043b\u044c\u0448\u0435 \u043d\u0435 \u0434\u043e\u043b\u0436\u0435\u043d \u043f\u043e\u044f\u0432\u043b\u044f\u0442\u044c\u0441\u044f.") & vbCrLf
    s = s & U("5. \u041f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0436\u0443\u0440\u043d\u0430\u043b:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   \u041d\u0435 \u0434\u043e\u043b\u0436\u043d\u043e \u0431\u044b\u0442\u044c \u0441\u0442\u0440\u043e\u043a\u0438 'WARN [PRJ_GetApprentice]'.") & vbCrLf & vbCrLf
    s = s & U("\u0423\u0421\u0422\u0420\u0410\u041d\u0415\u041d\u0418\u0415 \u041d\u0415\u041f\u041e\u041b\u0410\u0414\u041e\u041a") & vbCrLf
    s = s & U("\u0412: \u041f\u0440\u0435\u0434\u0443\u043f\u0440\u0435\u0436\u0434\u0435\u043d\u0438\u0435 \u0432\u0441\u0451 \u0435\u0449\u0451 \u043f\u043e\u044f\u0432\u043b\u044f\u0435\u0442\u0441\u044f \u043f\u043e\u0441\u043b\u0435 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438.") & vbCrLf
    s = s & U("\u041e: \u0412\u044b \u043f\u0440\u043e\u043f\u0443\u0441\u0442\u0438\u043b\u0438 \u0428\u0430\u0433 3 (\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0430\u0446\u0438\u044f). Apprentice Server \u0442\u0440\u0435\u0431\u0443\u0435\u0442") & vbCrLf
    s = s & U("   \u044f\u0432\u043d\u043e\u0439 COM-\u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0430\u0446\u0438\u0438.") & vbCrLf & vbCrLf
    s = s & U("\u0412: \u041a\u043e\u043c\u0430\u043d\u0434\u0430 \u0440\u0435\u0433\u0438\u0441\u0442\u0440\u0430\u0446\u0438\u0438 \u0441\u043e\u043e\u0431\u0449\u0430\u0435\u0442 '\u0444\u0430\u0439\u043b \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d'.") & vbCrLf
    s = s & U("\u041e: \u041d\u0435\u0432\u0435\u0440\u043d\u044b\u0439 \u043d\u043e\u043c\u0435\u0440 \u0432\u0435\u0440\u0441\u0438\u0438 \u0432 \u043f\u0443\u0442\u0438. \u041f\u0440\u043e\u0432\u0435\u0440\u044c\u0442\u0435 \u0444\u0430\u043a\u0442\u0438\u0447\u0435\u0441\u043a\u0443\u044e") & vbCrLf
    s = s & U("   \u043f\u0430\u043f\u043a\u0443 \u0432 C:\Program Files\Autodesk\.") & vbCrLf & vbCrLf
    s = s & U("\u0412: \u042f \u043d\u0435 \u043c\u043e\u0433\u0443 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u044c Apprentice Server (\u043d\u0435\u0442 \u043f\u0440\u0430\u0432 \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0430\u0442\u043e\u0440\u0430).") & vbCrLf
    s = s & U("\u041e: Batch PDF Publisher \u0432\u0441\u0451 \u0440\u0430\u0432\u043d\u043e \u0440\u0430\u0431\u043e\u0442\u0430\u0435\u0442, \u043d\u043e \u0432 10-50 \u0440\u0430\u0437 \u043c\u0435\u0434\u043b\u0435\u043d\u043d\u0435\u0435.") & vbCrLf
    s = s & U("   \u041e\u0431\u0440\u0430\u0442\u0438\u0442\u0435\u0441\u044c \u0432 \u0432\u0430\u0448 \u043e\u0442\u0434\u0435\u043b \u0418\u0422.") & vbCrLf & vbCrLf
    s = s & U("\u0412: \u041a\u0430\u043a \u043f\u043e\u0434\u0430\u0432\u0438\u0442\u044c \u043f\u0440\u0435\u0434\u0443\u043f\u0440\u0435\u0436\u0434\u0435\u043d\u0438\u0435 \u0431\u0435\u0437 \u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0438?") & vbCrLf
    s = s & U("\u041e: \u0423\u0434\u0430\u043b\u0438\u0442\u0435 \u044d\u0442\u043e\u0442 \u043a\u043b\u044e\u0447 \u0440\u0435\u0435\u0441\u0442\u0440\u0430 \u0438 \u043f\u0435\u0440\u0435\u0437\u0430\u043f\u0443\u0441\u0442\u0438\u0442\u0435 Inventor:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_RU = s
End Function

Private Function ApprenticeReadme_PL() As String
    Dim s As String
    s = U("Batch PDF Publisher - Przewodnik instalacji Apprentice Server") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("DLACZEGO JEST TO POTRZEBNE?") & vbCrLf
    s = s & U("Autodesk Inventor 2026 i nowsze nie do\u0142\u0105czaj\u0105 ju\u017c Apprentice Server") & vbCrLf
    s = s & U("do g\u0142\u00f3wnego instalatora. Jest to teraz osobne pobranie, kt\u00f3re nale\u017cy") & vbCrLf
    s = s & U("zainstalowa\u0107 i zarejestrowa\u0107 r\u0119cznie. Bez niego Batch PDF Publisher") & vbCrLf
    s = s & U("korzysta z pe\u0142nego silnika Inventor, kt\u00f3ry jest 10-50 razy wolniejszy") & vbCrLf
    s = s & U("podczas odczytu liczby arkuszy z plik\u00f3w IDW.") & vbCrLf & vbCrLf
    s = s & U("PRZED ROZPOCZ\u0118CIEM") & vbCrLf
    s = s & U("- Zainstaluj wszystkie oczekuj\u0105ce aktualizacje Windows i uruchom ponownie.") & vbCrLf
    s = s & U("- Uprawnienia administratora s\u0105 potrzebne tylko do instalatora.") & vbCrLf
    s = s & U("- Wolne miejsce na dysku: co najmniej 2,2 GB (1,0 GB na dysku systemowym).") & vbCrLf
    s = s & U("- Zamknij wszystkie aplikacje Autodesk (Inventor, Vault, Design Assistant).") & vbCrLf & vbCrLf
    s = s & U("KROK 1 - POBRANIE") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("Wybierz wersj\u0119 odpowiadaj\u0105c\u0105 Twojemu Inventorowi:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("KROK 2 - INSTALACJA") & vbCrLf
    s = s & U("1. Kliknij dwukrotnie pobrany plik .sfx.exe.") & vbCrLf
    s = s & U("2. Post\u0119puj zgodnie z instrukcjami instalatora.") & vbCrLf
    s = s & U("3. Po zako\u0144czeniu kliknij Zako\u0144cz.") & vbCrLf
    s = s & U("4. W razie potrzeby uruchom ponownie system.") & vbCrLf & vbCrLf
    s = s & U("KROK 3 - REJESTRACJA (WYMAGANE)") & vbCrLf
    s = s & U("Otw\u00f3rz wiersz polecenia (cmd.exe) JAKO ZWYK\u0141Y U\u017bYTKOWNIK, nie jako") & vbCrLf
    s = s & U("administrator, i wykonaj:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Je\u015bli zainstalowa\u0142e\u015b Apprentice Server 2026, zast\u0105p 2027 przez 2026") & vbCrLf
    s = s & U("w powy\u017cszej \u015bcie\u017cce.") & vbCrLf & vbCrLf
    s = s & U("KROK 4 - WERYFIKACJA") & vbCrLf
    s = s & U("Uruchom Design Assistant 2027. W menu Pomoc wybierz Informacje o") & vbCrLf
    s = s & U("Design Assistant. Okno powinno pokazywa\u0107:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("KROK 5 - URUCHOM PONOWNIE INVENTOR I PRZETESTUJ") & vbCrLf
    s = s & U("1. Ca\u0142kowicie zamknij Inventor.") & vbCrLf
    s = s & U("2. Uruchom Inventor.") & vbCrLf
    s = s & U("3. Uruchom Batch PDF Publisher (Run_BatchPDFPublisher).") & vbCrLf
    s = s & U("4. Okno ostrze\u017cenia Apprentice nie powinno si\u0119 ju\u017c pojawia\u0107.") & vbCrLf
    s = s & U("5. Sprawd\u017a dziennik:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   Nie powinno by\u0107 wiersza 'WARN [PRJ_GetApprentice]'.") & vbCrLf & vbCrLf
    s = s & U("ROZWI\u0104ZYWANIE PROBLEM\u00d3W") & vbCrLf
    s = s & U("P: Okno ostrze\u017cenia nadal si\u0119 pojawia po instalacji.") & vbCrLf
    s = s & U("O: Pomini\u0119to Krok 3 (rejestracja). Apprentice Server wymaga jawnej") & vbCrLf
    s = s & U("   rejestracji COM.") & vbCrLf & vbCrLf
    s = s & U("P: Polecenie rejestracji zg\u0142asza 'nie znaleziono pliku'.") & vbCrLf
    s = s & U("O: Numer wersji w \u015bcie\u017cce jest nieprawid\u0142owy. Sprawd\u017a rzeczywisty") & vbCrLf
    s = s & U("   folder w C:\Program Files\Autodesk\.") & vbCrLf & vbCrLf
    s = s & U("P: Nie mog\u0119 zainstalowa\u0107 Apprentice Server (brak uprawnie\u0144 administratora).") & vbCrLf
    s = s & U("O: Batch PDF Publisher nadal dzia\u0142a, ale jest 10-50x wolniejszy.") & vbCrLf
    s = s & U("   Popro\u015b dzia\u0142 IT o instalacj\u0119.") & vbCrLf & vbCrLf
    s = s & U("P: Jak ukry\u0107 ostrze\u017cenie bez instalacji?") & vbCrLf
    s = s & U("O: Usu\u0144 ten klucz rejestru i uruchom ponownie Inventor:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_PL = s
End Function

Private Function ApprenticeReadme_CS() As String
    Dim s As String
    s = U("Batch PDF Publisher - Pr\u016fvodce instalac\u00ed Apprentice Server") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("PRO\u010c JE TO POT\u0158EBA?") & vbCrLf
    s = s & U("Autodesk Inventor 2026 a nov\u011bj\u0161\u00ed ji\u017e neobsahuj\u00ed Apprentice Server") & vbCrLf
    s = s & U("v hlavn\u00edm instal\u00e1toru. Nyn\u00ed je to samostatn\u00e9 sta\u017een\u00ed, kter\u00e9 je t\u0159eba") & vbCrLf
    s = s & U("ru\u010dn\u011b nainstalovat a zaregistrovat. Bez n\u011bj Batch PDF Publisher") & vbCrLf
    s = s & U("pou\u017e\u00edv\u00e1 pln\u00fd engine Inventor, kter\u00fd je p\u0159i \u010dten\u00ed po\u010dtu list\u016f z IDW") & vbCrLf
    s = s & U("soubor\u016f 10-50x pomalej\u0161\u00ed.") & vbCrLf & vbCrLf
    s = s & U("NE\u017d ZA\u010cNETE") & vbCrLf
    s = s & U("- Nainstalujte v\u0161echny \u010dekaj\u00edc\u00ed aktualizace Windows a restartujte.") & vbCrLf
    s = s & U("- Pr\u00e1va spr\u00e1vce pot\u0159ebujete pouze pro instal\u00e1tor.") & vbCrLf
    s = s & U("- Voln\u00e9 m\u00edsto na disku: alespo\u0148 2,2 GB (1,0 GB na syst\u00e9mov\u00e9m disku).") & vbCrLf
    s = s & U("- Zav\u0159ete v\u0161echny aplikace Autodesk (Inventor, Vault, Design Assistant).") & vbCrLf & vbCrLf
    s = s & U("KROK 1 - STA\u017dEN\u00cd") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("Vyberte verzi odpov\u00eddaj\u00edc\u00ed va\u0161emu Inventoru:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("KROK 2 - INSTALACE") & vbCrLf
    s = s & U("1. Dvakr\u00e1t klikn\u011bte na sta\u017een\u00fd soubor .sfx.exe.") & vbCrLf
    s = s & U("2. Postupujte podle pokyn\u016f instal\u00e1toru.") & vbCrLf
    s = s & U("3. Po dokon\u010den\u00ed klikn\u011bte na Dokon\u010dit.") & vbCrLf
    s = s & U("4. V p\u0159\u00edpad\u011b pot\u0159eby restartujte syst\u00e9m.") & vbCrLf & vbCrLf
    s = s & U("KROK 3 - REGISTRACE (POVINN\u00c9)") & vbCrLf
    s = s & U("Otev\u0159ete p\u0159\u00edkazov\u00fd \u0159\u00e1dek (cmd.exe) JAKO B\u011a\u017dN\u00dd U\u017dIVATEL, ne jako") & vbCrLf
    s = s & U("spr\u00e1vce, a spus\u0165te:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Pokud jste nainstalovali Apprentice Server 2026, nahra\u010fte 2027 \u010d\u00edslem") & vbCrLf
    s = s & U("2026 v uveden\u00e9 cest\u011b.") & vbCrLf & vbCrLf
    s = s & U("KROK 4 - OV\u011a\u0158EN\u00cd") & vbCrLf
    s = s & U("Spus\u0165te Design Assistant 2027. V nab\u00eddce N\u00e1pov\u011bda vyberte O aplikaci") & vbCrLf
    s = s & U("Design Assistant. Okno by m\u011blo zobrazovat:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("KROK 5 - RESTARTUJTE INVENTOR A ZNOVU OTESTUJTE") & vbCrLf
    s = s & U("1. Zcela zav\u0159ete Inventor.") & vbCrLf
    s = s & U("2. Spus\u0165te Inventor.") & vbCrLf
    s = s & U("3. Spus\u0165te Batch PDF Publisher (Run_BatchPDFPublisher).") & vbCrLf
    s = s & U("4. Dialog upozorn\u011bn\u00ed Apprentice by se ji\u017e nem\u011bl zobrazovat.") & vbCrLf
    s = s & U("5. Zkontrolujte protokol:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   Nem\u011bl by tam b\u00fdt \u0159\u00e1dek 'WARN [PRJ_GetApprentice]'.") & vbCrLf & vbCrLf
    s = s & U("ODSTRA\u0147OV\u00c1N\u00cd PROBL\u00c9M\u016e") & vbCrLf
    s = s & U("O: Dialog upozorn\u011bn\u00ed se st\u00e1le zobrazuje po instalaci.") & vbCrLf
    s = s & U("Odp: P\u0159esko\u010dili jste Krok 3 (registrace). Apprentice Server vy\u017eaduje") & vbCrLf
    s = s & U("     explicitn\u00ed registraci COM.") & vbCrLf & vbCrLf
    s = s & U("O: P\u0159\u00edkaz registrace hl\u00e1s\u00ed 'soubor nenalezen'.") & vbCrLf
    s = s & U("Odp: \u010c\u00edslo verze v cest\u011b je nespr\u00e1vn\u00e9. Zkontrolujte skute\u010dnou slo\u017eku") & vbCrLf
    s = s & U("     v C:\Program Files\Autodesk\.") & vbCrLf & vbCrLf
    s = s & U("O: Nemohu nainstalovat Apprentice Server (chyb\u00ed pr\u00e1va spr\u00e1vce).") & vbCrLf
    s = s & U("Odp: Batch PDF Publisher st\u00e1le funguje, ale je 10-50x pomalej\u0161\u00ed.") & vbCrLf
    s = s & U("     Po\u017e\u00e1dejte sv\u00e9 IT odd\u011blen\u00ed o instalaci.") & vbCrLf & vbCrLf
    s = s & U("O: Jak potla\u010d\u00edm upozorn\u011bn\u00ed bez instalace?") & vbCrLf
    s = s & U("Odp: Odstra\u0148te tento kl\u00ed\u010d registru a restartujte Inventor:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_CS = s
End Function

Private Function ApprenticeReadme_JA() As String
    Dim s As String
    s = U("Batch PDF Publisher - Apprentice Server \u30a4\u30f3\u30b9\u30c8\u30fc\u30eb \u30ac\u30a4\u30c9") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("\u306a\u305c\u5fc5\u8981\u304b\uff1f") & vbCrLf
    s = s & U("Autodesk Inventor 2026 \u4ee5\u964d\u3067\u306f\u3001Apprentice Server \u306f\u30e1\u30a4\u30f3") & vbCrLf
    s = s & U("\u30a4\u30f3\u30b9\u30c8\u30fc\u30e9\u30fc\u306b\u542b\u307e\u308c\u306a\u304f\u306a\u308a\u307e\u3057\u305f\u3002\u73fe\u5728\u306f\u5225\u9014\u30c0\u30a6\u30f3\u30ed\u30fc\u30c9") & vbCrLf
    s = s & U("\u3057\u3066\u624b\u52d5\u3067\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u304a\u3088\u3073\u767b\u9332\u3059\u308b\u5fc5\u8981\u304c\u3042\u308a\u307e\u3059\u3002\u306a\u3044\u5834\u5408\u3001") & vbCrLf
    s = s & U("Batch PDF Publisher \u306f\u5b8c\u5168\u306a Inventor \u30a8\u30f3\u30b8\u30f3\u3092\u4f7f\u7528\u3057\u3001IDW \u30d5\u30a1\u30a4\u30eb") & vbCrLf
    s = s & U("\u304b\u3089\u30b7\u30fc\u30c8\u6570\u3092\u8aad\u307f\u53d6\u308b\u969b\u306b 10\u301c50 \u500d\u9045\u304f\u306a\u308a\u307e\u3059\u3002") & vbCrLf & vbCrLf
    s = s & U("\u958b\u59cb\u3059\u308b\u524d\u306b") & vbCrLf
    s = s & U("- \u4fdd\u7559\u4e2d\u306e Windows \u66f4\u65b0\u3092\u3059\u3079\u3066\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u3057\u3001\u518d\u8d77\u52d5\u3057\u307e\u3059\u3002") & vbCrLf
    s = s & U("- \u7ba1\u7406\u8005\u6a29\u9650\u306f\u30a4\u30f3\u30b9\u30c8\u30fc\u30e9\u30fc\u306b\u306e\u307f\u5fc5\u8981\u3067\u3059\u3002") & vbCrLf
    s = s & U("- \u7a7a\u304d\u30c7\u30a3\u30b9\u30af\u5bb9\u91cf: \u5c11\u306a\u304f\u3068\u3082 2.2 GB (\u30b7\u30b9\u30c6\u30e0 \u30c9\u30e9\u30a4\u30d6\u306b 1.0 GB)\u3002") & vbCrLf
    s = s & U("- \u3059\u3079\u3066\u306e Autodesk \u30a2\u30d7\u30ea\u30b1\u30fc\u30b7\u30e7\u30f3 (Inventor, Vault, Design Assistant) \u3092\u9589\u3058\u307e\u3059\u3002") & vbCrLf & vbCrLf
    s = s & U("\u30b9\u30c6\u30c3\u30d7 1 - \u30c0\u30a6\u30f3\u30ed\u30fc\u30c9") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("Inventor \u306e\u30d0\u30fc\u30b8\u30e7\u30f3\u306b\u5408\u3063\u305f\u30d0\u30fc\u30b8\u30e7\u30f3\u3092\u9078\u629e\u3057\u307e\u3059:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("\u30b9\u30c6\u30c3\u30d7 2 - \u30a4\u30f3\u30b9\u30c8\u30fc\u30eb") & vbCrLf
    s = s & U("1. \u30c0\u30a6\u30f3\u30ed\u30fc\u30c9\u3057\u305f .sfx.exe \u30d5\u30a1\u30a4\u30eb\u3092\u30c0\u30d6\u30eb\u30af\u30ea\u30c3\u30af\u3057\u307e\u3059\u3002") & vbCrLf
    s = s & U("2. \u30a4\u30f3\u30b9\u30c8\u30fc\u30e9\u30fc\u306e\u6307\u793a\u306b\u5f93\u3044\u307e\u3059\u3002") & vbCrLf
    s = s & U("3. \u5b8c\u4e86\u3057\u305f\u3089 [\u5b8c\u4e86] \u3092\u30af\u30ea\u30c3\u30af\u3057\u307e\u3059\u3002") & vbCrLf
    s = s & U("4. \u8981\u6c42\u3055\u308c\u305f\u3089\u518d\u8d77\u52d5\u3057\u307e\u3059\u3002") & vbCrLf & vbCrLf
    s = s & U("\u30b9\u30c6\u30c3\u30d7 3 - \u767b\u9332 (\u5fc5\u9808)") & vbCrLf
    s = s & U("\u30b3\u30de\u30f3\u30c9 \u30d7\u30ed\u30f3\u30d7\u30c8 (cmd.exe) \u3092\u4e00\u822c\u30e6\u30fc\u30b6\u30fc\u3068\u3057\u3066\u958b\u304d\u3001") & vbCrLf
    s = s & U("\u7ba1\u7406\u8005\u3068\u3057\u3066\u3067\u306f\u306a\u304f\u3001\u6b21\u3092\u5b9f\u884c\u3057\u307e\u3059:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Apprentice Server 2026 \u3092\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u3057\u305f\u5834\u5408\u306f\u3001\u4e0a\u8a18\u306e") & vbCrLf
    s = s & U("\u30d1\u30b9\u306e 2027 \u3092 2026 \u306b\u7f6e\u304d\u63db\u3048\u307e\u3059\u3002") & vbCrLf & vbCrLf
    s = s & U("\u30b9\u30c6\u30c3\u30d7 4 - \u691c\u8a3c") & vbCrLf
    s = s & U("Design Assistant 2027 \u3092\u8d77\u52d5\u3057\u307e\u3059\u3002[\u30d8\u30eb\u30d7] \u30e1\u30cb\u30e5\u30fc\u3067") & vbCrLf
    s = s & U("[Design Assistant \u306b\u3064\u3044\u3066] \u3092\u9078\u629e\u3057\u307e\u3059\u3002\u6b21\u304c\u8868\u793a\u3055\u308c\u307e\u3059:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("\u30b9\u30c6\u30c3\u30d7 5 - Inventor \u3092\u518d\u8d77\u52d5\u3057\u3066\u518d\u30c6\u30b9\u30c8") & vbCrLf
    s = s & U("1. Inventor \u3092\u5b8c\u5168\u306b\u9589\u3058\u307e\u3059\u3002") & vbCrLf
    s = s & U("2. Inventor \u3092\u8d77\u52d5\u3057\u307e\u3059\u3002") & vbCrLf
    s = s & U("3. Batch PDF Publisher (Run_BatchPDFPublisher) \u3092\u5b9f\u884c\u3057\u307e\u3059\u3002") & vbCrLf
    s = s & U("4. Apprentice \u8b66\u544a\u30c0\u30a4\u30a2\u30ed\u30b0\u306f\u8868\u793a\u3055\u308c\u306a\u304f\u306a\u308a\u307e\u3059\u3002") & vbCrLf
    s = s & U("5. \u30ed\u30b0\u3092\u78ba\u8a8d\u3057\u307e\u3059:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   'WARN [PRJ_GetApprentice]' \u306e\u884c\u304c\u306a\u3044\u306f\u305a\u3067\u3059\u3002") & vbCrLf & vbCrLf
    s = s & U("\u30c8\u30e9\u30d6\u30eb\u30b7\u30e5\u30fc\u30c6\u30a3\u30f3\u30b0") & vbCrLf
    s = s & U("Q: \u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u5f8c\u306b\u3082\u8b66\u544a\u30c0\u30a4\u30a2\u30ed\u30b0\u304c\u8868\u793a\u3055\u308c\u307e\u3059\u3002") & vbCrLf
    s = s & U("A: \u30b9\u30c6\u30c3\u30d7 3 (\u767b\u9332) \u3092\u30b9\u30ad\u30c3\u30d7\u3057\u305f\u53ef\u80fd\u6027\u304c\u3042\u308a\u307e\u3059\u3002Apprentice Server \u306f") & vbCrLf
    s = s & U("   \u660e\u793a\u7684\u306a COM \u767b\u9332\u304c\u5fc5\u8981\u3067\u3059\u3002") & vbCrLf & vbCrLf
    s = s & U("Q: \u767b\u9332\u30b3\u30de\u30f3\u30c9\u304c '\u30d5\u30a1\u30a4\u30eb\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093' \u3092\u5831\u544a\u3057\u307e\u3059\u3002") & vbCrLf
    s = s & U("A: \u30d1\u30b9\u5185\u306e\u30d0\u30fc\u30b8\u30e7\u30f3\u756a\u53f7\u304c\u9593\u9055\u3063\u3066\u3044\u307e\u3059\u3002C:\Program Files\Autodesk\ \u4e0b\u306e") & vbCrLf
    s = s & U("   \u5b9f\u969b\u306e\u30d5\u30a9\u30eb\u30c0\u30fc\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002") & vbCrLf & vbCrLf
    s = s & U("Q: Apprentice Server \u3092\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u3067\u304d\u307e\u305b\u3093 (\u7ba1\u7406\u8005\u6a29\u9650\u306a\u3057)\u3002") & vbCrLf
    s = s & U("A: Batch PDF Publisher \u306f\u5f15\u304d\u7d9a\u304d\u52d5\u4f5c\u3057\u307e\u3059\u304c\u300110-50 \u500d\u9045\u304f\u306a\u308a\u307e\u3059\u3002") & vbCrLf
    s = s & U("   IT \u90e8\u9580\u306b\u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u3092\u4f9d\u983c\u3057\u3066\u304f\u3060\u3055\u3044\u3002") & vbCrLf & vbCrLf
    s = s & U("Q: \u30a4\u30f3\u30b9\u30c8\u30fc\u30eb\u305b\u305a\u306b\u8b66\u544a\u3092\u975e\u8868\u793a\u306b\u3067\u304d\u307e\u3059\u304b\uff1f") & vbCrLf
    s = s & U("A: \u3053\u306e\u30ec\u30b8\u30b9\u30c8\u30ea \u30ad\u30fc\u3092\u524a\u9664\u3057\u3066 Inventor \u3092\u518d\u8d77\u52d5\u3057\u307e\u3059:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_JA = s
End Function

Private Function ApprenticeReadme_KO() As String
    Dim s As String
    s = U("Batch PDF Publisher - Apprentice Server \uc124\uce58 \uac00\uc774\ub4dc") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("\uc65c \ud544\uc694\ud55c\uac00\uc694?") & vbCrLf
    s = s & U("Autodesk Inventor 2026 \uc774\uc0c1\uc5d0\uc11c\ub294 Apprentice Server\uac00 \uba54\uc778") & vbCrLf
    s = s & U("\uc124\uce58 \ud504\ub85c\uadf8\ub7a8\uc5d0 \ud3ec\ud568\ub418\uc9c0 \uc54a\uc2b5\ub2c8\ub2e4. \uc774\uc81c\ub294 \ubcc4\ub3c4\ub85c \ub2e4\uc6b4\ub85c\ub4dc\ud558\uc5ec") & vbCrLf
    s = s & U("\uc218\ub3d9\uc73c\ub85c \uc124\uce58\ud558\uace0 \ub4f1\ub85d\ud574\uc57c \ud569\ub2c8\ub2e4. \uc5c6\ub294 \uacbd\uc6b0 Batch PDF Publisher\ub294") & vbCrLf
    s = s & U("\uc804\uccb4 Inventor \uc5d4\uc9c4\uc744 \uc0ac\uc6a9\ud558\uba70 IDW \ud30c\uc77c\uc5d0\uc11c \uc2dc\ud2b8 \uc218\ub97c \uc77d\uc744 \ub54c") & vbCrLf
    s = s & U("10-50\ubc30 \ub290\ub9bd\ub2c8\ub2e4.") & vbCrLf & vbCrLf
    s = s & U("\uc2dc\uc791\ud558\uae30 \uc804\uc5d0") & vbCrLf
    s = s & U("- \ubcf4\ub958 \uc911\uc778 Windows \uc5c5\ub370\uc774\ud2b8\ub97c \ubaa8\ub450 \uc124\uce58\ud558\uace0 \ub2e4\uc2dc \uc2dc\uc791\ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("- \uad00\ub9ac\uc790 \uad8c\ud55c\uc740 \uc124\uce58 \uad00\ub9ac\uc790\uc5d0\uac8c\ub9cc \ud544\uc694\ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("- \uc5ec\uc720 \ub514\uc2a4\ud06c \uacf5\uac04: \ucd5c\uc18c 2.2 GB (\uc2dc\uc2a4\ud15c \ub4dc\ub77c\uc774\ube0c\uc5d0 1.0 GB).") & vbCrLf
    s = s & U("- \ubaa8\ub4e0 Autodesk \uc571 (Inventor, Vault, Design Assistant)\uc744 \ub2eb\uc2b5\ub2c8\ub2e4.") & vbCrLf & vbCrLf
    s = s & U("\ub2e8\uacc4 1 - \ub2e4\uc6b4\ub85c\ub4dc") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("Inventor \ubc84\uc804\uacfc \uc77c\uce58\ud558\ub294 \ubc84\uc804\uc744 \uc120\ud0dd\ud558\uc138\uc694:") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("\ub2e8\uacc4 2 - \uc124\uce58") & vbCrLf
    s = s & U("1. \ub2e4\uc6b4\ub85c\ub4dc\ud55c .sfx.exe \ud30c\uc77c\uc744 \ub354\ube14 \ud074\ub9ad\ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("2. \uc124\uce58 \uad00\ub9ac\uc790\uc758 \uc9c0\uc2dc\ub97c \ub530\ub985\ub2c8\ub2e4.") & vbCrLf
    s = s & U("3. \uc644\ub8cc \ud6c4 \ub9c8\uce68\uc744 \ud074\ub9ad\ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("4. \ud544\uc694\ud55c \uacbd\uc6b0 \uc2dc\uc2a4\ud15c\uc744 \ub2e4\uc2dc \uc2dc\uc791\ud569\ub2c8\ub2e4.") & vbCrLf & vbCrLf
    s = s & U("\ub2e8\uacc4 3 - \ub4f1\ub85d (\ud544\uc218)") & vbCrLf
    s = s & U("\uba85\ub839 \ud504\ub86c\ud504\ud2b8 (cmd.exe)\ub97c \uc77c\ubc18 \uc0ac\uc6a9\uc790\ub85c \uc5f4\uace0 \uad00\ub9ac\uc790\uac00") & vbCrLf
    s = s & U("\uc544\ub2cc \uc0c1\ud0dc\uc5d0\uc11c \ub2e4\uc74c\uc744 \uc2e4\ud589\ud569\ub2c8\ub2e4:") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("Apprentice Server 2026\uc744 \uc124\uce58\ud55c \uacbd\uc6b0 \uc704 \uacbd\ub85c\uc758 2027\uc744") & vbCrLf
    s = s & U("2026\uc73c\ub85c \ubc14\uafb8\uc138\uc694.") & vbCrLf & vbCrLf
    s = s & U("\ub2e8\uacc4 4 - \ud655\uc778") & vbCrLf
    s = s & U("Design Assistant 2027\uc744 \uc2dc\uc791\ud569\ub2c8\ub2e4. \ub3c4\uc6c0\ub9d0 \uba54\ub274\uc5d0\uc11c") & vbCrLf
    s = s & U("Design Assistant \uc815\ubcf4\ub97c \uc120\ud0dd\ud569\ub2c8\ub2e4. \ub2e4\uc74c\uc774 \ud45c\uc2dc\ub418\uc5b4\uc57c \ud569\ub2c8\ub2e4:") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("\ub2e8\uacc4 5 - Inventor \uc7ac\uc2dc\uc791 \ubc0f \uc7ac\ud14c\uc2a4\ud2b8") & vbCrLf
    s = s & U("1. Inventor\ub97c \uc644\uc804\ud788 \ub2eb\uc2b5\ub2c8\ub2e4.") & vbCrLf
    s = s & U("2. Inventor\ub97c \uc2dc\uc791\ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("3. Batch PDF Publisher (Run_BatchPDFPublisher)\ub97c \uc2e4\ud589\ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("4. Apprentice \uacbd\uace0 \ub300\ud654\uc0c1\uc790\uac00 \ub354 \uc774\uc0c1 \ud45c\uc2dc\ub418\uc9c0 \uc54a\uc544\uc57c \ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("5. \ub85c\uadf8\ub97c \ud655\uc778\ud569\ub2c8\ub2e4:") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   'WARN [PRJ_GetApprentice]' \uc904\uc774 \uc5c6\uc5b4\uc57c \ud569\ub2c8\ub2e4.") & vbCrLf & vbCrLf
    s = s & U("\ubb38\uc81c \ud574\uacb0") & vbCrLf
    s = s & U("Q: \uc124\uce58 \ud6c4\uc5d0\ub3c4 \uacbd\uace0 \ub300\ud654\uc0c1\uc790\uac00 \ud45c\uc2dc\ub429\ub2c8\ub2e4.") & vbCrLf
    s = s & U("A: \ub2e8\uacc4 3 (\ub4f1\ub85d)\uc744 \uac74\ub108\ub6f0\uc168\uc2b5\ub2c8\ub2e4. Apprentice Server\ub294 \uba85\uc2dc\uc801\uc778") & vbCrLf
    s = s & U("   COM \ub4f1\ub85d\uc774 \ud544\uc694\ud569\ub2c8\ub2e4.") & vbCrLf & vbCrLf
    s = s & U("Q: \ub4f1\ub85d \uba85\ub839\uc774 '\ud30c\uc77c\uc744 \ucc3e\uc744 \uc218 \uc5c6\uc74c'\uc744 \ubcf4\uace0\ud569\ub2c8\ub2e4.") & vbCrLf
    s = s & U("A: \uacbd\ub85c\uc758 \ubc84\uc804 \ubc88\ud638\uac00 \uc7d8\ubabb\ub418\uc5c8\uc2b5\ub2c8\ub2e4. C:\Program Files\Autodesk\ \uc544\ub798\uc758") & vbCrLf
    s = s & U("   \uc2e4\uc81c \ud3f4\ub354\ub97c \ud655\uc778\ud558\uc138\uc694.") & vbCrLf & vbCrLf
    s = s & U("Q: Apprentice Server\ub97c \uc124\uce58\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4 (\uad00\ub9ac\uc790 \uad8c\ud55c \uc5c6\uc74c).") & vbCrLf
    s = s & U("A: Batch PDF Publisher\ub294 \uacc4\uc18d \uc791\ub3d9\ud558\uc9c0\ub9cc 10-50\ubc30 \ub290\ub9bd\ub2c8\ub2e4.") & vbCrLf
    s = s & U("   IT \ubd80\uc11c\uc5d0 \uc124\uce58\ub97c \uc694\uccad\ud558\uc138\uc694.") & vbCrLf & vbCrLf
    s = s & U("Q: \uc124\uce58 \uc5c6\uc774 \uacbd\uace0\ub97c \uc5b5\uc81c\ud558\ub824\uba74?") & vbCrLf
    s = s & U("A: \uc774 \ub808\uc9c0\uc2a4\ud2b8\ub9ac \ud0a4\ub97c \uc0ad\uc81c\ud558\uace0 Inventor\ub97c \ub2e4\uc2dc \uc2dc\uc791\ud558\uc138\uc694:") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_KO = s
End Function

Private Function ApprenticeReadme_ZH() As String
    Dim s As String
    s = U("Batch PDF Publisher - Apprentice Server \u5b89\u88c5\u6307\u5357") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("\u4e3a\u4f55\u9700\u8981\u5b83\uff1f") & vbCrLf
    s = s & U("Autodesk Inventor 2026 \u53ca\u66f4\u9ad8\u7248\u672c\u4e0d\u518d\u5c06 Apprentice Server \u5305\u542b\u5728") & vbCrLf
    s = s & U("\u4e3b\u5b89\u88c5\u7a0b\u5e8f\u4e2d\u3002\u5b83\u73b0\u5728\u662f\u4e00\u4e2a\u5355\u72ec\u7684\u4e0b\u8f7d\uff0c\u5fc5\u987b\u624b\u52a8\u5b89\u88c5\u548c") & vbCrLf
    s = s & U("\u6ce8\u518c\u3002\u5982\u679c\u4e0d\u5b89\u88c5\uff0cBatch PDF Publisher \u5c06\u4f7f\u7528\u5b8c\u6574\u7684 Inventor \u5f15\u64ce\uff0c") & vbCrLf
    s = s & U("\u4ece IDW \u6587\u4ef6\u8bfb\u53d6\u56fe\u7eb8\u6570\u91cf\u65f6\u901f\u5ea6\u6162 10-50 \u500d\u3002") & vbCrLf & vbCrLf
    s = s & U("\u5f00\u59cb\u4e4b\u524d") & vbCrLf
    s = s & U("- \u5b89\u88c5\u6240\u6709\u6302\u8d77\u7684 Windows \u66f4\u65b0\u5e76\u91cd\u65b0\u542f\u52a8\u3002") & vbCrLf
    s = s & U("- \u4ec5\u5b89\u88c5\u7a0b\u5e8f\u9700\u8981\u7ba1\u7406\u5458\u6743\u9650\u3002") & vbCrLf
    s = s & U("- \u7a7a\u95f2\u78c1\u76d8\u7a7a\u95f4\uff1a\u81f3\u5c11 2.2 GB\uff08\u7cfb\u7edf\u9a71\u52a8\u5668\u4e0a 1.0 GB\uff09\u3002") & vbCrLf
    s = s & U("- \u5173\u95ed\u6240\u6709 Autodesk \u5e94\u7528\u7a0b\u5e8f\uff08Inventor\u3001Vault\u3001Design Assistant\uff09\u3002") & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9aa4 1 - \u4e0b\u8f7d") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("\u9009\u62e9\u4e0e\u60a8\u7684 Inventor \u5339\u914d\u7684\u7248\u672c\uff1a") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9aa4 2 - \u5b89\u88c5") & vbCrLf
    s = s & U("1. \u53cc\u51fb\u4e0b\u8f7d\u7684 .sfx.exe \u6587\u4ef6\u3002") & vbCrLf
    s = s & U("2. \u6309\u7167\u5b89\u88c5\u7a0b\u5e8f\u63d0\u793a\u8fdb\u884c\u64cd\u4f5c\u3002") & vbCrLf
    s = s & U("3. \u5b8c\u6210\u540e\u5355\u51fb\u201c\u5b8c\u6210\u201d\u3002") & vbCrLf
    s = s & U("4. \u5982\u6709\u63d0\u793a\uff0c\u8bf7\u91cd\u65b0\u542f\u52a8\u3002") & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9aa4 3 - \u6ce8\u518c\uff08\u5fc5\u9700\uff09") & vbCrLf
    s = s & U("\u4f5c\u4e3a\u666e\u901a\u7528\u6237\uff08\u800c\u975e\u7ba1\u7406\u5458\uff09\u6253\u5f00\u547d\u4ee4\u63d0\u793a\u7b26 (cmd.exe)\uff0c") & vbCrLf
    s = s & U("\u7136\u540e\u8fd0\u884c\uff1a") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("\u5982\u679c\u60a8\u5b89\u88c5\u7684\u662f Apprentice Server 2026\uff0c\u8bf7\u5c06\u4e0a\u8ff0\u8def\u5f84\u4e2d") & vbCrLf
    s = s & U("\u7684 2027 \u66ff\u6362\u4e3a 2026\u3002") & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9aa4 4 - \u9a8c\u8bc1") & vbCrLf
    s = s & U("\u542f\u52a8 Design Assistant 2027\u3002\u5728\u201c\u5e2e\u52a9\u201d\u83dc\u5355\u4e2d\uff0c\u9009\u62e9\u201c\u5173\u4e8e") & vbCrLf
    s = s & U("Design Assistant\u201d\u3002\u5bf9\u8bdd\u6846\u5e94\u663e\u793a\uff1a") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9aa4 5 - \u91cd\u65b0\u542f\u52a8 Inventor \u5e76\u91cd\u65b0\u6d4b\u8bd5") & vbCrLf
    s = s & U("1. \u5b8c\u5168\u5173\u95ed Inventor\u3002") & vbCrLf
    s = s & U("2. \u542f\u52a8 Inventor\u3002") & vbCrLf
    s = s & U("3. \u8fd0\u884c Batch PDF Publisher (Run_BatchPDFPublisher)\u3002") & vbCrLf
    s = s & U("4. Apprentice \u8b66\u544a\u5bf9\u8bdd\u6846\u5e94\u4e0d\u518d\u51fa\u73b0\u3002") & vbCrLf
    s = s & U("5. \u68c0\u67e5\u65e5\u5fd7\uff1a") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   \u4e0d\u5e94\u6709 'WARN [PRJ_GetApprentice]' \u884c\u3002") & vbCrLf & vbCrLf
    s = s & U("\u6545\u969c\u6392\u9664") & vbCrLf
    s = s & U("\u95ee\uff1a\u5b89\u88c5\u540e\u4ecd\u7136\u51fa\u73b0\u8b66\u544a\u5bf9\u8bdd\u6846\u3002") & vbCrLf
    s = s & U("\u7b54\uff1a\u60a8\u8df3\u8fc7\u4e86\u6b65\u9aa4 3\uff08\u6ce8\u518c\uff09\u3002Apprentice Server \u9700\u8981\u663e\u5f0f\u7684 COM") & vbCrLf
    s = s & U("     \u6ce8\u518c\u3002") & vbCrLf & vbCrLf
    s = s & U("\u95ee\uff1a\u6ce8\u518c\u547d\u4ee4\u62a5\u544a\u201c\u672a\u627e\u5230\u6587\u4ef6\u201d\u3002") & vbCrLf
    s = s & U("\u7b54\uff1a\u8def\u5f84\u4e2d\u7684\u7248\u672c\u53f7\u9519\u8bef\u3002\u8bf7\u68c0\u67e5 C:\Program Files\Autodesk\ \u4e0b") & vbCrLf
    s = s & U("     \u7684\u5b9e\u9645\u6587\u4ef6\u5939\u3002") & vbCrLf & vbCrLf
    s = s & U("\u95ee\uff1a\u6211\u65e0\u6cd5\u5b89\u88c5 Apprentice Server\uff08\u65e0\u7ba1\u7406\u5458\u6743\u9650\uff09\u3002") & vbCrLf
    s = s & U("\u7b54\uff1aBatch PDF Publisher \u4ecd\u53ef\u8fd0\u884c\uff0c\u4f46\u901f\u5ea6\u6162 10-50 \u500d\u3002") & vbCrLf
    s = s & U("     \u8bf7\u8054\u7cfb\u60a8\u7684 IT \u90e8\u95e8\u8fdb\u884c\u5b89\u88c5\u3002") & vbCrLf & vbCrLf
    s = s & U("\u95ee\uff1a\u5982\u4f55\u5728\u4e0d\u5b89\u88c5\u7684\u60c5\u51b5\u4e0b\u5c4f\u853d\u8b66\u544a\uff1f") & vbCrLf
    s = s & U("\u7b54\uff1a\u5220\u9664\u6b64\u6ce8\u518c\u8868\u952e\u5e76\u91cd\u65b0\u542f\u52a8 Inventor\uff1a") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_ZH = s
End Function

Private Function ApprenticeReadme_ZHT() As String
    Dim s As String
    s = U("Batch PDF Publisher - Apprentice Server \u5b89\u88dd\u6307\u5357") & vbCrLf
    s = String(60, "-") & vbCrLf & vbCrLf
    s = s & U("\u70ba\u4f55\u9700\u8981\u5b83\uff1f") & vbCrLf
    s = s & U("Autodesk Inventor 2026 \u53ca\u66f4\u65b0\u7248\u672c\u4e0d\u518d\u5c07 Apprentice Server \u5305\u542b\u5728") & vbCrLf
    s = s & U("\u4e3b\u5b89\u88dd\u7a0b\u5f0f\u4e2d\u3002\u5b83\u73fe\u5728\u662f\u4e00\u500b\u55ae\u7368\u7684\u4e0b\u8f09\uff0c\u5fc5\u9808\u624b\u52d5\u5b89\u88dd\u548c") & vbCrLf
    s = s & U("\u767b\u9304\u3002\u5982\u679c\u4e0d\u5b89\u88dd\uff0cBatch PDF Publisher \u5c07\u4f7f\u7528\u5b8c\u6574\u7684 Inventor \u5f15\u64ce\uff0c") & vbCrLf
    s = s & U("\u5f9e IDW \u6a94\u6848\u8b80\u53d6\u5716\u7d19\u6578\u91cf\u6642\u901f\u5ea6\u6162 10-50 \u500d\u3002") & vbCrLf & vbCrLf
    s = s & U("\u958b\u59cb\u4e4b\u524d") & vbCrLf
    s = s & U("- \u5b89\u88dd\u6240\u6709\u64f1\u7f6e\u7684 Windows \u66f4\u65b0\u4e26\u91cd\u65b0\u555f\u52d5\u3002") & vbCrLf
    s = s & U("- \u50c5\u5b89\u88dd\u7a0b\u5f0f\u9700\u8981\u7cfb\u7d71\u7ba1\u7406\u54e1\u6b0a\u9650\u3002") & vbCrLf
    s = s & U("- \u7a7a\u9592\u78c1\u789f\u7a7a\u9593\uff1a\u81f3\u5c11 2.2 GB\uff08\u7cfb\u7d71\u9a45\u52d5\u5668\u4e0a 1.0 GB\uff09\u3002") & vbCrLf
    s = s & U("- \u95dc\u9589\u6240\u6709 Autodesk \u61c9\u7528\u7a0b\u5f0f\uff08Inventor\u3001Vault\u3001Design Assistant\uff09\u3002") & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9a5f 1 - \u4e0b\u8f09") & vbCrLf
    s = s & APP_DOWNLOAD_URL & vbCrLf
    s = s & U("\u9078\u64c7\u8207\u60a8\u7684 Inventor \u5339\u914d\u7684\u7248\u672c\uff1a") & vbCrLf
    s = s & U("- Inventor 2027  ->  Apprentice Server 2027") & vbCrLf
    s = s & U("- Inventor 2026  ->  Apprentice Server 2026") & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9a5f 2 - \u5b89\u88dd") & vbCrLf
    s = s & U("1. \u96d9\u64ca\u4e0b\u8f09\u7684 .sfx.exe \u6a94\u6848\u3002") & vbCrLf
    s = s & U("2. \u6309\u7167\u5b89\u88dd\u7a0b\u5f0f\u63d0\u793a\u9032\u884c\u64cd\u4f5c\u3002") & vbCrLf
    s = s & U("3. \u5b8c\u6210\u5f8c\u6309\u4e00\u4e0b\u300c\u5b8c\u6210\u300d\u3002") & vbCrLf
    s = s & U("4. \u5982\u6709\u63d0\u793a\uff0c\u8acb\u91cd\u65b0\u555f\u52d5\u3002") & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9a5f 3 - \u767b\u9304\uff08\u5fc5\u9808\uff09") & vbCrLf
    s = s & U("\u4ee5\u4e00\u822c\u4f7f\u7528\u8005\uff08\u800c\u975e\u7cfb\u7d71\u7ba1\u7406\u54e1\uff09\u958b\u555f\u547d\u4ee4\u63d0\u793a\u5b57\u5143 (cmd.exe)\uff0c") & vbCrLf
    s = s & U("\u7136\u5f8c\u57f7\u884c\uff1a") & vbCrLf & vbCrLf
    s = s & "    """ & APP_REG_PATH_2027 & """ /install" & vbCrLf & vbCrLf
    s = s & U("\u5982\u679c\u60a8\u5b89\u88dd\u7684\u662f Apprentice Server 2026\uff0c\u8acb\u5c07\u4e0a\u8ff0\u8def\u5f91\u4e2d") & vbCrLf
    s = s & U("\u7684 2027 \u53d6\u4ee3\u70ba 2026\u3002") & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9a5f 4 - \u9a57\u8b49") & vbCrLf
    s = s & U("\u555f\u52d5 Design Assistant 2027\u3002\u5728\u300c\u8aaa\u660e\u300d\u529f\u80fd\u8868\u4e2d\uff0c\u9078\u53d6\u300c\u95dc\u65bc") & vbCrLf
    s = s & U("Design Assistant\u300d\u3002\u5c0d\u8a71\u65b9\u584a\u61c9\u986f\u793a\uff1a") & vbCrLf
    s = s & "    Build: 192   Release: 2027" & vbCrLf & vbCrLf
    s = s & U("\u6b65\u9a5f 5 - \u91cd\u65b0\u555f\u52d5 Inventor \u4e26\u91cd\u65b0\u6e2c\u8a66") & vbCrLf
    s = s & U("1. \u5b8c\u5168\u95dc\u9589 Inventor\u3002") & vbCrLf
    s = s & U("2. \u555f\u52d5 Inventor\u3002") & vbCrLf
    s = s & U("3. \u57f7\u884c Batch PDF Publisher (Run_BatchPDFPublisher)\u3002") & vbCrLf
    s = s & U("4. Apprentice \u8b66\u544a\u5c0d\u8a71\u65b9\u584a\u61c9\u4e0d\u518d\u51fa\u73fe\u3002") & vbCrLf
    s = s & U("5. \u6aa2\u67e5\u65e5\u8a8c\uff1a") & vbCrLf
    s = s & "     %TEMP%\VBADR\BatchPDF_Log_<PID>.txt" & vbCrLf
    s = s & U("   \u4e0d\u61c9\u6709 'WARN [PRJ_GetApprentice]' \u884c\u3002") & vbCrLf & vbCrLf
    s = s & U("\u6545\u969c\u6392\u9664") & vbCrLf
    s = s & U("\u554f\uff1a\u5b89\u88dd\u5f8c\u4ecd\u7136\u51fa\u73fe\u8b66\u544a\u5c0d\u8a71\u65b9\u584a\u3002") & vbCrLf
    s = s & U("\u7b54\uff1a\u60a8\u8df3\u904e\u4e86\u6b65\u9a5f 3\uff08\u767b\u9304\uff09\u3002Apprentice Server \u9700\u8981\u986f\u5f0f\u7684 COM") & vbCrLf
    s = s & U("     \u767b\u9304\u3002") & vbCrLf & vbCrLf
    s = s & U("\u554f\uff1a\u767b\u9304\u547d\u4ee4\u5831\u544a\u300c\u627e\u4e0d\u5230\u6a94\u6848\u300d\u3002") & vbCrLf
    s = s & U("\u7b54\uff1a\u8def\u5f91\u4e2d\u7684\u7248\u672c\u865f\u932f\u8aa4\u3002\u8acb\u6aa2\u67e5 C:\Program Files\Autodesk\ \u4e0b") & vbCrLf
    s = s & U("     \u7684\u5be6\u969b\u8cc7\u6599\u593e\u3002") & vbCrLf & vbCrLf
    s = s & U("\u554f\uff1a\u6211\u7121\u6cd5\u5b89\u88dd Apprentice Server\uff08\u7121\u7cfb\u7d71\u7ba1\u7406\u54e1\u6b0a\u9650\uff09\u3002") & vbCrLf
    s = s & U("\u7b54\uff1aBatch PDF Publisher \u4ecd\u53ef\u904b\u884c\uff0c\u4f46\u901f\u5ea6\u6162 10-50 \u500d\u3002") & vbCrLf
    s = s & U("     \u8acb\u806f\u7d61\u60a8\u7684 IT \u90e8\u9580\u9032\u884c\u5b89\u88dd\u3002") & vbCrLf & vbCrLf
    s = s & U("\u554f\uff1a\u5982\u4f55\u5728\u4e0d\u5b89\u88dd\u7684\u60c5\u6cc1\u4e0b\u5c4f\u853d\u8b66\u544a\uff1f") & vbCrLf
    s = s & U("\u7b54\uff1a\u522a\u9664\u6b64\u767b\u9304\u6a5f\u78bc\u4e26\u91cd\u65b0\u555f\u52d5 Inventor\uff1a") & vbCrLf
    s = s & "     HKCU\Software\VB and VBA Program Settings\VBADR01\Prefs\ApprenticeWarnAck" & vbCrLf
    ApprenticeReadme_ZHT = s
End Function

' Copy-to-clipboard support

' HELP_ApprenticeCopyHint - one-line hint to append at the end of the MsgBox,
'   telling the user that Yes copies the full guide to clipboard.
Public Function HELP_ApprenticeCopyHint(ByVal sLang As String) As String
    Select Case LCase$(sLang)
        Case "de":  HELP_ApprenticeCopyHint = U("Klicken Sie auf Ja, um die vollst\u00e4ndige Anleitung in die Zwischenablage zu kopieren.")
        Case "fr":  HELP_ApprenticeCopyHint = U("Cliquez sur Oui pour copier le guide complet dans le presse-papiers.")
        Case "it":  HELP_ApprenticeCopyHint = U("Fare clic su S\u00ec per copiare la guida completa negli appunti.")
        Case "es":  HELP_ApprenticeCopyHint = U("Haga clic en S\u00ed para copiar la gu\u00eda completa al portapapeles.")
        Case "pt":  HELP_ApprenticeCopyHint = U("Clique em Sim para copiar o guia completo para a \u00e1rea de transfer\u00eancia.")
        Case "ru":  HELP_ApprenticeCopyHint = U("\u041d\u0430\u0436\u043c\u0438\u0442\u0435 \u0414\u0430, \u0447\u0442\u043e\u0431\u044b \u0441\u043a\u043e\u043f\u0438\u0440\u043e\u0432\u0430\u0442\u044c \u043f\u043e\u043b\u043d\u043e\u0435 \u0440\u0443\u043a\u043e\u0432\u043e\u0434\u0441\u0442\u0432\u043e \u0432 \u0431\u0443\u0444\u0435\u0440 \u043e\u0431\u043c\u0435\u043d\u0430.")
        Case "pl":  HELP_ApprenticeCopyHint = U("Kliknij Tak, aby skopiowa\u0107 pe\u0142ny przewodnik do schowka.")
        Case "cs":  HELP_ApprenticeCopyHint = U("Klikn\u011bte na Ano pro zkop\u00edrov\u00e1n\u00ed \u00fapln\u00e9ho pr\u016fv\u00f3dce do schr\u00e1nky.")
        Case "ja":  HELP_ApprenticeCopyHint = U("\u300c\u306f\u3044\u300d\u3092\u30af\u30ea\u30c3\u30af\u3059\u308b\u3068\u3001\u5b8c\u5168\u306a\u30ac\u30a4\u30c9\u304c\u30af\u30ea\u30c3\u30d7\u30dc\u30fc\u30c9\u306b\u30b3\u30d4\u30fc\u3055\u308c\u307e\u3059\u3002")
        Case "ko":  HELP_ApprenticeCopyHint = U("\uc608\ub97c \ud074\ub9ad\ud558\uba74 \uc804\uccb4 \uac00\uc774\ub4dc\uac00 \ud074\ub9bd\ubcf4\ub4dc\uc5d0 \ubcf5\uc0ac\ub429\ub2c8\ub2e4.")
        Case "zh":  HELP_ApprenticeCopyHint = U("\u5355\u51fb\u201c\u662f\u201d\u5c06\u5b8c\u6574\u6307\u5357\u590d\u5236\u5230\u526a\u8d34\u677f\u3002")
        Case "zht": HELP_ApprenticeCopyHint = U("\u6309\u4e00\u4e0b\u300c\u662f\u300d\u5c07\u5b8c\u6574\u6307\u5357\u8907\u88fd\u5230\u526a\u8cbc\u7c3f\u3002")
        Case Else:  HELP_ApprenticeCopyHint = "Click Yes to copy the full guide to clipboard."
    End Select
End Function

' HELP_CopyApprenticeREADMEToClipboard   put the full guide into the clipboard.
Public Function HELP_CopyApprenticeREADMEToClipboard(ByVal sLang As String) As Boolean
    Dim sBody As String
    HELP_CopyApprenticeREADMEToClipboard = False
    On Error GoTo Fail
    sBody = HELP_ApprenticeREADMEBody(sLang)
    If Len(sBody) = 0 Then Exit Function
    HELP_CopyApprenticeREADMEToClipboard = PRJ_PutClipboard(sBody)
    Exit Function
Fail:
    HELP_CopyApprenticeREADMEToClipboard = False
End Function

' END: modHelp v9.19.9
