# Outlook Quick Move Add-in

Dieses Projekt liefert ein Outlook-VSTO-Add-in für klassisches Outlook (Windows), das markierte E-Mails per Shortcut in einen Zielordner verschiebt. Der Quick-Move-Dialog unterstützt eine Live-Suche über alle Ordner in allen eingebundenen Postfächern (inkl. Shared Mailboxes und Zusatz-Stores) und ist vollständig per Tastatur bedienbar.

## Funktionsumfang

* Ribbon-Tab **Quick Move** mit:
  * **Quick Move öffnen**
  * **Einstellungen**
  * **Letzte Ziele** (Dropdown)
  * **Favoriten** (Dropdown)
* Konfigurierbarer Shortcut (Standard: `Ctrl+Shift+M`), mit Validierung.
* Quick-Move-Dialog mit Live-Suche, Ergebnislimit (Top 50), Tastatursteuerung (Pfeile, Enter, Esc, Ctrl+Enter).
* Ordner aus allen Stores inkl. Shared Mailboxes (sofern Outlook Zugriff gewährt).
* Favoriten- und „Letzte Ziele“-Priorisierung in der Trefferliste.
* Einstellungsdialog mit Favoritenverwaltung, Anzahl „letzte Ziele“, Filteroptionen und Ordner-Cache-Refresh.
* Settings werden pro Benutzer unter `%AppData%\QuickMoveOutlook\settings.json` gespeichert.
* Logging nach `%AppData%\QuickMoveOutlook\quickmove.log`.

## Projektstruktur

* `outlook-extension/ThisAddIn.cs` – Einstiegspunkt, Events, Dialogsteuerung, Move-Logik.
* `outlook-extension/Ribbon/QuickMoveRibbon.cs` – Ribbon-Definition und Callback-Logik.
* `outlook-extension/Services/*` – FolderCache, Suche, Settings, Hotkey, Logging.
* `outlook-extension/UI/*` – Quick-Move-Dialog, Favoriten-Picker, Einstellungen.
* `outlook-extension/Models/*` – Datenmodelle (Settings, FolderInfo).

## Installation / Voraussetzungen

* Windows 10/11
* Outlook Desktop (classic)
* .NET Framework 4.8.1
* VSTO Runtime (wird im Projekt als Bootstrapper referenziert)

### Installation

1. Projekt in Visual Studio öffnen (`outlook-extension.sln`).
2. Build für `AnyCPU` (Debug/Release) ausführen.
3. Das Add-in wird beim Debuggen automatisch registriert; für produktive Installation ist ein MSI-Installer vorgesehen (siehe Lastenheft).

## Bedienung

### Quick Move per Shortcut

1. E-Mail(s) im Explorer markieren (oder Mail im Inspector geöffnet).
2. Shortcut drücken (Standard `Ctrl+Shift+M`).
3. Im Suchfeld den Zielordner tippen.
4. Mit `Enter` verschieben, mit `Ctrl+Enter` verschieben und Dialog geöffnet lassen.
5. `Esc` schließt ohne Aktion.

### Quick Move per Ribbon

* **Quick Move öffnen** startet den Dialog.
* **Letzte Ziele** und **Favoriten** verschieben sofort in den gewählten Ordner.
* **Einstellungen** öffnet die Konfiguration.

## Einstellungen

* **Shortcut**: Einfach im Feld drücken (z. B. `Ctrl+Shift+M`).
* **Favoriten**: Über die Suche hinzufügen, entfernen, Reihenfolge ändern.
* **Anzahl letzte Ziele**: Begrenzung der Liste.
* **Nur Unterordner von Posteingang anzeigen**: Filter in der Suche.
* **Archiv/Online-Archive anzeigen**: Schaltet die Archiv-Stores in der Ordnerliste ein/aus.
* **Ordnerliste neu laden**: Cache der Ordnerliste aktualisieren (z. B. bei neuen Postfächern).

## Cache/Performance

* Ordner werden beim Add-in-Start gelesen und in einem Cache gehalten.
* Änderungen an Stores (Hinzufügen/Entfernen) triggern einen Refresh.
* Suche arbeitet im Cache und liefert max. 50 Treffer mit Priorisierung:
  1. Favoriten
  2. Letzte Ziele
  3. Exakte Ordnernamen
  4. Pfad-Teiltreffer

## Tests

Das Add-in erfordert Outlook Desktop. Ein automatisierter Integrationstest ist hier nicht enthalten. Für lokale Tests:

1. Outlook starten.
2. Add-in laden (Debug-Session in Visual Studio).
3. Funktionalität gemäß Lastenheft prüfen (Shortcut, Live-Suche, Move, Ribbon).
