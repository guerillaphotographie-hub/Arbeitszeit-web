# Arbeitszeit – Web-App (iPhone, Android, Computer)

Start/Stopp, Pause, Monatsansicht, jeden Tag editierbar, Standardzeit
7:30–16:00 mit 30 min Pause, Design Hell/Dunkel/System und Excel-Export im Layout deiner Vorlage.
Die App läuft im Browser, lässt sich wie eine App auf den Home-Bildschirm legen und funktioniert danach auch offline.
Auf dem Computer zeigt sie sich in einer breiten Ansicht mit Tabelle, auf dem Handy in der schmalen Ansicht; der Wechsel passiert automatisch.

## Auf dem Handy installieren

**iPhone:** Die Adresse in **Safari** öffnen (nicht in Chrome). Unten auf das **Teilen-Symbol** (Quadrat mit Pfeil) tippen,
dann **Zum Home-Bildschirm** und **Hinzufügen**. Danach startet die App vom Home-Bildschirm wie eine normale App, ohne Browserleiste.

**Android:** Die Adresse in Chrome öffnen, Menü (drei Punkte) → **App installieren** bzw. **Zum Startbildschirm hinzufügen**.

**Beim ersten Start:** Oben rechts **Einstellungen** öffnen und Firma, Mitarbeiter und Projektkürzel eintragen
(z. B. Projekt `STL`). Dort kannst du auch Standardzeiten, das Design (System, Hell, Dunkel) und die Ansicht wählen.

## Mobile und Desktop-Ansicht

- **Automatisch (Voreinstellung):** Ab einer Fensterbreite von 900 Pixeln zeigt die App die **Desktop-Ansicht**: links Start/Stopp, Monatssummen und die Buttons,
  rechts der Monat als Tabelle mit den Spalten Datum, Zeit, Pause, Netto und Stichworte. Darunter (Handy, schmales Fenster) gilt die **mobile Ansicht** mit einer Spalte.
  Die App wechselt von selbst, auch wenn du das Fenster am Computer größer oder kleiner ziehst oder das Handy drehst.
- **Manuell umschalten:** Der kleine Button oben rechts (neben „Einstellungen“) schaltet der Reihe nach **Auto → Mobil → Desktop → Auto**.
  Die Wahl bleibt gespeichert. Dasselbe lässt sich in den Einstellungen unter „Ansicht“ einstellen.
- **Desktop-Ansicht auf dem Handy:** Die Seite wird dann breit gerechnet und verkleinert dargestellt, mit zwei Fingern kannst du hineinzoomen.
  Beim Umschalten lädt sich die Seite einmal neu, das ist Absicht. Deine Daten bleiben erhalten.
- **Am Computer** funktionieren Tastatur und Maus: Enter bestätigt Dialoge, Escape schließt sie, Zeilen lassen sich mit Tab ansteuern.

## Bedienung

- **Start / Stopp**, **Pause / Weiter**: Die Zeit wird über den Startzeitpunkt berechnet und läuft deshalb auch weiter, wenn die App geschlossen oder das Handy gesperrt ist.
  Startest du am selben Tag nach einem Stopp noch einmal, zählt die Lücke automatisch als Pause.
- **Monatsansicht**: Mit ‹ › den Monat wechseln, Antippen des Monatsnamens springt zu heute.
- **Liste oder Kalender**: Der kleine Button rechts neben dem Monatsnamen schaltet zwischen der Listenansicht und der Kalenderansicht um (Wochenraster Mo–So mit Netto-Stunden und Stichwort je Tag, am Computer zusätzlich mit den Zeiten). Ein Tipp auf einen Kalendertag öffnet wie gewohnt die Bearbeitung. Die Wahl bleibt gespeichert.
- **Tag antippen**: Start, Ende, Pause und Stichworte bearbeiten oder manuell eintragen. „Standardzeit einsetzen“ trägt 7:30–16:00 mit 30 min Pause ein,
  „Keine Arbeitszeit“ ist für Urlaub oder Feiertage, „Leeren“ löscht den Tag.
- **Mo–Fr füllen**: trägt die Standardzeit für alle leeren Werktage bis heute ein. Feiertage werden nicht erkannt, bitte danach einzeln anpassen.
- **Excel exportieren**: Vorher fragt die App nach dem Dateinamen. Vorgabe ist
  `Stunden_Projekt_Mitarbeiter_Monat_Jahr.xlsx`, aus den Einstellungen gebildet (z. B. `Stunden_STL_Max_Mustermann_10_2026.xlsx`). Du kannst den Namen
  ändern, die Endung `.xlsx` wird automatisch ergänzt, „Standardname“ stellt die Vorgabe wieder her.
  Auf dem iPhone öffnet „Teilen / Speichern“ das Teilen-Menü, dort **In Dateien sichern** wählen oder die Datei direkt per Mail oder WhatsApp senden.

## Wichtig zu den Daten

- Alle Eingaben (auch Firma und Name) bleiben **nur im Browser deines Geräts**. Sie werden nirgends hochgeladen und liegen nicht im GitHub-Repository.
- Jedes Gerät hat seine eigenen Daten. Es gibt keine Synchronisierung zwischen iPhone, Android und Computer.
- Wenn du die Website-Daten des Browsers löschst, sind die Einträge weg. Exportiere deshalb regelmäßig die Monatsauswertung als Excel-Datei.
- Updates: Geänderte Dateien im Repository einfach hochladen (gleicher Dateiname überschreibt). Die App holt sich die neue Version beim nächsten Start mit Internetverbindung.
