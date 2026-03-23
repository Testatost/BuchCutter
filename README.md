# BuchCutter

![alt text](https://github.com/Testatost/BuchCutter/blob/main/Programm-Screenshot.png?raw=true)

**BuchCutter** ist ein Desktop-Tool zum Zuschneiden und Aufteilen von Buchseiten, Scans, Bildern und PDF-Seiten.  


## Vorschau

BuchCutter bietet eine direkte Vorschau mit interaktiver Bearbeitung:

- Crop-Rahmen mit Ziehpunkten
- verschiebbarer und drehbarer Trennbalken
- Zoom per Mausrad
- Navigation zwischen geladenen Seiten
- Verarbeitung einzelner oder aller Einträge

## Unterstützte Dateiformate

**Eingabe:**
- PNG
- JPG / JPEG
- BMP
- TIF / TIFF
- PDF

**Ausgabe:**
- JPEG
- PNG
- TIFF
- BMP
- PDF

## Installation

### 1. Repository klonen

```bash
git clone <https://github.com/Testatost/Bild-Crop-Trenner.git>
cd BuchCutter
```

### 2. Requirements
```bash
pip install pyside6 pillow pymupdf
```

### Bedienung
Dateien laden

Über „Bilder / PDFs laden“ können Bilder oder PDFs importiert werden.
PDFs werden automatisch in einzelne Seiten aufgeteilt und als separate Einträge angezeigt.

## Crop-Bereich

Mit „Crop-Bereich“ kann ein Zuschneidebereich aktiviert werden.
Dieser lässt sich direkt in der Vorschau aufziehen, verschieben und anpassen.

## Trennbalken

Mit „Trennbalken“ kann eine Seite in mehrere Teile zerlegt werden.
Das ist besonders praktisch für Buchseiten mit zwei Textspalten oder getrennten Inhaltsbereichen.

## Smart Split

Wenn „Smart Split“ aktiviert ist, versucht BuchCutter, den Trennbalken automatisch an einer passenden inhaltlichen Trennlinie auszurichten.

## Farbmodus

Für ausgewählte Einträge kann zwischen folgenden Modi umgeschaltet werden:

Farbig (RGB)
Grau (S/W)

Zusätzlich kann ein Kontrastmodus aktiviert werden.

## Verarbeitung
Einmal bearbeiten verarbeitet nur den aktuell ausgewählten Eintrag
Alle bearbeiten verarbeitet alle aktivierten Einträge
Stopp bricht eine laufende Stapelverarbeitung ab
Ausgabeordner

Wenn kein Speicherordner gewählt wurde, speichert BuchCutter standardmäßig im Ordner der Quelldatei.

Die Ausgaben werden in Unterordnern organisiert:

Crop-Ordner/
Trenn-Ordner/

Innerhalb dieser Ordner werden die Dateien zusätzlich nach Ausgabeformat sortiert.

### Projektstruktur
```bash
BuchCutter/
├── main.py
├── logo.png
├── icon.ico
├── README.md
└── ...
```
________________________________
Disclaimer: Das Programm wurde unteranderem m.H. von ChatGPT 5 bearbeitet.
