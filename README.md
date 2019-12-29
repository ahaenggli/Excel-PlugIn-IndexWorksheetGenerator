# Excel-PlugIn-IndexWorksheetGenerator (Deutsch)
Das PlugIn dient dazu, mittels F5-Taste ein Inhaltsverzeichnis (Übersichtsblatt, Index, Summary, Help-Seite,...) für die geöffnete Excel-Mappe zu erstellen.
Im Verzeichnis werden sämtliche sichtbaren Register aufgeführt und sind direkt verlinkt. Optional können weitere Attribute wie Beschreibung, Erstelldatum, Status, uvw. gesetzt werden.
Die Standard-Einstellungen lassen sich via F5 > Button "config" einsehen und editieren.

## Wie installiere ich ein Excel-Add-In?
PowerShell: 
```PowerShell
Invoke-WebRequest "https://github.com/ahaenggli/Excel-PlugIn-TableOfContents/releases/latest/download/TableOfContentsWorksheetGenerator.xlam" -OutFile $env:APPDATA"\Microsoft\AddIns\TableOfContentsWorksheetGenerator.xlam"
$xl=New-Object -ComObject excel.application
$xl.Application.AddIns | ?{$_.Name -eq 'TableOfContentsWorksheetGenerator.xlam'} | %{$_.Installed=$true}
$xl.Quit()
```
Manuell in Excel:  
https://support.office.com/de-de/article/add-or-remove-add-in-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460  

## Wie kann ich das Inhaltsverzeichnis erstellen?
Einfach die F5-Taste anklicken.

## Kann ich das Inhaltsverzeichnis aktualisieren? (Ich habe einen Blattnamen geändert oder so ...)
Einfach die F5-Taste anklicken.

## Kann ich die Eigenschaften direkt im Übersichtsblatt bearbeiten?
Natürlich.

## Kann ich die Standard-Einstellungen ändern?
Die Standard-Einstellungen lassen sich via F5 > Button "config" einsehen und editieren.

## Was kann angepasst werden?
Der Name des Übersichtsblattes, welche Eigenschaften auf jedem Blatt verfügbar sind, welche Attribute Index stehen sollen und natürlich deren Reihenfolge.

## Ich habe F5 für das Excel-GoTo verwendet
Stattdessen kann [CTRL]+G verwendet werden.

# Excel PlugIn IndexWorksheetGenerator (English)
The plug-in is used to create a table of contents (summary sheet, index, help page, ...) for the opened Excel workbook by using the F5 key.
In the table of contents sheet are all other visible sheets listed and directly linked. 
Optionally, further attributes such as description, date of creation, status, and much more can be added.
The default settings can be viewed and edited via F5 > "config"-button.

## How to install an Excel add-in
PowerShell: 
```PowerShell
Invoke-WebRequest "https://github.com/ahaenggli/Excel-PlugIn-TableOfContents/releases/latest/download/TableOfContentsWorksheetGenerator.xlam" -OutFile $env:APPDATA"\Microsoft\AddIns\TableOfContentsWorksheetGenerator.xlam"
$xl=New-Object -ComObject excel.application
$xl.Application.AddIns | ?{$_.Name -eq 'TableOfContentsWorksheetGenerator.xlam'} | %{$_.Installed=$true}
$xl.Quit()
```
Excel:  
https://support.office.com/en-us/article/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460

## How can I generate the table of contents worksheet?
Just click your F5 key.

## Can I refresh the table of contents sheet? (I changed a sheet name or so...)
Sure, just click your F5 key in the updated sheet.
Tip: If you open your index sheet and press F5, it refreshes without dialog.

## Can I edit the properties directly in the summary sheet?
Sure.

## Can I change the default properties?
Sure, just click F5 and use the config button to customize it.

## What can be customized?
The name of the index sheet, wich properties are available on each sheet, which properties you want in the index sheet and of course their order.

## I used F5 for the Excel's GoTo
You can use [CTRL]+G instead.
