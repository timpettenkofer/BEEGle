# Biomass EEG Evaluation Tool (BEEva)
Automatisierter Abgleich von Daten aus dem Marktstammdatenregister mit der Zuschlagsliste aus einer EEG-Ausschreibung

Software zur internen Nutzung des Bundesverbands Bioenergie e.V. Die Nutzung erfolgt auf eigene Gefahr. Jegliche Haftung ist ausgeschlossen. Der Quellcode unterliegt der GPL-3.0 Lizenz.

BEEva vergleicht den Auszug aus dem Marktstammdatenregister mit den gemeldeten Anlagen für feste Biomasse ab 1 MW mit den Anlagen, die einen Zuschlag bekommen haben, und zwar mit Blick auf die Zuschlagsnummer. 

Wenn ein Eintrag in beiden Listen dieselbe Zuschlagsnummer hat, wird er zusammengeführt und in einer neuen Liste abgelegt. 

Achtung: Ein bekanntes Problem beim MaStR ist, dass die Daten nicht immer aktuell sind. Das liegt daran, dass die Unternehmen die Daten selbst einpflegen und die BNetzA keine Zeit hat sich darum zu kümmern, dass die Daten wirklich stimmen. 

Benutzung:
1. Zunächst fragt die Software nach dem Pfad zum Auszug aus dem Marktstammdatenregister. Diese Datei muss zwingend im .xlsx-Format vorliegen. Wenn die Daten über die Webseite der BNetzA bezogen werden, liegen sie zunächst im .csv-Format vor und müssen erst mit Hilfe von Excel umgewandelt werden.
Der Filterlink für den Abruf der aktuellen MaStR-Daten findet sich unter diesem Link.

2. Im zweiten Schritt fragt die Software nach dem Pfad zur Liste mit den Zuschlägen der Biomasseausschreibung. Die Liste für die Zuschläge der Ausschreibung im Oktober 2024 findet sich beispielsweise unter diesem Link. Ach diese Datei muss im .xlsx-Format vorliegen.

3. Als Drittes nennt man den Pfad und die Bezeichnung der Ausgabedatei. Diese erfolgt automatisch .xlsx-Datei.

