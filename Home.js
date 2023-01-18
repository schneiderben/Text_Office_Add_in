(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Benachrichtigungsmechanismus initialisieren und ausblenden
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // Wenn nicht Excel 2016 verwendet wird, Fallbacklogik verwenden.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("Dieses Beispiel zeigt den Wert der Zellen an, die Sie in der Tabelle ausgewählt haben.");
                $('#button-text').text("Anzeigen");
                $('#button-desc').text("Zeigt die Auswahl an.");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("Dieses Beispiel hebt den größten Wert aus den Zellen hervor, die Sie in der Tabelle ausgewählt haben.");
            $('#button-text').text("Hervorheben");
            $('#button-desc').text("Hebt die größte Zahl hervor.");
                
            loadSampleData();

            // Fügt einen Klickereignishandler für die Hervorhebungsschaltfläche hinzu.
            $('#highlight-button').click(hightlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Führt einen Batchvorgang für das Excel-Objektmodell aus.
        Excel.run(function (ctx) {
            // Erstellt ein Proxyobjekt für die aktive Blattvariable
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Reiht einen Befehl zum Schreiben der Beispieldaten in das Arbeitsblatt in die Warteschlange ein.
            sheet.getRange("B3:D5").values = values;

            // Führt die in die Warteschlange eingereihten Befehle aus und gibt eine Zusage zum Angeben des Abschlusses der Aufgabe zurück.
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function hightlightHighestValue() {
        // Führt einen Batchvorgang für das Excel-Objektmodell aus.
        Excel.run(function (ctx) {
            // Erstellt ein Proxyobjekt für den ausgewählten Bereich und lädt seine Eigenschaften
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Führt den in die Warteschlange eingereihten Befehl aus und gibt eine Zusage zum Angeben des Abschlusses der Aufgabe zurück.
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Sucht nach der hervorzuhebenden Zelle.
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Hebt die Zelle hervor.
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('Der ausgewählte Text lautet:', '"' + result.value + '"');
                } else {
                    showNotification('Fehler', result.error.message);
                }
            });
    }

    // Eine Hilfsfunktion zur Behandlung von Fehlern.
    function errorHandler(error) {
        // Stellen Sie immer sicher, dass kumulierte Fehler abgefangen werden, die bei der Ausführung von "Excel.run" auftreten.
        showNotification("Fehler", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Eine Hilfsfunktion zum Anzeigen von Benachrichtigungen.
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
