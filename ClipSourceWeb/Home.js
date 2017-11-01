﻿
(function () {
    "use strict";

    var messageBanner;

    var clips = [];

    var clipIndex = 0;

    var bibliographyAdded = false;

    var currentCitationNumber = 1;

    var currentFootNoteNumber = 1;


    // Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialisiert den FabricUI-Benachrichtigungsmechanismus und blendet ihn aus.
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // Wenn nicht Word 2016 verwendet wird, Fallbacklogik verwenden.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                return;
            }

            //loadSampleData();

            console.log("INIT");



            // Fügt einen Klickereignishandler für die Hervorhebungsschaltfläche hinzu.
            getClipSource();
            //getClipCitations();

            $('#back-button').on('click', showPreviousClip);
            $('#forward-button').on('click', showNextClip);


            $('#source-footnote-button').on('click', insertFoot);
            $('#cite-button').on('click', insertCitation);
            
            $('#citation-bibliography-button').on('click', insertBibliography);






            Word.run(function (context) {
                // Erstellt ein Proxyobjekt für den Dokumenttext.
                var body = context.document.body;


                // Synchronisiert den Dokumentzustand durch Ausführen von in die Warteschlange eingereihten Befehlen und gibt eine Zusage zum Anzeigen des Abschlusses der Aufgabe zurück.
                return context.sync();
            })
                .catch(errorHandler);

        });
    };

    function getSuperScript() {
        switch (currentFootNoteNumber) {
            case 1: return '\u00B9';
            case 2: return '\u00b2';
            case 3: return '\u00b3';
            case 4: return '\u2074';
            case 5: return '\u2075';
            case 6: return '\u2076';
            case 7: return '\u2077';
            case 8: return '\u2078';
            case 9: return '\u2079';
            case 0: return '\u2080';
        }
    };

    function getCitationNumber() {
        return ' [' + currentCitationNumber + ']';
    }

    function insertCreditline() {
        Word.run(function (context) {

            var range = context.document.getSelection();

            var creditline = range.insertText(
                clips[clipIndex]['citations'] + ' ' + 'Quelle: ' + clips[clipIndex]['source'],
                Word.InsertLocation.replace);

            creditline.styleBuiltIn = Word.Style.caption;

            return context.sync();
        })
            .catch(errorHandler);
    }

    function insertBibliographyAfterCitation(forCitation) {

        if ($('#citation-bibliography-button').text() == 'Creditline erstellen') {
            insertCreditline();
        }
        else {

            Word.run(function (context) {
                if (bibliographyAdded) {
                    var range = context.document.getSelection();
                    var bibControls = context.document.contentControls.getByTag("bibliography");
                    context.load(bibControls);

                    forCitation.insertText(getCitationNumber(), 'after');

                    return context.sync()
                        .then(function () {


                            var citation = bibControls.items[0].insertParagraph('[' + currentCitationNumber + ']' + ' ' + clips[clipIndex]['citations'], 'end');
                            citation.spaceAfter = 20;
                            //citation.insertText('\n', 'end');
                            citation.styleBuiltIn = Word.Style.bibliography;
                            currentCitationNumber++;


                            return context.sync();

                        });
                }
                else {
                    var range = context.document.body.paragraphs.getLast();
                    var number = forCitation.insertText(getCitationNumber(), 'after');
                 

                    number.insertBreak(Word.BreakType.page, "after");
                    var heading = range.insertText('Literaturverzeichnis', Word.InsertLocation.end);
                    bibliographyAdded = true;

                    var bibliography = heading.insertContentControl('end');
                    //range.insertBreak(Word.BreakType.page, "after");

                    bibliography.tag = 'bibliography';
                    bibliography.styleBuiltIn = Word.Style.bibliography;
                    var citation = bibliography.insertParagraph('[' + currentCitationNumber + ']' + ' ' + clips[clipIndex]['citations'], 'end');
                    citation.spaceAfter = 20;
                    //citation.insertText('\n', 'end');
                    citation.styleBuiltIn = Word.Style.bibliography;
                    currentCitationNumber++;
                    $('#citation-bibliography-button').text('Zum Literaturverzeichnis hinzufügen');

                    heading.styleBuiltIn = Word.Style.heading2;




                    return context.sync();
                }


            })
                .catch(errorHandler);
        }
    }

    function insertBibliography() {

        if ($('#citation-bibliography-button').text() == 'Creditline erstellen') {
            insertCreditline();
        }
        else {

            Word.run(function (context) {
                if (bibliographyAdded) {
                    var range = context.document.getSelection();
                    var bibControls = context.document.contentControls.getByTag("bibliography");
                    context.load(bibControls);

                    range.insertText(getCitationNumber(), 'after');

                    return context.sync()
                        .then(function () {


                            var citation = bibControls.items[0].insertParagraph('[' + currentCitationNumber + ']' + ' ' + clips[clipIndex]['citations'], 'end');
                            citation.spaceAfter = 20;
                            //citation.insertText('\n', 'end');
                            citation.styleBuiltIn = Word.Style.bibliography;
                            currentCitationNumber++;


                            return context.sync();

                        });
                }
                else {
                    var range = context.document.getSelection();
                    range.insertText(getCitationNumber());
                    if (!bibliographyAdded) {

                        var heading = range.insertText('Literaturverzeichnis', Word.InsertLocation.end);
                        heading.styleBuiltIn = Word.Style.heading2;
                        bibliographyAdded = true;
                    }

                    //var picture = context.document.body.paragraphs.getFirst().insertBreak(Word.BreakType.page, "before");
                    var bibliography = range.insertContentControl('after');
                    bibliography.insertBreak(Word.BreakType.page, "before");


                    bibliography.tag = 'bibliography';
                    bibliography.styleBuiltIn = Word.Style.bibliography;
                    var citation = bibliography.insertParagraph('[' + currentCitationNumber + ']' + ' ' + clips[clipIndex]['citations'], 'end');
                    citation.spaceAfter = 20;
                    //citation.insertText('\n', 'end');
                    citation.styleBuiltIn = Word.Style.bibliography;
                    currentCitationNumber++;
                    $('#citation-bibliography-button').text('Zum Literaturverzeichnis hinzufügen');



                    return context.sync();
                }


            })
                .catch(errorHandler);
        }
    }

    function insertCitation() {
        Word.run(function (context) {

            var range = context.document.getSelection();

            /*if (clips[clipIndex]['citations'] != 'no citations in clipboard') {
                var citation = range.insertText('"' + clips[clipIndex]['content']  + '"',
                    Word.InsertLocation.replace);
                citation.styleBuiltIn = Word.Style.quote;
                insertBibliography();
            }
            else {
                var citation = range.insertText(
                    '"' + clips[clipIndex]['content'] + '"' + getSuperScript(),
                    Word.InsertLocation.replace);
                citation.styleBuiltIn = Word.Style.quote;
                insertFoot();

            }*/

            var citation = range.insertText(
                '"' + clips[clipIndex]['content'] + '"',
                Word.InsertLocation.replace);
            citation.styleBuiltIn = Word.Style.quote;

            citation.tag = 'citation';

            //var citationControls = context.document.contentControls.getByTag("citation");

            //citation.insertText('[1]', 'after');
            insertBibliographyAfterCitation(citation);

            return context.sync();
        })
            .catch(errorHandler);
    }

    function insertFoot() {
        Word.run(function (context) {

            // Create a proxy sectionsCollection object.
            var mySections = context.document.sections;

            var range = context.document.getSelection();

            range.insertText(getSuperScript(), 'end');

            // Queue a commmand to load the sections.
            context.load(mySections, 'body/style');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                // Create a proxy object the primary footer of the first section. 
                // Note that the footer is a body object.

                var myFooter = mySections.items[0].getFooter("primary");

                //myFooter.clear();

                // Queue a command to insert text at the end of the footer.
                myFooter.insertParagraph(getSuperScript(1) + ' ' + 'Quelle: ' + clips[clipIndex]['source'], Word.InsertLocation.end);
                myFooter.spaceAfter = 10;
                //myFooter.insertBreak('after');

                // Queue a command to wrap the header in a content control.
                myFooter.insertContentControl();

                myFooter.styleBuiltIn = Word.Style.footnoteText;

                currentFootNoteNumber++;


                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log("Added a footer to the first section.");
                });
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function showNextClip() {
        if (clips.length > 0) {

            if (clipIndex != clips.length - 1) {
                clipIndex = clipIndex + 1;
            }

            if (clipIndex > -1) {
                var source = clips[clipIndex]['source'];
                var content = clips[clipIndex]['content'];
                var citation = clips[clipIndex]['citations'];

                console.log(source);
                console.log(content);
                console.log(citation);

                $('#txtCitation').text(citation);
                $('#txtContent').text(content);
                $('#txtSource').text(source);

            }

        }
    }

    function showPreviousClip() {
        if (clips.length > 0) {

            if (clipIndex - 1 > -1) {
                clipIndex = clipIndex - 1;
            }

            if (clipIndex > -1) {
                var source = clips[clipIndex]['source'];
                var content = clips[clipIndex]['content'];
                var citation = clips[clipIndex]['citations'];

                console.log(source);
                console.log(content);
                console.log(citation);

                $('#txtCitation').text(citation);
                $('#txtContent').text(content);
                $('#txtSource').text(source);

            }

        }
    }

    function handlePaste() {
        console.log("PASTE");
    }

    var interval;

    function responseToString(response) {
        console.log(response[1]);


    }

    function getClipSource() {
        var url = 'https://localhost:5000/source.py';
        $.ajaxSetup({
            cache: false
        });
        $.ajax({
            type: 'GET',

            url: url,
            success: function (response) {
                var json = JSON.parse(response);
                var dictObject = {}

                var source = json.source;

                var content = json.content;

                var dict = {}

                dict['source'] = source;
                dict['content'] = content;

                //console.log(source);       

                //console.log(response[1]);



                //interval = setTimeout(getClipSource, 5000);

                $.ajax({
                    type: 'GET',
                    url: 'https://localhost:5000/citations.py',
                    success: function (response) {

                        console.log(content);
                        if (content == 'image with metadata') {
                            if (clips.slice(-1)[0]['source'] == source) {
                                console.log("was same image");
                            }
                            else {
                                var json = JSON.parse(response);
                                var name = json.citations.name;
                                var description = json.citations.description;
                                var author = json.citations.author;
                                var copyright = json.citations.copyright;

                                var meta_data_string = description + ' ' + author + '. ' + copyright + '.';

                                dict['citations'] = meta_data_string;

                                $('#txtCitation').text(meta_data_string);
                                $('#txtContent').text(content);
                                $('#txtSource').text(source);

                                $('#citation-bibliography-button').text('Creditline erstellen');

                                clips.push(dict);
                                clipIndex = clips.length - 1;
                            }
                            
                       
                        }
                        else {
                            if (bibliographyAdded) {
                                $('#citation-bibliography-button').text('Zum Literaturverzeichnis hinzufügen');
                            }
                            else {
                                $('#citation-bibliography-button').text('Literaturverzeichnis erstellen');
                            }

                            var json = JSON.parse(response);
                            var citation = json.APA;
                            console.log(citation);


                            dict['citations'] = citation;



                            if (clips.length == 0) {
                                console.log("length is 0");
                                $('#txtCitation').text(citation);
                                $('#txtContent').text(content);
                                $('#txtSource').text(source);
                                clips.push(dict);

                                clipIndex = clips.length - 1;
                                console.log(clipIndex);
                            }
                            else if (clips.slice(-1)[0]['source'] == content || clips.slice(-1)[0]['citations'] == content || (clips.slice(-1)[0]['content'] == content) && clips.slice(-1)[0]['source'] == source) {
                                console.log(clips.slice(-1)[0]['content']);
                            }
                            else {
                                $('#txtCitation').text(citation);
                                $('#txtContent').text(content);
                                $('#txtSource').text(source);
                                clips.push(dict);
                                clipIndex = clips.length - 1;
                            }
                        }
                        citation_interval = setTimeout(getClipSource, 5000);
                    },
                    error: function (response) {

                        citation_interval = setTimeout(getClipSource, 5000);
                        console.log(console.error(response));

                    }
                });

            },

            error: function (response) {
                console.log(console.error(response));
                citation_interval = setTimeout(getClipSource, 5000);

                return response;
            }
        });
        /*var source = $.get(url, function (responseText) {
            console.log(responseText);
        });*/


    }

    var citation_interval
    function getClipCitations() {
        var url = 'https://localhost:5000/citations.py';
        console.log("Getting Citations");
        $.ajaxSetup({
            cache: false
        });
        $.ajax({
            type: 'GET',
            url: url,
            success: function (response) {
                var json = JSON.parse(response);
                var citation = json.APA;
                console.log(citation);
                //console.log(response[1]);
                $('#txtCitation').text(citation);

                citation_interval = setTimeout(getClipCitations, 5000);

                console.log(response);
                return response;
            },
            error: function (response) {

                citation_interval = setTimeout(getClipCitations, 5000);
                console.log(console.error(response));

                return response;
            }
        });
        /*var source = $.get(url, function (responseText) {
            console.log(responseText);
        });*/


    }

    function inserText(text) {
        Word.run(function (context) {
            // Erstellt ein Proxyobjekt für den Dokumenttext.
            var body = context.document.body;

            // Reiht einen Befehl zum Löschen des Inhalts des Texts in die Warteschlange ein.
            body.clear();
            // Reiht einen Befehl zum Einfügen von Text am Ende des Word-Dokumenttexts in die Warteschlange ein.

            body.insertText(
                text,
                Word.InsertLocation.end);
            body.insertContentControl();

            // Synchronisiert den Dokumentzustand durch Ausführen von in die Warteschlange eingereihten Befehlen und gibt eine Zusage zum Anzeigen des Abschlusses der Aufgabe zurück.
            return context.sync();
        })
            .catch(errorHandler);
    }

    function insertTestCitation() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;

            // Queue a command to load content control properties.
            context.load(thisDocument, 'contentControls/id, contentControls/text, contentControls/tag');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                if (thisDocument.contentControls.items.length !== 0) {
                    for (var i = 0; i < thisDocument.contentControls.items.length; i++) {
                        console.log(thisDocument.contentControls.items[i].id);
                        console.log(thisDocument.contentControls.items[i].text);
                        console.log(thisDocument.contentControls.items[i].tag);
                    }
                } else {
                    console.log('No content controls in this document.');
                }
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    // Add data (HTML) to the current document selection
    function addHtml() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the paragraphs collection.
            var paragraphs = context.document.body.paragraphs;

            // Queue a commmand to load the style property for the top 2 paragraphs.
            context.load(paragraphs, { select: 'style', top: 2 });

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                // Queue a a set of commands to get the OOXML of the first paragraph.
                var ooxml = paragraphs.items[0].getOoxml();

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Paragraph OOXML: ' + ooxml.value);
                });
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function insertFootnote() {
        // Run a batch operation against the Word object model.
        console.log("insert footnote");
        Word.run(function (context) {

            // Create a proxy sectionsCollection object.
            var mySections = context.document.sections;


            // Queue a commmand to load the sections.
            context.load(mySections, 'body/style');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                // Create a proxy object the primary footer of the first section. 
                // Note that the footer is a body object.

                var myFooter = mySections.items[0].getFooter("primary");

                // Queue a command to insert text at the end of the footer.
                myFooter.insertText(source, Word.InsertLocation.end);

                // Queue a command to wrap the header in a content control.
                myFooter.insertContentControl();

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log("Added a footer to the first section.");
                });
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    function loadSampleData() {
        console.log("Loading sample");
        // Führt einen Batchvorgang für das Word-Objektmodell aus.
        Word.run(function (context) {
            // Erstellt ein Proxyobjekt für den Dokumenttext.
            var body = context.document.body;

            // Reiht einen Befehl zum Löschen des Inhalts des Texts in die Warteschlange ein.
            body.clear();
            // Reiht einen Befehl zum Einfügen von Text am Ende des Word-Dokumenttexts in die Warteschlange ein.

            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);
            body.insertContentControl();

            // Synchronisiert den Dokumentzustand durch Ausführen von in die Warteschlange eingereihten Befehlen und gibt eine Zusage zum Anzeigen des Abschlusses der Aufgabe zurück.
            return context.sync();
        })
            .catch(errorHandler);
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Reiht einen Befehl zum Abrufen der aktuellen Auswahl in die Warteschlange ein und
            // erstellt dann ein Proxybereichsobjekt mit den Ergebnissen.
            var range = context.document.getSelection();

            // Diese Variable enthält die Suchergebnisse für das längste Wort.
            var searchResults;

            // Reiht einen Befehl in die Warteschlange ein, um das Bereichsauswahlergebnis zu laden.
            context.load(range, 'text');

            // Synchronisiert den Zustand des Dokuments durch Ausführen der in die Warteschlange eingereihten Befehle
            // und gibt eine Zusage zum Angeben des Abschlusses der Aufgabe zurück.
            return context.sync()
                .then(function () {
                    // Ruft das längste Wort aus der Auswahl ab.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Reiht einen Suchbefehl in die Warteschlange ein.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Reiht einen Befehl zum Laden der Eigenschaft "font" der Ergebnisse in die Warteschlange ein.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Reiht einen Befehl zum Hervorheben der Suchergebnisse in die Warteschlange ein.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Gelb
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('Der ausgewählte Text lautet:', '"' + result.value + '"');
                } else {
                    showNotification('Fehler:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Fehler:", error);
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

    function initializeUI() {
        const textFieldElements = document.querySelectorAll(".ms-TextField");
        for (let i = 0; i < textFieldElements.length; i++) {
            new fabric['TextField'](textFieldElements[i]);
        }

        const dropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
        for (let i = 0; i < dropdownHTMLElements.length; ++i) {
            new fabric['Dropdown'](dropdownHTMLElements[i]);
        }

        const checkBoxElements = document.querySelectorAll(".ms-CheckBox");
        acceptCheckBox = new fabric['CheckBox'](checkBoxElements[0]);

        const choiceFieldGroupElements = document.querySelectorAll(".ms-ChoiceFieldGroup");
        for (let i = 0; i < choiceFieldGroupElements.length; i++) {
            new fabric['ChoiceFieldGroup'](choiceFieldGroupElements[i]);
        }

        const toggleElements = document.querySelectorAll(".ms-Toggle");
        for (let i = 0; i < toggleElements.length; i++) {
            new fabric['Toggle'](toggleElements[i]);
        }
    }

})();
