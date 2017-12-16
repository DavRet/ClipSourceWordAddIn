
(function () {


    var messageBanner;

    // Array of clipboard contents
    var clips = [];

    // Index of current clipboard content (gets updated continuously)
    var clipIndex = 0;

    // Variable for checking if bibliography was already added
    var bibliographyAdded = false;

    // Citation iteration number
    var currentCitationNumber = 1;
    // Footnote iteration number
    var currentFootNoteNumber = 1;
    // Image iteration number
    var currentImageNumber = 1;
    // Interval time for checking if new sources are available
    var intervalTime = 1000;
    // Current image as Base64
    var base64image;

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

            // Loads sample data for user study (not used anymore)
            //loadSampleData();

            // Checks of new clipboard sources in given interval
            getClipSource();
        
            // Setup click listeners
            $('#back-button').on('click', showPreviousClip);
            $('#forward-button').on('click', showNextClip);
            $('#source-footnote-button').on('click', insertFoot);
            $('#cite-button').on('click', insertCitation);
            $('#citation-bibliography-button').on('click', insertBibliography);
            $('#image-container').hide();
            $('#insert-image-button').on('click', insertImage);


            Word.run(function (context) {
                // Erstellt ein Proxyobjekt für den Dokumenttext.
                var body = context.document.body;
                // Synchronisiert den Dokumentzustand durch Ausführen von in die Warteschlange eingereihten Befehlen und gibt eine Zusage zum Anzeigen des Abschlusses der Aufgabe zurück.
                return context.sync();
            })
                .catch(errorHandler);

        });
    };

    /*
    Gets images as Base64 image
    */
    function getBase64Image(img) {
        var canvas = document.createElement("canvas");
        canvas.width = img.width;
        canvas.height = img.height;
        var ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0);
        var dataURL = canvas.toDataURL("image/png");
        return dataURL.replace(/^data:image\/(png|jpg);base64,/, "");
    }

    /*
    Sets the current Base64 image
    */
    function setBase64Image(url) {
        toDataURL(url, function (dataUrl) {
            console.log('BASE64 RESULT:', dataUrl);
            var base64 = dataUrl.split(',')[1];
            base64image = base64;
        })    
    }

    /*
    URL to data for Base64 image
    */
    function toDataURL(url, callback) {
        url = 'https:' + url.split(':')[1];
        console.log("TO DATA URL", url);
        var xhr = new XMLHttpRequest();
        xhr.onload = function () {
            var reader = new FileReader();
            reader.onloadend = function () {
                callback(reader.result);
            }
            reader.readAsDataURL(xhr.response);
        };
        xhr.open('GET', url);
        xhr.responseType = 'blob';
        xhr.send();
    }

    /*
    Inserts images in the document with caption
    */
    function insertImage() {
    
        Word.run(function (context) {
            // Gets current cursor position
            var range = context.document.getSelection();
            // Inserts Base64 image 
            var image = range.insertParagraph("", "after").insertInlinePictureFromBase64(base64image, "start");

            // Sets image hyperlink to source URL
            var imageSource = clips[clipIndex]['source'];
            image.hyperlink = imageSource;

            if (imageSource.indexOf('data:image/') != -1) {
                imageSource = 'No URL found';
            }
            // Builds the caption text 
            var captionText = ''
            // Caption text for normal images
            if (clips[clipIndex]['content'] == 'image') {
                captionText = "Figure " + getImageCaptionNumber() + ": Quelle: " + imageSource;
            }
            // Caption text for images with metadata
            else if (clips[clipIndex]['content'] == 'image with metadata') {
                captionText = "Figure " + getImageCaptionNumber() + ': ' + clips[clipIndex]['citations'] + ' ' + 'Quelle: ' + imageSource;
            }
            // Insert caption under the image
            var caption = image.insertText(captionText, 'after');
            // Use built in "caption" style for caption
            caption.styleBuiltIn = Word.Style.caption;
            // Increase the current image number
            currentImageNumber++;
            return context.sync();
        })
            .catch(errorHandler);
    }

    
    /*
    Get superscript numbers up to 10 for footnotes
    */
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

    /*
    Return current image number for captions
    */
    function getImageCaptionNumber() {
        return currentImageNumber;
    }
    /*
    Return current citation number for placeholder
    */
    function getCitationNumber() {
        return ' [' + currentCitationNumber + ']';
    }

    /*
    Inserts image creditline without image
    */
    function insertCreditline() {
        Word.run(function (context) {
            var range = context.document.getSelection();
            var imageSource = clips[clipIndex]['source'];

            if (imageSource.indexOf('data:image/') != -1) {
                imageSource = 'No URL found';
            }
              
            var creditline = range.insertText(
                clips[clipIndex]['citations'] + ' ' + 'Quelle: ' + imageSource,
                Word.InsertLocation.replace);

            creditline.styleBuiltIn = Word.Style.caption;

            return context.sync();
        })
            .catch(errorHandler);
    }

    /*
    Insert bibliography after inserting citation
    */
    function insertBibliographyAfterCitation(forCitation) {

        // If current clipboard content is image with metadata, insert creditline instead of bibliography
        if ($('#citation-bibliography-button').text() == 'Creditline erstellen') {
            insertCreditline();
        }
        else {

            Word.run(function (context) {
                // If bibliographgy already exists, get it by it's tag and insert citation source in it
                if (bibliographyAdded) {
                    var range = context.document.getSelection();
                    var bibControls = context.document.contentControls.getByTag("bibliography");
                    context.load(bibControls);

                    forCitation.insertText(getCitationNumber(), 'after');

                    return context.sync()
                        .then(function () {
                            // Add citation to bibliography and define it's style
                            var citation = bibControls.items[0].insertParagraph('[' + currentCitationNumber + ']' + ' ' + clips[clipIndex]['citations'], 'end');
                            citation.spaceAfter = 20;
                            //citation.insertText('\n', 'end');
                            citation.styleBuiltIn = Word.Style.bibliography;
                            currentCitationNumber++;                 
                            return context.sync();
                        });
                }
                else {
                    // If bibliography was not already added, insert it and put citation source in it
                    
                    // The range differs from the other insertBibliography function
                    var range = context.document.body.paragraphs.getLast();
                    var number = forCitation.insertText(getCitationNumber(), 'after');
                 

                    number.insertBreak(Word.BreakType.page, "after");
                    var heading = range.insertText('Literaturverzeichnis', Word.InsertLocation.end);
                    bibliographyAdded = true;

                    var bibliography = heading.insertContentControl('end');
                    //range.insertBreak(Word.BreakType.page, "after");

                    // Give tag to bibliography, so we can access it later

                    bibliography.tag = 'bibliography';
                    bibliography.styleBuiltIn = Word.Style.bibliography;
                    var citation = bibliography.insertParagraph('[' + currentCitationNumber + ']' + ' ' + clips[clipIndex]['citations'], 'end');
                    citation.spaceAfter = 20;
                    //citation.insertText('\n', 'end');
                    citation.styleBuiltIn = Word.Style.bibliography;
                    currentCitationNumber++;

                    // Change button label after bibliography is added
                    $('#citation-bibliography-button').text('Zum Literaturverzeichnis hinzufügen');

                    heading.styleBuiltIn = Word.Style.heading2;

                    return context.sync();
                }


            })
                .catch(errorHandler);
        }
    }

    /*
    Insert bibliography with current clipboard content's source
    */
    function insertBibliography() {

        // If current clipboard content is image with metadata, insert creditline instead of bibliography
        if ($('#citation-bibliography-button').text() == 'Creditline erstellen') {
            insertCreditline();
        }
        else {

            Word.run(function (context) {
                // If bibliographgy already exists, get it by it's tag and insert citation source in it
                if (bibliographyAdded) {
                    var range = context.document.getSelection();
                    var bibControls = context.document.contentControls.getByTag("bibliography");
                    context.load(bibControls);

                    range.insertText(getCitationNumber(), 'after');

                    return context.sync()
                        .then(function () {

                            // Add citation to bibliography and define it's style
                            var citation = bibControls.items[0].insertParagraph('[' + currentCitationNumber + ']' + ' ' + clips[clipIndex]['citations'], 'end');
                            citation.spaceAfter = 20;
                            citation.styleBuiltIn = Word.Style.bibliography;
                            currentCitationNumber++;
                            return context.sync();

                        });
                }
                else {
                    // If bibliography was not already added, insert it and put citation source in it
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

                    // Give tag to bibliography, so we can access it later
                    bibliography.tag = 'bibliography';
                    bibliography.styleBuiltIn = Word.Style.bibliography;
                    var citation = bibliography.insertParagraph('[' + currentCitationNumber + ']' + ' ' + clips[clipIndex]['citations'], 'end');
                    citation.spaceAfter = 20;
                    //citation.insertText('\n', 'end');
                    citation.styleBuiltIn = Word.Style.bibliography;
                    currentCitationNumber++;
                    // Change button label after bibliography is added
                    $('#citation-bibliography-button').text('Zum Literaturverzeichnis hinzufügen');

                    return context.sync();
                }


            })
                .catch(errorHandler);
        }
    }

    /*
    Inserts ciatation at selection and adds source to bibliography
    */
    function insertCitation() {
        Word.run(function (context) {
            // Get current selection
            var range = context.document.getSelection();
            // Get clipboard content
            var content = $.trim(clips[clipIndex]['content']);
            // Insert content as citation
            var citation = range.insertText(
                '"' + content + '"',
                Word.InsertLocation.replace);
            // Give "Quote" style to inserted citation
            citation.styleBuiltIn = Word.Style.quote;

            citation.tag = 'citation';

            // Insert bibliography with citation or add citation to bibliography, if it already exists
            insertBibliographyAfterCitation(citation);

            return context.sync();
        })
            .catch(errorHandler);
    }

    /*
    Inserts footnote, code was already in example code, then altered
    */
    function insertFoot() {
        Word.run(function (context) {

            // Create a proxy sectionsCollection object.
            var mySections = context.document.sections.getFirst();

            var range = context.document.getSelection();

            range.insertText(getSuperScript(), 'end');

            // Queue a commmand to load the sections.
            context.load(mySections, 'body/style');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                // Create a proxy object the primary footer of the first section. 
                // Note that the footer is a body object.

                var myFooter = mySections.getFooter("primary");

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

    /*
    Shows next clipboard content, after clicking on "Vorwärts" button
    */
    function showNextClip() {
        if (clips.length > 0) {

            if (clipIndex != clips.length - 1) {
                clipIndex = clipIndex + 1;
            }

            if (clipIndex > -1) {
                var source = clips[clipIndex]['source'];
                var content = clips[clipIndex]['content'];
                var citation = clips[clipIndex]['citations'];

                if (content == 'image' || content == 'image with metadata') {
             
                    $('#content-container').hide();
                    $('#image-container').show();
                    $("#preview-image").attr("src", source);
                    $('#txtCitation').text(citation);
                    $('#txtContent').text(content);
                    $('#txtSource').text(source);

                    if (source.indexOf('data:image/') == -1) {
                        setBase64Image(source);
                    }
                    else {
                        base64image = source.split(',')[1];
                    }
                }
                else {
                    $('#content-container').show();
                    $('#image-container').hide();
                    $('#txtCitation').text(citation);
                    $('#txtContent').text(content);
                    $('#txtSource').text(source);
                }
            }
        }
    }

    /*
    Shows previous clipboard content, after clicking on "Zurück" button
    */
    function showPreviousClip() {
        if (clips.length > 0) {

            if (clipIndex - 1 > -1) {
                clipIndex = clipIndex - 1;
            }

            if (clipIndex > -1) {
                var source = clips[clipIndex]['source'];
                var content = clips[clipIndex]['content'];
                var citation = clips[clipIndex]['citations'];

                if (content == 'image' || content == 'image with metadata') {
                
                    $('#content-container').hide();
                    $('#image-container').show();
                    $("#preview-image").attr("src", source);
                    $('#txtCitation').text(citation);
                    $('#txtContent').text(content);
                    $('#txtSource').text(source);
                    if (source.indexOf('data:image/') == -1) {
                        setBase64Image(source);
                    }
                    else {
                        base64image = source.split(',')[1];
                    }                }
                else {
                    $('#content-container').show();
                    $('#image-container').hide();
                    $('#txtCitation').text(citation);
                    $('#txtContent').text(content);
                    $('#txtSource').text(source);
                }

            }

        }
    }
    var interval;
    /*
    Gets the clipboard's SOURCE and CITATION formats with the help of the python flask server
    */
    function getClipSource() {
        var sourceUrl = 'https://localhost:5000/source.py';
        var citationsUrl = 'https://localhost:5000/citations.py';

        // We don't want any old item in the cache
        $.ajaxSetup({
            cache: false
        });
        // AJAX query to server, first we want to get the sources
        $.ajax({
            type: 'GET',
            url: sourceUrl,
            success: function (response) {
                // Parse the JSON response
                var json = JSON.parse(response);
                var dictObject = {}

                var source = json.source;

                var content = json.content;

                // Dictionary for saving sources and content
                var dict = {}
                dict['source'] = source;
                dict['content'] = content;

                // Second AJAX query, this time we want to get the citations
                $.ajax({
                    type: 'GET',
                    url: citationsUrl,
                    success: function (response) {

                        console.log(content);

                        // Check if content is an image
                        if (content == 'image') {
                            var json = JSON.parse(response);

                            if (clips.length == 0) {
                                // Show/Hide UI elements if it's an image and set the text of these elements
                                $('#content-container').hide();
                                $('#image-container').show();
                                $("#preview-image").attr("src", source);
                                $('#txtCitation').text(json.APA);
                                $('#txtContent').text(content);
                                $('#txtSource').text(source);

                                if (source.indexOf('data:image/') == -1) {
                                    setBase64Image(source);
                                }
                                else {
                                    base64image = source.split(',')[1];
                                }

                                // Push the citations to the dictionary and iterate clip index
                                dict['citations'] = citation;
                                clips.push(dict);

                                clipIndex = clips.length - 1;
                            }
                            // Check if it was the same content as before. If so, do nothing.
                            if (clips.slice(-1)[0]['source'] == source) {
                                console.log("was same image");
                            }
                            else {
                                // Show/Hide UI elements if it's an image and set the text of these elements
                                $('#content-container').hide();
                                $('#image-container').show();
                                $("#preview-image").attr("src", source);
                                $('#txtCitation').text(json.APA);
                                $('#txtContent').text(content);
                                $('#txtSource').text(source);
                               
                                if (source.indexOf('data:image/') == -1) {
                                    setBase64Image(source);
                                }
                                else {
                                    base64image = source.split(',')[1];
                                }

                                // Push the citations to the dictionary and iterate clip index
                                dict['citations'] = citation;
                                clips.push(dict);
                                clipIndex = clips.length - 1;

                            }

                        }
                        // Check if content was image with metadata
                        else if (content == 'image with metadata') {

                            // Check if it was the same content as before. If so, do nothing.
                            if (clips.length > 0 && clips.slice(-1)[0]['source'] == source) {
                                console.log("was same image");
                            }
                            else {
                                // Get the metadata out of the JSON object
                                var json = JSON.parse(response);
                                var name = json.citations.name;
                                var description = json.citations.description;
                                var author = json.citations.author;
                                var copyright = json.citations.copyright;

                                // Build metadata string
                                var meta_data_string = name + ', ' + description + ', ' + author + ', ';

                                // Set metadata string in dictionary
                                dict['citations'] = meta_data_string;

                                // Set the text of the UI elements
                                $('#txtCitation').text(meta_data_string);
                                $('#txtContent').text(content);
                                $('#txtSource').text(source);

                                // Hide/show UI elements
                                $('#content-container').hide();
                                $('#image-container').show();
                                $("#preview-image").attr("src", source);

                                // Set the Base64 image, which is needed for inserting images into the Word document
                                setBase64Image(source);

                                $('#citation-bibliography-button').text('Creditline erstellen');

                                // Push dictionary to clip array
                                clips.push(dict);
                                clipIndex = clips.length - 1;
                            }
                            
                       
                        }
                        else {
                            // In this case, there are no images, just text

                            // Change the text of the bibliography button 
                            if (bibliographyAdded) {
                                $('#citation-bibliography-button').text('Zum Literaturverzeichnis hinzufügen');
                            }
                            else {
                                $('#citation-bibliography-button').text('Literaturverzeichnis erstellen');
                            }

                           
                            // Parse the JSON object
                            var json = JSON.parse(response);

                            // Get the APA citation
                            var citation = json.APA;

                            // Sometimes there are unexpected newlines in the citation, replace them
                            try {
                                citation = citation.replace('\n', '');
                            }
                            catch (err) {
                                console.log('could not replace newlines in citation', err.message);
                            }

                            // Add the citation to the dictionary
                            dict['citations'] = citation;


                            // Push the dictionary to the clips array
                            if (clips.length == 0) {
                                $('#txtCitation').text(citation);
                                $('#txtContent').text(content);
                                $('#txtSource').text(source);
                                clips.push(dict);
                                clipIndex = clips.length - 1;
                            }
                            // If content was same as before, do nothing
                            else if (clips.slice(-1)[0]['source'] == content || clips.slice(-1)[0]['citations'] == content || (clips.slice(-1)[0]['content'] == content) && clips.slice(-1)[0]['source'] == source) {
                                console.log(clips.slice(-1)[0]['content']);
                            }
                            else {
                                // Set the text of UI elements and push dictionary in clips array
                                $('#txtCitation').text(citation);
                                $('#txtContent').text(content);
                                $('#txtSource').text(source);
                                clips.push(dict);
                                clipIndex = clips.length - 1;
                            }

                            
                            // Make sure to show or hide image container
                            if (clips[clipIndex]['content'] == 'image with metadata' || clips[clipIndex]['content'] == 'image') {
                                console.log("is image!");
                                $('#content-container').hide();
                                $('#image-container').show();
                            }
                            else  {
                                $('#content-container').show();
                                $('#image-container').hide();
                            }
                          
                           

                        }
                        // Set timeout of interval length and then call this whole function again to check for new clipboard data
                        citation_interval = setTimeout(getClipSource, intervalTime);
                    },
                    error: function (response) {
                        // If something goes wrong, call the function again and check for new clipboard data anyway
                        citation_interval = setTimeout(getClipSource, intervalTime);
                        console.log(console.error(response));
                    }
                });

            },

            error: function (response) {
                console.log(console.error(response));
                // If something goes wrong, call the function again and check for new clipboard data anyway
                citation_interval = setTimeout(getClipSource, intervalTime);

                return response;
            }
        });
       
    }

   
    /*
    Loads sample data. This was used for the user study and is not used anymore. However, it shows some useful functions for inserting headings into the document
    */
    function loadSampleData() {
        console.log("Loading sample");
        Word.run(function (context) {
            var body = context.document.body;

            body.clear();

            var tiger = body.insertParagraph(
                "Tiger",
                Word.InsertLocation.end);
            tiger.styleBuiltIn = Word.Style.heading2;

            var elefant = body.insertParagraph(
                "Elefant",
                Word.InsertLocation.end);
            elefant.styleBuiltIn = Word.Style.heading2;

            var affe = body.insertParagraph(
                "Affe",
                Word.InsertLocation.end);
            affe.styleBuiltIn = Word.Style.heading2;

            var hund = body.insertParagraph(
                "Hund",
                Word.InsertLocation.end);
            hund.styleBuiltIn = Word.Style.heading2;

            var pdf = body.insertParagraph(
                "PDFs",
                Word.InsertLocation.end);
            pdf.styleBuiltIn = Word.Style.heading2;
        
            return context.sync();
        })
            .catch(errorHandler);
    }

    // Was already included in the sample code, used for debugging
    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Fehler:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Was already included in the example code, used for debugging
    // Eine Hilfsfunktion zum Anzeigen von Benachrichtigungen.
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

})();
