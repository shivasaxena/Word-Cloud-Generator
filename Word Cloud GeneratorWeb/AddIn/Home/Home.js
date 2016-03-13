/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            draw();

            // code to generate a new word cloud 
            $("#cloud-from-selection").on('click', function () {
                Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                   function (result) {
                       if (result.status === Office.AsyncResultStatus.Succeeded) {
                           if (result.value == "") {
                               app.showNotification('No data was selected. Please select some data in document');
                           } else {
                               draw(result.value);
                           }

                       } else {
                           app.showNotification('Error:', result.error.message);
                       }
                   }
                );
            });

            $.support.cors = true;
            $('#insert-cloud-to-document').on('click', function () {
                var $canvas = $('#word-cloud-placeholder');
                var imageData = $canvas[0].toDataURL();
                if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {
                    console.log("ImageCoercion supported");
                    imageBase64Data = imageBase64Data.replace(/^data:image\/(png|jpg);base64,/, "");// Remove the extra metadata added by toDataURL
                    setBase64Image(imageData);
                } else {
                    console.log("ImageCoercion not supported");
                    $.ajax({
                        type: 'post',
                        url: 'https://metalop.com/Word-Cloud-Generator/image-url-generator.php',
                        data: {
                            image: imageData
                        },
                        error: function (e) {
                            console.error(e);
                        },
                        success: function (response) {
                            console.log(response);
                            var imageHTML = "<img " +
                                "src='" + response + "' img/>";
                            console.log(imageHTML);
                            setHTMLImage(imageHTML)
                        }
                    });
                }
               
               

            })

        });
    };

    function setHTMLImage(imageHTML) {
        Office.context.document.setSelectedDataAsync(
            imageHTML, { coercionType: "html" },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    app.showNotification('Error: ' + asyncResult.error.message);
                    console.log(asyncResult);
                }
            });
    }

    function setBase64Image(imageBase64Data) {
        Office.context.document.setSelectedDataAsync(imageBase64Data, {
            coercionType: Office.CoercionType.Image,
            imageLeft: 50,
            imageTop: 50,
            imageWidth: 250,
            imageHeight: 250
        },
       function (asyncResult) {
           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
               console.log("Action failed with error: " + asyncResult.error.message);
           }
       });
    }

    function draw(words) {
        var list = [['This', 20],
            ['is ', 18],
            ['the', 16],
            ['defaut', 15],
            ['Word', 22],
            ['Cloud',20 ]
        ];

        if (words) {
            var newList = [];
            var regex = /[^\s\.,!?]+/g;
            var individualWords = words.match(regex); // Split individual words from sentences
            if (individualWords) {
                individualWords.forEach(function (word) {
                    var wordEntry = [];
                    wordEntry.push(word);
                    wordEntry.push(Math.floor(Math.random() * 15) + 1);
                    newList.push(wordEntry);
                });
                list = newList;
            } else {
                app.showNotification('No data was selected. Please select some data in document');
            }
        }

        WordCloud(document.getElementById('word-cloud-placeholder'), {
            list: list,
            fontFamily: 'Finger Paint, cursive, sans-serif',
            gridSize: 16,
            weightFactor: 2,
           
            color: '#f0f0c0',
            backgroundColor: '#001f00',
            shuffle: false,
            rotateRatio: 0
        });

    }
    
})();