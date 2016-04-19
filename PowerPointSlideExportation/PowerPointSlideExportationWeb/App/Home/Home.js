﻿/// <reference path="../App.js" />-

(function () {
    "use strict";
    var BaseURL = 'https://testanddebug.azurewebsites.net/PowerPoint/';
    var SendFileUrl = 'https://testanddebug.azurewebsites.net/PowerPoint/publish';
    var SignInUrl = 'https://testanddebug.azurewebsites.net/PowerPoint/Test/';
    var GetTokenUrl = 'https://testanddebug.azurewebsites.net/PowerPoint/Token/';
    var sliceSize = (256 * 1024);
    var interval;
    var history = [];

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $.support.cors = true;

            $(".block-ui").hide();

            $(document)
                .ajaxStart(function() {
                    $(".block-ui").show();
                })
                .ajaxStop(function() {
                    $(".block-ui").hide('slow');
                });

            $.ajax({
                url: GetTokenUrl,
                method: 'GET'
            }).done(function (result) {
                if (result['code'] === 'Success') {
                    Office.context.document.settings.set('token', result['token']);
                } else if (result['code'] === 'NeedResync') {
                    window.location.reload;
                } else {
                    app.showNotification('Request error:', result['error']);
                }
            }).fail(function(result) {
                app.showNotification('Token Error', 'Token not recieved, reload');
            });

            $('#signup-btn').click(showSignUp);

            $('#signin-btn').click(function () {

                history.push(showWelcome);

                $('.welcome-container').hide();
                $('.navbar-header').hide();

                $('.waiting-container').show();
                $('.back-container').show();

                openSignIn();
                interval = setInterval(function () {
                    // wait for token
                    $.ajax({
                        url: GetTokenUrl,
                        method: 'GET'
                    }).done(function (result) {
                        if (result['code'] === 'Success') {
                            Office.context.document.settings.set('token', result['token']);

                            clearInterval(interval);
                            $('.waiting-container').hide();
                            showPublish();
                        } else if (result['code'] === 'NeedResync') {
                            window.location.reload;
                        } else {
                            app.showNotification('Request error:', result['error']);
                        }
                    });
                }, 120000);
            });

            $('#publish-btn').click(function () {

                history.push(showPublish);

                $('.publish-container').hide();
                $('.publishing-container').show();

                $('.success-title').hide();
                $('.failed-title').hide();

                sendFile();
            });

            $('#back').click(goBack);

            $('#btn-signup-submit').click(function (e) {
                e.preventDefault();
                postSignUp();
                return false;
            });

            showWelcome();
        });
    };

    function openSignIn()
    {
        window.open(SignInUrl + Office.context.document.settings.get('token'));
    }

    function b64EncodeUnicode(str) {
        return btoa(encodeURIComponent(str).replace(/%([0-9A-F]{2})/g, function (match, p1) {
            return String.fromCharCode('0x' + p1);
        }));
    }

    function closeFile(state) {

        // Close the file when you're done with it.
        state.file.closeAsync(function (result) {

            // If the result returns as a success, the
            // file has been successfully closed.
            if (result.status === "succeeded") {
                //updateStatus("File closed.");
            }
        });
    }

    function sendSlice(slice, state) {
        var data = slice.data;

        // If the slice contains data, create an HTTP request.
        if (data) {

            // Encode the slice data, a byte array, as a Base64 string.
            var fileData = b64EncodeUnicode(data);

            $.ajax({
                method: 'POST',
                url: SendFileUrl,
                beforeSend: function (xhr) {
                    xhr.setRequestHeader("Slice-Number", slice.index);
                    xhr.setRequestHeader("Slice-Total", state.sliceCount);
                    xhr.setRequestHeader("File-Name", Office.context.document.settings.get('token'));
                },
                data: fileData
            }).done(function (response) {
                if (result['code'] === 'Success') {
                    state.counter++;

                    $('.publishing-percent').html(Math.floor((state.counter * 100) / state.sliceCount) + '%');
                    $('#publishing-progress').val(state.counter);

                    if (state.counter < state.sliceCount) {
                        getSlice(state); // Continue upload
                    } else {
                        // end OK
                        $('#publishing-progress').val(state.counter);
                        closeFile(state);
                        $('.success-title').show();
                    }
                } else if (result['code'] === 'NeedResync') {
                    window.location.reload;
                } else {
                    $('.failed-title').show();
                    $('#publishing-progress').addClass('upload-failed');
                    closeFile(state);

                    app.showNotification('Request error:', result['error']);
                }

            }).fail(function(response) {
                $('.failed-title').show();
                $('#publishing-progress').addClass('upload-failed');
                closeFile(state);
            });
        }
    }

    function getSlice(state)
    {

        state.file.getSliceAsync(state.counter, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                sendSlice(result.value, state);
            }
            else {
                $('.success-title').hide();
                $('.failed-title').show();
                $('#publishing-progress').addClass('upload-failed');
            }
        });
    }

    function sendFile() {

        Office.context.document.getFileAsync("compressed", { sliceSize: sliceSize },
            function (result) {

                if (result.status === Office.AsyncResultStatus.Succeeded) {

                    // Get the File object from the result.
                    var myFile = result.value;

                    var state = {
                        file: myFile,
                        counter: 0,
                        sliceCount: myFile.sliceCount
                    };

                    $('#publishing-progress')
                        .removeClass('upload-failed')
                        .val(0)
                        .attr('max', myFile.sliceCount);

                    getSlice(state);
                }
                else {
                    $('.success-title').hide();
                    $('.failed-title').show();
                    $('#publishing-progress').addClass('upload-failed');
                }
            });
    }

    // Get a slice from the file and then call sendSlice.
    function goBack() {
        var fun = history.pop();
        fun();
        clearInterval(interval);
    }

    function showPublish() {
        history.push(showWelcome);

        $('.publish-container').show();
        $('.publishing-container').hide();
    }

    function postSignUp() {
        var destinationUrl = BaseURL + 'SignUp';
        var $form = $('form');
        $form.find('#token').val(Office.context.document.settings.get('token'));

        $.ajax({
                url: destinationUrl,
                method: 'POST',
                data: $form.serializeArray()
            })
            .done(function (result) {
                if (result['code'] === 'Success') {
                    // save token and go to publish
                    Office.context.document.settings.set('token', result['token']);
                    $('.signup-container').hide();

                    showPublish();
                } else if (result['code'] === 'NeedResync') {
                    window.location.reload;
                } else {
                    app.showNotification('Request error:', result['error']);
                }
            })
            .fail(function (jqXHR, textStatus, errorThrown) {
                //$('.failed-container').show();
                app.showNotification('Request error:', 'An error has ocurred, please try again');
            })
            .always(function(result) {
                
            });
    }

    function showSignUp() {
        history.push(showWelcome);

        $('.welcome-container').hide();
        $('.navbar-header').hide();
        $('.back-container').show();
        $('.signup-container').show();
    }

    function showWelcome() {
        // Get welcome screen token
        // Office.context.document.settings.set('token', $body.find('#welcome-token').val());
        $('.welcome-container').show();
        $('.navbar-header').show('slow');

        $('.back-container').hide();
        $('.waiting-container').hide();
        $('.publish-container').hide();
        $('.publishing-container').hide();
        $('.signup-container').hide();
        $('.failed-container').hide();
    }
})();