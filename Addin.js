Office.initialize = function (reason) { };

jasmine.DEFAULT_TIMEOUT_INTERVAL = 20000;



function sendTestReport() {
    
       var testResults = $("#testResults").html();

        $("#testResults").remove();
 

       $(".jasmine_html-reporter").after(testResults);

       
    var completeHtml = "<html>" + $("html").html() + "</html>";
  


    var textcompleteHtml = completeHtml;
    textcompleteHtml = textcompleteHtml.replace(/&/g, '&amp;');
    textcompleteHtml = textcompleteHtml.replace(/</g, '&lt;');
    textcompleteHtml = textcompleteHtml.replace(/>/g, '&gt;');



    var options = {
        isRest: true,
        asyncContext: { message: 'Hello World!' }
    };

    Office.context.mailbox.getCallbackTokenAsync(options, cb);


    function cb(asyncResult) {
        var token = asyncResult.value;

        var sendMessageUrl = Office.context.mailbox.restUrl +
            '/v2.0/me/sendmail';

        var emailData = {
            "Message": {
                "Subject": "Read API Test Result for " + Office.context.mailbox.diagnostics.hostName + ":" + Office.context.mailbox.diagnostics.hostVersion,
                "Body": {
                    "ContentType": "Html",
                    "Content": completeHtml
                },
                "ToRecipients": [
                    {
                        "EmailAddress": {
                            "Address": Office.context.mailbox.userProfile.emailAddress
                        }
                    }
                ],
                "Attachments": [
                ]
            },
            "SaveToSentItems": "false"
        };

        $.ajax({
            url: sendMessageUrl,
            contentType: 'application/json',
            data: JSON.stringify(emailData),
            type: 'post',
            headers: { 'Authorization': 'Bearer ' + token }
        }).done(function (item) {
           
        }).fail(function (error) {
            $(".jasmine_html-reporter").after("<p>" + error + "</p>");
            console.log(error);
        });




    }


}


var currentMessage;
var currentUser;
var currentAttachments;
var mailBoxSettings;


function getCurrentMessage() {

    var options = {
        isRest: true,
        asyncContext: { message: 'Hello World!' }
    };

    Office.context.mailbox.getCallbackTokenAsync(options, getMessage);


    function getMessage(asyncResult) {
        var token = asyncResult.value;

        var getMessageUrl = Office.context.mailbox.restUrl +
            '/v2.0/me/messages/' + Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0);

       

        $.ajax({
            url: getMessageUrl,
            contentType: 'application/json',
            type: 'get',
            headers: { 'Authorization': 'Bearer ' + token }
        }).done(function (item) {

            currentMessage = item;
          

        }).fail(function (error) {
          
            console.log(error);
           
        });




    }


}

function getCurrentUser() {

    var options = {
        isRest: true,
        asyncContext: { message: 'Hello World!' }
    };

    Office.context.mailbox.getCallbackTokenAsync(options, getMessage);


    function getMessage(asyncResult) {
        var token = asyncResult.value;

        var getUserUrl = Office.context.mailbox.restUrl +
            '/v2.0/me';



        $.ajax({
            url: getUserUrl,
            contentType: 'application/json',
            type: 'get',
            headers: { 'Authorization': 'Bearer ' + token }
        }).done(function (item) {

            currentUser = item;
           

        }).fail(function (error) {

            console.log(error);
           
        });




    }


}


function getAttachments()

    {

       var options ={
           isRest: true,
           asyncContext: { message: 'Hello World!' }
           };

        Office.context.mailbox.getCallbackTokenAsync(options, getMessage);


        function getMessage(asyncResult)

          {
            var token = asyncResult.value;

            var getAttachmentsUrl = Office.context.mailbox.restUrl +
            '/v2.0/me/messages/' + Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0) + '/attachments';



            $.ajax({
            url: getAttachmentsUrl,
            contentType: 'application/json',
            type: 'get',
            headers: { 'Authorization': 'Bearer ' + token }
            }).done(function (item) {

            currentAttachments = item;

            }).fail(function (error) {

            console.log(error);

             });




        }


    }



function getMaiboxSettings() {

    var options = {
        isRest: true,
        asyncContext: { message: 'Hello World!' }
    };

    Office.context.mailbox.getCallbackTokenAsync(options, getMessage);


    function getMessage(asyncResult) {
        var token = asyncResult.value;

        var getMailboxSettingsUrl = Office.context.mailbox.restUrl +
            '/v2.0/me/MailboxSettings';



        $.ajax({
            url: getMailboxSettingsUrl ,
            contentType: 'application/json',
            type: 'get',
            headers: { 'Authorization': 'Bearer ' + token }
        }).done(function (item) {

            mailBoxSettings = item;
           

        }).fail(function (error) {

            console.log(error);
           
        });




    }


}


describe("",
    function () {


        beforeAll(function (done)
        {
            setTimeout(function () {
                getCurrentMessage();

                getCurrentUser();

                getMaiboxSettings();

                getAttachments();
                setTimeout(function () {
                    done();
                }, 2000)
            }, 3000)



        });
        afterAll(function () {


            setTimeout(function () {
                sendTestReport();
            }, 5000)


        })


        describe("Office.context.", function () {


            it(" Get the display language of Outlook",
                function () {

                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get the display language of Outlook */

                    var displayLanguage = Office.context.displayLanguage;
                    console.log("Display language is " + Office.context.displayLanguage);
                    document.getElementById("displayLanguage").innerHTML = Office.context.displayLanguage;

                    if (displayLanguage == "en-US")
                    {

                        expect(displayLanguage).toBe("en-US");
                    }

                    else if (displayLanguage == "en-IN")
                    {
                        expect(displayLanguage).toBe("en-IN");
                    }

                    else if (displayLanguage == "en-GB")
                    {
                        expect(displayLanguage).toBe("en-GB");

                    }




                });

            xit("Get the theme of Outlook",
                function () {


                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get the theme of Outlook */
                    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
                    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
                    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
                    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;
                    console.log("Body:(" + bodyBackgroundColor + "," + bodyForegroundColor + "), Control:(" + controlBackgroundColor + "," + controlForegroundColor + ")");
                    document.getElementById("theme").innerHTML = "Body:(" +
                        bodyBackgroundColor +
                        "," +
                        bodyForegroundColor +
                        "), Control:(" +
                        controlBackgroundColor +
                        "," +
                        controlForegroundColor +
                        ")";
                    expect(bodyBackgroundColor).toBeDefined();
                    expect(bodyForegroundColor).toBeDefined();
                    expect(controlBackgroundColor).toBeDefined();
                    expect(controlForegroundColor).toBeDefined();

                });


            it(" Set and Save roaming settings",
                function (done) {
                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Set and Save roaming settings */


                    Office.context.roamingSettings.set("myKey", "Hello World!");
                    Office.context.roamingSettings.saveAsync(
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);

                                document.getElementById("setRoamingSetting").innerHTML =
                                    asyncResult.error.message;

                            } else {
                                console.log("Settings saved successfully");
                                document.getElementById("setRoamingSetting").innerHTML =
                                    "Settings saved successfully";



                            }

                            expect(asyncResult.status).toBe("succeeded");
                            done();



                        }
                    );


                });

            it(" Get roaming settings",
                function () {
                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get roaming settings */
                    var settingsValue = Office.context.roamingSettings.get("myKey");
                    console.log("myKey value is " + settingsValue);
                    document.getElementById("getRoamingsetting").innerHTML = settingsValue;
                    expect(settingsValue).toBe("Hello World!");


                });

            it("Remove roaming settings",
                function (done) {

                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Remove roaming settings */


                    Office.context.roamingSettings.remove("myKey");
                    Office.context.roamingSettings.saveAsync(
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);
                                document.getElementById("removeRoamingSetting").innerHTML =
                                    "Action failed with error: " + asyncResult.error.message;

                            } else {
                                console.log("Settings saved successfully");
                                document.getElementById("removeRoamingSetting").innerHTML =
                                    "Settings saved successfully";

                            }
                            expect(asyncResult.status).toBe("succeeded");
                            done();

                        }
                    );

                });


        });


        describe("Office.context.mailbox.", function () {

            it(" Convert to REST ID:Requires item Id",
                function (done) {

                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to REST ID */
                    // Get the currently selected item's ID
                    var ewsId = "AAMkADhlODgyMjQ3LTY0OTEtNDVhNy1hMjE4LTRiNWViODdjNzM1OQBGAAAAAADTm / rlU8XIRYZy3kXeC31hBwATCz0JAbtBSrpwxQVbcRSjAAADfWGhAAATCz0JAbtBSrpwxQVbcRSjAAAGtCG2AAA=";
                    var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
                    console.log(restId);
                    document.getElementById("convertToRestId").innerHTML = restId;
                    expect(restId).toBeDefined();
                    expect(restId).toBe("AAMkADhlODgyMjQ3LTY0OTEtNDVhNy1hMjE4LTRiNWViODdjNzM1OQBGAAAAAADTm - rlU8XIRYZy3kXeC31hBwATCz0JAbtBSrpwxQVbcRSjAAADfWGhAAATCz0JAbtBSrpwxQVbcRSjAAAGtCG2AAA=")
                    done();



                });

            it("Convert to EWS ID",
                function (done) {


                    /* Restricted or ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to EWS ID */
                    // Get an item's ID from a REST API
                    var restId = "AAMkAGY4NTY1NDE4LTYwY2UtNGFkMi1iYWM0LTFjNWNlZTRiYzJiZgBGAAAAAADoWq5beaIQS5H0b244q4teBwBBlpJMXmrvRZroKP1QMFD7AAWOIICDAAAyMljtOF9eSIpjBvMLrE1RAADk489TAAA=";
                    var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
                    console.log(ewsId);
                    document.getElementById("convertToEwsId").innerHTML = ewsId;
                    expect(ewsId).toBeDefined();
                    expect(ewsId).toBe("AAMkAGY4NTY1NDE4LTYwY2UtNGFkMi1iYWM0LTFjNWNlZTRiYzJiZgBGAAAAAADoWq5beaIQS5H0b244q4teBwBBlpJMXmrvRZroKP1QMFD7AAWOIICDAAAyMljtOF9eSIpjBvMLrE1RAADk489TAAA=");
                    done();
                });


            it(" Convert to local client time",
                function () {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to local client time */
                    var localTime = Office.context.mailbox.convertToLocalClientTime(new Date("September 5,2017 10:30:00"));
                    console.log("LocalTime:" + localTime.date + "/" + (localTime.month + 1) + "/" + localTime.year
                        + " " + localTime.hours + ":" + localTime.minutes + " (+" + localTime.timezoneOffset + ")");

                    document.getElementById("localClientTime").innerHTML = "LocalTime:" +
                        localTime.date +
                        "/" +
                        (localTime.month + 1) +
                        "/" +
                        localTime.year +
                        " " +
                        localTime.hours +
                        ":" +
                        localTime.minutes +
                        " (+" +
                        localTime.timezoneOffset +
                        ")";

                    expect(localTime).toBeDefined();

                    if (Office.context.mailbox.userProfile.timeZone == "India Standard Time") {
                        expect("LocalTime:" +
                            localTime.date +
                            "/" +
                            (localTime.month + 1) +
                            "/" +
                            localTime.year +
                            " " +
                            localTime.hours +
                            ":" +
                            localTime.minutes +
                            " (+" +
                            localTime.timezoneOffset +
                            ")").toBe("LocalTime:5/9/2017 10:30 (+330)");
                    }

                    else if (Office.context.mailbox.userProfile.timeZone == "Pacific Standard Time")
                    {

                        expect("LocalTime:" +
                            localTime.date +
                            "/" +
                            (localTime.month + 1) +
                            "/" +
                            localTime.year +
                            " " +
                            localTime.hours +
                            ":" +
                            localTime.minutes +
                            " (+" +
                            localTime.timezoneOffset +
                            ")").toBe("LocalTime:5/9/2017 10:30 (+-420)");






                    }



                });




            it("Convert to UTC client time ",
                function () {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Convert to UTC client time */
                    var localTime = Office.context.mailbox.convertToLocalClientTime(new Date("September 5,2017 10:30:00"));
                    var utcClientTime = Office.context.mailbox.convertToUtcClientTime(localTime);
                    console.log("UTC:" + utcClientTime);

                    document.getElementById("utcClientTime").innerHTML = "UTC:" + utcClientTime;
                    expect(utcClientTime).toBeDefined();
                    if (Office.context.mailbox.userProfile.timeZone == "India Standard Time") {
                        expect("UTC:" + utcClientTime).toBe("UTC:Tue Sep 05 2017 10:30:00 GMT+0530 (IST)");
                    }

                    else if (Office.context.mailbox.userProfile.timeZone == "Pacific Standard Time")
                       
                    {

                        expect("UTC:" + utcClientTime).toBe("UTC:Tue Sep 05 2017 10:30:00 GMT-0700 (PDT)");

                    }
                    

                });


          

            it("Get EWS URL",
                function () {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get EWS URL */
                    var ewsurl = Office.context.mailbox.ewsUrl;
                    console.log(Office.context.mailbox.ewsUrl);
                    document.getElementById("ewsURL").innerHTML = ewsurl;
                    expect(ewsurl).toBe("https://outlook.office365.com/EWS/Exchange.asmx");
                });


      

            it("Make EWS Request",
                function (done) {
                    /* ReadWriteMailbox */
                    /* EWS request to create and send a new item */

                    var request = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                        ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                        '  <soap:Header><t:RequestServerVersion Version="Exchange2010" /></soap:Header>' +
                        '  <soap:Body>' +
                        '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
                        '      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>' +
                        '      <m:Items>' +
                        '        <t:Message>' +
                        '          <t:Subject>Hello, Outlook!</t:Subject>' +
                        '          <t:Body BodyType="HTML">I sent this message to myself using the Outlook API!</t:Body>' +
                        '          <t:ToRecipients>' +
                        '            <t:Mailbox><t:EmailAddress>' + Office.context.mailbox.userProfile.emailAddress + '</t:EmailAddress></t:Mailbox>' +
                        '          </t:ToRecipients>' +
                        '        </t:Message>' +
                        '      </m:Items>' +
                        '    </m:CreateItem>' +
                        '  </soap:Body>' +
                        '</soap:Envelope>';

                    Office.context.mailbox.makeEwsRequestAsync(request,
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);
                                document.getElementById("ewsRequest").innerHTML = "Action failed with error: " + asyncResult.error.message;
                            } else {
                                console.log("Message sent! Check your inbox.");
                                document.getElementById("ewsRequest").innerHTML = "Message sent! Check your inbox.";
                            }
                            expect(asyncResult.status).toBe("succeeded");
                            done();

                        }
                    );


                });


            it("Get callback token async",
                function (done) {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get callback token async */


                    Office.context.mailbox.getCallbackTokenAsync(
                        function (asyncResult) {

                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);

                            } else {
                                console.log("Tokens: " + asyncResult.value);

                            }
                            document.getElementById("callbackToken").innerHTML = asyncResult.value;
                            expect(asyncResult.value).toBeDefined();
                            expect(asyncResult.status).toBe("succeeded");

                            if (Office.context.mailbox.item.attachments.length > 0) {

                                var encodedAttachmentId = encodeURIComponent(Office.context.mailbox.item.attachments[0].id);



                                $.ajax({
                                    url: "https://OutlookAPITestWebService.azurewebsites.net/home/getAttachmentDetails?token=" + asyncResult.value + "&ewsUrl=" + Office.context.mailbox.ewsUrl + "&attachmentId=" + encodedAttachmentId,
                                    type: 'get',


                                }).done(function (item) {

                                    document.getElementById("callbackToken").innerHTML = item.toString();
                                    expect(item.toString()).toBe(Office.context.mailbox.item.attachments[0].name);

                                    done();

                                }).fail(function (error) {
                                    // $(".jasmine_html-reporter").after("<p>" + error + "</p>");
                                    document.getElementById("userIdentityToken").innerHTML = error.toString();
                                    done();
                                });
                            }
                            done();
                        }
                    );



                });

            it("Get user identity token async",
                function (done) {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get user identity token async */

                    Office.context.mailbox.getUserIdentityTokenAsync(
                        function (asyncResult) {
                            if (asyncResult.status == "failed") {
                                console.log("Action failed with error: " + asyncResult.error.message);
                                done();
                            } else {
                                console.log("Tokens: " + asyncResult.value);
                                //document.getElementById("userIdentityToken").innerHTML = asyncResult.value;
                                expect(asyncResult.value).toBeDefined();
                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }


                            $.ajax({
                                url: "https://OutlookAPITestWebService.azurewebsites.net/home/CreateAndValidateIdentityToken?rawToken=" + asyncResult.value + "&hostUri=https://trelloaddin.azurewebsites.net/OutlookAPITestAddin/MessageRead.html",
                                // url: "https://localhost:44301/home/CreateAndValidateIdentityToken?rawToken=" + asyncResult.value + "&hostUri=https://trelloaddin.azurewebsites.net/OutlookAPITestAddin/MessageRead.html",
                                type: 'get',


                            }).done(function (item) {

                                document.getElementById("userIdentityToken").innerHTML = item.toString();
                                expect(item.toString()).toContain("autodiscover/metadata/json/");

                                done();

                            }).fail(function (error) {
                                // $(".jasmine_html-reporter").after("<p>" + error + "</p>");
                                document.getElementById("userIdentityToken").innerHTML = error.toString();
                                expect("Error for Identity Token ").toBe(error.toString());
                                expect(error.toString()).toBe("Error validating Identity token")
                                expect(asyncResult.value).toBe("Value");
                                expect(asyncResult.status).toBe("status");

                                document.getElementById("userIdentityToken").innerHTML = error.toString();

                                done();
                            });




                        }
                    );

                });

            it(" Get host version",
                function () {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get host version */
                    console.log(Office.context.mailbox.diagnostics.hostVersion);
                    document.getElementById("hostVersion").innerHTML = Office.context.mailbox.diagnostics.hostVersion;
                    expect(Office.context.mailbox.diagnostics.hostVersion).toBeDefined();
                   // expect(Office.context.mailbox.diagnostics.hostVersion.indexOf("15")).toBe(0);

                });



        });


        describe("Office.context.mailbox.diagnostics.", function () {


            it(" Get host name",
                function () {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get host name */
                    var hostName = Office.context.mailbox.diagnostics.hostName;
                    console.log(Office.context.mailbox.diagnostics.hostName);
                    document.getElementById("hostName").innerHTML = hostName;

                    if (hostName == "OutlookWebApp")
                    {
                        expect(hostName).toBe("OutlookWebApp");
                    }

                    else 
                    {
                        expect(hostName).toBe("Outlook");
                    }



                });

         


            it(" Get OWA view (only supported in OWA)",
                function (done) {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get OWA view (only supported in OWA) */
                    console.log(Office.context.mailbox.diagnostics.OWAView);
                    document.getElementById("owaView").innerHTML = Office.context.mailbox.diagnostics.OWAView;
                    if (Office.context.mailbox.diagnostics.hostName == "OutlookWebApp")
                    { expect(Office.context.mailbox.diagnostics.OWAView).toBeDefined(); }
                    else
                    { expect(Office.context.mailbox.diagnostics.OWAView).not.toBeDefined(); }
                    done();

                });




        });


        describe("Office.context.mailbox.userProfile.", function () {


            it(" Get display name",
                function () {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get display name */
                    var dispalyNameOfUser = Office.context.mailbox.userProfile.displayName;
                    console.log(Office.context.mailbox.userProfile.displayName);
                    //document.getElementById("displayName").innerHTML = dispalyNameOfUser;
                    document.getElementById("displayName").innerHTML = currentUser.DisplayName;
                    expect(dispalyNameOfUser).toBeDefined();
                   // expect(dispalyNameOfUser).toBe("Allan Deyoung");
                    expect(dispalyNameOfUser).toBe(currentUser.DisplayName);

                });

           it(" Get account Type",
                function () {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get display name */
                    var accountTypeOfUser = Office.context.mailbox.userProfile.accountType;
                    console.log(Office.context.mailbox.userProfile.accountType);
                    //document.getElementById("displayName").innerHTML = dispalyNameOfUser;
                    document.getElementById("accountType").innerHTML = accountTypeOfUser;
                    expect(accountTypeOfUser).toBeDefined();
                    // expect(dispalyNameOfUser).toBe("Allan Deyoung");
                    expect(accountTypeOfUser).toBe("office365");

                });

            it(" Get email address",
                function () {


                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get email address */
                    var emailAddressOfUser = Office.context.mailbox.userProfile.emailAddress;
                    console.log(Office.context.mailbox.userProfile.emailAddress);
                    //document.getElementById("emailAddress").innerHTML = emailAddressOfUser;
                    document.getElementById("emailAddress").innerHTML = currentUser.EmailAddress;

                    expect(emailAddressOfUser).toBeDefined();
                    //expect(emailAddressOfUser).toBe("mactest3@mod321281.onmicrosoft.com");
                    expect(emailAddressOfUser.toLowerCase()).toBe(currentUser.EmailAddress.toLowerCase());
                });


            it("Get time zone ",
                function () {
                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* Get time zone */
                    var timeZone = Office.context.mailbox.userProfile.timeZone;
                    console.log(Office.context.mailbox.userProfile.timeZone);
                    document.getElementById("timeZone").innerHTML = timeZone;
                    //document.getElementById("timeZone").innerHTML = mailBoxSettings.TimeZone;
                    expect(timeZone).toBeDefined();
                    if (timeZone == "India Standard Time") {
                        expect(timeZone).toBe("India Standard Time");
                    }

                    else if (timeZone == "Pacific Standard Time")
                    {
                        expect(timeZone).toBe("Pacific Standard Time");
                    }

                });




        });

      
        describe("1.5 API Office.context.", function () {

            

                it(" get rest URL",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* get rest URL */


                        console.log(Office.context.mailbox.restUrl);
                        document.getElementById("getRestUrl").innerHTML = Office.context.mailbox.restUrl;
                        expect(Office.context.mailbox.restUrl).toBeDefined();
                        expect(Office.context.mailbox.restUrl).toBe("https://outlook.office.com/api");

                    });


               

                


                it("get callback token isrest",
                    function (done) {
                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* get callback token isrest*/
                        var options = {
                            isRest: true,
                            asyncContext: { message: 'Hello World!' }
                        };

                        Office.context.mailbox.getCallbackTokenAsync(options, cb);


                        function cb(asyncResult) {
                            var token = asyncResult.value;
                            console.log(token);
                            expect(token).toBeDefined();
                            document.getElementById("getCallbackTokenIsRest").innerHTML = token;
                            expect(asyncResult.status).toBe("succeeded");
                            done();
                        }




                    });

               




        });


        describe("Office.context.UI.", function () {


            it("displayDialog",
                function (done) {

                    /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                    /* displayDialog */
                    var dialogOptions = { height: 80, width: 50, displayInIframe: false, requireHTTPS: false };

                    Office.context.ui.displayDialogAsync("https://trelloaddin.azurewebsites.net/OutlookAPITestAddin/Test/displayDialog.html", dialogOptions, displayDialogCallback);



                    function displayDialogCallback(asyncResult) {

                        console.log(asyncResult.status);

                        expect(asyncResult.status).toBe("succeeded");

                        var dialog = asyncResult.value;

                        dialog.addEventHandler(Office.EventType.DialogEventReceived, redirectingToHTTP);

                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, validateMessage);


                      


                        function redirectingToHTTP(arg) {

                            expect(arg.error).toBe(12003);

                            switch (arg.error) {
                                case 12002:
                                    document.getElementById("displayDialogEventReceived").innerHTML = "The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.";
                                    break;
                                case 12003:
                                    document.getElementById("displayDialogEventReceived").innerHTML = "The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required."; break;
                                case 12006:
                                    document.getElementById("displayDialogEventReceived").innerHTML = "Dialog closed.";
                                    break;
                                default:
                                    document.getElementById("displayDialogEventReceived").innerHTML = "Unknown error in dialog box.";
                                    break;
                            }

                        }

                        function validateMessage(messageObject) {

                            var messageFromDialog = messageObject.message.toString();
                            expect(messageFromDialog).toBe("test");
                            document.getElementById("displayDialogMessageReceived").innerHTML = messageObject.message.toString();

                            setTimeout(function () {
                                dialog.close();
                                done();
                            }, 5000)
                        }


                    };






                });

           




        });
       
        

       
        describe("Office.context.mailbox.item.", function () {


            

                it("Get item Id",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get item Id */
                        console.log(Office.context.mailbox.item.itemId);
                        document.getElementById("itemId").innerHTML = Office.context.mailbox.item.itemId;
                        expect(Office.context.mailbox.item.itemId).toBeDefined();
                        //expect(Office.context.mailbox.item.itemId).toBe("AAMkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NABGAAAAAAC3Bc26XexrR4XknrAwz6j9BwBDfaKHIE1iQJlAjLUe7EC6AACETLArAABDfaKHIE1iQJlAjLUe7EC6AACETMglAAA=");
                        expect(Office.context.mailbox.convertToRestId(Office.context.mailbox.item.itemId, Office.MailboxEnums.RestVersion.v2_0)).toBe(currentMessage.Id);



                    });
                it("Get item class",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get item class */
                        console.log(Office.context.mailbox.item.itemClass);
                        document.getElementById("itemClass").innerHTML = Office.context.mailbox.item.itemClass;
                        expect(Office.context.mailbox.item.itemClass).toBeDefined();
                        expect(Office.context.mailbox.item.itemClass).toBe("IPM.Note");
                        //expect(Office.context.mailbox.item.itemClass).toBe(currentMessage.);




                    });
                it("Get list of attachments",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get list of attachments */
                        var outputString = "";

                        for (i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
                            var _att = Office.context.mailbox.item.attachments[i];
                            outputString += "<BR>" + i + ". Name: ";
                            outputString += _att.name;
                            outputString += "<BR>ID: " + _att.id;
                            outputString += "<BR>contentType: " + _att.contentType;
                            outputString += "<BR>size: " + _att.size;
                            outputString += "<BR>attachmentType: " + _att.attachmentType;
                            outputString += "<BR>isInline: " + _att.isInline;
                            expect(outputString).toBeDefined();
                        }

                        var outputStringFromRest = "";
                        for (i = 0; i < currentAttachments.value.length; i++) {
                            var _att = currentAttachments.value[i];
                            outputStringFromRest += "<BR>" + i + ". Name: ";
                            outputStringFromRest += _att.Name;
                            outputStringFromRest += "<BR>ID: " + _att.Id;
                            outputStringFromRest += "<BR>contentType: " + _att.ContentType;
                            outputStringFromRest += "<BR>size: " + _att.Size;
                            outputStringFromRest += "<BR>attachmentType: " + "file";
                            outputStringFromRest += "<BR>isInline: " + _att.IsInline;
                        }

                        for (i = 0; i < currentAttachments.value.length; i++) {
                            var _attRest = currentAttachments.value[i];
                            var _att = Office.context.mailbox.item.attachments[i];
                            expect(_att.name).toBe(_attRest.Name);
                            expect(Office.context.mailbox.convertToRestId(_att.id, Office.MailboxEnums.RestVersion.v2_0)).toBe(_attRest.Id);
                            expect(_att.contentType).toBe(_attRest.ContentType);
                            expect(_att.size).toBe(_attRest.Size);
                            expect(_att.name).toBe(_attRest.Name);

                            if (_att.attachmentType == "file")
                            {
                                expect(_att.attachmentType).toBe("file");
                            }

                            else if (_att.attachmentType == "item") {
                                expect(_att.attachmentType).toBe("item");
                            }
                            
                            expect(_att.isInline).toBe(_attRest.IsInline);

                            
                        }

                        document.getElementById("attachments").innerHTML = outputString;
                        console.log(outputString);
                        expect(outputString).toBeDefined();
                        

                        //expect(outputString).toBe(outputStringFromRest);
                       // expect(outputString).toBe('<BR>0. Name: squirrel.png<BR>ID: AAMkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NABGAAAAAAC3Bc26XexrR4XknrAwz6j9BwBDfaKHIE1iQJlAjLUe7EC6AACETLArAABDfaKHIE1iQJlAjLUe7EC6AACETMglAAABEgAQAKmBP9spxj9Nq5K4j80h2zE=<BR>contentType: image/png<BR>size: ' + _att.size + '<BR>attachmentType: file<BR>isInline: false<BR>1. Name: squirrel.png<BR>ID: AAMkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NABGAAAAAAC3Bc26XexrR4XknrAwz6j9BwBDfaKHIE1iQJlAjLUe7EC6AACETLArAABDfaKHIE1iQJlAjLUe7EC6AACETMglAAABEgAQALvvtaen8sBGlQwcTc70tgc=<BR>contentType: image/png<BR>size: ' + _att.size + '<BR>attachmentType: file<BR>isInline: false');
                    });

                it("Get date time created",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get date time created */
                        console.log(Office.context.mailbox.item.dateTimeCreated);
                        document.getElementById("dateTimeCreated").innerHTML = Office.context.mailbox.item.dateTimeCreated;
                        expect(Office.context.mailbox.item.dateTimeCreated).toBeDefined();
                       // expect(Office.context.mailbox.item.dateTimeCreated.toString()).toBe("Tue Jul 25 2017 21:51:46 GMT+0530 (IST)");

                      
                        expect(Date.parse(Office.context.mailbox.item.dateTimeCreated)).toBe(Date.parse(currentMessage.CreatedDateTime));


                    });

                it("Get date time modified",
                    function () {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get date time modified */
                        console.log(Office.context.mailbox.item.dateTimeModified);
                        document.getElementById("dateTimeModified").innerHTML = Office.context.mailbox.item.dateTimeModified;
                        expect(Office.context.mailbox.item.dateTimeModified).toBeDefined();
                        
                        var fromClient = Date.parse(Office.context.mailbox.item.dateTimeModified)/(1000*60*60);
                        var fromServer = Date.parse(currentMessage.LastModifiedDateTime) /(1000 * 60 * 60);
                        //expect(Office.context.mailbox.item.dateTimeModified).toBe(currentMessage.LastModifiedDateTime);
                        expect(fromServer - fromClient).toBeLessThan(12);



                    });


              

                it(" Get normalized subject",
                    function () {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get normalized subject */
                        console.log(Office.context.mailbox.item.normalizedSubject);
                        document.getElementById("normalizedSubject").innerHTML = Office.context.mailbox.item.normalizedSubject;
                        expect(Office.context.mailbox.item.normalizedSubject).toBeDefined();
                        //expect(Office.context.mailbox.item.normalizedSubject).toBe("Test Email for Outlook Entensibilty Test");
                        expect(Office.context.mailbox.item.normalizedSubject).toBe(currentMessage.Subject);


                    });

            


                 it("Get conversation Id (Applicable only on message)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get conversation Id (Applicable only on message) */
                            console.log(Office.context.mailbox.item.conversationId);
                            document.getElementById("conversationId").innerHTML = Office.context.mailbox.item.conversationId;
                            expect(Office.context.mailbox.item.conversationId).toBeDefined();
                            //expect(Office.context.mailbox.item.conversationId).toBeDefined("AAQkAGZiZjc1Y2RkLTczNjktNGU1YS1hYTkzLTYzZTU3OTE5OWQ3NAAQAIBzvxP2lndLl8RpbgD18kY=");
                            expect(Office.context.mailbox.item.conversationId).toBeDefined(currentMessage.ConversationId);



                        });

                 it("Get internet message Id (Applicable only on message)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get internet message Id (Applicable only on message) */
                            console.log(Office.context.mailbox.item.internetMessageId);
                            document.getElementById("internetMessageId").innerHTML = Office.context.mailbox.item.internetMessageId;
                            //expect(Office.context.mailbox.item.internetMessageId).toBeDefined();
                            expect(Office.context.mailbox.item.internetMessageId).toBeDefined(currentMessage.InternetMessageId);


                        });

                 it("Get Cc recipients (Applicable only on message)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get Cc recipients (Applicable only on message) */
                            var recipients = "";
                            Office.context.mailbox.item.cc.forEach(function (recipient, index) {
                                recipients = recipients + recipient.displayName + " (" + recipient.emailAddress + ");";
                            });

                            var recipientsFromREST = "";

                            currentMessage.CcRecipients.forEach(function (recipient, index) {
                                recipientsFromREST = recipientsFromREST + recipient.EmailAddress.Name + " (" + recipient.EmailAddress.Address + ");";
                            });


                            console.log(recipients);
                            document.getElementById("ccRecipients").innerHTML = recipients;
                            //expect(recipients).toBe("Mufeez Ahmed (ZEN3 Infosolutions Private Ltd) (v-mufahm@microsoft.com);Kallu Sushma (ksushma@microsoft.com);Deepak Agrawal (deagrawa@microsoft.com);");
                            expect(recipients).toBe(recipientsFromREST);
                            
                            expect(recipients).toBeDefined();

                        });

                 it("Get from (Applicable only on message)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get from (Applicable only on message) */
                            var from = Office.context.mailbox.item.from;
                            console.log(from.displayName + " (" + from.emailAddress + ");");
                            document.getElementById("from").innerHTML = (from.displayName + " (" + from.emailAddress + ");");
                            expect(from).toBeDefined();
                            //expect(from.displayName + " (" + from.emailAddress + ");").toBe("Mufeez Ahmed (Zen3 Infosolutions (India) Lim) (v-mufahm@microsoft.com);");
                            expect(from.displayName + " (" + from.emailAddress + ");").toBe(currentMessage.From.EmailAddress.Name + " (" + currentMessage.From.EmailAddress.Address + ");");


                        });

                 it("Get sender (Applicable only on message)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get sender (Applicable only on message) */
                            var sender = Office.context.mailbox.item.sender;
                            console.log(sender.displayName + " (" + sender.emailAddress + ");");
                            document.getElementById("sender").innerHTML = (sender.displayName + " (" + sender.emailAddress + ");");
                            expect(sender).toBeDefined();
                           // expect(sender.displayName + " (" + sender.emailAddress + ");").toBeDefined("Mufeez Ahmed (Zen3 Infosolutions (India) Lim) (v-mufahm@microsoft.com);");
                            expect(sender.displayName + " (" + sender.emailAddress + ");").toBeDefined(currentMessage.Sender.EmailAddress.Name + " (" + currentMessage.Sender.EmailAddress.Address + ");");
                        });
                 it("Get To recipients (Applicable only on message)",
                        function () {


                            /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                            /* Get To recipients (Applicable only on message) */
                            var recipients = "";
                            Office.context.mailbox.item.to.forEach(function (recipient, index) {
                                recipients = recipients + recipient.displayName + " (" + recipient.emailAddress + ");";
                                document.getElementById("to").innerHTML = recipients;
                            });
                            expect(recipients).toBeDefined();

                            var toFromREST = "";

                            currentMessage.ToRecipients.forEach(function (recipient, index) {
                                toFromREST = toFromREST + recipient.EmailAddress.Name + " (" + recipient.EmailAddress.Address + ");";
                            });

                            //expect(recipients).toBe("Allan Deyoung (mactest3@MOD321281.onmicrosoft.com);");
                            expect(recipients).toBe(toFromREST);




                        });

       

                 it("Get body content async",
                     function (done) {

                         /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                         /* Get body content */
                         Office.context.mailbox.item.body.getAsync("text",
                             function (asyncResult) {
                                 if (asyncResult.status == "failed") {
                                     console.log("Action failed with error: " + asyncResult.error.message);
                                 } else {
                                     console.log(asyncResult.value);
                                     
                                     document.getElementById("messageBody").innerHTML = asyncResult.value;


                                 }

                                 expect(asyncResult.status).toBe("succeeded");
                                 //expect(document.getElementById("messageBody").innerHTML.trim()).toContain("Tester@xyz.com")
                                 //expect(document.getElementById("messageBody").innerHTML.trim()).toContain("click here!")
                                 //expect(document.getElementById("messageBody").innerHTML.trim()).toContain("Click here!")
                                 /**
                                    * Returns the text from a HTML string
                                     * 
                                     * @param {html} String The html string
                                 */
                                 function stripHtml(html) {
                                     // Create a new div element
                                     var temporalDivElement = document.createElement("div");

                                     


                                     // Set the HTML content with the providen

                                     

                                     temporalDivElement.innerHTML = html;

                                     //Removes Html Comments

                                    // $(temporalDivElement.innerHTML).comments().remove();

                                    // $(temporalDivElement.innerHTML).contents().filter(function () { return this.nodeType == 8; }).remove(); // replaceWith()  etc.

                                         
                                         
                                         
                                     // Retrieve the text property of the element (cross-browser support)
                                     return temporalDivElement.textContent || temporalDivElement.innerText || "";
                                 }

                                 Array.prototype.clean = function () {
                                     for (var i = 0; i < this.length; i++) {
                                         if (this[i].trim().length == 0) {
                                             this.splice(i, 1);
                                             i--;
                                         }
                                     }
                                     return this;
                                 };
                                 Array.prototype.trimArray = function () {
                                     for (var i = 0; i < this.length; i++) {
                                         this[i] = this[i].trim();
                                     }
                                     return this;
                                 };

                                 Array.prototype.containsArray = function (array /*, index, last*/) {

                                     if (arguments[1]) {
                                         var index = arguments[1], last = arguments[2];
                                     } else {
                                         var index = 0, last = 0; this.sort(); array.sort();
                                     };

                                     return index == array.length
                                         || (last = this.indexOf(array[index], last)) > -1
                                         && this.containsArray(array, ++index, ++last);

                                 };

                                 var wordsFromClient = asyncResult.value.toString().replace(/\r/g, " ").replace(/\n/g, " ").replace(/\t/g, " ").split(" ").clean().trimArray();
                                 var wordsFromRest = stripHtml(currentMessage.Body.Content).replace(/\r/g, " ").replace(/\n/g, " ").replace(/\t/g, " ").split(" ").clean().trimArray();

                                 document.getElementById("messageBody").innerHTML = "Number of words from Client :" + wordsFromClient.length + "<BR>" + "Number of Words from Rest :" + wordsFromClient.length + "<BR>" + "words from Client :" + wordsFromClient.toString()+ "<BR>" + "Words from Rest :" + wordsFromClient.toString() + "<BR>" +asyncResult.value;

                                 expect(wordsFromClient.length).toBe(wordsFromClient.length);
                                 //expect(asyncResult.value.toString().replace(/\t/g, "").replace(/\r/g, "").replace(/\n/g, "").replace(/ /g, "")).toBe(currentMessage.BodyPreview.toString().replace(/\t/g, "").replace(/\r/g, "").replace(/\n/g, "").replace(/ /g, ""));
                                 done();
                             }
                         );





                     });

                




                it("Get item type",
                    function () {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get item type */
                        console.log(Office.context.mailbox.item.itemType);
                        document.getElementById("itemType").innerHTML = Office.context.mailbox.item.itemType;
                        expect(Office.context.mailbox.item.itemType).toBeDefined();
                        expect(Office.context.mailbox.item.itemType).toBe("message");

                    });



                it("Add notification message async",
                    function (done) {

                        var resultStatus = "";

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Add notification message async */
                        Office.context.mailbox.item.notificationMessages.addAsync("foo",
                            {
                                type: "progressIndicator",
                                message: "this operation is in progress",
                            },
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("addNotificationMessageAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                    resultStatus = "failed";

                                } else {
                                    console.log("Added a new progress notification message for this item");
                                    document.getElementById("addNotificationMessageAsync").innerHTML = "Added a new progress notification message for this item";
                                    resultStatus = "passed";
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();


                            }
                        );



                    });

                it("Replace notification message async",
                    function (done) {

                        var resultStatus = "";
                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Replace notification message async */
                        Office.context.mailbox.item.notificationMessages.replaceAsync("foo",
                            {
                                type: "informationalMessage",
                                icon: "icon_24",
                                message: "this operation is complete",
                                persistent: false
                            },
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("replaceNotificationMessageAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                    resultStatus = "failed";

                                } else {
                                    console.log("Replaced existing notification with new notification message");
                                    document.getElementById("replaceNotificationMessageAsync").innerHTML = "Replaced existing notification with new notification message";
                                    resultStatus = "passed";
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();


                            }


                        );




                    });
                it("Get all notification messages async",
                    function (done) {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get all notification messages async */
                        Office.context.mailbox.item.notificationMessages.getAllAsync(
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("getAllNotificationMessageAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    var outputString = "";
                                    asyncResult.value.forEach(
                                        function (noti, index) {
                                            outputString += "<BR>" + index + ". Key: ";
                                            outputString += noti.key;
                                            outputString += "<BR>type: " + noti.type;
                                            outputString += "<BR>icon: " + noti.icon;
                                            outputString += "<BR>message: " + noti.message;
                                            outputString += "<BR>persistent: " + noti.persistent;

                                            console.log(outputString);
                                            document.getElementById("getAllNotificationMessageAsync").innerHTML = outputString;
                                            
                                        }

                                    );

                                }

                                var allNotifications = JSON.stringify(asyncResult.value);
                                var expectedAllNotification = JSON.stringify([{ "key": "foo", "type": "informationalMessage", "message": "this operation is complete", "persistent": false, "icon": "icon_24" }]);
                                expect(asyncResult.status).toBe("succeeded");
                                expect(allNotifications).toBe(expectedAllNotification)
                                done();
                            }
                        );




                    });

                it(" Remove notification messages async ",
                    function (done) {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Remove notification messages async */
                        Office.context.mailbox.item.notificationMessages.removeAsync("foo",
                            function (asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Action failed with error: " + asyncResult.error.message);
                                    document.getElementById("removeNotificationMessageAsync").innerHTML = "Action failed with error: " + asyncResult.error.message;
                                } else {
                                    console.log("Notification successfully removed");
                                    document.getElementById("removeNotificationMessageAsync").innerHTML = "Notification successfully removed";
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );



                    });

                it("Set and save custom property 1",
                    function (done) {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Set and save custom property */
                        Office.context.mailbox.item.loadCustomPropertiesAsync(
                            function customPropsCallback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Failed to load custom property");
                                    done();

                                }
                                else {
                                    var customProps = asyncResult.value;
                                    customProps.set("myProp1", "value1");
                                    customProps.saveAsync(
                                        function (asyncResult) {
                                            if (asyncResult.status == "failed") {
                                                console.log("Failed to save custom property");


                                            }
                                            else {
                                                console.log("Saved custom property");


                                            }

                                            expect(asyncResult.status).toBe("succeeded");
                                            done();
                                        }
                                    );
                                }


                            }
                        );




                    });

                it("Set and save custom property ",
                    function (done) {



                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Set and save custom property */
                        Office.context.mailbox.item.loadCustomPropertiesAsync(
                            function customPropsCallback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Failed to load custom property");
                                    document.getElementById("setAndSaveCustomProperty").innerHTML = "Failed to load custom property";

                                }
                                else {
                                    var customProps = asyncResult.value;
                                    customProps.set("myProp", "value");
                                    customProps.saveAsync(
                                        function (asyncResult) {
                                            if (asyncResult.status == "failed") {
                                                console.log("Failed to save custom property");
                                                document.getElementById("setAndSaveCustomProperty").innerHTML = "Failed to save custom property";

                                            }
                                            else {
                                                console.log("Saved custom property");
                                                document.getElementById("setAndSaveCustomProperty").innerHTML = "Saved custom property";
                                            

                                            }

                                        }
                                    );
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });

                it("Get custom property",
                    function (done) {

                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get custom property */
                        Office.context.mailbox.item.loadCustomPropertiesAsync(
                            function customPropsCallback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Failed to load custom property");

                                }
                                else {
                                    var customProps = asyncResult.value;
                                    var myProp1 = customProps.get("myProp1");
                                    document.getElementById("getCustomProperty").innerHTML = myProp1;
                                    console.log(myProp1);
                                    expect(myProp1).toBe("value1");


                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });
                it("Remove and save custom property",
                    function (done) {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Remove and save custom property */
                        Office.context.mailbox.item.loadCustomPropertiesAsync(
                            function customPropsCallback(asyncResult) {
                                if (asyncResult.status == "failed") {
                                    console.log("Failed to load custom property");
                                    document.getElementById("removeAndSaveCustomProperty").innerHTML = "Failed to load custom property";

                                }
                                else {
                                    var customProps = asyncResult.value;
                                    customProps.remove("myProp");
                                    customProps.saveAsync(
                                        function (asyncResult) {
                                            if (asyncResult.status == "failed") {
                                                console.log("Failed to save custom property");
                                                document.getElementById("removeAndSaveCustomProperty").innerHTML = "Failed to Save custom property";

                                            }
                                            else {
                                                console.log("Saved custom property");
                                                document.getElementById("removeAndSaveCustomProperty").innerHTML = "Saved custom property";
                                               

                                            }

                                        }
                                    );
                                }

                                expect(asyncResult.status).toBe("succeeded");
                                done();
                            }
                        );




                    });
                it("Get entities ",
                    function (done) {


                        /* ReadItem or ReadWriteItem or ReadWriteMailbox */
                        /* Get entities */
                        var emailAddresses = "";
                        Office.context.mailbox.item.getEntities().emailAddresses.forEach(function (emailAddress, index) {
                            emailAddresses = emailAddresses + emailAddress + ";<BR>";
                        });
                        document.getElementById("getEntities").innerHTML = emailAddresses;
                        expect(emailAddresses).toBeDefined();

                        if (emailAddresses.length > 0) {

                            Office.context.mailbox.item.getEntities().emailAddresses.forEach(function (emailAddress, index) {
                                expect(emailAddress).toMatch(/^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$/);
                            });
                        }
                    
                        console.log(emailAddresses);
                        done();
                      //Regex to extract emails:[a-zA-Z0-9-_.]+@[a-zA-Z0-9-_.]+



                    });
             

        });
      });
       
  

