function sendPwEmail() {
    // Open Google Sheets document
    var sheetId = "Replace_with_Sheet_id" // Make sure you replace this with ID
    var ss = SpreadsheetApp.openById(sheetId);
    var userSheet = ss.getSheetByName("Sheet1"); // Make sure sheet is named correctly
    var userDataRange = userSheet.getDataRange();
    var userData = userDataRange.getValues();

    // Go through all users in .csv file(First row is header, so start from the second row)
    for (var i = 1; i < userData.length; i++) {
        var firstName = userData[i][0]; // First Column (FirstName)
        var lastName = userData[i][1]; // Second Column (LastName)
        var schoolMail = userData[i][2]; // Third Column (SchoolEmail)
        var tempPassword = userData[i][3]; // forth Column (tempPassword)
        var personalMail = userData[i][9]; // Tenth Column (PersonalEmail)

        // Checks if the personalMail is empty
        if (!personalMail) {
            Logger.log(`Row ${i + 1}: address is missing.`);
            continue; // Continue with next account if personalMail is empty
        }

        // OneTimeSecret API
        var url = "https://onetimesecret.com/api/v1/share";
        var options = {
            "method": "post",
            "payload": {
                "secret": tempPassword,
                "api_key": "your_api_key" // Replace with your OneTimeSecret API key
            }
        };

        var response = UrlFetchApp.fetch(url, options);
        var json = JSON.parse(response.getContentText());
        var secretLink = "https://onetimesecret.com/secret/" + json.secret_key;

        // Email subject and message
        var subject = "Your new email account has been created"
        var message = `
            <html>
                <body>
                    <p>Hey <strong>${firstName} ${lastName}</strong>,</p>

                    <p>Your new email address has been created. Here is your new account and password.</p>
                    <h3><strong>NB! Immediately save the password in a safe place, because the password will be shown to you on the link only once and then the link will be destroyed!</strong></h3>
                    
                    <p><strong>Email:</strong> ${schoolMail}</p>
                    <p><strong>Password:</strong> <a href="${secretLink}" target="_blank">Click here to show password</a></p>

                    <p>If you got any problems or any questions email us email@email.com</p> <!--Replace with your email-->

                </body>
            </html>
        `;

        // Send email
        MailApp.sendEmail({
            to: personalMail,
            subject: subject,
            htmlBody: message
        });
    }
}