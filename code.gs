function checkRedirectsSimple() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(header => header.toLowerCase().replace(/_/g, '').replace(/\s/g, '')); // Normalize headers
    console.log(headers)

    // Dynamically find column indices
    const oldUrlIndex = headers.indexOf('oldurl');
    const expectedNewUrlIndex = headers.indexOf('newurl');

    if (oldUrlIndex === -1 || expectedNewUrlIndex === -1) {
        throw new Error('Required columns "Old URL" or "New URL" not found.');
    }

    // Add column headers if they don't exist
    const statusIndex = headers.indexOf('status');
    const redirectChainIndex = headers.indexOf('redirectchain');
    const errorsIndex = headers.indexOf('errors');
    const finalUrlIndex = headers.indexOf('finalurl');
    const finalStatusIndex = headers.indexOf('finalstatus');

    if (statusIndex === -1) sheet.getRange(1, headers.length + 1).setValue('Status');
    if (redirectChainIndex === -1) sheet.getRange(1, headers.length + 2).setValue('Redirect Chain');
    if (errorsIndex === -1) sheet.getRange(1, headers.length + 3).setValue('Errors');
    if (finalUrlIndex === -1) sheet.getRange(1, headers.length + 4).setValue('Final URL');
    if (finalStatusIndex === -1) sheet.getRange(1, headers.length + 5).setValue('Final Status');

    const data2 = sheet.getDataRange().getValues();

    // Re-read the headers in case new columns were added
    const updatedHeaders = data2[0].map(header => header.toLowerCase().replace(/_/g, '').replace(/\s/g, '')); // Normaliz headers

    console.log(updatedHeaders)

    console.log(updatedHeaders.indexOf('status'))
    const statusCol = updatedHeaders.indexOf('status') + 1;
    const redirectChainCol = updatedHeaders.indexOf('redirectchain') + 1;
    const errorsCol = updatedHeaders.indexOf('errors') + 1;
    const finalUrlCol = updatedHeaders.indexOf('finalurl') + 1;
    const finalStatusCol = updatedHeaders.indexOf('finalstatus') + 1;

    for (let i = 1; i < data.length; i++) {  // Start from row 2 to skip header row
        const oldUrl = data[i][oldUrlIndex]; // Dynamically get column B
        const expectedNewUrl = data[i][expectedNewUrlIndex]; // Dynamically get column D
        let redirectChain = [];
        let errors = [];
        let statusMessage = 'Does not redirect successfully';
        let statusCode = '';
        let finalUrl = '';
        let finalStatus = '';

        if (oldUrl && expectedNewUrl) {  // Only proceed if both URLs are present
            let redirectUrl = oldUrl;
            let finalUrlReached = false;

            while (redirectUrl) {
                try {
                    console.log(redirectUrl)
                    const response = UrlFetchApp.fetch(redirectUrl, {followRedirects: false});
                    statusCode = response.getResponseCode();  // Get the status code
                    console.log(statusCode)

                    if(response.getHeaders()['Location']===undefined){
                          console.log("break")
                          errors.push(`Error: Redirected but final URL did not match expected URL`);
                          break;
                        }

                    redirectChain.push(`${redirectUrl}`);
                    
                    // If it's a redirect status (301 or 302), get the 'Location' header
                    if (statusCode) {
                        let newLocation = response.getHeaders()['Location'];

                        // Handle relative URLs in the 'Location' header
                        if (newLocation && newLocation.startsWith('/')) {
                            const baseUrl = getBaseUrl(redirectUrl);  // Get the base URL of the old URL
                            newLocation = baseUrl + newLocation;  // Combine with base URL
                        }

                        if(newLocation===undefined){
                          console.log("break")
                          errors.push(`Error: Redirected but final URL did not match expected URL`);
                          break;
                        }

                        redirectUrl = newLocation;
                        
                        // Capture the final redirect URL
                        finalUrl = redirectUrl;
                        console.log(finalUrl)

                        console.log(redirectUrl)
                        console.log(expectedNewUrl)

                        finalStatus = statusCode;

                        // Check if the redirect destination matches the Expected New URL
                        if (redirectUrl === expectedNewUrl) {
                            console.log("yes")
                            //finalUrlReached = true;
                            statusMessage = "Redirects successfully";
                            finalStatus = statusCode;  // Capture the status code for the final URL
                            break;
                        }
                    } else {
                        console.log("bbb")
                        finalStatus = statusCode;
                        redirectUrl = null;  // Stop if not a redirect status
                    }

                } catch (error) {
                    const response = UrlFetchApp.fetch(redirectUrl, {muteHttpExceptions: true});
                    statusCode = response.getResponseCode();
                    if (statusCode === 404) {
                    console.log("404")
                    finalStatus = statusCode;
                    }
                    console.log("ccc")
                     // Handle errors by extracting the HTTP status code and error message
                    const errorMessage = error.message;
                    const statusMatch = errorMessage.match(/returned code (\d+)/);
                    if (statusMatch) {
                        const statusCode = statusMatch[1];  // Extract status code (e.g., 404)
                        errors.push(`Error: HTTP ${statusCode}`);  // Only include the HTTP status code and error message
                    } else {
                        errors.push(`Error: ${error.message}`);  // Include other errors if necessary
                    }
                    redirectUrl = null;  // Stop if there's an error
                }
            }
            console.log("eee")

            // Ensure the final URL is added to the redirect chain
            if (finalUrl) {
                console.log("aaa")
                redirectChain.push(`${finalUrl}`);
            }

            // Save results to the sheet
            sheet.getRange(i + 1, statusCol).setValue(statusMessage);
            sheet.getRange(i + 1, redirectChainCol).setValue(redirectChain.join(' -> '));
            sheet.getRange(i + 1, errorsCol).setValue(errors.join(', '));
            sheet.getRange(i + 1, finalUrlCol).setValue(finalUrl);
            sheet.getRange(i + 1, finalStatusCol).setValue(finalStatus);

        }
    }
}

// Helper function to get the base URL (e.g., 'http://example.com') from a full URL
function getBaseUrl(url) {
    const match = url.match(/^https?:\/\/([^\/]+)/);  // Matches http://example.com or https://example.com
    return match ? match[0] : '';
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Redirect Checker')
      .addItem('Check Redirects', 'checkRedirectsSimple')
      .addToUi();
}
