/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

/**
 * Append a hard-coded webex meeting information into the user's email body
 * @param event {Office.AddinCommands.Event}
 */
function changeMailBody(event) {
    Office.context.mailbox.item.body.getAsync("html", function (asyncResult) {
        var result = asyncResult.value;
        var index = result.indexOf('</body>');
        var hardcodeEmailBody = '<br><div><a href=\"https://www.bing.com\" style=\"color:#005E7D; text-decoration:none;\">https://www.bing.com</a></div>';
        var alternateResult = index != -1 ? result.substring(0, index) + hardcodeEmailBody + result.substring(index) : result + hardcodeEmailBody;
        Office.context.mailbox.item.body.setAsync(
            alternateResult,
            {
                coercionType: Office.CoercionType.Html,
                asyncContext: {},
            },
            function (asyncResult) { event.completed(); });
    });

}

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.changeMailBody = changeMailBody;
