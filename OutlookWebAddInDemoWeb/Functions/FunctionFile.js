Office.initialize = function () {
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
        type: "informationalMessage",
        icon: icon,
        message: text,
        persistent: false
    });
}

function defaultStatus(event) {
    statusUpdate("icon16", "Hello World!");
}


function action(event) {
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(function (asyncResult) {
        var customProps = asyncResult.value;
        var myProps = customProps.get("myProps");
      
        customProps.set("myProps", "Property saved in Command button");
        customProps.saveAsync(function (result) {
           
           // item.saveAsync(function (result) {

            event.completed();

            //});


        });
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
g.action = action;
