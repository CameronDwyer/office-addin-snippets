
function get() {
    var settingName = $("#settingName").val();
    var settingValue = Office.context.roamingSettings.get(settingName);
    $("#settingValue").val(settingValue);
    console.log(`The value of setting "${settingName}" is "${settingValue}".`);
}

function set() {
    var settingName = $("#settingName").val();
    var settingValue = $("#settingValue").val();
    Office.context.roamingSettings.set(settingName, settingValue);
    console.log(`Setting "${settingName}" set to value "${settingValue}".`);
}

function save() {
    // Save settings in the mailbox to make it available in future sessions.
    Office.context.roamingSettings.saveAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        } else {
        console.log(`Settings saved with status: ${result.status}`);
        }
    });
}

Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
        $("#get").click(get);
        $("#set").click(set);
        $("#save").click(save);
    });
});