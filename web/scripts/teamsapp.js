(function () {
    'use strict';

    // Call the initialize API first
    microsoftTeams.initialize();

    // Check the initial theme user chose and respect it
    microsoftTeams.getContext(function (context) {
        if (context && context.theme) {
            setTheme(context.theme);
        }
    });

    // Handle theme changes
    microsoftTeams.registerOnThemeChangeHandler(function (theme) {
        setTheme(theme);
    });

    // Save configuration changes
    microsoftTeams.settings.registerOnSaveHandler(function (saveEvent) {
        // Let the Microsoft Teams platform know what you want to load based on
        // what the user configured on this page
        microsoftTeams.settings.setSettings({
            contentUrl: createTabUrl(), // Mandatory parameter
            entityId: createTabUrl(), // Mandatory parameter
        });

        // Tells Microsoft Teams platform that we are done saving our settings. Microsoft Teams waits
        // for the app to call this API before it dismisses the dialog. If the wait times out, you will
        // see an error indicating that the configuration settings could not be saved.
        saveEvent.notifySuccess();
    });

    // Logic to let the user configure what they want to see in the tab being loaded
    /*document.addEventListener('DOMContentLoaded', function () {
        var newTab = document.getElementById('newTab');
        if (newTab) {
            newTab.onclick = function () {
                microsoftTeams.settings.setValidityState(true);
            };
        }
    });*/
    document.addEventListener('fullscreenchange',
    function()
    {var fullscreencontrol = document.getElementById('main').contentWindow.document.getElementById('fullscreen-controls');
    if (document.fullscreenEnabled)
    {
        fullscreencontrol.removeAttribute('hidden');
    }
    else
    {
        fullscreencontrol.addAttribute('hidden');
    }
    });

    // Set the desired theme
    function setTheme(theme) {
        if (theme) {
            // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
            document.body.className =
                'theme-' + (theme === 'default' ? 'light' : theme);
        }
    }

    // Create the URL that Microsoft Teams will load in the tab. You can compose any URL even with query strings.
    //function createTabUrl() {return (window.location.protocol + '//' + window.location.host + '/' +selectedTab);}

})();
