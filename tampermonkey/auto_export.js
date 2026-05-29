// ==UserScript==
// @name         Auto Export Article
// @namespace    kb-tools
// @version      1.0
// @description  Automatically clicks Export Article on the word export page
// @match        *://your-intranet/modules.php?name=Knowledge&file=wordexport*
// @grant        none
// ==/UserScript==

// Update the @match directive above to point to your knowledge base's word export URL.
// The script expects a page containing an input[type="submit"] with value "Export Article".

(function() {
    'use strict';
    const btn = document.querySelector('input[type="submit"][value="Export Article"]');
    if (btn) {
        btn.click();
        setTimeout(() => window.close(), 1000);
    }
})();
