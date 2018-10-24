/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */



$(document).ready(() => {
    $('#run').click(run);
});
  
// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

async function run() {
    Office.context.mailbox.item.body.setSelectedDataAsync($('#app-body').html(),
        {coercionType: Office.CoercionType.Html}, function(result) {
          if (result.status == 'failed') {
            showError('Could not insert Gist: ' + result.error.message);
          }
      });
}