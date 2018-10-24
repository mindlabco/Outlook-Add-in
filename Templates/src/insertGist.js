(function(){
    'use strict';
  
    var config;
    var settingsDialog;
    var count = 0;
    Office.initialize = function(reason){
      config = getConfig();
  
      jQuery(document).ready(function(){
        // Check if add-in is configured
        if (config && config.gitHubUserName) {
          // If configured load the gist list
          loadGists(config.gitHubUserName);
        } else {
          // Not configured yet
          $('#not-configured').show();
        }
  
        // When insert button is clicked, build the content
        // and insert into the body.
        $('#insert-button').on('click', function(){
          var gistId = $('.ms-ListItem.is-selected').children('.gist-id').val();
          getGist(gistId, function(gist, error) {
            if (gist) {
              buildBodyContent(gist, function (content, error) {
                if (content) {
                  if(count == 0)
                  {
                    var mailtxt = "<p>Hi,</p><p>Please find Gists as below.</p><p>..."+content+"...</p>";
                    count++;
                  }
                  else
                  {
                    var mailtxt = "<p>..."+content + "...</p>";
                  }
                  Office.context.mailbox.item.body.setSelectedDataAsync(mailtxt,
                    {coercionType: Office.CoercionType.Html}, function(result) {
                      if (result.status == 'failed') {
                        showError('Could not insert Gist: ' + result.error.message);
                      }
                  });
                } else {
                  showError('Could not create insertable content: ' + error);
                } 
              });
            } else {
              showError('Could not retreive Gist: ' + error);
            }
          });
        });
  
        // When the settings icon is clicked, open the settings dialog
        $('#settings-icon').on('click', function(){
          // Display settings dialog
          var url = new URI('../src/setting.html').absoluteTo(window.location).toString();
          if (config) {
            // If the add-in has already been configured, pass the existing values
            // to the dialog
            url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
          }
  
          var dialogOptions = { width: 20, height: 40, displayInIframe: true };
          
          Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
            settingsDialog = result.value;
            settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
            settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
          });
        })
      });
    };
    
    function loadGists(user) {
      $('#error-display').hide();
      $('#not-configured').hide();
      $('#gist-list-container').show();
  
      getUserGists(user, function(gists, error) {
        if (error) {
  
        } else {
          buildGistList($('#gist-list'), gists, onGistSelected);
        }
      });
    }
  
    function onGistSelected() {
      $('.ms-ListItem').removeClass('is-selected');
      $(this).addClass('is-selected');
      $('#insert-button').removeAttr('disabled');
    }
  
    function showError(error) {
      $('#not-configured').hide();
      $('#gist-list-container').hide();
      $('#error-display').text(error);
      $('#error-display').show();
    }
  
    function receiveMessage(message) {
      config = JSON.parse(message.message);
      setConfig(config, function(result) {
        settingsDialog.close();
        settingsDialog = null;
        loadGists(config.gitHubUserName);
      });
    }
  
    function dialogClosed(message) {
      settingsDialog = null;
    }
  })();