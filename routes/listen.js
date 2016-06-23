/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
var express = require('express');
var router = express.Router();
var io = require('../helpers/socketHelper.js');
var requestHelper = require('../helpers/requestHelper.js');
var dbHelper = new (require('../helpers/dbHelper'))();
var http = require('http');
var clientStateValueExpected = require('../constants').subscriptionConfiguration.clientState;
var util = require('util'); 

/* Default listen route */
router.post('/', function (req, res, next) {
  var status;
  var clientStatesValid;
  var i;
  var resource;
  var subscriptionId;
  // If there's a validationToken parameter in the query string,
  // then this is the request that Office 365 sends to check
  // that this is a valid endpoint.
  // Just send the validationToken back.
  if (req.query && req.query.validationToken) {
    res.send(req.query.validationToken);
    // Send a status of 'Ok'
    status = 200;
  } else {
    clientStatesValid = false;

    // First, validate all the clientState values in array
    for (i = 0; i < req.body.value.length; i++) {
      if (req.body.value[i].clientState !== clientStateValueExpected) {
        // If just one clientState is invalid, we discard the whole batch
        clientStatesValid = false;
        break;
      } else {
        clientStatesValid = true;
      }
    }

    // If all the clientStates are valid, then
    // process the notification
    if (clientStatesValid) {
      for (i = 0; i < req.body.value.length; i++) {
        resource = req.body.value[i].resource;
        subscriptionId = req.body.value[i].subscriptionId;
        processNotification(subscriptionId, resource, res, next);
      }
      // Send a status of 'Accepted'
      status = 202;
    } else {
      // Since the clientState field doesn't have the expected value,
      // this request might NOT come from Microsoft Graph.
      // However, you should still return the same status that you'd
      // return to Microsoft Graph to not alert possible impostors
      // that you have discovered them.
      status = 202;
    }
  }
  res.status(status).end(http.STATUS_CODES[status]);
});

// Get subscription data from the database
// Retrieve the actual mail message data from Office 365.
// Send the message data to the socket.
function processNotification(subscriptionId, resource, res, next) {
  dbHelper.getSubscription(subscriptionId, function (dbError, subscriptionData) {
    if (subscriptionData) {

      requestHelper.getData(
        '/beta/' + resource, subscriptionData.accessToken,
        function (requestError, endpointData) {
          if (endpointData) {

            //patchEmail(endpointData, subscriptionData.accessToken );
            postNewEmail(endpointData, subscriptionData.accessToken ); 

            io.to(subscriptionId).emit('notification_received', endpointData);
          } else if (requestError) {
            res.status(500);
            next(requestError);
          }
        }
      );
    } else if (dbError) {
      res.status(500);
      next(dbError);
    }
  });
}


function postNewEmail (mailData, token) {
  //console.log(util.inspect(mailData, {showHidden:true, depth:null}))


  var newEmail = {
    "attendees": mailData.toRecipients,
    "start": mailData.startDateTime,
    "end": mailData.endDateTime,
    "subject": "Mobile Friendly Dial In - from Hack Pack",
    "body": {
            contentType: "html",
            content: '"' + "<html>Woot -> <a href='tel:" + collectNumber(mailData, token) + "%23" + "'>Click here for mobile dial-in</a></html>" + '"'
    }
  }; 

  requestHelper.postData( '/beta/me/events',
        token, 
        JSON.stringify(newEmail),
        function (requestError, subscriptionData) {
          if (subscriptionData) {
            console.log("done");
            
          } else if (requestError) {
            console.log(requestError);
          }
        }); 
}

function collectNumber (mailData, token) {
  var emailbody = mailData.body.content;
  var skypecheck = emailbody.indexOf("This is an online meeting for Skype for Business, the professional meetings and communications app formerly known as Lync."); 
  var addincheck = emailbody.indexOf("Quick dial-in link for mobile users");
  var joinbyphonepos = emailbody.indexOf("Join by Phone</span>");

if (skypecheck > 0 && addincheck < 0 && typeof mailData.meetingMessageType != 'undefined') {   

  var confidpos = emailbody.indexOf("Conference ID: ") + 15;    
  //console.log("confidpos: " + confidpos);
  var confid = emailbody.substr(confidpos, 9).toString();      
  var telregex = / (((\(\d{3}\) ?)|(\d{3}-))?\d{3}-\d{4})|(\d{11})/;
  //console.log(emailbody.match(telregex));
  //console.log(emailbody);
  var tel = telregex.exec(emailbody)[0].toString(); // Matches first phone number only. 
  // console.log('Tel before replaces: ' + tel);
  tel = tel.replace('-', '');
  tel = tel.replace('(', '');
  tel = tel.replace(')', '');
  tel = tel.replace(' ', '');
  tel = tel.replace(' ', '');
  var mobiletel = tel + ',,,' + confid;     
  //var newemailbody = emailbody.replace("\r\n</body>\r\n</html>\r\n", "<div><span>Quick dial-in link for mobile users: " + mobiletel + "</div></span>\r\n</body>\r\n</html>\r\n");
 return mobiletel; 
}



}

// PATCH - send event update to Graph
function patchEmail (mailData, token) {

  var emailbody = mailData.body.content;
  var skypecheck = emailbody.indexOf("This is an online meeting for Skype for Business, the professional meetings and communications app formerly known as Lync."); 
  var addincheck = emailbody.indexOf("Quick dial-in link for mobile users");
  var joinbyphonepos = emailbody.indexOf("Join by Phone</span>");

if (skypecheck > 0 && addincheck < 0 && typeof mailData.meetingMessageType != 'undefined') {   

  var confidpos = emailbody.indexOf("Conference ID: ") + 15;    
  //console.log("confidpos: " + confidpos);
  var confid = emailbody.substr(confidpos, 9).toString();      
  var telregex = / (((\(\d{3}\) ?)|(\d{3}-))?\d{3}-\d{4})|(\d{11})/;
  //console.log(emailbody.match(telregex));
  //console.log(emailbody);
  var tel = telregex.exec(emailbody)[0].toString(); // Matches first phone number only. 
  // console.log('Tel before replaces: ' + tel);
  tel = tel.replace('-', '');
  tel = tel.replace('(', '');
  tel = tel.replace(')', '');
  tel = tel.replace(' ', '');
  tel = tel.replace(' ', '');
  var mobiletel = tel + ',,,' + confid;     
  var newemailbody = emailbody.replace("\r\n</body>\r\n</html>\r\n", "<div><span>Quick dial-in link for mobile users: " + mobiletel + "</div></span>\r\n</body>\r\n</html>\r\n");


requestHelper.patchData(
        '/beta/me/messages/' + mailData.id,
        token,
        JSON.stringify({body: {
            contentType: "html",
            content: '"' + "<html>hello</html>" + '"'
        }}),
        function (requestError, subscriptionData) {
          if (subscriptionData) {
            console.log("done");
            
          } else if (requestError) {
            console.log(requestError);
          }
        }
   );

 }
else {}
}
  

module.exports = router;
