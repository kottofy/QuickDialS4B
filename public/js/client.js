/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
var socket = io.connect('http://localhost:3001'); // eslint-disable-line no-undef
var subscriptionId;
var userId;

// Socket `notification_received` event handler.
socket.on('notification_received', function (mailData) {


  var emailbody = mailData.body.content;
  var skypecheck = emailbody.indexOf("This is an online meeting for Skype for Business, the professional meetings and communications app formerly known as Lync."); 
  var addincheck = emailbody.indexOf("Quick dial-in link for mobile users");
  var joinbyphonepos = emailbody.indexOf("Join by Phone</span>");

if (skypecheck > 0 && addincheck < 0 && typeof mailData.meetingMessageType != 'undefined') {   

  var confidpos = emailbody.indexOf("Conference ID: ") + 15;    
  console.log("confidpos: " + confidpos);
  var confid = emailbody.substr(confidpos, 9).toString();    
  //var telpos = emailbody.indexOf("<a href=\"tel:") + 14;   
  //var telregex = /((\(\d{3}\) ?)|(\d{3}-))?\d{3}-\d{4}/;
  var telregex = /^(\(?\+?[0-9]*\)?)?[0-9_\- \(\)]*$/;
  var tel = emailbody.match(telregex)[0].toString(); // Matches first phone number only. 
  var mobiletel = tel + ',,,' + confid;     
  var newemailbody = emailbody.replace("\r\n</body>\r\n</html>\r\n", "<div><span>Quick dial-in link for mobile users: " + mobiletel + "</div></span>\r\n</body>\r\n</html>\r\n");
  console.log('ConfID: ' + confid);
  console.log('Tel: ' + tel);
  console.log('New tel number: ' + mobiletel); 
 }

else {}

  var listItem;
  var primaryText;
  var secondaryText;

  listItem = document.createElement('div');
  listItem.className = 'ms-ListItem is-selectable';
  listItem.onclick = function () {
    window.open(mailData.webLink, 'outlook');
  };

  primaryText = document.createElement('span');
  primaryText.className = 'ms-ListItem-primaryText';
  primaryText.innerText = mailData.sender.emailAddress.name;
  secondaryText = document.createElement('span');
  secondaryText.className = 'ms-ListItem-secondaryText';
  secondaryText.innerText = mailData.subject;
  listItem.appendChild(primaryText);
  listItem.appendChild(secondaryText);

  document.getElementById('notifications').appendChild(listItem);
});

// When the page first loads, create the socket room.
subscriptionId = getQueryStringParameter('subscriptionId');
socket.emit('create_room', subscriptionId);
document.getElementById('subscriptionId').innerHTML = subscriptionId;

// The page also needs to send the userId to properly
// sign out the user.
userId = getQueryStringParameter('userId');
document.getElementById('userId').innerHTML = userId;
document.getElementById('signOutButton').onclick = function () {
  location.href = '/signout/' + subscriptionId;
};

function getQueryStringParameter(paramToRetrieve) {
  var params = document.URL.split('?')[1].split('&');
  var i;
  var singleParam;

  for (i = 0; i < params.length; i = i + 1) {
    singleParam = params[i].split('=');
    if (singleParam[0] === paramToRetrieve) {
      return singleParam[1];
    }
  }
  return null;
}
