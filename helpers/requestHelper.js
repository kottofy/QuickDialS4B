/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
var https = require('https');
var host = 'graph.microsoft.com';

/**
 * Generates a POST request (of Content-type ```application/json```)
 * @param {string} path the path, relative to the host, to which this request will be sent
 * @param {string} token the access token with which the request should be authenticated
 * @param {string} data the data which will be 'POST'ed
 * @param {callback} callback
 */
function postData(path, token, data, callback) {
  var options = {
    host: host,
    path: path,
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + token,
      'Content-Length': data.length
    }
  };

  var req = https.request(options, function (res) {
    var subscriptionData = '';
    res.on('data', function (chunk) {
      subscriptionData += chunk;
    });
    res.on('end', function () {
      if (res.statusCode === 201) {
        callback(null, JSON.parse(subscriptionData));
      } else {
        callback(JSON.parse(subscriptionData), null);
      }
    });
  });

  req.write(data);
  req.end();

  req.on('error', function (error) {
    callback(error, null);
  });
}

/**
 * Generates a GET request (of Content-type ```application/json```)
 * @param {string} path the path, relative to the host, to which this request will be sent
 * @param {string} token the acess token with which the request should be authenticated
 * @param {callback} callback
 */
function getData(path, token, callback) {
  var options = {
    host: host,
    path: path,
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json;odata.metadata=minimal;' +
              'odata.streaming=true;IEEE754Compatible=false',
      Authorization: 'Bearer ' + token
    }
  };

  var req = https.request(options, function (res) {
    var endpointData = '';
    res.on('data', function (chunk) {
      endpointData += chunk;
    });
    res.on('end', function () {
      if (res.statusCode === 200) {
        callback(null, JSON.parse(endpointData));
      } else {
        callback(JSON.parse(endpointData), null);
      }
    });
  });

  req.write('');
  req.end();

  req.on('error', function (error) {
    callback(error, null);
  });
}

/**
 * Generates a DELETE request
 * @param {string} path the path, relative to the host, to which this request will be sent
 * @param {string} token the acess token with which the request should be authenticated
 * @param {callback} callback
 */
function deleteData(path, token, callback) {
  var options = {
    host: host,
    path: path,
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain',
      'X-HTTP-Method': 'DELETE',
      Authorization: 'Bearer ' + token
    }
  };

  var req = https.request(options, function (res) {
    var endpointData = '';
    res.on('data', function (chunk) {
      endpointData += chunk;
    });
    res.on('end', function () {
      callback(null);
    });
  });

  req.end();

  req.on('error', function (error) {
    callback(error);
  });
}

/**
 * Generates a POST request (of Content-type ```application/json```)
 * @param {string} path the path, relative to the host, to which this request will be sent
 * @param {string} token the access token with which the request should be authenticated
 * @param {string} data the data which will be 'POST'ed
 * @param {callback} callback
 */
function patchData(path, token, data, callback) {
  console.log("PATH: " + path);
  console.log("DATA: " + data + "\n");
  var options = {
    host: host,
    path: path,
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + token,
      'Content-Length': data.length
    }
  };

  var req = https.request(options, function (res) {
    var subscriptionData = '';
    res.on('data', function (chunk) {
      subscriptionData += chunk;
    });
    res.on('end', function () {
      if (res.statusCode === 201) {
        callback(null, JSON.parse(subscriptionData));
      } else {
        callback(JSON.parse(subscriptionData), null);
      }
    });
  });

  req.write(data);
  req.end();

  req.on('error', function (error) {
    callback(error, null);
  });
}


exports.patchData = patchData;
exports.postData = postData;
exports.getData = getData;
exports.deleteData = deleteData;
