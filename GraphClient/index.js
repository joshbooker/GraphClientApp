const config = {
     "TENANT": process.env["TENANT"],
     "GRAPH_URL": process.env["GRAPH_URL"],
     "CLIENT_ID": process.env["CLIENT_ID"],
     "CLIENT_SECRET": process.env["CLIENT_SECRET"],
     "USERS_ENDPOINT": process.env["USERS_ENDPOINT"],
     "FROM_USERID": process.env["FROM_USERID"]
};

var adal = require('adal-node');
var fetch = require('node-fetch');

module.exports = function (context, req) {
    req.body = JSON.parse(req.body);
    if (req.query.method || (req.body && req.body.method)){
        var method = (req.query.method || req.body.method );//handle no method
        switch (method) {
            case 'getToken':
                    context.log('getToken entered');
                    getToken()
                    .then(function(tokenRes){
                        context.res = {
                            body: tokenRes
                        };
                        context.done(); 
                    })
                    .catch(function(err){
                        context.res = {
                            body: JSON.parse(JSON.stringify(err))
                        };
                        context.done();        
                    });
                break;
            case 'sendMail':
                context.log('sendMail entered');
                var subject = (req.query.subject || (req.body && req.body.subject));
                var bodyContentType = (req.query.bodyContentType || (req.body && req.body.bodyContentType));
                var bodyContent = (req.query.bodyContent || (req.body && req.body.bodyContent));
                var toRecipientEmails = (req.query.toRecipientEmails || (req.body && req.body.toRecipientEmails));
                var ccRecipientEmails = (req.query.ccRecipientEmails || (req.body && req.body.ccRecipientEmails));
                var bccRecipientEmails = (req.query.bccRecipientEmails || (req.body && req.body.bccRecipientEmails));
                var saveToSentItems = (req.query.saveToSentItems || (req.body && req.body.saveToSentItems));
                var fromUserId = (req.query.fromUserId || (req.body && req.body.fromUserId));
                context.log('sendMail var loaded');
                sendMail(subject, bodyContentType, bodyContent, toRecipientEmails, ccRecipientEmails, bccRecipientEmails, saveToSentItems, fromUserId)     
                .then(function(response){
                    context.log("res: " + JSON.stringify(response));
                        context.res = {
                            status: response.status,
                            statusText: response.statusText,
                            body: {
                                status: response.status,
                                statusText: response.statusText,
                                body: JSON.stringify(response)
                                }
                        };
                        context.done(); 
                    })
                    .catch(function(err){
                        context.log("err: " + JSON.stringify(err));
                        context.res = {
                            status: 400,
                            body: "err: " + JSON.stringify(err)
                        };
                        context.done();        
                    });
                break;
            case 'valueN':
                //Statements executed when the result of expression matches valueN
                context.done(); 
                break;
            default:
                context.res = {
                    status: 400,
                    body: "Please pass a valid method on the query string or in the request body"
                };
                context.done(); 
                break;
        }
    } else {
        context.res = {
            status: 400,
            body: "Please pass a method on the query string or in the request body"
        };
        context.done(); 
    }
};

function getToken() {
    return new Promise((resolve, reject) => {
        const authContext = new adal.AuthenticationContext(`https://login.microsoftonline.com/${config.TENANT}`);
        authContext.acquireTokenWithClientCredentials(config.GRAPH_URL, config.CLIENT_ID, config.CLIENT_SECRET, (err, tokenRes) => {
            if (err) {
                reject(err);
            }
            resolve(tokenRes);
        });
    });
}

    /*
    * Send Email on behalf or current user
    * 
    * @param {any} subject - subject
    * @param {any} bodyContentType - "text" or "HTML"
    * @param {any} bodyContent - bodyContent
    * @param {any} toRecipientEmail - TO email address
    * @param {any} ccRecipientEmail - CC email address
    * @param {any} saveToSentItems - true or false
    */
function sendMail(subject, bodyContentType, bodyContent, toRecipientEmails, ccRecipientEmails, bccRecipientEmails, saveToSentItems, fromUserId) {
    //return new Promise((resolve, reject) => {
    var body = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": bodyContentType,
                "content": bodyContent
            },
            "toRecipients": [
            ],
            "ccRecipients": [
            ],
            "bccRecipients": [
            ]
        },
        "saveToSentItems": saveToSentItems
    };
    toRecipientEmails.forEach(function(email) {
        body.message.toRecipients.push({"emailAddress": {"address": email}});
    });
    ccRecipientEmails.forEach(function(email) {
        body.message.ccRecipients.push({"emailAddress": {"address": email}});
    });
    bccRecipientEmails.forEach(function(email) {
        body.message.bccRecipients.push({"emailAddress": {"address": email}});
    });
    fromUserId = fromUserId ? fromUserId : config.FROM_USERID
    var endpoint = config.GRAPH_URL + config.USERS_ENDPOINT + "/" + fromUserId + "/sendMail";
    return callGraphApi(endpoint, "POST", body)
            .then(function (response) {
                return Promise.resolve(response);
            })
            .catch(function (error) {
                return Promise.reject("e3" + error);
            });
}
   /*
    * Get an access token then Call a Web API .
    * 
    * @param {any} endpoint - Web API endpoint
    * @param {any} method - http method ["GET", "POST"]
    * @param {any} body - http body for POST
    * @param {object} headers - optional - new Headers()
    */
function callGraphApi(endpoint, method, body, headers) {
    return getToken()
        .then(function (tokenRes) {
            return new Promise((resolve, reject) => {
                if (!headers) { 
                    var bearer = "Bearer " + tokenRes.accessToken;
                    headers = {
                        "Authorization": bearer,
                        "Content-Type": "application/json"
                    };
                }
                var options = {
                    method: method,
                    headers: headers
                };
                if (body) { options.body = JSON.stringify(body) };
                fetch(endpoint, options)
                    .then(function (response) {
                        resolve(response);
                    })
                    .catch(function (error) {
                        reject("e2" + error);
                    });
            });
        })
        .catch(function (err) {
            return "e1" + err;
        });
}