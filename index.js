const express = require("express");
const msal = require('@azure/msal-node');
const { response } = require("express");
var querystring = require('querystring');
var https = require('https');
var base64 = require('base64-min');


var unirest = require('unirest');

const SERVER_PORT = process.env.PORT || 3000;

// Create Express App and Routes
const app = express();

var accessToken;

const config = {
    auth: {
        clientId: "8a14009d-8dcf-4dae-96b9-a142cfd8d1f6",
        authority: "https://login.microsoftonline.com/common",
        clientSecret: "_a3~n6-GxPTvZ9njnwXR4it.n2zVZi.y.x"
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
};

// Create msal application object
const cca = new msal.ConfidentialClientApplication(config);

app.get('/', (req, res) => {
    const authCodeUrlParameters = {
        scopes: ["APIConnectors.Read.All",
            "APIConnectors.ReadWrite.All",
            "Mail.Read",
            "Mail.Read.Shared",
            "Mail.ReadBasic",
            "Mail.ReadWrite",
            "Mail.ReadWrite.Shared",
            "Mail.Send",
            "Mail.Send.Shared",
            "User.Read",
            "User.Read.All"],
        redirectUri: "http://localhost:3000/redirect",
    };

    cca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
        res.redirect(response);
    }).catch((error) => console.log(error));
    });

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["APIConnectors.Read.All",
            "APIConnectors.ReadWrite.All",
            "Mail.Read",
            "Mail.Read.Shared",
            "Mail.ReadBasic",
            "Mail.ReadWrite",
            "Mail.ReadWrite.Shared",
            "Mail.Send",
            "Mail.Send.Shared",
            "User.Read",
            "User.Read.All"],
        redirectUri: "http://localhost:3000/redirect",
    };

    cca.acquireTokenByCode(tokenRequest).then((response) => {
        console.log("\nResponse: \n:", response);
        // console.log(">"+ response.accessToken)
        accessTokens = response.accessToken;
       // performRequest(path);
        fleskRequest();
        res.sendStatus(200);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});

// const account = cca.getAllAccounts();
// console.log(account);

app.get('/refresh', async (req, res) => {
    // console.log({ account });
    console.log(cca.getTokenCache());
    let msalTokenCache = cca.getTokenCache();
    const cachedAccounts = await msalTokenCache.getAllAccounts();
    console.log("----> ", cachedAccounts);
    if (cachedAccounts.length > 0) {
        const accessTokenRequest = {
            scopes: ["user.read"],
            account: cachedAccounts[0]
        }
        cca.acquireTokenSilent(accessTokenRequest).then(accessTokenResponse => {
            console.log("Response from refresh :", accessTokenResponse);
            var accessTokens = accessTokenResponse.accessToken;
            res.sendStatus(200);
        }).catch((error) => {
            console.log(error);
            res.status(500).send(error);
        });
    } else {
        console.log("No cached accounts has been found.");
    }
});

// publicClientApplication.acquireTokenSilent(accessTokenRequest).then(function(accessTokenResponse) {
//     // Acquire token silent success
//     let accessToken = accessTokenResponse.accessToken;
//     // Call your API with token
//     callApi(accessToken);
// }).catch(function (error) {
//     //Acquire token silent failure, and send an interactive request
//     if (error instanceof InteractionRequiredAuthError) {
//         publicClientApplication.acquireTokenPopup(accessTokenRequest).then(function(accessTokenResponse) {
//             // Acquire token interactive success
//             let accessToken = accessTokenResponse.accessToken;
//             // Call your API with token
//             callApi(accessToken);
//         }).catch(function(error) {
//             // Acquire token interactive failure
//             console.log(error);
//         });
//     }
//     console.log(error);
// });



var fs = require('fs');


var messageid = 'AAMkAGRjMDczMDE0LTBlMzgtNGE5NS1iNDZiLWZiNDQxNzkzZTEzNwBGAAAAAAAfMlnUxxOwQIMw1t1YPXuxBwBtOM7YAAIKS7FQ07h9Mh8mAAAAAAEMAABtOM7YAAIKS7FQ07h9Mh8mAAAZAGKWAAA=';
var attachmentId= 'AAMkAGRjMDczMDE0LTBlMzgtNGE5NS1iNDZiLWZiNDQxNzkzZTEzNwBGAAAAAAAfMlnUxxOwQIMw1t1YPXuxBwBtOM7YAAIKS7FQ07h9Mh8mAAAAAAEMAABtOM7YAAIKS7FQ07h9Mh8mAAAZAGKWAAABEgAQAAf8B4xREMZLnrTy4VcpEvU=';
var host = 'graph.microsoft.com';
var path = `/v1.0/me/messages/${messageid}/attachments/${attachmentId}`;

performRequest(path);




// function performRequest(endpoint, method, data, success) {
//     console.log('Landed Herer', accessTokens);
//   var dataString = JSON.stringify(data);

//   var method = 'GET';
 
//   var headers = {
//       Authorization: 'Bearer '+ accessTokens,
    
//   };
  
//   if (method == 'GET' && false) {
//     endpoint += '?' + querystring.stringify(data);
//   }

//   var options = {
//     hostname: host,
//     path: endpoint,
//     port: 443,
//     method: method,
//     headers: headers
//   };

//   console.log(options);

//   var req = https.request(options, function(res) {
//     res.setEncoding('utf-8');

//     var responseString = '';
    
//     res.on('data', d => {
//         responseString += d;

//       })

//     res.on('end', function(re) {
//         const data = JSON.parse(responseString).contentBytes;
//         console.log('end', data);
//         const type = JSON.parse(responseString).contentType;
//         const ext= type.split("/")[1];
//         console.log('type-->', ext);
//         const decode = base64.decodeToFile(data,`attachments/test.${ext}`) ;
           
//     });
//   });

  
//     req.on('error', error => {
//         console.error(error)
//     })

//     req.write('');
//     req.end();

// }


// // var API_KEY = 'dzZZfsTuWMskKzlX6j5S';
// //   var FD_ENDPOINT = 'newaccount1625481250603';
// //   var PATH = '/api/v2/tickets/21';
// //   var enocoding_method = 'base64';
// //   var auth = 'Basic ' + Buffer.from(API_KEY + ':' + 'X').toString(enocoding_method);
// //   var URL =  'https://' + FD_ENDPOINT + '.freshdesk.com' + PATH;


// // function fleskRequest() {

// // var fields = {
// //   'email': 'email@yourdomain.com',
// //   'subject': 'Tic0ket subject',
// //   'description': 'Ticket description.',
// //   'status': 2,
// //   'priority': 1
// // }

// // var headers = {
// //   'Authorization': auth
// // }


// // unirest.get(URL)
// //   .headers(headers)
// //   .field(fields)
// //   .attach('attachments[]', fs.createReadStream('./attachments/test'))
// //   .end(function(response){
// //     console.log(response.body)
// //     console.log("Response Status : " + response.status)
// //     if(response.status == 201){
// //       console.log("Location Header : "+ response.headers['location'])
// //     }
// //     else{
// //       console.log("X-Request-Id :" + response.headers['x-request-id']);
// //     }
// //   });


  
// // }



// var axios = require('axios');
// var config_1 = {
//   method: 'get',
//   url: "https://geekfactoryassist.freshdesk.com/api/v2/tickets/4791",
//   headers: { 
//     'Authorization': 'Basic RjZBSjJpd1c5NnNXbjZhZzVGSjp4', 
//     //'Cookie': '_x_w=31_1; _x_m=x_c'
//   }
// };
// axios(config_1)
// .then(function (response) {
//   console.log(JSON.stringify(response.data));
//   const type_1= JSON.stringify(response.data.attachments[0].content_type);
//   const end= type_1.split("/")[1].replace('"','');
//   const path_1= (JSON.stringify(response.data.attachments[0].attachment_url));
//   console.log(path_1)
//   var ret = path_1.replace('"','');
//   var ret1 = ret.replace('"','');
//  console.log(ret1);

//  function saveImageToDisk(url, localPath) {
  
//     var file = fs.createWriteStream(localPath);
//     var request = https.get(url, function(response) {
//     response.pipe(file);
//     });
//     }

// saveImageToDisk(ret1, "./attach/"+ Date.now()+`.${end}`);

//  })
// .catch(function (error) {
//   console.log(error);
// });


// accessToken="eyJ0eXAiOiJKV1QiLCJub25jZSI6IlphYTNuc25TSlZ1cThMdE5icGQwSndaVHZ2ZFhyS2UzdGFEN2M0dGZvRXciLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9jN2Q4NTM5Yi1jMDE4LTRmNzMtOGNlZi1jOGNmYzNjYWEzMDYvIiwiaWF0IjoxNjI3NTgxMTU0LCJuYmYiOjE2Mjc1ODExNTQsImV4cCI6MTYyNzU4NTA1NCwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsidXJuOnVzZXI6cmVnaXN0ZXJzZWN1cml0eWluZm8iLCJ1cm46bWljcm9zb2Z0OnJlcTEiLCJ1cm46bWljcm9zb2Z0OnJlcTIiLCJ1cm46bWljcm9zb2Z0OnJlcTMiLCJjMSIsImMyIiwiYzMiLCJjNCIsImM1IiwiYzYiLCJjNyIsImM4IiwiYzkiLCJjMTAiLCJjMTEiLCJjMTIiLCJjMTMiLCJjMTQiLCJjMTUiLCJjMTYiLCJjMTciLCJjMTgiLCJjMTkiLCJjMjAiLCJjMjEiLCJjMjIiLCJjMjMiLCJjMjQiLCJjMjUiXSwiYWlvIjoiQVVRQXUvOFRBQUFBMGhGRytoR2s0VURxb0ZnTGdITHYwYU5rQnFBbVFMWFUyK0pQLzBHRlBxdE5ReGdWQmtSMHU3Z0hqeVc4RTJ5STVXSlozRHZxRkk3UHdEdU1aelZWU3c9PSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiUG93ZXIgQXV0b21hdGUgSW50ZWdyYXRpb24gLSBGRCIsImFwcGlkIjoiOGExNDAwOWQtOGRjZi00ZGFlLTk2YjktYTE0MmNmZDhkMWY2IiwiYXBwaWRhY3IiOiIxIiwiZmFtaWx5X25hbWUiOiJHb3N1IiwiZ2l2ZW5fbmFtZSI6IlNhYWdhcmlrYSIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjE4My44My4xNTMuMjM3IiwibmFtZSI6IlNhYWdhcmlrYSBHb3N1Iiwib2lkIjoiYTQ3N2I1OTktMzY4My00OGE0LWI2NGQtZDc2ZTU3YzBmMjU3IiwicGxhdGYiOiI1IiwicHVpZCI6IjEwMDMyMDAxNEEzNkU2QzUiLCJyaCI6IjAuQVhFQW0xUFl4eGpBYzAtTTc4alB3OHFqQnAwQUZJclBqYTVObHJtaFFzX1kwZlp4QUZjLiIsInNjcCI6IkFQSUNvbm5lY3RvcnMuUmVhZC5BbGwgQVBJQ29ubmVjdG9ycy5SZWFkV3JpdGUuQWxsIE1haWwuUmVhZCBNYWlsLlJlYWQuU2hhcmVkIE1haWwuUmVhZEJhc2ljIE1haWwuUmVhZFdyaXRlIE1haWwuUmVhZFdyaXRlLlNoYXJlZCBNYWlsLlNlbmQgTWFpbC5TZW5kLlNoYXJlZCBvcGVuaWQgcHJvZmlsZSBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IjdtcmlyUUNlTXZHWk1KX3hka1VNRHlqOS1DME5LUlV6Q2Rib0d0cUtFblUiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiQVMiLCJ0aWQiOiJjN2Q4NTM5Yi1jMDE4LTRmNzMtOGNlZi1jOGNmYzNjYWEzMDYiLCJ1bmlxdWVfbmFtZSI6IkpvaG5AdmlydHVhbGFnZW50YXVzLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6IkpvaG5AdmlydHVhbGFnZW50YXVzLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6InJScG1GeVpTNkVlUU1UVFQ3LWxwQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfc3QiOnsic3ViIjoiTm1rNFZlNFliQ3JqUVVZckFzby1JREhMWUExWUxUOFZiN3BzT05rZTRYYyJ9LCJ4bXNfdGNkdCI6MTYyMjQ0ODc3Nn0.jyCsgzeOwTMFTE8fbpsqFrf4FRLw_-ixJ6oTVzgBg-_H9iqtiaaYORBlYHL2Zbxu-6JyUE2UBgab09bVZW8_UMDIZmLFFzWG7O_IYqBLR6gvXz4nGJRch_evaod4JiaPgkj0OKSE6b0t3bS_18fZEyDCQ6Pzt1w22t6LD50XPJqMjTvL-8etIYcxnojuWDFR_jOUmj_1yNP7RJCvsQGAIy10wcCJ0hZylp6qB8UEAtxUNxxWMY34inIymo7H-dAhGfvGMTCguBRBVrT27uXsWzgAZE3Fe8yvWjentqWNwdn6W6koYWNcxHuhKdUU40tfdQgPpnWLgoToSjKEnZBf6A"


var contents = fs.readFileSync("./attachments/test.jpeg", { encoding: 'base64' });
  


function performRequest(endpoint, method, data, success) {
            
var options = {
                hostname: 'graph.microsoft.com', 
                path: endpoint, 
                port: 443, 
                method: 'POST', 
                headers: {
                    'Content-Type' : "application/json",

                    Authorization: 'Bearer ' + "eyJ0eXAiOiJKV1QiLCJub25jZSI6IlphYTNuc25TSlZ1cThMdE5icGQwSndaVHZ2ZFhyS2UzdGFEN2M0dGZvRXciLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9jN2Q4NTM5Yi1jMDE4LTRmNzMtOGNlZi1jOGNmYzNjYWEzMDYvIiwiaWF0IjoxNjI3NTgxMTU0LCJuYmYiOjE2Mjc1ODExNTQsImV4cCI6MTYyNzU4NTA1NCwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsidXJuOnVzZXI6cmVnaXN0ZXJzZWN1cml0eWluZm8iLCJ1cm46bWljcm9zb2Z0OnJlcTEiLCJ1cm46bWljcm9zb2Z0OnJlcTIiLCJ1cm46bWljcm9zb2Z0OnJlcTMiLCJjMSIsImMyIiwiYzMiLCJjNCIsImM1IiwiYzYiLCJjNyIsImM4IiwiYzkiLCJjMTAiLCJjMTEiLCJjMTIiLCJjMTMiLCJjMTQiLCJjMTUiLCJjMTYiLCJjMTciLCJjMTgiLCJjMTkiLCJjMjAiLCJjMjEiLCJjMjIiLCJjMjMiLCJjMjQiLCJjMjUiXSwiYWlvIjoiQVVRQXUvOFRBQUFBMGhGRytoR2s0VURxb0ZnTGdITHYwYU5rQnFBbVFMWFUyK0pQLzBHRlBxdE5ReGdWQmtSMHU3Z0hqeVc4RTJ5STVXSlozRHZxRkk3UHdEdU1aelZWU3c9PSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiUG93ZXIgQXV0b21hdGUgSW50ZWdyYXRpb24gLSBGRCIsImFwcGlkIjoiOGExNDAwOWQtOGRjZi00ZGFlLTk2YjktYTE0MmNmZDhkMWY2IiwiYXBwaWRhY3IiOiIxIiwiZmFtaWx5X25hbWUiOiJHb3N1IiwiZ2l2ZW5fbmFtZSI6IlNhYWdhcmlrYSIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjE4My44My4xNTMuMjM3IiwibmFtZSI6IlNhYWdhcmlrYSBHb3N1Iiwib2lkIjoiYTQ3N2I1OTktMzY4My00OGE0LWI2NGQtZDc2ZTU3YzBmMjU3IiwicGxhdGYiOiI1IiwicHVpZCI6IjEwMDMyMDAxNEEzNkU2QzUiLCJyaCI6IjAuQVhFQW0xUFl4eGpBYzAtTTc4alB3OHFqQnAwQUZJclBqYTVObHJtaFFzX1kwZlp4QUZjLiIsInNjcCI6IkFQSUNvbm5lY3RvcnMuUmVhZC5BbGwgQVBJQ29ubmVjdG9ycy5SZWFkV3JpdGUuQWxsIE1haWwuUmVhZCBNYWlsLlJlYWQuU2hhcmVkIE1haWwuUmVhZEJhc2ljIE1haWwuUmVhZFdyaXRlIE1haWwuUmVhZFdyaXRlLlNoYXJlZCBNYWlsLlNlbmQgTWFpbC5TZW5kLlNoYXJlZCBvcGVuaWQgcHJvZmlsZSBVc2VyLlJlYWQgVXNlci5SZWFkLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IjdtcmlyUUNlTXZHWk1KX3hka1VNRHlqOS1DME5LUlV6Q2Rib0d0cUtFblUiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiQVMiLCJ0aWQiOiJjN2Q4NTM5Yi1jMDE4LTRmNzMtOGNlZi1jOGNmYzNjYWEzMDYiLCJ1bmlxdWVfbmFtZSI6IkpvaG5AdmlydHVhbGFnZW50YXVzLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6IkpvaG5AdmlydHVhbGFnZW50YXVzLm9ubWljcm9zb2Z0LmNvbSIsInV0aSI6InJScG1GeVpTNkVlUU1UVFQ3LWxwQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjYyZTkwMzk0LTY5ZjUtNDIzNy05MTkwLTAxMjE3NzE0NWUxMCIsImI3OWZiZjRkLTNlZjktNDY4OS04MTQzLTc2YjE5NGU4NTUwOSJdLCJ4bXNfc3QiOnsic3ViIjoiTm1rNFZlNFliQ3JqUVVZckFzby1JREhMWUExWUxUOFZiN3BzT05rZTRYYyJ9LCJ4bXNfdGNkdCI6MTYyMjQ0ODc3Nn0.jyCsgzeOwTMFTE8fbpsqFrf4FRLw_-ixJ6oTVzgBg-_H9iqtiaaYORBlYHL2Zbxu-6JyUE2UBgab09bVZW8_UMDIZmLFFzWG7O_IYqBLR6gvXz4nGJRch_evaod4JiaPgkj0OKSE6b0t3bS_18fZEyDCQ6Pzt1w22t6LD50XPJqMjTvL-8etIYcxnojuWDFR_jOUmj_1yNP7RJCvsQGAIy10wcCJ0hZylp6qB8UEAtxUNxxWMY34inIymo7H-dAhGfvGMTCguBRBVrT27uXsWzgAZE3Fe8yvWjentqWNwdn6W6koYWNcxHuhKdUU40tfdQgPpnWLgoToSjKEnZBf6A", 
                    body: JSON.stringify({
                        "message": {
                            "attachments": [
                                {
                                    "contentType": "image/jpeg",
                                    "@odata.type": "microsoft.graph.fileAttachment",
                                    "name": "12345",
                                    "contentBytes": contents
                                }
                            ],
                            "ccRecipients": [
                                {
                                    "emailAddress": {
                                        "address": "sarf@geekfactory.tech"
                                    }
                                }
                            ],
                            "toRecipients": [
                                {
                                    "emailAddress": {
                                        "address": "janani@geekfactory.tech",
                                        "name": "janani"
                                    }
                                }
                            ]
                        },
                        "comment": "hello"
                    })
                    // body :
                }
            };

console.log("options -->", options);
        var req = https.request(options, function(res) {
                console.log("res", res.statusCode);
                res.setEncoding('utf-8');
                var responseString = '';
                res.on('data', d => {
                    responseString += d;
                })
                res.on('end', function (re) {
                    console.log("responseString", responseString);
                    // const data = JSON.parse(responseString).contentBytes;
                    // const type = JSON.parse(responseString).contentType;
                    // const ext = type.split("/")[1];
                    // console.log('type-->', ext);
                    // base64.decodeToFile(data, `attachments/${request.attach_name}.${ext}`);
                    // resolve(`${request.attach_name}.${ext}`);
                    // res.sendStatus(200);
                });
            });
            req.on('error', error => {
                console.error(error)
            })
            req.write('');
            req.end();


        }




app.listen(SERVER_PORT, () => console.log(`Msal Node Auth Code Sample app listening on port ${SERVER_PORT}`))







// "APIConnectors.Read.All",
// "APIConnectors.ReadWrite.All",
// "Mail.Read",
// "Mail.Read.Shared",
// "Mail.ReadBasic",
// "Mail.ReadWrite",
// "Mail.ReadWrite.Shared",
// "Mail.Send",
// "Mail.Send.Shared",
// "User.Read",
// "User.Read.All"