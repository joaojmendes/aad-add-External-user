
// Add Guest User to AAD

import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import * as request from "request-promise";
import * as sgMail from '@sendgrid/mail';

 
// Env Vars
const TENANT = getEnviromentVariable("TENANT");
const CLIENT_ID = getEnviromentVariable("CLIENT_ID");
const CLIENT_SECRET = getEnviromentVariable("CLIENT_SECRET");
const GROUP_ID = getEnviromentVariable("GROUP_ID");
const SENDGRID_KEY = getEnviromentVariable("SENDGRID_KEY");

// Interfaces
interface IAADToken {
    token_type: string;
    expires_in: number;
    ext_expires_in: number;
    access_token: string;
  }

// Add Guest User to a Group 
const httpTrigger: AzureFunction = async function(context: Context, req: HttpRequest): Promise<void> {
  context.log("HTTP trigger function processed a request.");
  const userId = req.query.userId || (req.body && req.body.userId);
  const groupId = req.query.groupId || (req.body && req.body.groupId);

 // check request parameters
  if (userId && groupId) {
    try {
      // run Main function 
        await run();
        context.log(`User ${userId} was add to group id : ${GROUP_ID} `);
        // Send Email 
        await sendEmail();
        context.res = {
            // status: 200, /* Defaults to 200 */          
            body: `User ${userId} was add to group id : ${GROUP_ID} `
          };
    } catch (error) {
        context.res = {
            status: 400,
            body: error.message
          };
    }
    
  } else {
    context.log("Please pass a userId  on the query string or in the request body")
    context.res = {
      status: 400,
      body: "Please pass a userId  on the query string or in the request body"
    };
  }
  // Run Main Function
  async function run():Promise<void> {
      try {
        // Get Access Token
        const accessToken = await getAccessToken();
        if (accessToken){
           // Create Invitation 
           let options = {
                method: 'POST',
                resolveWithFullResponse: true,
                url: 'https://graph.microsoft.com/beta/invitations',
                headers: {
                    'Authorization': 'Bearer ' + accessToken,
                    'content-type': 'application/json'
                },
                body: JSON.stringify({
                    "invitedUserDisplayName": userId,
                    "invitedUserEmailAddress": userId,
                    "inviteRedirectUrl": "https://URL-TO-SITE",
                    "sendInvitationMessage": false
                })
            }        
           const invitationResponse = await request(options);
            
             if (invitationResponse.statusCode == 201){
               // Add addUser to O365 Group 
               const result = JSON.parse(invitationResponse.body);
             const invitedUserId: string  = result.invitedUser.id;

                let options = {
                    method: 'POST',
                    resolveWithFullResponse: true,
                    url: `https://graph.microsoft.com/v1.0/groups/${groupId}/members/$ref`,
                    headers: {
                        'Authorization': 'Bearer ' + accessToken,
                        'content-type': 'application/json'
                    },
                    body: JSON.stringify({
                        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${invitedUserId}`
                    })
                };
                const response = await request(options);        
             }
        }
      } catch (error) {
          context.log(error);
          throw new Error(error);
      }
  }

  // Send Confirmation Email.
  async function sendEmail(){
      try {
          sgMail.setApiKey(SENDGRID_KEY);
          const fromUser = 'joao.mendes@devjjm.onmicrosoft.com';
          const msg = {
            to: userId,
            from: fromUser,
            subject: 'Access to Teams ',
            text: 'Access to Teams Confirm Message',
            html: '<strong>Access to Teams Confirm Message</strong>',
          };
          
          await sgMail.send(msg);
      } catch (error) {
          context.log(error);
      }
  }
  // Get Access Token 
  async function getAccessToken(): Promise<string> {
    try {
      
      let options = {
        method: 'POST',
        uri: `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,        
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        form: {
          grant_type: 'client_credentials',
          client_id: `${CLIENT_ID}`,
          client_secret: `${CLIENT_SECRET}`,
          scope: 'https://graph.microsoft.com/.default'
        }
      };
      const results = await request(options);
      const aadToken: IAADToken = JSON.parse(results);
      return aadToken.access_token;
    } catch (error) {
      throw new Error(`Error getting MSgraph token: ${error.message}`);
    }
  }
}
  
// Get Env Var
function getEnviromentVariable(name: string): string {
  return process.env[name];
}

export default httpTrigger;
