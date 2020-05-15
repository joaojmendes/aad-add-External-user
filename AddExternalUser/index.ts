
// Add Guest User to AAD

import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import * as request from "request-promise";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
 
// Env Vars
const TENANT = getEnviromentVariable("TENANT");
const CLIENT_ID = getEnviromentVariable("CLIENT_ID");
const CLIENT_SECRET = getEnviromentVariable("CLIENT_SECRET");
const GROUP_ID = getEnviromentVariable("GROUP_ID");


// Interfaces
interface IAADToken {
    token_type: string;
    expires_in: number;
    ext_expires_in: number;
    access_token: string;
  }

// response of function 
interface IReturnResp {
  groupId: string;
  invitation: MicrosoftGraph.Invitation,
}

// Add Guest User to a Group 
const httpTrigger: AzureFunction = async function(context: Context, req: HttpRequest): Promise<void> {
  context.log("HTTP trigger function processed a request.");
  const userId = (req.body && req.body.userId);
  const groupId =  (req.body && req.body.groupId);
  const userName =  (req.body && req.body.userName);
  

let invitedUserMessageInfo: MicrosoftGraph.InvitedUserMessageInfo = {} as MicrosoftGraph.InvitedUserMessageInfo;

 // check request parameters
  if (userId && groupId) {
    try {

      
      // run Main function 
     const returnResp: IReturnResp = await run();
        context.log(`User ${userId} was add to group id : ${GROUP_ID} `);
        context.res = {
            // status: 200, /* Defaults to 200 */          
            body: returnResp
          };
    } catch (error) {
        context.res = {
            status: 400,
            body: error.message
          };
    }  
  } else {
    context.log("Please pass a userId and GroupId on the query string or in the request body")
    context.res = {
      status: 400,
      body: "Please pass a userId abd GroupId on the query string or in the request body"
    };
  }

  // Run Main Function
  async function run():Promise<IReturnResp> {
      try {
        // Get Access Token
        const accessToken:string = await getAccessToken();
       
        if (accessToken){
          const groupUrl:string = await getGroupUrl();
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
                    "invitedUserDisplayName": userName,
                    "invitedUserEmailAddress": userId,
                    "inviteRedirectUrl": groupUrl,
                    "sendInvitationMessage": false             
                })
            }  
           // POST request      
           const invitationResponse = await request(options);
             // If Invite Created 
             if (invitationResponse.statusCode == 201){
            // Add addUser to O365 Group 
             const invitation:MicrosoftGraph.Invitation = JSON.parse(invitationResponse.body);
             const invitedUserId: string  = invitation.invitedUser.id;

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
                // POST request
                const response = await request(options);  
                return { groupId : groupId, invitation: invitation };      
             }
        }
      } catch (error) {
          context.log(error);
          throw new Error(error);
      }
  }

  async function getGroupUrl(): Promise<string> {
      try {
        const accessToken = await getAccessToken();
        let options = {
          method: 'GET',
          resolveWithFullResponse: true,
          url: `https://graph.microsoft.com/v1.0/groups/${groupId}/sites/root/weburl`,
          headers: {
              'Authorization': 'Bearer ' + accessToken,
              'content-type': 'application/json'
          }          
      };
      // POST request
      const response = await request(options);  

      return  JSON.parse(response.body).value;   
      } catch (error) {
        context.log(error);
        throw new Error(error);
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
