const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const {
  handleMessageExtensionQueryWithSSO,
  OnBehalfOfUserCredential,
} = require("@microsoft/teamsfx");
require("isomorphic-fetch");
const config = require("./config");

const oboAuthConfig = {
  authorityHost: config.authorityHost,
  clientId: config.clientId,
  tenantId: config.tenantId,
  clientSecret: config.clientSecret,
};

const initialLoginEndpoint = `https://${config.botDomain}/auth-start.html`;

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  async handleTeamsMessagingExtensionQuery(context, query) {
    try {
        // Use the SSO authentication method to handle the query.
        return await handleMessageExtensionQueryWithSSO(
            context,
            oboAuthConfig, // The configuration for on-behalf-of authentication.
            initialLoginEndpoint, // The endpoint for initial login.
            ["User.Read.All", "User.Read"], // The scopes required for the Graph API calls.
            async (token) => { // The callback to run after authentication.
                // Create a credential object with the SSO token and the authentication configuration.
                const credential = new OnBehalfOfUserCredential(
                    token.ssoToken,
                    oboAuthConfig
                );
                // Initialize an empty array to hold the attachments.
                const attachments = [];
                // Create an authentication provider with the credential and the required scopes.
                const authProvider = new TokenCredentialAuthenticationProvider(
                    credential,
                    {
                        scopes: ["User.Read"],
                    }
                );
                // Initialize a Graph client with the authentication provider.
                const graphClient = Client.initWithMiddleware({
                    authProvider: authProvider,
                });

                const { parameters } = query;

                const skills = getParameterByName(parameters, "Skill");
                const country = getParameterByName(parameters, "Location");
                const availabilityParam = getParameterByName(parameters, "Availability");

                var availability;

                if (availabilityParam == "true") {
                  availability = true;
                }
                else if (availabilityParam == "false") {
                  availability = false;
                }
                else {
                  availability = undefined;
                }

                function constructSearchObject(skills, country, availability) {
                const filterObject = {};

                if (country) {
                  filterObject.country = country;
                }

                if (skills) {
                  filterObject.skills = skills;
                }

                if (availability != undefined) {
                  filterObject.availability = availability;
                }

                return filterObject;
              }

              const searchObject = constructSearchObject(skills, country, availability);

              var candidateData = [];

              async function fetchCandidatesFromGraph(graphClient, filters) {
                try {
                  const filteredUsers = [];
                  const users = await graphClient
                    .api(`/users`)
                    .get();
          
                  for (const user of users.value){
                    const id = user.id;
                    const userProfileResponse = await graphClient
                      .api(`/users/${id}/?$select=id,displayName,skills,officeLocation`)
                      .get();
          
                    if (userProfileResponse.skills.some(skill => filters.skills.toLowerCase().includes(skill.toLowerCase()))) {
                      const userPhoto = await graphClient
                        .api(`/users/${id}/photo/$value`)
                        .responseType('arraybuffer')
                        .get();
                      if (userPhoto) {
                        const userPhotoBase64 = Buffer.from(userPhoto).toString('base64');
                        userProfileResponse.photo = `data:image/jpeg;base64,${userPhotoBase64}`;
                      }
                      else {
                        userProfileResponse.photo = "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg";
                      }
                      filteredUsers.push(userProfileResponse);
                    }
                  }
          
                  return filteredUsers;
          
                } catch (error) {
                  console.error("Error fetching candidates from Microsoft Graph:", error);
                  return [];
                }
              }
              
              // Fetch candidates based on applied filters.
              const candidatesFromGraph = await fetchCandidatesFromGraph(graphClient, searchObject);
              
              candidateData = candidatesFromGraph;
              console.log("Candidates:");
              candidateData.map((result) => console.log(result.displayName));

              // Create Adaptive Card object
              candidateData.map((result) => {
                // var availability = result.availability._ ? "Yes" : "No"
                const resultCard = CardFactory.adaptiveCard({
                  "type": "AdaptiveCard",
                  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                  "version": "1.4",
                  "body": [
                    {
                      "type": "TextBlock",
                      "text": "Expert Finder",
                      "wrap": true,
                      "size": "Large",
                      "weight": "Bolder",
                      "separator": true
                    },
                    {
                      "type": "ColumnSet",
                      "columns": [
                        {
                          "type": "Column",
                          "items": [
                            {
                              "type": "Image",
                              "url": result.photo,
                              "altText": "profileImage",
                              "size": "Small",
                              "style": "Person"
                            }
                          ],
                          "width": "auto"
                        },
                        {
                          "type": "Column",
                          "items": [
                            {
                              "type": "TextBlock",
                              "weight": "Bolder",
                              "text": `${result.displayName}`,
                              "wrap": true,
                              "spacing": "None",
                              "horizontalAlignment": "Left",
                              "maxLines": 0,
                              "size": "Medium"
                            }
                          ],
                          "width": "stretch",
                          "spacing": "Medium",
                          "verticalContentAlignment": "Center"
                        }
                      ]
                    },
                    {
                      "type": "FactSet",
                      "facts": [
                        {
                          "title": "Skills:",
                          "value": `${result.skills}`
                        },
                        {
                          "title": "Location:",
                          "value": `${result.officeLocation}`
                        },
                        {
                          "title": "Available:",
                          "value": "Yes",
                        }
                      ]
                    }
                  ],
                  
                });

                const previewCard = CardFactory.heroCard(
                  result.displayName,
                  result.skills.join(', ')
                );

                attachments.push({ ...resultCard, preview: previewCard });
              });
                // Return the attachments as a result to the Teams client.
                return {
                    composeExtension: {
                        type: "result",
                        attachmentLayout: "list",
                        attachments: attachments
                    },
                };
            }
        );
      } catch (error) {
          console.error("Error handling Teams messaging extension query:", error);
          // Return an error message to the Teams client.
          return {
              composeExtension: {
                  type: "message",
                  text: "An error occurred while processing your request. Please try again later."
              },
          };
      }
  }
}

const getParameterByName = (parameters, name) => {
  const param = parameters.find((p) => p.name === name);
  return param ? param.value : "";
};

module.exports.TeamsBot = TeamsBot;