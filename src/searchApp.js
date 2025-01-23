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
                  const skillsArray = Array.isArray(filters.skills) ? filters.skills : filters.skills.split(',');
                  const skillsSet = new Set(skillsArray.map(skill => skill.toLowerCase()));
                  // const skillsSet = new Set(filters.skills.map(skill => skill.toLowerCase()));
                  console.log("Skills Set:", skillsSet);
                  console.log("filters:", filters);
                  let nextLink = '/users';
              
                  while (nextLink) {
                    const usersResponse = await graphClient.api(nextLink).get();
                    const users = usersResponse.value;
              
                    // Fetch profiles and photos concurrently
                    const userPromises = users.map(async (user) => {
                      const id = user.id;
              
                      try {
                        // Fetch user profile
                        const userProfileResponse = await graphClient
                          .api(`/users/${id}/?$select=id,displayName,skills,officeLocation`)
                          .get();
              
                        const userSkills = userProfileResponse.skills || [];
                        // const hasMatchingSkill = userSkills.some(skill => skillsSet.has(skill.toLowerCase()));
                        const hasMatchingSkill = userSkills.some(skill => {
                          return Array.from(skillsSet).some(searchSkill => skill.toLowerCase().includes(searchSkill));
                        });
              
                        if (!hasMatchingSkill) return null; // Skip if no matching skills
              
                        // Fetch user photo
                        try {
                          const userPhoto = await graphClient
                            .api(`/users/${id}/photo/$value`)
                            .responseType('arraybuffer')
                            .get();
              
                          const userPhotoBase64 = Buffer.from(userPhoto).toString('base64');
                          userProfileResponse.photo = `data:image/jpeg;base64,${userPhotoBase64}`;
                        } catch (photoError) {
                          if (photoError.statusCode === 404) {
                            userProfileResponse.photo =
                              'https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg';
                          } else {
                            console.error(`Error fetching photo for user ${id}:`, photoError);
                          }
                        }
              
                        return userProfileResponse;
                      } catch (profileError) {
                        console.error(`Error fetching profile for user ${id}:`, profileError);
                        return null;
                      }
                    });
              
                    // Wait for all promises to resolve
                    const resolvedUsers = await Promise.all(userPromises);
              
                    // Filter out null values
                    filteredUsers.push(...resolvedUsers.filter(user => user !== null));
              
                    // Check for pagination
                    nextLink = usersResponse['@odata.nextLink'];
                  }
              
                  return filteredUsers;
                } catch (error) {
                  console.error('Error fetching candidates from Microsoft Graph:', error);
                  return [];
                }
              }
              
              
              const candidatesFromGraph = await fetchCandidatesFromGraph(graphClient, searchObject);
              candidateData = candidatesFromGraph;
              console.log("Candidates:");
              candidateData.map((result) => console.log(result.displayName));

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
                              "url": `${result.photo}`,
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
                          "value": `${result.skills.join(', ')}`
                        },
                        {
                          "title": "Location:",
                          "value": `${result.officeLocation}`,
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