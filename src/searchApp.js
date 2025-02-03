const { ApplicationBuilder } = require('@microsoft/teams-ai');
const { CardFactory, MemoryStorage } = require('botbuilder');
const { GraphClient } = require('./graphClient');
const config = require("./config");


const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        clientSecret: config.clientSecret
    }
};

const storage = new MemoryStorage();

const app = new ApplicationBuilder()
    .withStorage(storage)
    .withAuthentication(null, {
        settings: {
            graph: {
                scopes: ['User.Read', 'User.Read.All'],
                signInLink: `https://${config.botDomain}/auth-start.html`,
                endOnInvalidMessage: true,
                msalConfig: msalConfig
            }
        },
        autoSignIn: (context) => {
            return Promise.resolve(context.activity?.value.commandId !== 'signOutCommand');
        }
    })
    .build();

app.messageExtensions.query('Search', async (_context, state, query) => {
      const searchQuery = query.parameters.queryText || '';
      const token = state.temp.authTokens['graph'];
      if (!token) {
          throw new Error('No auth token found in state. Authentication failed.');
      }

      const parameters  = query.parameters;

      const skills = parameters.Skill || "";
      const country = parameters.Location || "";
      const availabilityParam = parameters.Availability || "";

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
    const users = await getUserDetailsFromGraph(token, searchObject);

    // const users = await getUserDetailsFromGraph(token, searchQuery);
    const results = [];

    users.forEach(user => {
        const previewCard = CardFactory.thumbnailCard(
            user.displayName,
            user.skills.join(', ')
        );

        const resultCard = CardFactory.adaptiveCard({
            type: "AdaptiveCard",
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            version: "1.4",
            body: [
                {
                    type: "TextBlock",
                    text: "Expert Finder",
                    wrap: true,
                    size: "Large",
                    weight: "Bolder",
                    separator: true
                },
                {
                    type: "ColumnSet",
                    columns: [
                        {
                            type: "Column",
                            items: [
                                {
                                    type: "Image",
                                    url: user.profilePhoto,
                                    altText: "profileImage",
                                    size: "Small",
                                    style: "Person"
                                }
                            ],
                            width: "auto"
                        },
                        {
                            type: "Column",
                            items: [
                                {
                                    type: "TextBlock",
                                    weight: "Bolder",
                                    text: user.displayName,
                                    wrap: true,
                                    spacing: "None",
                                    horizontalAlignment: "Left",
                                    size: "Medium"
                                }
                            ],
                            width: "stretch",
                            spacing: "Medium",
                            verticalContentAlignment: "Center"
                        }
                    ]
                },
                {
                    type: "FactSet",
                    facts: [
                        { title: "Skills:", value: user.skills.join(', ') },
                        { title: "Location:", value: user.officeLocation },
                        { title: "Available:", value: user.availability }
                    ]
                }
            ]
        });

        const attachment = {
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: resultCard.content,
            preview: previewCard,
        };

        results.push(attachment);
    });

    return {
        attachmentLayout: 'list',
        type: 'result',
        attachments: results,
        composeExtension: {
            type: 'invokeAction',
        }
    };
});

async function getUserDetailsFromGraph(token, searchQuery) {
    const graphClient = new GraphClient(token);
    console.log('Getting candidates from Graph API...');
    // const users = await graphClient.fetchCandidatesFromGraph({ skills: searchQuery });
    const users = await graphClient.fetchCandidatesFromGraph(searchQuery);

    return users.map(user => ({
        displayName: user.displayName,
        skills: user.skills,
        profilePhoto: user.profilePhoto,
        officeLocation: user.officeLocation,
        availability: user.availability
    }));
}

module.exports = { app };
