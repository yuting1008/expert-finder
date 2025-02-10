const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

class GraphClient {
    constructor(token) {
        if (!token || !token.trim()) {
            throw new Error('GraphClient: Invalid token received.');
        }

        this._token = token;

        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this._token);
            }
        });
    }

    async fetchWithRetry(url, retries = 3, delay = 500) {
        for (let i = 0; i < retries; i++) {
            try {
                return await this.graphClient.api(url).get();
            } catch (error) {
                console.error(`Fetch failed for ${url}. Retrying in ${delay}ms...`);
                await new Promise(res => setTimeout(res, delay));
            }
        }
        throw new Error(`Failed to fetch ${url} after ${retries} retries.`);
    }

    async fetchCandidatesFromGraph(filters) {
        try {
            const filteredUsers = [];
            // const skillsArray = Array.isArray(filters.skills) ? filters.skills : filters.skills.split(',');
            const skillsArray = Array.isArray(filters.skills) ? filters.skills : (filters.skills ? filters.skills.split(',') : []);
            const skillsSet = new Set(skillsArray.map(skill => skill.toLowerCase()));

            console.log("filters:", filters);
            let nextLink = '/users';

            while (nextLink) {
                const usersResponse = await this.graphClient.api(nextLink).get();
                const users = usersResponse.value;

                const userPromises = users.map(async (user) => {
                    const { id, displayName } = user;
                    try {
                        const userProfileResponse = await this.fetchWithRetry(`/users/${id}/?$select=id,displayName,skills,officeLocation`);
                        const userSkills = userProfileResponse.skills || [];
                        const hasMatchingSkill = userSkills.some(skill => {
                            return Array.from(skillsSet).some(searchSkill => skill.toLowerCase().includes(searchSkill));
                        });

                        if (!hasMatchingSkill) return null;

                        try {
                            const graphPhotoEndpoint = `https://graph.microsoft.com/v1.0/users/${id}/photo/$value`;
                            const graphRequestParams = {
                                method: 'GET',
                                headers: {
                                    'Content-Type': 'image/png',
                                    authorization: 'bearer ' + this._token
                                }
                            };

                            const response = await fetch(graphPhotoEndpoint, graphRequestParams);
                            const imageBuffer = await response.arrayBuffer();
                            const imageUri = 'data:image/jpeg;base64,' + Buffer.from(imageBuffer).toString('base64');
                            userProfileResponse.profilePhoto = imageUri;
                        } catch (photoError) {
                            if (photoError.statusCode === 404) {
                                userProfileResponse.profilePhoto = 'https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg';
                            } else {
                                console.error(`Error fetching photo for user ${displayName}:`, photoError);
                            }
                        }

                        userProfileResponse.availability = 'Yes';
                        return userProfileResponse;
                    } catch (profileError) {
                        console.error(`Error fetching profile for user ${displayName}:`, profileError);
                        return null;
                    }
                });

                const resolvedUsers = await Promise.all(userPromises);
                filteredUsers.push(...resolvedUsers.filter(user => user !== null));
                nextLink = usersResponse['@odata.nextLink'];
            }

            return filteredUsers;
        } catch (error) {
            console.error('Error fetching candidates from Microsoft Graph:', error);
            return [];
        }
    }
}

module.exports = { GraphClient };