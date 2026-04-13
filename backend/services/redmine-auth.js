const axios = require('axios');
const config = require('../config');

async function verifyRedmineCredentials(login, password) {
  try {
    const auth = Buffer.from(`${login}:${password}`).toString('base64');

    const response = await axios.get(`${config.REDMINE_URL}/users/current.json?include=memberships`, {
      headers: {
        'Authorization': `Basic ${auth}`,
        'Content-Type': 'application/json'
      },
      timeout: 5000
    });

    if (response.status === 200 && response.data.user) {
      const user = response.data.user;

      // Extraire le rôle sur le projet "Soutien des systèmes DSNA" (id: 1)
      const mcoMembership = (user.memberships || []).find(
        m => m.project && m.project.id === 1
      );
      const role = mcoMembership?.roles?.[0]?.name || null;

      return {
        success: true,
        user: {
          id: user.id,
          login: user.login,
          email: user.mail,
          firstname: user.firstname,
          lastname: user.lastname,
          role
        }
      };
    }

    return { success: false, error: 'Utilisateur non trouvé' };
  } catch (error) {
    console.error('Erreur auth Redmine:', error.message);

    if (error.response?.status === 401) {
      return { success: false, error: 'Identifiants incorrects' };
    }

    return { success: false, error: 'Erreur de connexion à Redmine' };
  }
}

module.exports = { verifyRedmineCredentials };
