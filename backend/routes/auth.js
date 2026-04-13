const express = require('express');
const jwt = require('jsonwebtoken');
const router = express.Router();
const config = require('../config');
const { verifyRedmineCredentials } = require('../services/redmine-auth');
const { validate } = require('../middleware/validate');
const { loginSchema } = require('../schemas/index');
const { upsertUser } = require('../services/database');

router.post('/login', validate(loginSchema), async (req, res) => {
  try {
    const { login, password } = req.body;

    if (!login || !password) {
      return res.status(400).json({ detail: 'Login et mot de passe requis' });
    }

    console.log('🔐 Tentative de connexion:', login);

    const result = await verifyRedmineCredentials(login, password);

    if (!result.success) {
      return res.status(401).json({ detail: result.error });
    }

    await upsertUser(result.user);

    const token = jwt.sign(
      {
        redmineId: result.user.id,
        login: result.user.login,
        email: result.user.email,
        name: `${result.user.firstname} ${result.user.lastname}`,
        role: result.user.role
      },
      config.JWT_SECRET,
      { expiresIn: config.JWT_EXPIRATION }
    );

    console.log('✅ Connexion réussie:', login, '| Rôle:', result.user.role);

    res.json({
      token,
      user: {
        login: result.user.login,
        email: result.user.email,
        firstname: result.user.firstname,
        lastname: result.user.lastname,
        role: result.user.role
      }
    });

  } catch (error) {
    console.error('❌ Erreur login:', error);
    res.status(500).json({ detail: 'Erreur serveur' });
  }
});

router.post('/verify', (req, res) => {
  try {
    const token = req.headers.authorization?.replace('Bearer ', '');

    if (!token) {
      return res.status(401).json({ valid: false });
    }

    const decoded = jwt.verify(token, config.JWT_SECRET);
    res.json({ valid: true, user: decoded });

  } catch (error) {
    res.status(401).json({ valid: false });
  }
});

module.exports = router;
