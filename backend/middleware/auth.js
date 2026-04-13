const jwt = require('jsonwebtoken');
const config = require('../config');

function authenticateToken(req, res, next) {
  const token = req.headers.authorization?.replace('Bearer ', '');

  if (!token) {
    return res.status(401).json({ detail: 'Token manquant' });
  }

  try {
    const decoded = jwt.verify(token, config.JWT_SECRET);
    req.user = decoded;
    next();
  } catch (error) {
    return res.status(401).json({ detail: 'Token invalide' });
  }
}

function rejectClient(req, res, next) {
  if (req.user.role === 'Client') {
    return res.status(403).json({ detail: 'Accès non autorisé pour votre profil' });
  }
  next();
}

module.exports = { authenticateToken, rejectClient };
