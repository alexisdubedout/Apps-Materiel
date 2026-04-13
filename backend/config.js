require('dotenv').config();

module.exports = {
  JWT_SECRET: process.env.JWT_SECRET,
  JWT_EXPIRATION: process.env.JWT_EXPIRATION || '24h',
  REDMINE_URL: process.env.REDMINE_URL,
  REDMINE_API_KEY: process.env.REDMINE_API_KEY,
  SHAREPOINT_DEVIS_URL: process.env.SHAREPOINT_DEVIS_URL,
  SHAREPOINT_REVUE_HEBDO_URL: process.env.SHAREPOINT_REVUE_HEBDO_URL,
  PORT: process.env.PORT || 10000,
  CORS_ORIGINS: (process.env.CORS_ORIGINS || '').split(',').filter(Boolean),
};
