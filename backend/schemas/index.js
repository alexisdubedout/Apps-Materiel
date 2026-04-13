const { z } = require('zod');

const loginSchema = z.object({
  login: z.string().min(1, 'Login requis').max(100),
  password: z.string().min(1, 'Mot de passe requis').max(200),
});

const mappingSchema = z.object({
  code: z.string().min(1, 'Code requis').max(50),
  description: z.string().max(200).optional().default(''),
  type: z.string().max(100).optional().default(''),
  typeEmplacement: z.string().max(100).optional().default(''),
  detail: z.string().max(200).optional().default(''),
  user: z.string().max(100).optional(),
});

const noteSchema = z.object({
  content: z.string().max(10000, 'Note trop longue'),
});

module.exports = { loginSchema, mappingSchema, noteSchema };
