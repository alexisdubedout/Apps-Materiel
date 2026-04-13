function validate(schema) {
  return (req, res, next) => {
    const result = schema.safeParse(req.body);
    if (!result.success) {
      const errors = result.error.errors.map(e => e.message).join(', ');
      return res.status(400).json({ detail: errors });
    }
    req.body = result.data;
    next();
  };
}

module.exports = { validate };
