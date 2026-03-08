/**
 * Simple bearer-token middleware for admin routes.
 * Set ADMIN_SECRET in your environment variables.
 * Pass the token as: Authorization: Bearer <ADMIN_SECRET>
 */
function requireAdmin(req, res, next) {
  const authHeader = req.headers['authorization'] || '';
  const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;

  const adminSecret = process.env.ADMIN_SECRET;
  if (!adminSecret) {
    console.error('ADMIN_SECRET environment variable is not set');
    return res.status(500).json({ error: 'Server misconfiguration' });
  }

  if (!token || token !== adminSecret) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  next();
}

module.exports = requireAdmin;
