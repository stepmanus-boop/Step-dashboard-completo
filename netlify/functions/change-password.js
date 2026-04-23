const { jsonResponse, requireSession, hashPassword, userPasswordMatches } = require('./_auth');
const { getUserById, updateUserPassword, isSupabaseConfigured } = require('./_supabase');

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return jsonResponse(200, { ok: true });
  }
  if (event.httpMethod !== 'POST') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }
  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  const auth = requireSession(event);
  if (!auth.ok) return auth.response;

  try {
    const session = auth.session;
    const body = JSON.parse(event.body || '{}');
    const currentPassword = String(body.currentPassword || '').trim();
    const newPassword = String(body.newPassword || '').trim();

    if (!currentPassword || !newPassword) {
      return jsonResponse(400, { ok: false, error: 'Informe a senha atual e a nova senha.' });
    }
    if (newPassword.length < 6) {
      return jsonResponse(400, { ok: false, error: 'A nova senha deve ter pelo menos 6 caracteres.' });
    }

    const user = await getUserById(session.sub || session.id);
    if (!user) {
      return jsonResponse(404, { ok: false, error: 'Usuário não encontrado para alterar a senha.' });
    }

    const validCurrent = await userPasswordMatches(currentPassword, user.passwordHash || '');
    if (!validCurrent) {
      return jsonResponse(401, { ok: false, error: 'A senha atual está incorreta.' });
    }

    const passwordHash = await hashPassword(newPassword);
    await updateUserPassword(user.id, passwordHash);

    return jsonResponse(200, { ok: true });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao alterar a senha.' });
  }
};
