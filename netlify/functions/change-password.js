const { jsonResponse, requireSession, hashPassword } = require('./_auth');
const { getUserById, updateUser, userPasswordMatches, isSupabaseConfigured } = require('./_supabase');

exports.handler = async (event) => {
  const sessionResult = requireSession(event);
  if (!sessionResult.ok) return sessionResult.response;

  if (event.httpMethod !== 'POST') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  try {
    const body = JSON.parse(event.body || '{}');
    const currentPassword = String(body.currentPassword || '').trim();
    const newPassword = String(body.newPassword || '').trim();

    if (!currentPassword || !newPassword) {
      return jsonResponse(400, { ok: false, error: 'Informe a senha atual e a nova senha.' });
    }
    if (newPassword.length < 6) {
      return jsonResponse(400, { ok: false, error: 'A nova senha deve ter pelo menos 6 caracteres.' });
    }
    if (currentPassword === newPassword) {
      return jsonResponse(400, { ok: false, error: 'A nova senha deve ser diferente da atual.' });
    }

    const user = await getUserById(sessionResult.session.sub);
    if (!user || !user.active) {
      return jsonResponse(404, { ok: false, error: 'Usuário não encontrado.' });
    }
    if (!userPasswordMatches(currentPassword, user.passwordHash)) {
      return jsonResponse(401, { ok: false, error: 'A senha atual está incorreta.' });
    }

    await updateUser(user.id, { passwordHash: hashPassword(newPassword) });
    return jsonResponse(200, { ok: true });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao alterar a senha.' });
  }
};
