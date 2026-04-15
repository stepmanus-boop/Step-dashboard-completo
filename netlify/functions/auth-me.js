const { jsonResponse, getSession } = require("./_auth");
const { isGithubConfigured } = require("./_githubStore");

exports.handler = async (event) => {
  const session = getSession(event);
  if (!session) {
    return jsonResponse(200, { ok: true, authenticated: false, publicAccess: true, githubSyncEnabled: isGithubConfigured() });
  }

  return jsonResponse(200, {
    ok: true,
    authenticated: true,
    githubSyncEnabled: isGithubConfigured(),
    user: {
      id: session.sub,
      name: session.name,
      username: session.username,
      role: session.role,
      sector: session.sector,
    },
  });
};
