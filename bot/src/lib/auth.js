// Teams SSO token validation.
//
// Tab calls microsoftTeams.authentication.getAuthToken() which returns a JWT
// with aud = api://botid-<botAppId> and tid = <kation tenant>. We verify the
// signature with Entra's JWKS endpoint and check aud + tid + token expiry.

const { createRemoteJWKSet, jwtVerify } = require('jose');

const TENANT = process.env.DAYREMINDERS_TENANT_ID;
const AUDIENCE = process.env.DAYREMINDERS_AUDIENCE;
const APP_ID = process.env.MicrosoftAppId;

let jwks;
function getJwks() {
  if (!jwks) {
    if (!TENANT) throw new Error('DAYREMINDERS_TENANT_ID is not set');
    jwks = createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${TENANT}/discovery/v2.0/keys`));
  }
  return jwks;
}

async function verifyTeamsToken(authHeader) {
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    const err = new Error('Missing Bearer token');
    err.status = 401;
    throw err;
  }
  const token = authHeader.slice('Bearer '.length).trim();
  if (!AUDIENCE) throw new Error('DAYREMINDERS_AUDIENCE is not set');

  const validAudiences = [AUDIENCE, APP_ID].filter(Boolean);

  try {
    const { payload } = await jwtVerify(token, getJwks(), {
      audience: validAudiences,
      issuer: [
        `https://login.microsoftonline.com/${TENANT}/v2.0`,
        `https://sts.windows.net/${TENANT}/`,
      ],
    });
    if (!payload.oid) {
      const err = new Error('Token has no oid claim');
      err.status = 401;
      throw err;
    }
    return {
      oid: payload.oid,
      tid: payload.tid,
      upn: payload.upn || payload.preferred_username || null,
      name: payload.name || null,
    };
  } catch (err) {
    // Surface the actual aud/iss for debugging when validation fails.
    if (err.code === 'ERR_JWT_CLAIM_VALIDATION_FAILED' || err.message?.includes('aud') || err.message?.includes('iss')) {
      try {
        const [, payloadB64] = token.split('.');
        const claims = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8'));
        const detail = `validation failed (got aud=${JSON.stringify(claims.aud)} iss=${claims.iss}, expected aud one of ${JSON.stringify(validAudiences)})`;
        const e = new Error(detail);
        e.status = 401;
        throw e;
      } catch (_) { /* fall through */ }
    }
    if (!err.status) err.status = 401;
    throw err;
  }
}

module.exports = { verifyTeamsToken };
