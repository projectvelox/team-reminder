// Shared CORS + security headers (v1.7.40).
//
// Previously every endpoint set `Access-Control-Allow-Origin: '*'`. That worked
// because all writes require a Teams SSO Bearer token (no cookies, so CSRF is
// N/A), but it's still overly permissive defense-in-depth-wise — any web page
// could read our public GET responses by attaching a stolen token.
//
// This module reflects the request's Origin only when it's on the allowlist.
// Same-origin / extension calls (no Origin header) are passed through. Adds
// nosniff + a tight Referrer-Policy on every response.

const ALLOWED_ORIGINS = new Set([
  'https://projectvelox.github.io',
  // Localhost ports for sideload-from-disk previews. Keep narrow.
  'http://localhost:3000',
  'http://127.0.0.1:3000',
  'http://localhost:5500',
  'http://127.0.0.1:5500',
]);

// Build a CORS+security header bundle for a given Functions request.
// `methods` defaults to the common verbs; callers can pass a narrower list.
function corsHeaders(request, methods) {
  const origin = request && request.headers ? request.headers.get('origin') : null;
  const allowOrigin = origin && ALLOWED_ORIGINS.has(origin) ? origin : 'https://projectvelox.github.io';
  return {
    'Access-Control-Allow-Origin': allowOrigin,
    // Vary so caches/CDNs don't serve one origin's response to another.
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': methods || 'GET, POST, PATCH, DELETE, PUT, OPTIONS',
    'Access-Control-Max-Age': '600',
    'Content-Type': 'application/json',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'no-referrer',
    'Cache-Control': 'no-store',
  };
}

// JSON response helper that always carries the CORS+security bundle.
function json(request, status, body, methods) {
  return { status, headers: corsHeaders(request, methods), body: JSON.stringify(body) };
}

module.exports = { corsHeaders, json, ALLOWED_ORIGINS };
