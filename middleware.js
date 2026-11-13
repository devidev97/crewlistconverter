import { authorize, unavailable } from "./server/basic-auth.mjs";

// Sin matcher: Vercel ejecuta este middleware en todas las rutas
// (HTML, estáticos y cualquier ruta futura) y antes de la caché.
export const config = {
  runtime: "nodejs",
};

function continueRequest() {
  return new Response(null, {
    headers: {
      "x-middleware-next": "1",
      "Cache-Control": "private, no-store",
      "CDN-Cache-Control": "no-store",
      "Vercel-CDN-Cache-Control": "no-store",
      Vary: "Authorization",
    },
  });
}

export default function middleware(request) {
  let decision;
  try {
    decision = authorize(request.headers.get("authorization"));
  } catch {
    decision = unavailable();
  }

  if (!decision.ok) {
    return new Response(decision.body, {
      status: decision.status,
      headers: decision.headers,
    });
  }

  return continueRequest();
}
