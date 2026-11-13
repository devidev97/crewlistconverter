import { defineConfig, loadEnv } from "vite";
import { authorize, unavailable } from "./server/basic-auth.mjs";

function basicAuth(env) {
  const markPrivate = (res) => {
    res.setHeader("Cache-Control", "private, no-store");
    res.setHeader("CDN-Cache-Control", "no-store");
    res.setHeader("Vercel-CDN-Cache-Control", "no-store");
    const vary = res.getHeader("Vary");
    if (!vary) res.setHeader("Vary", "Authorization");
    else if (!String(vary).toLowerCase().includes("authorization")) {
      res.setHeader("Vary", `${vary}, Authorization`);
    }
  };

  const withoutSharedCache = (headers) => {
    if (Array.isArray(headers)) {
      return headers.filter(([name]) => !/^(cache-control|cdn-cache-control|vercel-cdn-cache-control)$/i.test(name));
    }
    const copy = { ...headers };
    for (const name of Object.keys(copy)) {
      if (/^(cache-control|cdn-cache-control|vercel-cdn-cache-control)$/i.test(name)) delete copy[name];
    }
    return copy;
  };

  const gate = (req, res, next) => {
    let decision;
    try {
      const header = req.headers.authorization;
      decision = authorize(typeof header === "string" ? header : undefined, env);
    } catch {
      decision = unavailable();
    }

    if (!decision.ok) {
      res.statusCode = decision.status;
      for (const [name, value] of Object.entries(decision.headers)) res.setHeader(name, value);
      res.end(req.method === "HEAD" ? undefined : decision.body);
      return;
    }

    const writeHead = res.writeHead;
    res.writeHead = function (statusCode, statusMessage, headers) {
      markPrivate(res);
      if (arguments.length === 1) return writeHead.call(this, statusCode);
      if (typeof statusMessage === "object" && statusMessage != null) {
        return writeHead.call(this, statusCode, withoutSharedCache(statusMessage));
      }
      if (arguments.length >= 3) {
        return writeHead.call(this, statusCode, statusMessage, withoutSharedCache(headers ?? {}));
      }
      return writeHead.call(this, statusCode, statusMessage);
    };

    const end = res.end;
    res.end = function (...args) {
      if (!res.headersSent) markPrivate(res);
      return end.apply(this, args);
    };
    next();
  };

  return {
    name: "basic-auth",
    configureServer(server) {
      server.middlewares.use(gate);
    },
    configurePreviewServer(server) {
      server.middlewares.use(gate);
    },
  };
}

export default defineConfig(({ mode }) => {
  const fileEnv = loadEnv(mode, process.cwd(), "");
  const authEnv = {
    BASIC_AUTH_USER: process.env.BASIC_AUTH_USER ?? fileEnv.BASIC_AUTH_USER,
    BASIC_AUTH_PASSWORD: process.env.BASIC_AUTH_PASSWORD ?? fileEnv.BASIC_AUTH_PASSWORD,
  };

  return {
    plugins: [basicAuth(authEnv)],
    server: {
      host: "0.0.0.0",
      port: 5173,
      strictPort: true,
      allowedHosts: [".code-server.devidev.es"],
    },
    preview: {
      host: "0.0.0.0",
      port: 4173,
      strictPort: true,
      allowedHosts: [".code-server.devidev.es"],
    },
  };
});
