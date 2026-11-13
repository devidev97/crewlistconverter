import assert from "node:assert";
import { authorize } from "../server/basic-auth.mjs";

const env = { BASIC_AUTH_USER: "tripulacion", BASIC_AUTH_PASSWORD: "secreto" };
const header = `Basic ${Buffer.from("tripulacion:secreto", "utf8").toString("base64")}`;

assert.strictEqual(authorize(header, env).ok, true);
assert.strictEqual(authorize("Basic " + Buffer.from("tripulacion:otra", "utf8").toString("base64"), env).status, 401);
assert.strictEqual(authorize(undefined, env).status, 401);
assert.strictEqual(authorize(header, {}).status, 503);
assert.strictEqual(authorize(header).status, 503);

console.log("OK: autenticación básica, mismas variables que ADR.");
