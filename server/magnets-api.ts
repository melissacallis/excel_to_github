// PCentral magnet board API (deployed on Railway as the "magnets-api" function in project "pcentral-board").
// This file is a copy kept in the repo for reference; Railway runs it with Bun.
//
// Stores which "magnets" are placed on each patient's card, shared by every device.
// Data lives in one JSON file on a Railway volume mounted at /data.
//
// GET  /magnets          -> { magnets: { [patientKey]: Magnet[] }, updatedAt }
// PUT  /magnets/:key     -> body { magnets: Magnet[] }  (replaces that patient's magnets)
// GET  /health           -> "ok"
// Every /magnets call needs header  X-Board-Pin: <BOARD_PIN>

const DATA_DIR = Bun.env.DATA_DIR || "/data";
const FILE = `${DATA_DIR}/magnets.json`;
const PIN = Bun.env.BOARD_PIN || "";
const ALLOWED_ORIGINS = (Bun.env.ALLOWED_ORIGINS || "https://melissacallis.github.io")
  .split(",").map(s => s.trim()).filter(Boolean);

type Magnet = {
  type: string;
  placedAt: string;
  time?: string;
  date?: string;
  location?: string;
  note?: string;
};
type Store = { magnets: Record<string, Magnet[]>; updatedAt: string };

async function load(): Promise<Store> {
  const f = Bun.file(FILE);
  if (!(await f.exists())) return { magnets: {}, updatedAt: new Date(0).toISOString() };
  try {
    const data = await f.json();
    return { magnets: data.magnets || {}, updatedAt: data.updatedAt || new Date(0).toISOString() };
  } catch {
    return { magnets: {}, updatedAt: new Date(0).toISOString() };
  }
}

// Writes go one at a time so two quick taps can't overwrite each other.
let queue: Promise<unknown> = Promise.resolve();
function withLock<T>(fn: () => Promise<T>): Promise<T> {
  const next = queue.then(fn, fn);
  queue = next.catch(() => {});
  return next;
}

const TYPES = new Set(["offunit", "reassess", "pw", "full", "prn", "iop"]);
const str = (v: unknown, max = 120) => (typeof v === "string" ? v.slice(0, max) : undefined);

function clean(list: unknown): Magnet[] {
  if (!Array.isArray(list)) return [];
  const out: Magnet[] = [];
  for (const m of list.slice(0, 20)) {
    if (!m || typeof m !== "object") continue;
    const type = str((m as any).type, 20);
    if (!type || !TYPES.has(type)) continue;
    out.push({
      type,
      placedAt: str((m as any).placedAt, 40) || new Date().toISOString(),
      time: str((m as any).time, 20),
      date: str((m as any).date, 20),
      location: str((m as any).location, 120),
      note: str((m as any).note, 240),
    });
  }
  return out;
}

function cors(req: Request): Record<string, string> {
  const origin = req.headers.get("Origin") || "";
  const allow = ALLOWED_ORIGINS.includes(origin) || origin.startsWith("http://localhost") || origin === "null";
  return {
    "Access-Control-Allow-Origin": allow ? origin : ALLOWED_ORIGINS[0],
    "Access-Control-Allow-Methods": "GET, PUT, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-Board-Pin",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
}

function json(req: Request, body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store", ...cors(req) },
  });
}

Bun.serve({
  port: Number(Bun.env.PORT || 3000),
  async fetch(req) {
    const url = new URL(req.url);
    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: cors(req) });
    if (url.pathname === "/health") return new Response("ok", { headers: cors(req) });

    if (url.pathname === "/magnets" || url.pathname.startsWith("/magnets/")) {
      if (!PIN || req.headers.get("X-Board-Pin") !== PIN) {
        return json(req, { error: "Wrong or missing board PIN" }, 401);
      }

      if (req.method === "GET" && url.pathname === "/magnets") {
        return json(req, await load());
      }

      if (req.method === "PUT" && url.pathname.startsWith("/magnets/")) {
        const key = decodeURIComponent(url.pathname.slice("/magnets/".length)).slice(0, 80);
        if (!key) return json(req, { error: "Missing patient key" }, 400);
        let body: any;
        try { body = await req.json(); } catch { return json(req, { error: "Bad JSON" }, 400); }
        const list = clean(body?.magnets);
        const saved = await withLock(async () => {
          const store = await load();
          if (list.length) store.magnets[key] = list; else delete store.magnets[key];
          store.updatedAt = new Date().toISOString();
          await Bun.write(FILE, JSON.stringify(store));
          return store;
        });
        return json(req, saved);
      }
      return json(req, { error: "Not found" }, 404);
    }
    return new Response("PCentral magnet board API", { headers: cors(req) });
  },
});

console.log(`Magnet API listening; data file ${FILE}`);
