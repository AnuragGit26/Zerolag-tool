
const LEVELS = {
    trace: { rank: 10, label: 'TRACE', color: '#8899aa' },
    debug: { rank: 20, label: 'DEBUG', color: '#6c8ae4' },
    info: { rank: 30, label: 'INFO', color: '#2e8b57' },
    warn: { rank: 40, label: 'WARN', color: '#c88719' },
    error: { rank: 50, label: 'ERROR', color: '#d93025' },
    fatal: { rank: 60, label: 'FATAL', color: '#ffffff', bg: '#d93025' }
};

const DEFAULT_LEVEL = 'debug';
const STORAGE_KEY = 'logger_level_v1';

function loadLevel() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved && LEVELS[saved]) return saved;
    } catch {/* ignored */ }
    return DEFAULT_LEVEL;
}

let currentLevel = loadLevel();

// Capture originals immediately (before any wrapping)
const ORIGINAL = {
    log: console.log.bind(console),
    info: console.info.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console),
    debug: (console.debug || console.log).bind(console),
    trace: (console.trace || console.log).bind(console)
};

function setLevel(level) {
    if (!LEVELS[level]) return;
    currentLevel = level;
    try { localStorage.setItem(STORAGE_KEY, level); } catch { /* noop */ }
}

function levelEnabled(level) {
    return LEVELS[level].rank >= LEVELS[currentLevel].rank;
}

function formatTs(date = new Date()) {
    const h = String(date.getHours()).padStart(2, '0');
    const m = String(date.getMinutes()).padStart(2, '0');
    const s = String(date.getSeconds()).padStart(2, '0');
    const ms = String(date.getMilliseconds()).padStart(3, '0');
    return `${h}:${m}:${s}.${ms}`;
}

function styleFor(levelMeta) {
    const base = 'padding:2px 6px;border-radius:4px;font-weight:600;font-family:system-ui,Roboto,Helvetica,Arial,sans-serif;';
    const fg = `color:${levelMeta.color};`;
    const bg = levelMeta.bg ? `background:${levelMeta.bg};color:${levelMeta.color};` : 'background:#1e1f21;color:#fff;';
    return levelMeta.bg ? base + bg : base + `background:${hexWithAlpha(levelMeta.color, 0.15)};` + fg;
}

function hexWithAlpha(hex, alpha) {
    if (!/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(hex)) return hex;
    const full = hex.length === 4 ? '#' + [...hex.slice(1)].map(c => c + c).join('') : hex;
    const r = parseInt(full.slice(1, 3), 16);
    const g = parseInt(full.slice(3, 5), 16);
    const b = parseInt(full.slice(5, 7), 16);
    return `rgba(${r},${g},${b},${alpha})`;
}

function flattenArgs(args) {
    if (!args.length) return { msg: '', rest: [] };
    const rest = [...args];
    return { msg: rest.shift(), rest };
}

const LEVEL_TO_ORIG = { trace: 'debug', debug: 'debug', info: 'info', warn: 'warn', error: 'error', fatal: 'error' };
const MSG_COLOR_KEY = 'logger_msg_color_v1';

function resolveMessageColor() {
    try {
        const override = localStorage.getItem(MSG_COLOR_KEY);
        if (override) return override;
    } catch { /* ignore */ }
    let dark = false;
    try { dark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches; } catch { /* ignore */ }
    return dark ? '#e5e7eb' : '#222222';
}

let messageColor = resolveMessageColor();
function refreshMessageColor() { messageColor = resolveMessageColor(); }
const CONFIG_KEY = 'logger_config_v1';
const DEFAULT_CONFIG = {
    remoteEnabled: false,
    remoteUrl: '',
    bufferSize: 50,
    flushIntervalMs: 5000,
    maxRetries: 5,
    backoffBase: 1000,
    sampling: { trace: 0.05, debug: 1, info: 1, warn: 1, error: 1, fatal: 1 },
    dedupIntervalMs: 2000,
    piiMaskKeys: ['session', 'sid', 'token', 'password', 'pass', 'auth', 'authorization'],
    maxEventBytes: 8000,
    captureGlobalErrors: true
};

let config = (function loadConfig() {
    try { const raw = localStorage.getItem(CONFIG_KEY); if (raw) { const parsed = JSON.parse(raw); return { ...DEFAULT_CONFIG, ...parsed, sampling: { ...DEFAULT_CONFIG.sampling, ...(parsed.sampling || {}) } }; } } catch { /* ignore */ }
    return { ...DEFAULT_CONFIG };
})();

function persistConfig() { try { localStorage.setItem(CONFIG_KEY, JSON.stringify(config)); } catch { /* ignore */ } }

let correlationId = (function () { try { return localStorage.getItem('logger_correlation_id') || crypto.randomUUID(); } catch { return 'cid-' + Math.random().toString(36).slice(2); } })();
try { localStorage.setItem('logger_correlation_id', correlationId); } catch { /* ignore */ }

let globalContext = { app: 'extension', version: (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.getManifest && chrome.runtime.getManifest().version) || 'unknown' };
let seq = 0;
let remoteBuffer = [];
let flushTimer = null;
let retryCount = 0;
const dedupCache = new Map();

function nowIso() { return new Date().toISOString(); }

function sampleAllowed(level) { const rate = (config.sampling && config.sampling[level]) ?? 1; if (rate >= 1) return true; return Math.random() < rate; }

function maskValue(key, value) { if (value == null) return value; if (typeof value === 'string') { if (config.piiMaskKeys.some(k => key.toLowerCase().includes(k))) return '***'; return value; } if (Array.isArray(value)) return value.map(v => maskUnknown(v)); if (typeof value === 'object') return maskObject(value); return value; }
function maskObject(obj) { const out = Array.isArray(obj) ? [] : {}; try { for (const k in obj) { if (Object.prototype.hasOwnProperty.call(obj, k)) { out[k] = maskValue(k, obj[k]); } } } catch { return obj; } return out; }
function maskUnknown(v) { if (v && typeof v === 'object') return maskObject(v); return v; }

function serializeError(err) { if (!err) return err; if (err instanceof Error) { return { name: err.name, message: err.message, stack: err.stack }; } return err; }

function dedupKey(level, msg) { return level + ':' + msg; }

function processRemote(event) {
    if (!config.remoteEnabled || !config.remoteUrl) return;
    if (!sampleAllowed(event.levelLower)) return;
    const key = dedupKey(event.levelLower, event.msg);
    const existing = dedupCache.get(key);
    const now = Date.now();
    if (existing && (now - existing.ts) < config.dedupIntervalMs) {
        existing.count += 1;
        existing.event.repeat = existing.count;
        return; // don't push again
    } else {
        dedupCache.set(key, { ts: now, count: 1, event });
    }
    remoteBuffer.push(event);
    if (remoteBuffer.length >= config.bufferSize) flushRemote();
    scheduleFlush();
}

function scheduleFlush() { if (flushTimer || !config.remoteEnabled) return; flushTimer = setTimeout(() => { flushTimer = null; flushRemote(); }, config.flushIntervalMs); }

async function flushRemote(force = false) {
    if (!config.remoteEnabled || !config.remoteUrl) return;
    if (!remoteBuffer.length) return;
    const batch = remoteBuffer.splice(0, remoteBuffer.length);
    const payload = JSON.stringify(batch.slice(0, 500)); // cap batch size
    if (payload.length > config.maxEventBytes * batch.length) { /* size heuristics, optional */ }
    try {
        await fetch(config.remoteUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: payload, keepalive: true });
        retryCount = 0;
    } catch (e) {
        // simple backoff
        remoteBuffer.unshift(...batch);
        retryCount++;
        if (retryCount <= config.maxRetries) {
            const delay = Math.min(30000, config.backoffBase * Math.pow(2, retryCount - 1));
            setTimeout(() => flushRemote(true), delay);
        } else {

            remoteBuffer = remoteBuffer.slice(-config.bufferSize * 2);
        }
    }
}

if (config.captureGlobalErrors) {
    try {
        window.addEventListener('error', e => {
            logger.error('GlobalError', serializeError(e.error) || { message: e.message, filename: e.filename, lineno: e.lineno });
        });
        window.addEventListener('unhandledrejection', e => {
            logger.error('UnhandledRejection', serializeError(e.reason));
        });
    } catch { /* ignore */ }
}

window.addEventListener('beforeunload', () => { try { flushRemote(true); } catch { /* ignore */ } });

function buildEvent(meta, msg, rest, ctx) {
    let data = rest.length === 1 ? rest[0] : (rest.length ? rest : undefined);
    if (data && typeof data === 'object') data = maskUnknown(data);
    return {
        ts: nowIso(),
        tsm: Math.round(performance.now()),
        seq: ++seq,
        level: meta.label,
        levelLower: meta.label.toLowerCase(),
        msg: String(msg),
        data: data === undefined ? undefined : data,
        cid: correlationId,
        ctx: ctx ? { ...globalContext, ...ctx } : globalContext
    };
}

function print(level, args, extraCtx) {
    if (!LEVELS[level]) level = 'info';
    if (!levelEnabled(level)) return;
    const meta = LEVELS[level];
    const { msg, rest } = flattenArgs(args);
    const ts = formatTs();
    const header = `%c${meta.label}%c ${ts} %c${msg}`;
    const styles = [styleFor(meta), '', `color:${messageColor};font-weight:500;`];
    try {
        const method = LEVEL_TO_ORIG[level] || 'log';
        const fn = ORIGINAL[method] || ORIGINAL.log;
        fn(header, ...styles, ...rest);
    } catch (e) {
        ORIGINAL.log(meta.label, ts, msg, ...rest);
    }

    try { const event = buildEvent(meta, msg, rest, extraCtx); processRemote(event); } catch { /* ignore */ }
}

function time(label) {
    const start = performance.now();
    print('debug', [`⏱️ start ${label}`]);
    return function end(extra) {
        const dur = performance.now() - start;
        const msg = `⏱️ end   ${label} +${dur.toFixed(1)}ms`;
        if (extra) print('debug', [msg, extra]); else print('debug', [msg]);
        return dur;
    };
}

let beautified = false;
function installConsoleBeautifier() {
    if (beautified) return;
    beautified = true;
    try { Object.defineProperty(console, '__orig', { value: ORIGINAL, enumerable: false }); } catch { /* ignore */ }
    ORIGINAL.log('%cLogger active', 'background:#2e8b57;color:#fff;padding:2px 6px;border-radius:4px;');
    console.info = (...a) => print('info', a);
    console.warn = (...a) => print('warn', a);
    console.error = (...a) => print('error', a);
    console.debug = (...a) => print('debug', a);
    console.trace = (...a) => print('trace', a);
    console.log = (...a) => print('info', a);
}

function configure(partial) {
    if (!partial || typeof partial !== 'object') return config;
    if (partial.sampling) { partial.sampling = { ...config.sampling, ...partial.sampling }; }
    config = { ...config, ...partial };
    persistConfig();
    if (!config.remoteEnabled) { try { if (flushTimer) { clearTimeout(flushTimer); flushTimer = null; } } catch { /* ignore */ } }
    return config;
}

function setCorrelationId(id) { if (!id) return correlationId; correlationId = id; try { localStorage.setItem('logger_correlation_id', id); } catch { /* ignore */ } return correlationId; }
function getCorrelationId() { return correlationId; }

function withContext(ctx) {
    return {
        trace: (...a) => print('trace', a, ctx),
        debug: (...a) => print('debug', a, ctx),
        info: (...a) => print('info', a, ctx),
        warn: (...a) => print('warn', a, ctx),
        error: (...a) => print('error', a, ctx),
        fatal: (...a) => print('fatal', a, ctx),
        time: (label) => time(label),
        ctx: () => ctx
    };
}

function flush() { return flushRemote(true); }

export const logger = {
    setLevel,
    getLevel: () => currentLevel,
    enabled: levelEnabled,
    trace: (...a) => print('trace', a),
    debug: (...a) => print('debug', a),
    info: (...a) => print('info', a),
    warn: (...a) => print('warn', a),
    error: (...a) => print('error', a),
    fatal: (...a) => print('fatal', a),
    time,
    installConsoleBeautifier,
    configure,
    getConfig: () => ({ ...config }),
    withContext,
    setCorrelationId,
    getCorrelationId,
    flush,
    refreshMessageColor
};

export default logger;
