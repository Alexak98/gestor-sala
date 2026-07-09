/* =========================================================
 * GestorSala — Meeting room display
 * =========================================================
 * CONFIG: change these values when you connect n8n
 * ========================================================= */
const CONFIG = {
  ROOM_NAME: "Sala Reuniones",

  // n8n webhooks
  N8N_GET_EVENTS_URL: "https://n8n.yurest.dev/webhook/sala-eventos",
  N8N_CREATE_EVENT_URL: "https://n8n.yurest.dev/webhook/sala-reservar",
  N8N_END_EVENT_URL: "https://n8n.yurest.dev/webhook/sala-finalizar",
  N8N_CHECKIN_URL:   "https://n8n.yurest.dev/webhook/sala-checkin",
  N8N_DELETE_URL:    "https://n8n.yurest.dev/webhook/sala-eliminar",

  // Check-in window (minutes): si no se confirma dentro de este tiempo
  // desde el inicio del evento, se cancela automáticamente
  CHECKIN_WINDOW_MIN: 15,

  // Refresh intervals
  REFRESH_EVENTS_MS: 60_000, // re-fetch events every 60s
  OFFLINE_THRESHOLD_MS: 150_000, // show offline badge if no successful fetch in 2.5 min
  REFRESH_CLOCK_MS: 1_000,   // tick every second

  // Timeline (work hours shown on the mini agenda)
  DAY_START_HOUR: 8,
  DAY_END_HOUR: 18,

  // Feature thresholds
  ENDING_SOON_MIN: 2,        // amber countdown in last N minutes
  NEXT_WARN_MIN: 5,          // show "próxima en X min" banner when next is ≤N min away
  AGENDA_WEEK_DAYS: 7,       // "7 días" tab range

  // Locale
  LOCALE: "es-ES",
  TIMEZONE: "Europe/Madrid",

  // Use mock data (true = mocks, false = real n8n)
  USE_MOCK: false,

  // Lista de personas que pueden reservar — EDITA AQUÍ
  PEOPLE: [
    "Alex",
    "Alvaro Jareño",
    "Luis Bahamonde",
    "Miguel Vilata",
    "Pedro Martin",
    "Borja Pastor",
    "Santiago Andres",
    "Paula Schmidt",
    "Aline Perles",
    "Stefania Sulis",
    "Javier Molina",
    "Carlos Llopis",
    "Raquel Batalla",
    "Victor",
    "Pablo Claramunt",
    "Juan Daniel",
    "Luis Alejandro",
    "Edgar",
    "Mercedes",
    "Carlos Aparicio",
    "Hugo Zalazar",
    "Mario Labrandero",
    "Ivan Ramirez",
    "Rino Luigi",
    "Javier Feliu",
    "Rafael Gonzalez",
    "Maria Fernandez",
    "Marina Rubio",
  ],

  // Títulos sugeridos para reserva rápida (chips)
  QUICK_TITLES: ["Reunión", "Llamada", "Entrevista", "1:1", "Formación"],
};

/* =========================================================
 * Perf primitives
 * =========================================================
 * Hot-path caches: created once, reused on every tick.
 * - Intl formatters are otherwise rebuilt per call (~ms each).
 * - PEOPLE_LOWER turns the linear splitTitle() lookup into O(1).
 * - $() memoizes getElementById; the DOM here never re-mounts.
 * ========================================================= */
const FMT_TIME       = new Intl.DateTimeFormat(CONFIG.LOCALE, { hour: "2-digit", minute: "2-digit", hour12: false });
const FMT_DATE_FULL  = new Intl.DateTimeFormat(CONFIG.LOCALE, { weekday: "long", day: "numeric", month: "long" });
const FMT_DATE_SHORT = new Intl.DateTimeFormat(CONFIG.LOCALE, { day: "numeric", month: "short" });
const FMT_MONTH_YEAR = new Intl.DateTimeFormat(CONFIG.LOCALE, { month: "long", year: "numeric" });
const FMT_GROUP_HDR  = new Intl.DateTimeFormat(CONFIG.LOCALE, { weekday: "long", day: "numeric", month: "long" });

const PEOPLE_LOWER = new Set(CONFIG.PEOPLE.map(p => p.trim().toLowerCase()));

const _domCache = Object.create(null);
function $(id) {
  let el = _domCache[id];
  if (!el || !el.isConnected) el = _domCache[id] = document.getElementById(id);
  return el;
}

/* =========================================================
 * Mock data (used while USE_MOCK = true)
 * ========================================================= */
function buildMockEvents() {
  const now = new Date();
  const today = new Date(now);
  const tomorrow = new Date(now); tomorrow.setDate(tomorrow.getDate() + 1);

  const at = (d, h, m) => {
    const x = new Date(d); x.setHours(h, m, 0, 0); return x.toISOString();
  };

  return [
    {
      id: "mock-1",
      title: "Reunión de dirección",
      organizer: "Alexander Stokes",
      start: at(today, 13, 0),
      end: at(today, 14, 0),
    },
    {
      id: "mock-2",
      title: "Entrevista con Sydney Roy",
      organizer: "Henrietta Gardner",
      start: at(today, 14, 15),
      end: at(today, 15, 15),
    },
    {
      id: "mock-3",
      title: "Llamada comercial",
      organizer: "Martin Gutierrez",
      start: at(today, 15, 45),
      end: at(today, 16, 45),
    },
    {
      id: "mock-4",
      title: "Planificación de proyecto",
      organizer: "Susie Dunn",
      start: at(tomorrow, 11, 0),
      end: at(tomorrow, 12, 30),
    },
  ];
}

/* =========================================================
 * API — calls n8n, or returns mocks
 * ========================================================= */
async function fetchEvents() {
  if (CONFIG.USE_MOCK || !CONFIG.N8N_GET_EVENTS_URL) {
    return buildMockEvents();
  }
  const res = await fetch(CONFIG.N8N_GET_EVENTS_URL, { method: "GET" });
  if (!res.ok) throw new Error(`GET events ${res.status}`);
  const raw = await res.json();
  return normalizeEvents(raw);
}

async function createEvent({ title, startISO, endISO }) {
  if (CONFIG.USE_MOCK || !CONFIG.N8N_CREATE_EVENT_URL) {
    // Simulate
    await new Promise(r => setTimeout(r, 600));
    return { id: "mock-new", title, start: startISO, end: endISO, organizer: "Tú" };
  }
  const res = await fetch(CONFIG.N8N_CREATE_EVENT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ title, start: startISO, end: endISO }),
  });
  if (!res.ok) throw new Error(`POST create ${res.status}`);
  return await res.json();
}

async function deleteEvent(eventId) {
  if (CONFIG.USE_MOCK || !CONFIG.N8N_DELETE_URL) {
    await new Promise(r => setTimeout(r, 400));
    state.events = state.events.filter(e => e.id !== eventId);
    return { deleted: true, eventId };
  }
  const res = await fetch(CONFIG.N8N_DELETE_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ eventId }),
  });
  if (!res.ok) throw new Error(`POST delete ${res.status}`);
  return await res.json();
}

async function checkInEvent(eventId) {
  if (CONFIG.USE_MOCK || !CONFIG.N8N_CHECKIN_URL) {
    await new Promise(r => setTimeout(r, 400));
    const ev = state.events.find(e => e.id === eventId);
    if (ev) ev.description = `[CHECKED_IN:${new Date().toISOString()}]\n\n${ev.description || ""}`;
    return { id: eventId };
  }
  const res = await fetch(CONFIG.N8N_CHECKIN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ eventId }),
  });
  if (!res.ok) throw new Error(`POST checkin ${res.status}`);
  return await res.json();
}

async function endEvent(eventId) {
  if (CONFIG.USE_MOCK || !CONFIG.N8N_END_EVENT_URL) {
    await new Promise(r => setTimeout(r, 500));
    // Simulate: shorten mock event end to now — keep startMs/endMs in sync.
    const ev = state.events.find(e => e.id === eventId);
    if (ev) {
      const now = Date.now();
      ev.end = new Date(now).toISOString();
      ev.endMs = now;
    }
    return { id: eventId };
  }
  const res = await fetch(CONFIG.N8N_END_EVENT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ eventId }),
  });
  if (!res.ok) throw new Error(`POST end ${res.status}`);
  return await res.json();
}

/**
 * Normalize n8n response — supports either:
 *  - { items: [ {id, summary, organizer:{displayName}, start:{dateTime}, end:{dateTime}}, ... ] } (raw Google Calendar)
 *  - [ { id, title, organizer, start, end }, ... ] (already shaped)
 */
function normalizeEvents(raw) {
  const items = Array.isArray(raw) ? raw : (raw.events || raw.items || raw.data || []);
  const out = [];
  for (let i = 0; i < items.length; i++) {
    const ev = items[i];
    const start = ev.start?.dateTime || ev.start?.date || ev.start;
    const end   = ev.end?.dateTime   || ev.end?.date   || ev.end;
    if (!start || !end) continue;
    const startMs = Date.parse(start);
    const endMs   = Date.parse(end);
    if (Number.isNaN(startMs) || Number.isNaN(endMs)) continue;
    out.push({
      id: ev.id,
      title: ev.title || ev.summary || "(sin título)",
      organizer:
        (typeof ev.organizer === "string" && ev.organizer) ||
        ev.organizerName ||
        ev.organizer?.displayName ||
        ev.creator?.displayName ||
        "",
      start,
      end,
      startMs,
      endMs,
      description: ev.description || "",
    });
  }
  // Pre-sort once — getNextEvent/getCurrentEvent then run in O(n) without re-sorting per tick.
  out.sort((a, b) => a.startMs - b.startMs);
  return out;
}

function isCheckedIn(ev) {
  return /\[CHECKED_IN:/i.test(ev?.description || "");
}

/* =========================================================
 * Title + person helpers
 * ========================================================= */
/** Split "Título — Persona" into { cleanTitle, person } */
function splitTitle(full) {
  if (!full) return { cleanTitle: "", person: null };
  const idx = full.lastIndexOf(" — ");
  if (idx < 0) return { cleanTitle: full, person: null };
  const person = full.slice(idx + 3).trim();
  // O(1) via PEOPLE_LOWER Set — was O(n) per render (called twice per tick).
  if (!PEOPLE_LOWER.has(person.toLowerCase())) return { cleanTitle: full, person: null };
  return { cleanTitle: full.slice(0, idx).trim(), person };
}

function getInitials(name) {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map(w => w[0].toUpperCase())
    .join("");
}

function colorFromName(name) {
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = ((hash << 5) - hash + name.charCodeAt(i)) | 0;
  }
  const hue = Math.abs(hash) % 360;
  return `hsl(${hue}, 55%, 48%)`;
}

/* =========================================================
 * Booking constraints
 * ========================================================= */
const BOOKING_HOURS = { start: 8, end: 18 }; // hora de inicio permitida: 08:00–18:00

/* =========================================================
 * State
 * ========================================================= */
const state = {
  events: [],
  lastFetch: 0,
  lastFetchOk: 0,            // timestamp of last successful fetch
  fetchFailing: false,
  bookingDuration: 30,
  bookingPerson: null,
  autoCancelled: new Set(), // event IDs already auto-cancelled this session
  checkingIn: false,
  customDuration: false,     // true when user picked "Más…" (end time mode)
  quickTitleChoice: null,    // chip selected in title modal
  dayView: "today",          // "today" | "tomorrow" | "week"
};

/* =========================================================
 * Clock
 * ========================================================= */
let lastListMinute = -1;
let lastClockMinute = -1;
function tickClock() {
  // Cheapest possible early-out — when the kiosk screen is off, do nothing.
  // Saves ~3600 ticks/h of layout work + GC pressure.
  if (document.visibilityState === "hidden") return;

  const now = new Date();
  const mKey = now.getHours() * 60 + now.getMinutes();

  // Clock text only changes once a minute — don't reformat or touch the DOM
  // on the other 59 ticks per minute.
  if (mKey !== lastClockMinute) {
    lastClockMinute = mKey;
    $("clock-time").textContent = FMT_TIME.format(now);
    $("clock-date").textContent = capitalize(FMT_DATE_FULL.format(now));
  }

  // Status panel (countdown) still needs per-second freshness.
  renderStatus(now);

  // Event list — drop just-ended items once a minute.
  if (mKey !== lastListMinute) {
    lastListMinute = mKey;
    renderEventsList();
  }
}

function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

/* =========================================================
 * Status panel rendering
 * ========================================================= */
// Hot path — called every second. Uses pre-computed ms and the sorted invariant
// so we avoid `new Date()` × 2 × n per tick and skip the .filter().sort() chain entirely.
function getCurrentEvent(now = Date.now()) {
  const nowMs = typeof now === "number" ? now : now.getTime();
  const events = state.events;
  for (let i = 0, n = events.length; i < n; i++) {
    const ev = events[i];
    if (ev.startMs > nowMs) return null;          // sorted → no later event can match
    if (nowMs < ev.endMs)   return ev;
  }
  return null;
}

function getNextEvent(now = Date.now()) {
  const nowMs = typeof now === "number" ? now : now.getTime();
  const events = state.events;
  for (let i = 0, n = events.length; i < n; i++) {
    if (events[i].startMs > nowMs) return events[i];
  }
  return null;
}

// Per-second diff cache — every DOM write below is gated on a change vs prior tick.
// Touching .textContent / .style on identical values still triggers style recalc,
// so skipping them is a real win on a 24/7 kiosk.
const _renderPrev = {
  status: null, label: null, sublabel: null,
  countdown: null, ending: null, progress: null, progressHidden: null,
  bookedInfoHidden: null, bookedTitle: null, withHidden: null, personName: null,
  avatarText: null, avatarBg: null, avatarHidden: null, organizer: null,
  nextHidden: null, nextText: null,
};
function setText(el, val, key) {
  if (_renderPrev[key] === val) return;
  _renderPrev[key] = val;
  el.textContent = val;
}
function setHidden(el, hidden, key) {
  if (_renderPrev[key] === hidden) return;
  _renderPrev[key] = hidden;
  el.hidden = hidden;
}
function setStyle(el, prop, val, key) {
  if (_renderPrev[key] === val) return;
  _renderPrev[key] = val;
  el.style[prop] = val;
}
function setDataStatus(el, val, key) {
  if (_renderPrev[key] === val) return;
  _renderPrev[key] = val;
  el.dataset.status = val;
}

function renderStatus(nowArg) {
  const now = nowArg || new Date();
  const nowMs = now.getTime();
  const current = getCurrentEvent(nowMs);
  const next = getNextEvent(nowMs);

  const app = $("app");
  const statusLabel = $("status-label");
  const statusSublabel = $("status-sublabel");
  const progressBar = $("progress-bar");
  const progressFill = $("progress-bar-fill");

  if (current) {
    setDataStatus(app, "booked", "status");
    setText(statusLabel, "OCUPADA", "label");
    setText(statusSublabel, `Hasta las ${fmtTime(current.endMs)}`, "sublabel");

    // Countdown
    const remaining = Math.max(0, current.endMs - nowMs);
    setText($("countdown-time"), formatCountdown(remaining), "countdown");

    const endingSoon = remaining > 0 && remaining <= CONFIG.ENDING_SOON_MIN * 60_000;
    if (_renderPrev.ending !== endingSoon) {
      _renderPrev.ending = endingSoon;
      $("countdown-wrap").classList.toggle("countdown--ending", endingSoon);
    }

    // Progress bar — round to 0.1% so identical pixel-widths skip the layout flush.
    const totalMs = current.endMs - current.startMs;
    const pct = totalMs > 0 ? Math.min(1, 1 - remaining / totalMs) : 0;
    const pctStr = `${(Math.round(pct * 1000) / 10)}%`;
    setStyle(progressFill, "width", pctStr, "progress");
    setHidden(progressBar, false, "progressHidden");

    setHidden($("booked-info"), false, "bookedInfoHidden");

    const { cleanTitle, person } = splitTitle(current.title);
    setText($("booked-title"), cleanTitle || current.title, "bookedTitle");

    const withEl = $("booked-with");
    const nameEl = $("booked-person-name");
    const avatarEl = $("booked-avatar");
    if (person) {
      setHidden(withEl, false, "withHidden");
      setText(nameEl, person, "personName");
      setText(avatarEl, getInitials(person), "avatarText");
      setStyle(avatarEl, "background", colorFromName(person), "avatarBg");
      setHidden(avatarEl, false, "avatarHidden");
    } else {
      setHidden(withEl, true, "withHidden");
      if (current.organizer) {
        setText(avatarEl, getInitials(current.organizer), "avatarText");
        setStyle(avatarEl, "background", colorFromName(current.organizer), "avatarBg");
        setHidden(avatarEl, false, "avatarHidden");
      } else {
        setText(avatarEl, "", "avatarText");
        setStyle(avatarEl, "background", "rgba(255,255,255,0.15)", "avatarBg");
        setHidden(avatarEl, true, "avatarHidden");
      }
    }
    setText(
      $("booked-organizer"),
      (!person && current.organizer) ? `Organiza ${current.organizer}` : "",
      "organizer",
    );

    // "Próxima en X min" banner
    const nextEl = $("next-meeting");
    if (next) {
      const minsToNext = Math.round((next.startMs - nowMs) / 60_000);
      if (minsToNext > 0 && minsToNext <= CONFIG.NEXT_WARN_MIN) {
        setHidden(nextEl, false, "nextHidden");
        const nextSplit = splitTitle(next.title);
        const label = nextSplit.person || nextSplit.cleanTitle || next.title;
        setText($("next-meeting-text"), `${fmtTime(next.startMs)} · ${label} (en ${minsToNext} min)`, "nextText");
      } else {
        setHidden(nextEl, true, "nextHidden");
      }
    } else {
      setHidden(nextEl, true, "nextHidden");
    }

    handleCheckinState(current, now);
  } else {
    setDataStatus(app, "free", "status");
    setText(statusLabel, "LIBRE", "label");
    setHidden(progressBar, true, "progressHidden");

    let sublabel;
    if (next && isSameDay(new Date(next.startMs), now)) {
      const mins = Math.round((next.startMs - nowMs) / 60_000);
      if (mins <= 60)        sublabel = `Próxima reunión en ${formatMins(mins)}`;
      else if (mins <= 180)  sublabel = `Libre ${formatMins(mins)} · hasta las ${fmtTime(next.startMs)}`;
      else                   sublabel = `Libre hasta las ${fmtTime(next.startMs)}`;
    } else {
      sublabel = "Sin reuniones programadas hoy";
    }
    setText(statusSublabel, sublabel, "sublabel");

    setHidden($("booked-info"), true, "bookedInfoHidden");
    setHidden($("checkin-banner"), true, "checkinHidden");
    if (_renderPrev.ending !== false) {
      _renderPrev.ending = false;
      $("countdown-wrap")?.classList.remove("countdown--ending");
    }
  }
}

function handleCheckinState(ev, now) {
  const banner = $("checkin-banner");
  const sub = $("checkin-banner-sub");
  const endBtn = $("end-btn");
  const countdownWrap = $("countdown-wrap");

  if (isCheckedIn(ev)) {
    if (_renderPrev.checkinHidden !== true) { _renderPrev.checkinHidden = true; banner.hidden = true; }
    if (_renderPrev.endBtnHidden !== false) { _renderPrev.endBtnHidden = false; endBtn.hidden = false; }
    if (_renderPrev.countdownDisplay !== "") { _renderPrev.countdownDisplay = ""; countdownWrap.style.display = ""; }
    return;
  }

  const deadlineMs = ev.startMs + CONFIG.CHECKIN_WINDOW_MIN * 60_000;
  const remaining = deadlineMs - now.getTime();

  if (remaining > 0) {
    if (_renderPrev.checkinHidden !== false) { _renderPrev.checkinHidden = false; banner.hidden = false; }
    if (_renderPrev.endBtnHidden !== true)   { _renderPrev.endBtnHidden = true;   endBtn.hidden = true; }
    if (_renderPrev.countdownDisplay !== "none") { _renderPrev.countdownDisplay = "none"; countdownWrap.style.display = "none"; }
    const mm = String(Math.floor(remaining / 60000)).padStart(2, "0");
    const ss = String(Math.floor((remaining % 60000) / 1000)).padStart(2, "0");
    const txt = `Se cancelará en ${mm}:${ss} si no confirmas`;
    if (_renderPrev.checkinSub !== txt) { _renderPrev.checkinSub = txt; sub.textContent = txt; }
  } else {
    // Window elapsed. The server-side cron (every 5 min) is the main mechanism.
    // Client-side fallback: auto-cancel, but with guardrails against races.
    if (_renderPrev.checkinHidden !== true) { _renderPrev.checkinHidden = true; banner.hidden = true; }

    // If a check-in request is in flight, give it time to land.
    if (state.checkingIn) return;

    // 30s grace after the deadline in case the check-in request is pending
    // via another device, or the GET hasn't refreshed yet.
    const GRACE_MS = 30_000;
    if (remaining > -GRACE_MS) return;

    if (!state.autoCancelled.has(ev.id)) {
      state.autoCancelled.add(ev.id);
      autoCancelEvent(ev);
    }
  }
}

async function autoCancelEvent(ev) {
  try {
    await endEvent(ev.id);
    toast("Sala liberada: no se confirmó asistencia", "error");
    await loadEvents();
  } catch (err) {
    console.error("Auto-cancel failed", err);
    // Allow retry on next tick if it failed
    state.autoCancelled.delete(ev.id);
  }
}

async function handleCheckinClick() {
  const current = getCurrentEvent();
  if (!current || state.checkingIn) return;
  const btn = document.getElementById("checkin-btn");
  state.checkingIn = true;
  btn.disabled = true;
  const oldHtml = btn.innerHTML;
  btn.textContent = "Confirmando…";
  try {
    await checkInEvent(current.id);
    // Optimistic: mark locally so UI updates instantly
    current.description = `[CHECKED_IN:${new Date().toISOString()}]\n${current.description || ""}`;
    toast("Asistencia confirmada", "success");
    renderStatus();
    loadEvents(); // no await — refresh in background
  } catch (err) {
    toast("No se pudo confirmar. Reintenta.", "error");
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.innerHTML = oldHtml;
    state.checkingIn = false;
  }
}

function isSameDay(a, b) {
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

/* =========================================================
 * Events list rendering
 * ========================================================= */
function renderEventsList() {
  const list = $("events-list");
  const now = new Date();
  const nowMs = now.getTime();

  renderTimeline(now);

  const todayK = dayKey(now);
  const tomorrowK = dayKey(new Date(nowMs + 86_400_000));
  const weekCutoffMs = nowMs + CONFIG.AGENDA_WEEK_DAYS * 86_400_000;

  // Single linear pass — was 3 nested .filter() chains, each parsing Date strings.
  // state.events is pre-sorted, so groups stay sorted without re-sorting per group.
  const groups = new Map();
  const view = state.dayView;
  const events = state.events;
  for (let i = 0; i < events.length; i++) {
    const ev = events[i];
    if (ev.endMs <= nowMs) continue;
    const k = dayKey(new Date(ev.startMs));
    if (view === "today" && k !== todayK) continue;
    if (view === "tomorrow" && k !== tomorrowK) continue;
    if (view === "week" && ev.startMs >= weekCutoffMs) continue;
    let bucket = groups.get(k);
    if (!bucket) { bucket = []; groups.set(k, bucket); }
    bucket.push(ev);
  }

  if (!groups.size) {
    const emptyMsg =
      view === "today"    ? "Sin reuniones el resto del día" :
      view === "tomorrow" ? "Sin reuniones mañana" :
                            "Sin reuniones en los próximos días";
    list.innerHTML = `<div class="events-empty">${escapeHtml(emptyMsg)}</div>`;
    return;
  }

  // Build one string then a single innerHTML write — minimal layout passes.
  const sortedKeys = [...groups.keys()].sort();
  let html = "";
  for (let i = 0; i < sortedKeys.length; i++) {
    const key = sortedKeys[i];
    const header = key === todayK ? "HOY" : key === tomorrowK ? "MAÑANA" : groupHeaderFromKey(key);
    html += `<div class="events-group-header">${escapeHtml(header)}</div>`;
    const items = groups.get(key);
    for (let j = 0; j < items.length; j++) html += renderEventItem(items[j], nowMs);
  }
  list.innerHTML = html;
}

function renderEventItem(ev, nowMs) {
  const isCurrent = ev.startMs <= nowMs && nowMs < ev.endMs;
  // Cached Intl formatter — was instantiating one per call.
  const sStr = FMT_TIME.format(ev.startMs);
  const eStr = FMT_TIME.format(ev.endMs);
  const idAttr = escapeHtml(ev.id);

  return `
    <div class="event-item ${isCurrent ? "event-item--current" : ""}" data-event-id="${idAttr}">
      <div class="event-header">
        <div class="event-main">
          <div class="event-time">
            <span>${sStr}</span>
            <span class="event-time-arrow">→</span>
            <span>${eStr}</span>
          </div>
          <div class="event-title">${escapeHtml(ev.title)}</div>
          ${ev.organizer ? `<div class="event-organizer">${escapeHtml(ev.organizer)}</div>` : ""}
        </div>
        <button class="event-cancel-btn" data-cancel-id="${idAttr}" aria-label="Cancelar reunión" title="Cancelar reunión">
          <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round">
            <line x1="6" y1="6" x2="18" y2="18" />
            <line x1="18" y1="6" x2="6" y2="18" />
          </svg>
        </button>
      </div>
    </div>
  `;
}

/* =========================================================
 * Mini timeline of today
 * ========================================================= */
let _timelineHoursRendered = false;
function renderTimeline(now = new Date()) {
  const track = $("timeline-track");
  const hoursEl = $("timeline-hours");
  const subEl = $("events-header-sub");
  const timelineEl = $("timeline");
  if (!track || !hoursEl) return;

  if (state.dayView === "week") {
    timelineEl.hidden = true;
    if (subEl) subEl.textContent = `Próximos ${CONFIG.AGENDA_WEEK_DAYS} días`;
    return;
  }
  timelineEl.hidden = false;

  const targetDate = state.dayView === "tomorrow"
    ? new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1)
    : new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const targetKey = dayKey(targetDate);

  const startHour = CONFIG.DAY_START_HOUR;
  const endHour   = CONFIG.DAY_END_HOUR;
  const startMin  = startHour * 60;
  const endMin    = endHour * 60;
  const rangeMin  = endMin - startMin;

  // Hour labels are static — render only once over the lifetime of the page.
  if (!_timelineHoursRendered) {
    const frag = document.createDocumentFragment();
    for (let h = startHour; h <= endHour; h += 2) {
      const s = document.createElement("span");
      s.textContent = `${String(h).padStart(2, "0")}:00`;
      frag.appendChild(s);
    }
    hoursEl.replaceChildren(frag);
    _timelineHoursRendered = true;
  }

  // Wipe existing blocks — collect in a single pass, then detach via fragment swap.
  const oldBlocks = track.getElementsByClassName("timeline-block");
  while (oldBlocks.length) oldBlocks[0].remove();

  const nowMs = now.getTime();
  const blockFrag = document.createDocumentFragment();
  let busyMinutes = 0;
  const events = state.events;
  for (let i = 0; i < events.length; i++) {
    const ev = events[i];
    // events are sorted — once we pass the target day we can break.
    // (cheap key-check still needed because we span midnight by date string)
    const s = new Date(ev.startMs);
    if (dayKey(s) !== targetKey) continue;

    const sMin = s.getHours() * 60 + s.getMinutes();
    const e = new Date(ev.endMs);
    const eMin = e.getHours() * 60 + e.getMinutes();
    const clampedStart = sMin > startMin ? sMin : startMin;
    const clampedEnd   = eMin < endMin   ? eMin : endMin;
    if (clampedEnd <= clampedStart) continue;

    busyMinutes += clampedEnd - clampedStart;

    const leftPct  = ((clampedStart - startMin) / rangeMin) * 100;
    const widthPct = ((clampedEnd   - clampedStart) / rangeMin) * 100;

    const block = document.createElement("div");
    let cls = "timeline-block";
    if (state.dayView === "today" && ev.startMs <= nowMs && nowMs < ev.endMs) {
      cls += " timeline-block--current";
    }
    block.className = cls;
    block.style.cssText = `left:${leftPct}%;width:${widthPct}%`;
    block.title = `${ev.title} · ${fmtTime(ev.startMs)}–${fmtTime(ev.endMs)}`;
    block.dataset.eventId = ev.id;
    blockFrag.appendChild(block);
  }
  track.appendChild(blockFrag);

  const nowEl = $("timeline-now");
  if (state.dayView !== "today") {
    nowEl.hidden = true;
  } else {
    const nowMin = now.getHours() * 60 + now.getMinutes();
    if (nowMin < startMin || nowMin > endMin) {
      nowEl.hidden = true;
    } else {
      nowEl.hidden = false;
      nowEl.style.left = `${((nowMin - startMin) / rangeMin) * 100}%`;
    }
  }
  // dayEvents count for sub-header summary
  let dayEventCount = 0;
  for (let i = 0; i < events.length; i++) {
    if (dayKey(new Date(events[i].startMs)) === targetKey) dayEventCount++;
  }

  if (subEl) {
    const busyH = Math.floor(busyMinutes / 60);
    const busyM = busyMinutes % 60;
    const dayLabel = state.dayView === "tomorrow" ? "mañana" : "hoy";
    const txt = busyMinutes
      ? `${dayEventCount} ${dayEventCount === 1 ? "reunión" : "reuniones"} ${dayLabel} · ${busyH ? busyH + "h " : ""}${busyM}min ocupada`
      : `Sin reuniones ${dayLabel}`;
    subEl.textContent = txt;
  }
}

/**
 * Click on empty timeline region → open custom book modal with that start time.
 */
function handleTimelineClick(e) {
  if (state.dayView === "week") return; // timeline hidden anyway
  // Ignore clicks on busy blocks
  if (e.target.closest(".timeline-block")) return;
  const track = document.getElementById("timeline-track");
  const rect = track.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const pct = Math.max(0, Math.min(1, x / rect.width));
  const startMin = CONFIG.DAY_START_HOUR * 60;
  const endMin = CONFIG.DAY_END_HOUR * 60;
  const totalMin = startMin + pct * (endMin - startMin);
  const rounded = Math.round(totalMin / 15) * 15;
  const h = Math.floor(rounded / 60);
  const m = rounded % 60;

  const base = state.dayView === "tomorrow"
    ? new Date(Date.now() + 86_400_000)
    : new Date();
  const target = new Date(base.getFullYear(), base.getMonth(), base.getDate(), h, m, 0, 0);

  if (target.getTime() < Date.now()) {
    const now = new Date();
    now.setMinutes(Math.ceil(now.getMinutes() / 15) * 15, 0, 0);
    modal.open({ startAfter: now });
  } else {
    modal.open({ startAfter: target });
  }
}

function dayKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function groupHeaderFromKey(key) {
  const [y, m, d] = key.split("-").map(Number);
  return FMT_GROUP_HDR.format(new Date(y, m - 1, d)).toUpperCase();
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

/* =========================================================
 * Modal — book a meeting
 * ========================================================= */
const modal = {
  open(opts = {}) {
    document.getElementById("modal").hidden = false;
    // No autofocus — don't pop the tablet keyboard
    document.getElementById("modal-error").hidden = true;

    // Default date/time. If booking while BOOKED, suggest right after current ends.
    const now = new Date();
    let defaultStart;
    if (opts.startAfter) {
      defaultStart = new Date(opts.startAfter);
    } else {
      defaultStart = new Date(now);
    }
    // Round up to next 15-min slot
    defaultStart.setMinutes(Math.ceil(defaultStart.getMinutes() / 15) * 15, 0, 0);
    // Clamp to allowed booking hours (start time only)
    if (defaultStart.getHours() < BOOKING_HOURS.start) {
      defaultStart.setHours(BOOKING_HOURS.start, 0, 0, 0);
    } else if (defaultStart.getHours() > BOOKING_HOURS.end) {
      // Past end of day → next day at opening time
      defaultStart.setDate(defaultStart.getDate() + 1);
      defaultStart.setHours(BOOKING_HOURS.start, 0, 0, 0);
    }
    setSelectedDate(defaultStart);
    setTimeValue("book-time", toTimeInput(defaultStart));
    setTimeValue("book-end-time", "");

    // Reset duration mode to predefined (30 min selected)
    state.customDuration = false;
    state.bookingDuration = 30;
    const durWrap = document.getElementById("duration-options");
    const durCustom = document.getElementById("duration-custom");
    durWrap.hidden = false;
    durCustom.hidden = true;
    durWrap.querySelectorAll("button").forEach(b =>
      b.classList.toggle("selected", b.dataset.mins === "30")
    );

    mountPeoplePicker("people-picker", {
      selectable: true,
      onPick: (name) => {
        state.bookingPerson = name;
        // Clear invalid highlight and any stale error about missing person
        document.getElementById("people-field")?.classList.remove("is-invalid");
        const err = document.getElementById("modal-error");
        if (err && /quién reserva/i.test(err.textContent)) err.hidden = true;
        // Jump to date / start time / duration once a person is chosen
        if (name) {
          requestAnimationFrame(() => {
            const target = document.querySelector("#modal .modal-row");
            target?.scrollIntoView({ behavior: "smooth", block: "start" });
          });
        }
      },
    });
    // Clear any prior highlight
    document.getElementById("people-field")?.classList.remove("is-invalid");
  },
  close() {
    document.getElementById("modal").hidden = true;
    document.getElementById("book-title").value = "";
    state.bookingPerson = null;
  },
  showError(msg, { highlight } = {}) {
    const el = document.getElementById("modal-error");
    el.textContent = msg;
    el.hidden = false;
    el.scrollIntoView({ behavior: "smooth", block: "nearest" });
    if (highlight) {
      const field = document.getElementById(highlight);
      if (field) {
        field.classList.remove("is-invalid");
        // force reflow so the animation re-triggers
        void field.offsetWidth;
        field.classList.add("is-invalid");
        field.scrollIntoView({ behavior: "smooth", block: "center" });
        setTimeout(() => field.classList.remove("is-invalid"), 2400);
      }
    }
  },
};

function toDateInput(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function toTimeInput(d) {
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

/* =========================================================
 * Custom calendar date picker (tablet-friendly)
 * ========================================================= */
const calState = {
  selected: null,    // Date (day resolution)
  viewMonth: null,   // Date (first day of viewed month)
};

function setSelectedDate(d) {
  const day = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  calState.selected = day;
  calState.viewMonth = new Date(day.getFullYear(), day.getMonth(), 1);
  const label = day.toLocaleDateString(CONFIG.LOCALE, {
    weekday: "long", day: "numeric", month: "long",
  });
  document.getElementById("book-date-label").textContent = capitalize(label);
}

function getSelectedDateValue() {
  return calState.selected ? toDateInput(calState.selected) : "";
}

const dateModal = {
  open() {
    if (!calState.viewMonth) calState.viewMonth = new Date();
    if (!calState.selected) calState.selected = new Date();
    renderCalendar();
    document.getElementById("date-modal").hidden = false;
  },
  close() {
    document.getElementById("date-modal").hidden = true;
  },
};

function renderCalendar() {
  const view = calState.viewMonth;
  const title = view.toLocaleDateString(CONFIG.LOCALE, { month: "long", year: "numeric" });
  document.getElementById("cal-title").textContent = capitalize(title);

  const grid = document.getElementById("cal-grid");
  grid.innerHTML = "";

  const y = view.getFullYear();
  const m = view.getMonth();
  const firstOfMonth = new Date(y, m, 1);
  // Monday-first: getDay() Sun=0..Sat=6 → (dow+6)%7 gives Mon=0..Sun=6
  const leading = (firstOfMonth.getDay() + 6) % 7;
  const start = new Date(y, m, 1 - leading);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 0; i < 42; i++) {
    const d = new Date(start.getFullYear(), start.getMonth(), start.getDate() + i);
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "cal-day";
    btn.textContent = String(d.getDate());

    if (d.getMonth() !== m) btn.classList.add("cal-day--muted");
    if (sameDay(d, today)) btn.classList.add("cal-day--today");
    if (calState.selected && sameDay(d, calState.selected)) btn.classList.add("cal-day--selected");
    if (d < today) btn.disabled = true;

    btn.addEventListener("click", () => {
      setSelectedDate(d);
      renderCalendar();
      dateModal.close();
    });

    grid.appendChild(btn);
  }
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

/* =========================================================
 * Custom time picker (tablet-friendly)
 * ========================================================= */
const timePickerState = {
  inputId: null,
  hour: null,
  minute: null,
  minHour: 0,
  maxHour: 23,
};

function setTimeValue(inputId, hhmm) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.value = hhmm || "";
  const label = document.getElementById(`${inputId}-label`);
  if (label) label.textContent = hhmm || "—";
  input.dispatchEvent(new Event("input", { bubbles: true }));
}

const timePickerModal = {
  open(inputId, opts = {}) {
    const input = document.getElementById(inputId);
    if (!input) return;
    timePickerState.inputId = inputId;
    timePickerState.minHour = Number.isInteger(opts.minHour) ? opts.minHour : 0;
    timePickerState.maxHour = Number.isInteger(opts.maxHour) ? opts.maxHour : 23;

    let h = 9, m = 0;
    const current = (input.value || "").split(":").map(Number);
    if (!Number.isNaN(current[0]) && !Number.isNaN(current[1])) {
      h = current[0];
      m = Math.round(current[1] / 15) * 15;
      if (m === 60) { h = (h + 1) % 24; m = 0; }
    }
    if (h < timePickerState.minHour) { h = timePickerState.minHour; m = 0; }
    if (h > timePickerState.maxHour) { h = timePickerState.maxHour; m = 0; }
    timePickerState.hour = h;
    timePickerState.minute = m;

    document.getElementById("time-modal-title").textContent = opts.title || "Selecciona la hora";
    renderTimePicker();
    document.getElementById("time-modal").hidden = false;
  },
  close() {
    document.getElementById("time-modal").hidden = true;
    timePickerState.inputId = null;
  },
  confirm() {
    const id = timePickerState.inputId;
    if (!id) return this.close();
    const hh = String(timePickerState.hour).padStart(2, "0");
    const mm = String(timePickerState.minute).padStart(2, "0");
    setTimeValue(id, `${hh}:${mm}`);
    this.close();
  },
};

function renderTimePicker() {
  const { hour, minute } = timePickerState;
  document.getElementById("time-display-hour").textContent = String(hour).padStart(2, "0");
  document.getElementById("time-display-minute").textContent = String(minute).padStart(2, "0");

  const hoursGrid = document.getElementById("time-grid-hours");
  hoursGrid.innerHTML = "";
  const { minHour, maxHour } = timePickerState;
  for (let h = minHour; h <= maxHour; h++) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "time-cell" + (h === hour ? " time-cell--selected" : "");
    btn.textContent = String(h).padStart(2, "0");
    btn.addEventListener("click", () => {
      timePickerState.hour = h;
      renderTimePicker();
    });
    hoursGrid.appendChild(btn);
  }

  const minsGrid = document.getElementById("time-grid-minutes");
  minsGrid.innerHTML = "";
  [0, 15, 30, 45].forEach((m) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "time-cell" + (m === minute ? " time-cell--selected" : "");
    btn.textContent = `:${String(m).padStart(2, "0")}`;
    btn.addEventListener("click", () => {
      timePickerState.minute = m;
      renderTimePicker();
    });
    minsGrid.appendChild(btn);
  });
}

/* =========================================================
 * "Who's booking" selector (quick-book flow)
 * ========================================================= */
const whoModal = {
  open(mins) {
    state.bookingDuration = mins;
    document.getElementById("who-modal-text").textContent = `Reserva rápida de ${formatMins(mins)}`;
    mountPeoplePicker("who-people-picker", {
      selectable: false,
      onPick: (name) => {
        whoModal.close();
        titleModal.open(name);
      },
    });
    document.getElementById("who-modal").hidden = false;
  },
  close() {
    document.getElementById("who-modal").hidden = true;
  },
};

/* =========================================================
 * Title modal — asks for meeting title after duration + person
 * ========================================================= */
const titleModal = {
  open(person) {
    state.bookingPerson = person;
    const mins = state.bookingDuration;
    const label = formatMins(mins);
    document.getElementById("title-modal-text").textContent = `${person} · ${label}`;
    const input = document.getElementById("quick-title-input");
    input.value = "";
    document.getElementById("title-modal-error").hidden = true;
    renderTitleChips();
    updateTitleButton(null);
    document.getElementById("title-modal").hidden = false;
    // No autofocus — avoid popping the on-screen keyboard on tablets
  },
  close() {
    document.getElementById("title-modal").hidden = true;
    state.bookingPerson = null;
    state.quickTitleChoice = null;
  },
  showError(msg) {
    const el = document.getElementById("title-modal-error");
    el.textContent = msg;
    el.hidden = false;
  },
};

function renderTitleChips() {
  const host = document.getElementById("title-chips");
  host.innerHTML = "";
  for (const t of CONFIG.QUICK_TITLES) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "title-chip";
    b.textContent = t;
    b.addEventListener("click", () => {
      state.quickTitleChoice = t;
      document.getElementById("quick-title-input").value = "";
      host.querySelectorAll(".title-chip").forEach(x =>
        x.classList.toggle("selected", x === b)
      );
      updateTitleButton(t);
    });
    host.appendChild(b);
  }
}

function updateTitleButton(title) {
  const btn = document.getElementById("confirm-title");
  if (title) btn.textContent = `Reservar · ${title}`;
  else {
    const typed = document.getElementById("quick-title-input").value.trim();
    btn.textContent = typed ? `Reservar · ${typed}` : "Reservar sin título";
  }
}

async function handleTitleConfirm() {
  const person = state.bookingPerson;
  if (!person) { titleModal.close(); return; }
  const input = document.getElementById("quick-title-input");
  const typed = input.value.trim();
  const baseTitle = typed || state.quickTitleChoice || "Reunión rápida";

  const btn = document.getElementById("confirm-title");
  btn.disabled = true;
  btn.textContent = "Reservando…";
  try {
    await quickBookWithTitle(baseTitle, person);
    titleModal.close();
  } catch (err) {
    const msg = /Conflict with "(.+)"/.exec(err.message);
    titleModal.showError(msg ? `Se solapa con "${msg[1]}"` : "No se pudo reservar. Reintenta.");
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = "Reservar";
  }
}

/**
 * Renders a full people picker: search box + initial-letter filter + grid.
 * options.selectable — if true, keeps the selected button highlighted.
 */
// Sort + lowercase the people list ONCE at module load. Was re-sorted on every
// modal open + the filter() ran toLowerCase() per name per keystroke.
const _PEOPLE_SORTED = [...CONFIG.PEOPLE]
  .map(s => s.trim())
  .filter(Boolean)
  .sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }))
  .map(name => ({ name, lower: name.toLowerCase(), upper0: name[0].toUpperCase() }));
const _PEOPLE_INITIALS = [...new Set(_PEOPLE_SORTED.map(p => p.upper0))].sort();

function mountPeoplePicker(containerId, { selectable, onPick }) {
  const host = document.getElementById(containerId);
  host.replaceChildren();

  const people = _PEOPLE_SORTED;
  const initials = _PEOPLE_INITIALS;

  // Persistent "Seleccionado" chip (selectable mode only)
  let selectedChip = null;
  if (selectable) {
    selectedChip = document.createElement("div");
    selectedChip.className = "people-selected-chip";
    selectedChip.hidden = true;
    host.appendChild(selectedChip);
  }

  // Search input
  const search = document.createElement("input");
  search.type = "text";
  search.className = "people-search";
  search.placeholder = "Buscar por nombre…";
  host.appendChild(search);

  // Initials row
  const initialsRow = document.createElement("div");
  initialsRow.className = "people-initials";
  const allBtn = document.createElement("button");
  allBtn.type = "button";
  allBtn.textContent = "Todos";
  allBtn.style.width = "auto";
  allBtn.style.padding = "0 10px";
  allBtn.classList.add("active");
  allBtn.dataset.letter = "";
  initialsRow.appendChild(allBtn);
  for (const letter of initials) {
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = letter;
    b.dataset.letter = letter;
    initialsRow.appendChild(b);
  }
  host.appendChild(initialsRow);

  // Grid
  const grid = document.createElement("div");
  grid.className = "people-grid";
  host.appendChild(grid);

  let activeLetter = "";
  let activeQuery = "";
  let selected = null;

  function renderSelectedChip() {
    if (!selectedChip) return;
    if (!selected) {
      selectedChip.hidden = true;
      host.classList.remove("people-picker--collapsed");
      return;
    }
    selectedChip.hidden = false;
    host.classList.add("people-picker--collapsed");
    selectedChip.innerHTML = "";
    const label = document.createElement("span");
    label.className = "people-selected-label";
    label.textContent = "Seleccionado:";
    const name = document.createElement("strong");
    name.textContent = selected;
    const change = document.createElement("button");
    change.type = "button";
    change.className = "people-selected-change";
    change.textContent = "Cambiar";
    change.addEventListener("click", (e) => {
      e.stopPropagation();
      selected = null;
      onPick(null);
      renderSelectedChip();
      render();
    });
    const x = document.createElement("button");
    x.type = "button";
    x.className = "people-selected-clear";
    x.setAttribute("aria-label", "Quitar selección");
    x.innerHTML = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><line x1="6" y1="6" x2="18" y2="18"/><line x1="18" y1="6" x2="6" y2="18"/></svg>';
    x.addEventListener("click", (e) => {
      e.stopPropagation();
      selected = null;
      onPick(null);
      renderSelectedChip();
      render();
    });
    selectedChip.appendChild(label);
    selectedChip.appendChild(name);
    selectedChip.appendChild(change);
    selectedChip.appendChild(x);
  }

  // Delegated click on the grid — one listener instead of N (≈28).
  grid.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-person]");
    if (!btn) return;
    const name = btn.dataset.person;
    if (selectable) {
      selected = name;
      renderSelectedChip();
      const prev = grid.querySelector("button.selected");
      if (prev && prev !== btn) prev.classList.remove("selected");
      btn.classList.add("selected");
    }
    onPick(name);
  });

  function render() {
    const q = activeQuery.toLowerCase();
    const frag = document.createDocumentFragment();
    let count = 0;
    for (let i = 0, n = people.length; i < n; i++) {
      const p = people[i];
      if (activeLetter && p.upper0 !== activeLetter) continue;
      if (q && p.lower.indexOf(q) === -1) continue;
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = p.name;
      b.dataset.person = p.name;
      if (selectable && selected === p.name) b.className = "selected";
      frag.appendChild(b);
      count++;
    }
    if (!count) {
      const empty = document.createElement("div");
      empty.className = "people-empty";
      empty.textContent = "Sin resultados";
      frag.appendChild(empty);
    }
    grid.replaceChildren(frag);
  }

  // rAF-coalesce: rapid keystrokes only trigger one render per frame.
  let _renderPending = false;
  function scheduleRender() {
    if (_renderPending) return;
    _renderPending = true;
    requestAnimationFrame(() => { _renderPending = false; render(); });
  }

  search.addEventListener("input", (e) => {
    activeQuery = e.target.value;
    scheduleRender();
  });

  initialsRow.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-letter]");
    if (!btn) return;
    activeLetter = btn.dataset.letter;
    initialsRow.querySelectorAll("button").forEach(b =>
      b.classList.toggle("active", b === btn)
    );
    render();
  });

  render();
  renderSelectedChip();
}

async function handleBookConfirm() {
  if (!state.bookingPerson) {
    modal.showError("Selecciona quién reserva la sala", { highlight: "people-picker" });
    return;
  }
  const titleInput = document.getElementById("book-title");
  const baseTitle = titleInput.value.trim() || "Reunión rápida";
  const title = `${baseTitle} — ${state.bookingPerson}`;
  const mins = state.bookingDuration;
  if (!mins || mins <= 0) {
    modal.showError("Duración no válida. Revisa la hora de fin.");
    return;
  }

  const dateStr = getSelectedDateValue();
  const timeStr = document.getElementById("book-time").value;
  if (!dateStr || !timeStr) {
    modal.showError("Selecciona fecha y hora");
    return;
  }
  const [y, mo, d] = dateStr.split("-").map(Number);
  const [h, mi] = timeStr.split(":").map(Number);
  const start = new Date(y, mo - 1, d, h, mi, 0, 0);

  if (h < BOOKING_HOURS.start || h > BOOKING_HOURS.end) {
    modal.showError(`La hora de inicio debe estar entre las ${String(BOOKING_HOURS.start).padStart(2,"0")}:00 y las ${String(BOOKING_HOURS.end).padStart(2,"0")}:45`);
    return;
  }

  if (start.getTime() < Date.now() - 60_000) {
    modal.showError("No se puede reservar en el pasado");
    return;
  }

  const end = new Date(start.getTime() + mins * 60_000);

  const conflict = findConflict(start, end);
  if (conflict) {
    modal.showError(`Se solapa con "${conflict.title}" (${fmtDateTime(conflict.start)}–${fmtTime(conflict.end)})`);
    return;
  }

  const btn = document.getElementById("confirm-book");
  btn.disabled = true;
  btn.textContent = "Creando…";

  try {
    const created = await createEvent({ title, startISO: start.toISOString(), endISO: end.toISOString() });
    // Auto-check-in if starting now
    await autoCheckinAfterCreate(created, start);
    modal.close();
    toast("Reunión creada", "success");
    await loadEvents();
  } catch (err) {
    modal.showError("No se pudo crear la reunión. Reintenta.");
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = "Confirmar";
  }
}

function findConflict(start, end) {
  const startMs = start instanceof Date ? start.getTime() : start;
  const endMs   = end   instanceof Date ? end.getTime()   : end;
  const events = state.events;
  for (let i = 0; i < events.length; i++) {
    const ev = events[i];
    if (ev.startMs >= endMs) break; // sorted — no later event can overlap
    if (ev.endMs > startMs) return ev;
  }
  return null;
}

async function quickBookWithTitle(baseTitle, person) {
  const mins = state.bookingDuration;
  const now = new Date();
  const start = new Date(now); start.setSeconds(0, 0);
  const end = new Date(start.getTime() + mins * 60_000);

  const conflict = findConflict(start, end);
  if (conflict) {
    throw new Error(`Conflict with "${conflict.title}"`);
  }

  const title = `${baseTitle} — ${person}`;
  const created = await createEvent({ title, startISO: start.toISOString(), endISO: end.toISOString() });
  // Auto-check-in: the person is physically here
  await autoCheckinAfterCreate(created, start);
  toast(`Reservado ${formatMins(mins)} — ${person}`, "success");
  await loadEvents();
}

/**
 * Auto-check-in if the booking starts within ~1 min of now, since the
 * person is physically at the tablet. Fails silently.
 */
async function autoCheckinAfterCreate(createdEvent, start) {
  if (!createdEvent?.id) return;
  const isNow = Math.abs(Date.now() - start.getTime()) < 60_000;
  if (!isNow) return;
  try {
    await checkInEvent(createdEvent.id);
  } catch (err) {
    console.warn("Auto-checkin failed (non-blocking)", err);
  }
}

/* =========================================================
 * Modal — cancel (delete) any meeting from the list
 * ========================================================= */
const cancelModal = {
  pendingId: null,
  open(eventId) {
    const ev = state.events.find(e => e.id === eventId);
    if (!ev) return;
    this.pendingId = eventId;
    const from = fmtTime(ev.start), to = fmtTime(ev.end);
    document.getElementById("cancel-modal-text").textContent =
      `Se eliminará "${ev.title}" (${from}–${to}) del calendario.`;
    document.getElementById("cancel-modal-error").hidden = true;
    document.getElementById("cancel-modal").hidden = false;
  },
  close() {
    document.getElementById("cancel-modal").hidden = true;
    this.pendingId = null;
  },
  showError(msg) {
    const el = document.getElementById("cancel-modal-error");
    el.textContent = msg;
    el.hidden = false;
  },
};

async function handleCancelConfirm() {
  const id = cancelModal.pendingId;
  if (!id) { cancelModal.close(); return; }
  const btn = document.getElementById("confirm-cancel");
  btn.disabled = true;
  btn.textContent = "Cancelando…";
  try {
    await deleteEvent(id);
    cancelModal.close();
    toast("Reunión cancelada", "success");
    await loadEvents();
  } catch (err) {
    cancelModal.showError("No se pudo cancelar. Reintenta.");
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = "Sí, cancelar";
  }
}

/* =========================================================
 * Modal — end current meeting
 * ========================================================= */
const endModal = {
  open() {
    const current = getCurrentEvent();
    if (!current) return;
    const txt = `Se finalizará "${current.title}" ahora mismo, liberando la sala.`;
    document.getElementById("end-modal-text").textContent = txt;
    document.getElementById("end-modal-error").hidden = true;
    document.getElementById("end-modal").hidden = false;
  },
  close() {
    document.getElementById("end-modal").hidden = true;
  },
  showError(msg) {
    const el = document.getElementById("end-modal-error");
    el.textContent = msg;
    el.hidden = false;
  },
};

async function handleEndConfirm() {
  const current = getCurrentEvent();
  if (!current) { endModal.close(); return; }

  const btn = document.getElementById("confirm-end");
  btn.disabled = true;
  btn.textContent = "Finalizando…";

  try {
    await endEvent(current.id);
    endModal.close();
    toast("Reunión finalizada", "success");
    await loadEvents();
  } catch (err) {
    endModal.showError("No se pudo finalizar. Reintenta.");
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = "Sí, finalizar";
  }
}

function fmtTime(iso) {
  // Accepts ISO string OR a numeric ms timestamp (faster for pre-normalized events).
  const t = typeof iso === "number" ? iso : Date.parse(iso);
  return FMT_TIME.format(t);
}

function formatMins(mins) {
  if (mins < 60) return `${mins} min`;
  const h = Math.floor(mins / 60), m = mins % 60;
  if (m === 0) return h === 1 ? "1 hora" : `${h} horas`;
  return `${h}h ${m}min`;
}

// Live countdown: under 1h shows MM:SS, 1h+ shows H:MM:SS so the user can
// always read the remaining hours at a glance instead of e.g. "119:30".
function formatCountdown(ms) {
  const total = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(total / 3600);
  const m = Math.floor((total % 3600) / 60);
  const s = total % 60;
  const ss = String(s).padStart(2, "0");
  if (h > 0) {
    const mm = String(m).padStart(2, "0");
    return `${h}:${mm}:${ss}`;
  }
  const mm = String(m).padStart(2, "0");
  return `${mm}:${ss}`;
}

function fmtDateTime(iso) {
  const t = typeof iso === "number" ? iso : Date.parse(iso);
  return `${FMT_DATE_SHORT.format(t)} ${FMT_TIME.format(t)}`;
}

/* =========================================================
 * Toast
 * ========================================================= */
let toastTimer = null;
function toast(msg, kind = "") {
  const el = document.getElementById("toast");
  el.className = "toast" + (kind ? ` toast--${kind}` : "");
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.hidden = true; }, 2400);
}

/* =========================================================
 * Data loading
 * ========================================================= */
// Fetch lifecycle: in-flight lock prevents overlapping fetches when the network
// is slow, exponential backoff stops us hammering n8n when it's down, and we
// prune autoCancelled so a kiosk that runs for weeks doesn't accumulate IDs.
let _fetchInFlight = false;
let _fetchBackoffMs = 0;
let _fetchBackoffUntil = 0;
async function loadEvents({ force = false } = {}) {
  if (_fetchInFlight) return;
  if (!force && Date.now() < _fetchBackoffUntil) return;
  _fetchInFlight = true;
  try {
    const events = await fetchEvents();
    state.events = events;
    state.lastFetch = Date.now();
    state.lastFetchOk = Date.now();
    state.fetchFailing = false;
    _fetchBackoffMs = 0;
    _fetchBackoffUntil = 0;

    // Prune autoCancelled — drop IDs that the server no longer returns so the
    // Set can't grow indefinitely on a long-running kiosk.
    if (state.autoCancelled.size) {
      const live = new Set();
      for (let i = 0; i < events.length; i++) live.add(events[i].id);
      for (const id of state.autoCancelled) if (!live.has(id)) state.autoCancelled.delete(id);
    }

    renderEventsList();
    renderStatus();
    updateOfflineBadge();
  } catch (err) {
    console.error("Error cargando eventos:", err);
    state.fetchFailing = true;
    // Exponential backoff: 5s → 10s → 20s → … capped at 5 min.
    _fetchBackoffMs = Math.min(300_000, _fetchBackoffMs ? _fetchBackoffMs * 2 : 5_000);
    _fetchBackoffUntil = Date.now() + _fetchBackoffMs;
    updateOfflineBadge();
  } finally {
    _fetchInFlight = false;
  }
}

function updateOfflineBadge() {
  const badge = document.getElementById("offline-badge");
  if (!badge) return;
  const age = Date.now() - (state.lastFetchOk || 0);
  const isOffline = state.fetchFailing && age > CONFIG.OFFLINE_THRESHOLD_MS;
  badge.hidden = !isOffline;
  if (isOffline) {
    const mins = Math.floor(age / 60_000);
    document.getElementById("offline-text").textContent =
      mins ? `Sin conexión · hace ${mins} min` : "Sin conexión";
  }
}

/* =========================================================
 * Init
 * ========================================================= */
function init() {
  document.getElementById("room-name").textContent = CONFIG.ROOM_NAME;

  // Clock + periodic re-render
  tickClock();
  setInterval(tickClock, CONFIG.REFRESH_CLOCK_MS);

  // Events fetch loop
  loadEvents();
  setInterval(loadEvents, CONFIG.REFRESH_EVENTS_MS);

  // Quick booking buttons (15/30/60 min) → ask who first
  document.querySelectorAll("[data-quick-mins]").forEach(btn => {
    btn.addEventListener("click", () => {
      whoModal.open(Number(btn.dataset.quickMins));
    });
  });

  // "Otra duración" → open custom modal
  document.getElementById("quick-custom").addEventListener("click", () => modal.open());

  // Who modal close handlers
  document.querySelectorAll("[data-who-close]").forEach(el => {
    el.addEventListener("click", () => whoModal.close());
  });

  // Date picker
  document.getElementById("book-date-trigger").addEventListener("click", () => dateModal.open());
  document.querySelectorAll("[data-date-close]").forEach(el => {
    el.addEventListener("click", () => dateModal.close());
  });
  document.getElementById("cal-prev").addEventListener("click", () => {
    const v = calState.viewMonth;
    calState.viewMonth = new Date(v.getFullYear(), v.getMonth() - 1, 1);
    renderCalendar();
  });
  document.getElementById("cal-next").addEventListener("click", () => {
    const v = calState.viewMonth;
    calState.viewMonth = new Date(v.getFullYear(), v.getMonth() + 1, 1);
    renderCalendar();
  });
  document.getElementById("cal-today-btn").addEventListener("click", () => {
    setSelectedDate(new Date());
    renderCalendar();
    dateModal.close();
  });

  // Time picker
  document.querySelectorAll(".time-trigger[data-time-target]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.timeTarget;
      const opts = id === "book-end-time"
        ? { title: "Hora de fin" }
        : { title: "Hora de inicio", minHour: BOOKING_HOURS.start, maxHour: BOOKING_HOURS.end };
      timePickerModal.open(id, opts);
    });
  });
  document.querySelectorAll("[data-time-close]").forEach((el) => {
    el.addEventListener("click", () => timePickerModal.close());
  });
  document.getElementById("time-confirm").addEventListener("click", () => timePickerModal.confirm());

  // Title modal — close + confirm + Enter key
  document.querySelectorAll("[data-title-close]").forEach(el => {
    el.addEventListener("click", () => titleModal.close());
  });
  document.getElementById("confirm-title").addEventListener("click", handleTitleConfirm);
  document.getElementById("quick-title-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter") handleTitleConfirm();
  });
  // Update button label as user types / clear chip selection
  document.getElementById("quick-title-input").addEventListener("input", () => {
    if (state.quickTitleChoice) {
      state.quickTitleChoice = null;
      document.querySelectorAll("#title-chips .title-chip.selected")
        .forEach(c => c.classList.remove("selected"));
    }
    updateTitleButton(null);
  });

  // Modal — close handlers
  document.querySelectorAll("[data-close]").forEach(el => {
    el.addEventListener("click", () => modal.close());
  });

  // Duration selector (with "Más…" custom end-time mode)
  const durWrap = document.getElementById("duration-options");
  const durCustom = document.getElementById("duration-custom");
  const endTimeInput = document.getElementById("book-end-time");
  const durInfo = document.getElementById("duration-custom-info");

  durWrap.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-mins]");
    if (!btn) return;

    if (btn.dataset.mins === "more") {
      // Switch to custom end-time mode
      state.customDuration = true;
      durWrap.hidden = true;
      durCustom.hidden = false;
      // Default end = start + 2h15m, clamped to hour
      const start = getBookingStart() || new Date();
      const end = new Date(start.getTime() + 135 * 60_000);
      setTimeValue("book-end-time", toTimeInput(end));
      updateCustomDurationInfo();
      return;
    }

    durWrap.querySelectorAll("button").forEach(b => b.classList.remove("selected"));
    btn.classList.add("selected");
    state.bookingDuration = Number(btn.dataset.mins);
    state.customDuration = false;
  });

  document.getElementById("duration-back").addEventListener("click", () => {
    state.customDuration = false;
    durCustom.hidden = true;
    durWrap.hidden = false;
    // Reset to 30 min selected
    durWrap.querySelectorAll("button").forEach(b =>
      b.classList.toggle("selected", b.dataset.mins === "30")
    );
    state.bookingDuration = 30;
  });

  endTimeInput.addEventListener("input", updateCustomDurationInfo);

  function updateCustomDurationInfo() {
    const start = getBookingStart();
    if (!start) { durInfo.textContent = "Selecciona fecha y hora de inicio primero"; return; }
    const [h, m] = (endTimeInput.value || "").split(":").map(Number);
    if (Number.isNaN(h) || Number.isNaN(m)) { durInfo.textContent = "Introduce una hora de fin"; return; }
    const end = new Date(start); end.setHours(h, m, 0, 0);
    const mins = Math.round((end - start) / 60_000);
    if (mins <= 0) {
      durInfo.textContent = "La hora de fin debe ser posterior a la de inicio";
      durInfo.classList.add("is-error");
      state.bookingDuration = 0;
      return;
    }
    durInfo.classList.remove("is-error");
    const h2 = Math.floor(mins / 60), m2 = mins % 60;
    const parts = [];
    if (h2) parts.push(`${h2}h`);
    if (m2) parts.push(`${m2}min`);
    durInfo.textContent = `Duración: ${parts.join(" ")} (${fmtTime(start)} → ${fmtTime(end)})`;
    state.bookingDuration = mins;
  }

  function getBookingStart() {
    const dateStr = getSelectedDateValue();
    const timeStr = document.getElementById("book-time").value;
    if (!dateStr || !timeStr) return null;
    const [y, mo, d] = dateStr.split("-").map(Number);
    const [h, mi] = timeStr.split(":").map(Number);
    return new Date(y, mo - 1, d, h, mi, 0, 0);
  }

  // Recompute custom duration when start date/time changes
  document.getElementById("book-time").addEventListener("input", () => {
    if (state.customDuration) updateCustomDurationInfo();
  });

  // Confirm booking
  document.getElementById("confirm-book").addEventListener("click", handleBookConfirm);

  // End meeting button
  document.getElementById("end-btn").addEventListener("click", () => endModal.open());

  // Cancel buttons (delegated on the events list)
  document.getElementById("events-list").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-cancel-id]");
    if (!btn) return;
    e.stopPropagation();
    cancelModal.open(btn.dataset.cancelId);
  });
  document.querySelectorAll("[data-cancel-close]").forEach(el => {
    el.addEventListener("click", () => cancelModal.close());
  });
  document.getElementById("confirm-cancel").addEventListener("click", handleCancelConfirm);

  // Check-in button
  document.getElementById("checkin-btn").addEventListener("click", handleCheckinClick);

  // Cancel reservation from the check-in banner
  document.getElementById("checkin-cancel-btn").addEventListener("click", () => {
    const current = getCurrentEvent();
    if (!current) return;
    cancelModal.open(current.id);
  });

  // Update offline badge once a minute
  setInterval(updateOfflineBadge, 30_000);

  // "Reservar otra" from the BOOKED view — opens custom modal starting after the current event
  document.getElementById("schedule-other-btn").addEventListener("click", () => {
    const current = getCurrentEvent();
    modal.open(current ? { startAfter: current.end } : {});
  });

  // Timeline: click on free area opens booking modal preset to that time
  document.getElementById("timeline-track").addEventListener("click", handleTimelineClick);

  // Day tabs
  document.getElementById("day-tabs").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-day]");
    if (!btn) return;
    state.dayView = btn.dataset.day;
    document.querySelectorAll("#day-tabs button").forEach(b =>
      b.classList.toggle("active", b === btn)
    );
    renderEventsList();
  });
  document.querySelectorAll("[data-end-close]").forEach(el => {
    el.addEventListener("click", () => endModal.close());
  });
  document.getElementById("confirm-end").addEventListener("click", handleEndConfirm);

  // Keyboard: Esc closes modals
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      modal.close(); endModal.close(); whoModal.close();
      titleModal.close(); dateModal.close(); cancelModal.close();
    }
  });

  // Prevent tablet screen sleep hint (needs user interaction on most browsers).
  // We retain the wake-lock object so we can re-acquire it after visibilitychange,
  // since browsers drop it on tab/page hide.
  let _wakeLock = null;
  async function acquireWake() {
    if (!("wakeLock" in navigator)) return;
    try { _wakeLock = await navigator.wakeLock.request("screen"); }
    catch (_) { _wakeLock = null; }
  }
  document.addEventListener("click", function requestWakeOnce() {
    acquireWake();
    document.removeEventListener("click", requestWakeOnce);
  });

  // When the tablet wakes from sleep:
  //   1. The 60s fetch interval has been paused — force-refresh immediately so
  //      the screen doesn't show minutes-old data.
  //   2. Re-acquire the screen wake-lock (browsers drop it on visibility hide).
  //   3. Reset the diff cache so the next renderStatus repaints from scratch
  //      (the DOM may have been frozen mid-update).
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState !== "visible") return;
    for (const k in _renderPrev) _renderPrev[k] = null;
    lastClockMinute = -1;
    lastListMinute = -1;
    tickClock();
    loadEvents({ force: true });
    acquireWake();
  });
}

document.addEventListener("DOMContentLoaded", init);
