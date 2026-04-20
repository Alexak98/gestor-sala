/* =========================================================
 * GestorSala — Meeting room display
 * =========================================================
 * CONFIG: change these values when you connect n8n
 * ========================================================= */
const CONFIG = {
  ROOM_NAME: "Sala Reuniones",

  // n8n webhooks
  N8N_GET_EVENTS_URL: "https://n8n-soporte.data.yurest.dev/webhook/sala-eventos",
  N8N_CREATE_EVENT_URL: "https://n8n-soporte.data.yurest.dev/webhook/sala-reservar",
  N8N_END_EVENT_URL: "https://n8n-soporte.data.yurest.dev/webhook/sala-finalizar",
  N8N_CHECKIN_URL:   "https://n8n-soporte.data.yurest.dev/webhook/sala-checkin",
  N8N_DELETE_URL:    "https://n8n-soporte.data.yurest.dev/webhook/sala-eliminar",

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
    // Simulate: shorten mock event end to now
    const ev = state.events.find(e => e.id === eventId);
    if (ev) ev.end = new Date().toISOString();
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
  return items.map(ev => ({
    id: ev.id,
    title: ev.title || ev.summary || "(sin título)",
    organizer:
      ev.organizer ||
      ev.organizerName ||
      (ev.organizer && ev.organizer.displayName) ||
      (ev.creator && ev.creator.displayName) ||
      "",
    start: ev.start?.dateTime || ev.start?.date || ev.start,
    end: ev.end?.dateTime || ev.end?.date || ev.end,
    description: ev.description || "",
  })).filter(ev => ev.start && ev.end);
}

function isCheckedIn(ev) {
  return /\[CHECKED_IN:/i.test(ev?.description || "");
}

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
};

/* =========================================================
 * Clock
 * ========================================================= */
let lastListMinute = -1;
function tickClock() {
  const now = new Date();
  const time = now.toLocaleTimeString(CONFIG.LOCALE, {
    hour: "2-digit", minute: "2-digit", hour12: false,
  });
  const date = now.toLocaleDateString(CONFIG.LOCALE, {
    weekday: "long", day: "numeric", month: "long",
  });
  document.getElementById("clock-time").textContent = time;
  document.getElementById("clock-date").textContent = capitalize(date);

  // Re-render status every tick (cheap) to keep countdown fresh
  renderStatus();

  // Re-render the event list once per minute to drop just-ended events
  const mKey = now.getHours() * 60 + now.getMinutes();
  if (mKey !== lastListMinute) {
    lastListMinute = mKey;
    renderEventsList();
  }
}

function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

/* =========================================================
 * Status panel rendering
 * ========================================================= */
function getCurrentEvent(now = new Date()) {
  return state.events.find(ev => {
    const s = new Date(ev.start), e = new Date(ev.end);
    return s <= now && now < e;
  }) || null;
}

function getNextEvent(now = new Date()) {
  return state.events
    .filter(ev => new Date(ev.start) > now)
    .sort((a, b) => new Date(a.start) - new Date(b.start))[0] || null;
}

function renderStatus() {
  const now = new Date();
  const current = getCurrentEvent(now);
  const next = getNextEvent(now);
  const app = document.getElementById("app");
  const statusLabel = document.getElementById("status-label");
  const statusSublabel = document.getElementById("status-sublabel");
  const progressBar = document.getElementById("progress-bar");
  const progressFill = document.getElementById("progress-bar-fill");

  if (current) {
    app.dataset.status = "booked";
    statusLabel.textContent = "OCUPADA";
    statusSublabel.textContent = `Hasta las ${fmtTime(current.end)}`;

    // Countdown
    const endMs = new Date(current.end).getTime();
    const remaining = Math.max(0, endMs - now.getTime());
    const mm = String(Math.floor(remaining / 60000)).padStart(2, "0");
    const ss = String(Math.floor((remaining % 60000) / 1000)).padStart(2, "0");
    document.getElementById("countdown-time").textContent = `${mm}:${ss}`;

    // Progress bar (elapsed)
    const totalMs = new Date(current.end) - new Date(current.start);
    const pct = totalMs > 0 ? Math.min(1, 1 - remaining / totalMs) : 0;
    progressFill.style.width = `${pct * 100}%`;
    progressBar.hidden = false;

    document.getElementById("booked-info").hidden = false;
    document.getElementById("booked-title").textContent = current.title;
    document.getElementById("booked-organizer").textContent =
      current.organizer ? `Organiza ${current.organizer}` : "";

    // Check-in / auto-cancel logic
    handleCheckinState(current, now);
  } else {
    app.dataset.status = "free";
    statusLabel.textContent = "LIBRE";
    progressBar.hidden = true;

    if (next && isSameDay(new Date(next.start), now)) {
      const mins = Math.round((new Date(next.start) - now) / 60000);
      if (mins <= 60) {
        statusSublabel.textContent = `Próxima reunión en ${mins} min`;
      } else if (mins <= 180) {
        statusSublabel.textContent = `Libre ${formatMins(mins)} · hasta las ${fmtTime(next.start)}`;
      } else {
        statusSublabel.textContent = `Libre hasta las ${fmtTime(next.start)}`;
      }
    } else {
      statusSublabel.textContent = "Sin reuniones programadas hoy";
    }

    document.getElementById("booked-info").hidden = true;
    document.getElementById("checkin-banner").hidden = true;
  }
}

function handleCheckinState(ev, now) {
  const banner = document.getElementById("checkin-banner");
  const sub = document.getElementById("checkin-banner-sub");
  const endBtn = document.getElementById("end-btn");
  const countdownWrap = document.getElementById("countdown-wrap");

  if (isCheckedIn(ev)) {
    banner.hidden = true;
    endBtn.hidden = false;
    countdownWrap.style.display = "";
    return;
  }

  const startMs = new Date(ev.start).getTime();
  const deadlineMs = startMs + CONFIG.CHECKIN_WINDOW_MIN * 60_000;
  const remaining = deadlineMs - now.getTime();

  if (remaining > 0) {
    // Check-in pending: show banner, hide countdown + end btn to focus attention
    banner.hidden = false;
    endBtn.hidden = true;
    countdownWrap.style.display = "none";
    const mm = String(Math.floor(remaining / 60000)).padStart(2, "0");
    const ss = String(Math.floor((remaining % 60000) / 1000)).padStart(2, "0");
    sub.textContent = `Se cancelará en ${mm}:${ss} si no confirmas`;
  } else {
    // Window elapsed. The server-side cron (every 5 min) is the main mechanism.
    // Client-side fallback: auto-cancel, but with guardrails against races.
    banner.hidden = true;

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
  const list = document.getElementById("events-list");
  const now = new Date();

  // Sub-header with today's summary
  renderTimeline(now);

  // Hide past events (end already passed)
  const visible = state.events.filter(ev => new Date(ev.end) > now);

  if (!visible.length) {
    list.innerHTML = `<div class="events-empty">No hay eventos próximos</div>`;
    return;
  }

  const today = dayKey(now);
  const tomorrow = dayKey(new Date(now.getTime() + 86_400_000));

  const groups = new Map();
  for (const ev of visible) {
    const k = dayKey(new Date(ev.start));
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(ev);
  }

  const sortedKeys = [...groups.keys()].sort();
  const html = sortedKeys.map(key => {
    let header;
    if (key === today) header = "HOY";
    else if (key === tomorrow) header = "MAÑANA";
    else header = groupHeaderFromKey(key);

    const items = groups.get(key)
      .sort((a, b) => new Date(a.start) - new Date(b.start))
      .map(ev => renderEventItem(ev, now))
      .join("");

    return `<div class="events-group-header">${escapeHtml(header)}</div>${items}`;
  }).join("");

  list.innerHTML = html;
}

function renderEventItem(ev, now) {
  const s = new Date(ev.start), e = new Date(ev.end);
  const isCurrent = s <= now && now < e;
  const timeFmt = { hour: "2-digit", minute: "2-digit", hour12: false };
  const sStr = s.toLocaleTimeString(CONFIG.LOCALE, timeFmt);
  const eStr = e.toLocaleTimeString(CONFIG.LOCALE, timeFmt);
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
function renderTimeline(now = new Date()) {
  const track = document.getElementById("timeline-track");
  const hoursEl = document.getElementById("timeline-hours");
  const subEl = document.getElementById("events-header-sub");
  if (!track || !hoursEl) return;

  const startHour = CONFIG.DAY_START_HOUR;
  const endHour = CONFIG.DAY_END_HOUR;
  const startMin = startHour * 60;
  const endMin = endHour * 60;
  const rangeMin = endMin - startMin;

  // Hour labels every 2h
  hoursEl.innerHTML = "";
  for (let h = startHour; h <= endHour; h += 2) {
    const s = document.createElement("span");
    s.textContent = `${String(h).padStart(2, "0")}:00`;
    hoursEl.appendChild(s);
  }

  // Keep only the "now" element; wipe blocks
  track.querySelectorAll(".timeline-block").forEach(b => b.remove());

  // Today's events
  const todayKey = dayKey(now);
  const today = state.events
    .filter(ev => dayKey(new Date(ev.start)) === todayKey)
    .sort((a, b) => new Date(a.start) - new Date(b.start));

  let busyMinutes = 0;
  for (const ev of today) {
    const s = new Date(ev.start), e = new Date(ev.end);
    const sMin = s.getHours() * 60 + s.getMinutes();
    const eMin = e.getHours() * 60 + e.getMinutes();
    const clampedStart = Math.max(sMin, startMin);
    const clampedEnd = Math.min(eMin, endMin);
    if (clampedEnd <= clampedStart) continue;

    busyMinutes += clampedEnd - clampedStart;

    const leftPct = ((clampedStart - startMin) / rangeMin) * 100;
    const widthPct = ((clampedEnd - clampedStart) / rangeMin) * 100;

    const block = document.createElement("div");
    block.className = "timeline-block";
    const isCurrent = s <= now && now < e;
    if (isCurrent) block.classList.add("timeline-block--current");
    block.style.left = `${leftPct}%`;
    block.style.width = `${widthPct}%`;
    block.title = `${ev.title} · ${fmtTime(ev.start)}–${fmtTime(ev.end)}`;
    block.dataset.eventId = ev.id;
    track.appendChild(block);
  }

  // "Now" indicator
  const nowEl = document.getElementById("timeline-now");
  const nowMin = now.getHours() * 60 + now.getMinutes();
  if (nowMin < startMin || nowMin > endMin) {
    nowEl.hidden = true;
  } else {
    nowEl.hidden = false;
    nowEl.style.left = `${((nowMin - startMin) / rangeMin) * 100}%`;
  }

  // Sub-header summary
  if (subEl) {
    const busyH = Math.floor(busyMinutes / 60);
    const busyM = busyMinutes % 60;
    const txt = busyMinutes
      ? `${today.length} ${today.length === 1 ? "reunión" : "reuniones"} hoy · ${busyH ? busyH + "h " : ""}${busyM}min ocupada`
      : "Sin reuniones hoy";
    subEl.textContent = txt;
  }
}

/**
 * Click on empty timeline region → open custom book modal with that start time.
 */
function handleTimelineClick(e) {
  // Ignore clicks on busy blocks
  if (e.target.closest(".timeline-block")) return;
  const track = document.getElementById("timeline-track");
  const rect = track.getBoundingClientRect();
  const x = e.clientX - rect.left;
  const pct = Math.max(0, Math.min(1, x / rect.width));
  const startMin = CONFIG.DAY_START_HOUR * 60;
  const endMin = CONFIG.DAY_END_HOUR * 60;
  const totalMin = startMin + pct * (endMin - startMin);
  // Round to nearest 15 min
  const rounded = Math.round(totalMin / 15) * 15;
  const h = Math.floor(rounded / 60);
  const m = rounded % 60;

  const target = new Date();
  target.setHours(h, m, 0, 0);

  // If the target is in the past, snap to the next 15-min slot from now
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
  const date = new Date(y, m - 1, d);
  return date.toLocaleDateString(CONFIG.LOCALE, {
    weekday: "long", day: "numeric", month: "long",
  }).toUpperCase();
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
    setSelectedDate(defaultStart);
    document.getElementById("book-time").value = toTimeInput(defaultStart);

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
function mountPeoplePicker(containerId, { selectable, onPick }) {
  const host = document.getElementById(containerId);
  host.innerHTML = "";

  const people = [...CONFIG.PEOPLE]
    .map(s => s.trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));
  const initials = [...new Set(people.map(n => n[0].toUpperCase()))].sort();

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
    if (!selected) { selectedChip.hidden = true; return; }
    selectedChip.hidden = false;
    selectedChip.innerHTML = "";
    const label = document.createElement("span");
    label.className = "people-selected-label";
    label.textContent = "Seleccionado:";
    const name = document.createElement("strong");
    name.textContent = selected;
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
    selectedChip.appendChild(x);
  }

  function render() {
    grid.innerHTML = "";
    const q = activeQuery.toLowerCase();
    const filtered = people.filter(n => {
      if (activeLetter && !n.toUpperCase().startsWith(activeLetter)) return false;
      if (q && !n.toLowerCase().includes(q)) return false;
      return true;
    });
    if (!filtered.length) {
      const empty = document.createElement("div");
      empty.className = "people-empty";
      empty.textContent = "Sin resultados";
      grid.appendChild(empty);
      return;
    }
    for (const name of filtered) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = name;
      b.dataset.person = name;
      if (selectable && selected === name) b.classList.add("selected");
      b.addEventListener("click", () => {
        if (selectable) {
          selected = name;
          renderSelectedChip();
          grid.querySelectorAll("button").forEach(x =>
            x.classList.toggle("selected", x.dataset.person === name)
          );
        }
        onPick(name);
      });
      grid.appendChild(b);
    }
  }

  search.addEventListener("input", (e) => {
    activeQuery = e.target.value;
    render();
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
  return state.events.find(ev => {
    const s = new Date(ev.start), e = new Date(ev.end);
    return s < end && start < e;
  });
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
  toast(`Reservado ${mins} min — ${person}`, "success");
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
  return new Date(iso).toLocaleTimeString(CONFIG.LOCALE, {
    hour: "2-digit", minute: "2-digit", hour12: false,
  });
}

function formatMins(mins) {
  if (mins < 60) return `${mins} min`;
  const h = Math.floor(mins / 60), m = mins % 60;
  if (m === 0) return h === 1 ? "1 hora" : `${h} horas`;
  return `${h}h ${m}min`;
}

function fmtDateTime(iso) {
  const d = new Date(iso);
  const date = d.toLocaleDateString(CONFIG.LOCALE, { day: "numeric", month: "short" });
  return `${date} ${fmtTime(iso)}`;
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
async function loadEvents() {
  try {
    const events = await fetchEvents();
    state.events = events;
    state.lastFetch = Date.now();
    state.lastFetchOk = Date.now();
    state.fetchFailing = false;
    renderEventsList();
    renderStatus();
    updateOfflineBadge();
  } catch (err) {
    console.error("Error cargando eventos:", err);
    state.fetchFailing = true;
    updateOfflineBadge();
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
      endTimeInput.value = toTimeInput(end);
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

  // Prevent tablet screen sleep hint (needs user interaction on most browsers)
  if ("wakeLock" in navigator) {
    document.addEventListener("click", async function requestWakeOnce() {
      try { await navigator.wakeLock.request("screen"); } catch (_) {}
      document.removeEventListener("click", requestWakeOnce);
    });
  }
}

document.addEventListener("DOMContentLoaded", init);
