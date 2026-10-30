/*
 * ICS Import Card for Home Assistant
 * Pick (or drop) an .ics file, preview the events it contains and add the
 * ones you want to any existing Home Assistant calendar that supports creating
 * events (Local Calendar, Google, CalDAV, ...).
 *
 * Standalone: no dependency on any other card.
 */
const CARD_VERSION = "0.1.0";

const FEATURE_CREATE = 1;

const I18N = {
  en: {
    title: "Import an .ics file",
    drop: "Drop an .ics file here or",
    choose: "Choose a file",
    no_cal: "No calendar that accepts new events was found.",
    not_ics: "This file is not a valid iCalendar (.ics) file.",
    no_events: "No events found in this file.",
    found: (n) => `${n} event${n > 1 ? "s" : ""} found`,
    all: "Select all",
    target: "Add to calendar",
    cancel: "Cancel",
    import: (n) => `Import (${n})`,
    importing: (i, n) => `Importing ${i}/${n}…`,
    done: (ok, ko) => `${ok} event${ok > 1 ? "s" : ""} added${ko ? `, ${ko} failed` : ""}.`,
    another: "Import another file",
    untitled: "(untitled)",
    all_day: "All day",
    repeats: "Repeats",
    exceptions: "Changed or skipped occurrences will not be imported",
    override: "Change to a single occurrence",
    cancelled: "Cancelled",
  },
  fr: {
    title: "Importer un fichier .ics",
    drop: "Glissez un fichier .ics ici ou",
    choose: "Choisir un fichier",
    no_cal: "Aucun calendrier acceptant la création d'évènements n'a été trouvé.",
    not_ics: "Ce fichier n'est pas un fichier iCalendar (.ics) valide.",
    no_events: "Aucun évènement trouvé dans ce fichier.",
    found: (n) => `${n} évènement${n > 1 ? "s" : ""} trouvé${n > 1 ? "s" : ""}`,
    all: "Tout sélectionner",
    target: "Ajouter au calendrier",
    cancel: "Annuler",
    import: (n) => `Importer (${n})`,
    importing: (i, n) => `Import ${i}/${n}…`,
    done: (ok, ko) => `${ok} évènement${ok > 1 ? "s" : ""} ajouté${ok > 1 ? "s" : ""}${ko ? `, ${ko} en échec` : ""}.`,
    another: "Importer un autre fichier",
    untitled: "(sans titre)",
    all_day: "Toute la journée",
    repeats: "Répétition",
    exceptions: "Les occurrences modifiées ou supprimées ne seront pas importées",
    override: "Modification d'une seule occurrence",
    cancelled: "Annulé",
  },
};

// Outlook / Exchange write Windows zone names in TZID.
const WINDOWS_ZONES = {
  "Romance Standard Time": "Europe/Paris",
  "W. Europe Standard Time": "Europe/Berlin",
  "Central Europe Standard Time": "Europe/Budapest",
  "Central European Standard Time": "Europe/Warsaw",
  "GMT Standard Time": "Europe/London",
  "Greenwich Standard Time": "Atlantic/Reykjavik",
  "E. Europe Standard Time": "Europe/Chisinau",
  "FLE Standard Time": "Europe/Kiev",
  "Eastern Standard Time": "America/New_York",
  "Central Standard Time": "America/Chicago",
  "Mountain Standard Time": "America/Denver",
  "Pacific Standard Time": "America/Los_Angeles",
  "UTC": "UTC",
  "Coordinated Universal Time": "UTC",
};

const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const pad = (n) => String(n).padStart(2, "0");
const dayKey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const addDaysKey = (key, n) => {
  const [y, m, d] = key.split("-").map(Number);
  const u = new Date(Date.UTC(y, m - 1, d + n));
  return `${u.getUTCFullYear()}-${pad(u.getUTCMonth() + 1)}-${pad(u.getUTCDate())}`;
};
// ISO 8601 with the browser's UTC offset, e.g. 2026-11-05T10:00:00+01:00
const isoOffset = (d) => {
  const o = -d.getTimezoneOffset();
  const sign = o >= 0 ? "+" : "-";
  return `${dayKey(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}${sign}${pad(Math.floor(Math.abs(o) / 60))}:${pad(Math.abs(o) % 60)}`;
};

// ---------------------------------------------------------------- .ics parsing

const unescapeText = (v) => v.replace(/\\([\\;,nN])/g, (_, c) => (c === "n" || c === "N" ? "\n" : c));

// Split on `sep` but not inside double quotes.
const splitOutsideQuotes = (s, sep, max = Infinity) => {
  const out = [];
  let cur = "";
  let q = false;
  for (const ch of s) {
    if (ch === '"') q = !q;
    if (ch === sep && !q && out.length < max - 1) { out.push(cur); cur = ""; } else cur += ch;
  }
  out.push(cur);
  return out;
};

const parseLine = (line) => {
  const [head, value = ""] = splitOutsideQuotes(line, ":", 2);
  const [name, ...rawParams] = splitOutsideQuotes(head, ";");
  const params = {};
  rawParams.forEach((p) => {
    const i = p.indexOf("=");
    if (i > 0) params[p.slice(0, i).toUpperCase()] = p.slice(i + 1).replace(/^"|"$/g, "");
  });
  return { name: name.toUpperCase(), params, value };
};

// Offset (ms) of time zone `tz` at instant `ms`.
const tzOffset = (ms, tz) => {
  const parts = {};
  new Intl.DateTimeFormat("en-US", {
    timeZone: tz, hourCycle: "h23", year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit",
  }).formatToParts(new Date(ms)).forEach((p) => { parts[p.type] = Number(p.value); });
  const asUtc = Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second);
  return asUtc - (ms - (ms % 1000));
};

// Wall-clock time in zone `tz` -> Date. Throws on an unknown zone.
const zonedDate = (y, mo, d, h, mi, s, tz) => {
  const guess = Date.UTC(y, mo - 1, d, h, mi, s);
  const off = tzOffset(guess, tz);
  let t = guess - off;
  const off2 = tzOffset(t, tz);
  if (off2 !== off) t = guess - off2;
  return new Date(t);
};

// Returns { allDay: true, value: "YYYY-MM-DD" } or { allDay: false, value: Date }, or null.
const parseDateValue = (value, params) => {
  const v = value.trim();
  let m = v.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return { allDay: true, value: `${m[1]}-${m[2]}-${m[3]}` };
  m = v.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})?(Z)?$/);
  if (!m) return null;
  const [y, mo, d, h, mi, s] = [m[1], m[2], m[3], m[4], m[5], m[6] || "0"].map(Number);
  if (m[7]) return { allDay: false, value: new Date(Date.UTC(y, mo - 1, d, h, mi, s)) };
  const tzid = params.TZID && (WINDOWS_ZONES[params.TZID] || params.TZID.replace(/^\/+/, ""));
  if (tzid) {
    try { return { allDay: false, value: zonedDate(y, mo, d, h, mi, s, tzid) }; } catch (e) { /* unknown zone: floating */ }
  }
  return { allDay: false, value: new Date(y, mo - 1, d, h, mi, s) };
};

// "P1DT2H" -> { days, ms }
const parseDuration = (v) => {
  const m = v.trim().match(/^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/);
  if (!m) return null;
  const sign = m[1] === "-" ? -1 : 1;
  const days = sign * ((Number(m[2] || 0) * 7) + Number(m[3] || 0));
  const ms = sign * ((Number(m[4] || 0) * 3600 + Number(m[5] || 0) * 60 + Number(m[6] || 0)) * 1000);
  return { days, ms };
};

const parseICS = (text) => {
  const raw = String(text || "").replace(/^﻿/, "");
  if (!/BEGIN:VCALENDAR/i.test(raw)) throw new Error("not an iCalendar file");
  const lines = raw.replace(/\r?\n[ \t]/g, "").split(/\r?\n/).filter(Boolean);
  const events = [];
  let cur = null;
  let depth = 0; // nested components inside a VEVENT (VALARM, ...)
  for (const line of lines) {
    const { name, params, value } = parseLine(line);
    if (name === "BEGIN") {
      if (cur) depth++;
      else if (value.toUpperCase() === "VEVENT") cur = { props: {}, exceptions: false };
      continue;
    }
    if (name === "END") {
      if (cur && depth) depth--;
      else if (cur && value.toUpperCase() === "VEVENT") { events.push(cur); cur = null; }
      continue;
    }
    if (!cur || depth) continue;
    if (name === "EXDATE" || name === "RDATE") cur.exceptions = true;
    else if (!cur.props[name]) cur.props[name] = { params, value };
  }

  return events.map((c) => {
    const p = c.props;
    const start = p.DTSTART && parseDateValue(p.DTSTART.value, p.DTSTART.params);
    if (!start) return null;
    let end = p.DTEND && parseDateValue(p.DTEND.value, p.DTEND.params);
    if (end && end.allDay !== start.allDay) end = null;
    let endValue = end?.value;
    if (!endValue && p.DURATION) {
      const dur = parseDuration(p.DURATION.value);
      if (dur) endValue = start.allDay ? addDaysKey(start.value, dur.days || 1) : new Date(start.value.getTime() + dur.days * 864e5 + dur.ms);
    }
    if (!endValue || endValue <= start.value) {
      endValue = start.allDay ? addDaysKey(start.value, 1) : new Date(start.value.getTime() + 3600e3);
    }
    const text = (k) => (p[k] ? unescapeText(p[k].value).trim() : "");
    const override = !!p["RECURRENCE-ID"];
    const cancelled = text("STATUS").toUpperCase() === "CANCELLED";
    return {
      uid: text("UID"),
      summary: text("SUMMARY"),
      allDay: start.allDay,
      start: start.value,
      end: endValue,
      location: text("LOCATION"),
      description: text("DESCRIPTION"),
      rrule: p.RRULE ? p.RRULE.value.trim() : "",
      exceptions: c.exceptions,
      override,
      cancelled,
      selected: !override && !cancelled,
    };
  }).filter(Boolean);
};

// ---------------------------------------------------------------- card

class IcsImportCard extends HTMLElement {
  static parseICS(text) { return parseICS(text); }

  static getStubConfig() { return {}; }

  setConfig(config) {
    if (config.calendars && !Array.isArray(config.calendars)) throw new Error("calendars must be a list of calendar entities");
    this._config = config || {};
    this._state = this._state || { phase: "idle" };
    if (this._hass) this._render();
  }

  set hass(hass) {
    const first = !this._hass;
    this._hass = hass;
    const sig = this._calendars().map((c) => `${c.entity}:${c.name}`).join("|");
    if (first || sig !== this._calSig) { this._calSig = sig; this._render(); }
  }

  getCardSize() { return 3; }

  _t(key, ...args) {
    const lang = (this._hass?.locale?.language || this._hass?.language || "en").slice(0, 2);
    const v = (I18N[lang] || I18N.en)[key] ?? I18N.en[key];
    return typeof v === "function" ? v(...args) : v;
  }

  _lang() { return this._hass?.locale?.language || this._hass?.language || "en"; }

  _calendars() {
    if (!this._hass) return [];
    const wanted = this._config?.calendars;
    const ids = wanted?.length ? wanted : Object.keys(this._hass.states).filter((id) => id.startsWith("calendar."));
    return ids
      .map((id) => this._hass.states[id])
      .filter((st) => st && (st.attributes.supported_features & FEATURE_CREATE))
      .map((st) => ({ entity: st.entity_id, name: st.attributes.friendly_name || st.entity_id }));
  }

  _when(ev) {
    const lang = this._lang();
    if (ev.allDay) {
      const fmt = new Intl.DateTimeFormat(lang, { weekday: "short", day: "numeric", month: "short", year: "numeric" });
      const toDate = (k) => { const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d); };
      const last = addDaysKey(ev.end, -1);
      const s = fmt.format(toDate(ev.start));
      return last === ev.start ? `${s} · ${this._t("all_day")}` : `${s} → ${fmt.format(toDate(last))}`;
    }
    const day = new Intl.DateTimeFormat(lang, { weekday: "short", day: "numeric", month: "short", year: "numeric" });
    const time = new Intl.DateTimeFormat(lang, { hour: "2-digit", minute: "2-digit" });
    const sameDay = dayKey(ev.start) === dayKey(ev.end);
    return sameDay
      ? `${day.format(ev.start)} · ${time.format(ev.start)} – ${time.format(ev.end)}`
      : `${day.format(ev.start)} ${time.format(ev.start)} → ${day.format(ev.end)} ${time.format(ev.end)}`;
  }

  _render() {
    if (!this.shadowRoot) {
      this.attachShadow({ mode: "open" });
      this.shadowRoot.addEventListener("click", (e) => this._onClick(e));
      this.shadowRoot.addEventListener("change", (e) => this._onChange(e));
      this.shadowRoot.addEventListener("dragover", (e) => { e.preventDefault(); this.shadowRoot.querySelector(".drop")?.classList.add("over"); });
      this.shadowRoot.addEventListener("dragleave", () => this.shadowRoot.querySelector(".drop")?.classList.remove("over"));
      this.shadowRoot.addEventListener("drop", (e) => { e.preventDefault(); const f = e.dataTransfer?.files?.[0]; if (f) this._readFile(f); });
    }
    const st = this._state;
    const cals = this._calendars();
    if (st.cal && !cals.some((c) => c.entity === st.cal)) st.cal = null;
    if (!st.cal && cals.length) st.cal = cals[0].entity;
    let body;
    if (!cals.length) {
      body = `<div class="msg">${esc(this._t("no_cal"))}</div>`;
    } else if (st.phase === "preview" || st.phase === "importing") {
      const busy = st.phase === "importing";
      const n = st.events.filter((e) => e.selected).length;
      body = `
        <div class="file"><ha-icon icon="mdi:file-outline"></ha-icon><span>${esc(st.fileName)}</span><span class="sp"></span><span class="muted">${esc(this._t("found", st.events.length))}</span></div>
        <label class="row all"><input type="checkbox" data-all ${n === st.events.length ? "checked" : ""} ${busy ? "disabled" : ""}><span>${esc(this._t("all"))}</span></label>
        <div class="list">${st.events.map((ev, i) => `
          <label class="row ${ev.selected ? "" : "off"}">
            <input type="checkbox" data-idx="${i}" ${ev.selected ? "checked" : ""} ${busy ? "disabled" : ""}>
            <span class="ev">
              <span class="ev-title">${esc(ev.summary || this._t("untitled"))}${ev.rrule ? ` <ha-icon class="mini" icon="mdi:repeat" title="${esc(this._t("repeats"))}: ${esc(ev.rrule)}"></ha-icon>` : ""}</span>
              <span class="muted">${esc(this._when(ev))}</span>
              ${ev.location ? `<span class="muted"><ha-icon class="mini" icon="mdi:map-marker-outline"></ha-icon>${esc(ev.location)}</span>` : ""}
              ${ev.exceptions ? `<span class="warn">${esc(this._t("exceptions"))}</span>` : ""}
              ${ev.override ? `<span class="warn">${esc(this._t("override"))}</span>` : ""}
              ${ev.cancelled ? `<span class="warn">${esc(this._t("cancelled"))}</span>` : ""}
            </span>
          </label>`).join("")}</div>
        <label class="target"><span>${esc(this._t("target"))}</span>
          <select data-cal ${busy ? "disabled" : ""}>${cals.map((c) => `<option value="${esc(c.entity)}" ${c.entity === st.cal ? "selected" : ""}>${esc(c.name)}</option>`).join("")}</select>
        </label>
        ${st.error ? `<div class="err">${esc(st.error)}</div>` : ""}
        <div class="foot">
          ${busy ? `<span class="muted">${esc(this._t("importing", st.progress, n))}</span>` : `<button class="btn" data-action="reset">${esc(this._t("cancel"))}</button>`}
          <button class="btn primary" data-action="import" ${busy || !n ? "disabled" : ""}>${esc(this._t("import", n))}</button>
        </div>`;
    } else if (st.phase === "done") {
      const ko = st.results.filter((r) => r.error);
      body = `
        <div class="msg ${ko.length ? "" : "ok"}"><ha-icon icon="${ko.length ? "mdi:alert-circle-outline" : "mdi:check-circle-outline"}"></ha-icon>${esc(this._t("done", st.results.length - ko.length, ko.length))}</div>
        ${ko.length ? `<ul class="fails">${ko.map((r) => `<li><b>${esc(r.ev.summary || this._t("untitled"))}</b> – ${esc(r.error)}</li>`).join("")}</ul>` : ""}
        <div class="foot"><button class="btn primary" data-action="reset">${esc(this._t("another"))}</button></div>`;
    } else {
      body = `
        <div class="drop">
          <ha-icon icon="mdi:calendar-import"></ha-icon>
          <span>${esc(this._t("drop"))}</span>
          <button class="btn primary" data-action="pick">${esc(this._t("choose"))}</button>
          <input type="file" accept=".ics,text/calendar" hidden>
        </div>
        ${st.error ? `<div class="err">${esc(st.error)}</div>` : ""}`;
    }
    this.shadowRoot.innerHTML = `
      <style>
        ha-card { padding: 16px; }
        .head { font-size: 1.15em; font-weight: 500; margin-bottom: 12px; }
        .drop { display: flex; flex-direction: column; align-items: center; gap: 10px; padding: 24px 12px; border: 2px dashed var(--divider-color); border-radius: 12px; text-align: center; color: var(--secondary-text-color); }
        .drop.over { border-color: var(--primary-color); background: rgba(var(--rgb-primary-color, 3,169,244), .06); }
        .drop > ha-icon { --mdc-icon-size: 40px; color: var(--primary-color); }
        .btn { font: inherit; border: 1px solid var(--divider-color); background: none; color: var(--primary-text-color); border-radius: 18px; padding: 7px 16px; cursor: pointer; }
        .btn.primary { background: var(--primary-color); border-color: var(--primary-color); color: var(--text-primary-color, #fff); }
        .btn:disabled { opacity: .5; cursor: default; }
        .file { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
        .file span:nth-child(2) { font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .sp { flex: 1; }
        .muted { color: var(--secondary-text-color); font-size: .9em; }
        .list { max-height: 360px; overflow-y: auto; border-top: 1px solid var(--divider-color); border-bottom: 1px solid var(--divider-color); }
        .row { display: flex; align-items: flex-start; gap: 10px; padding: 8px 2px; cursor: pointer; }
        .row + .row { border-top: 1px solid var(--divider-color); }
        .row.all { padding-top: 0; }
        .row.off .ev { opacity: .55; }
        .row input { margin-top: 3px; accent-color: var(--primary-color); }
        .ev { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
        .ev-title { font-weight: 500; }
        .mini { --mdc-icon-size: 15px; vertical-align: -2px; margin-right: 2px; }
        .warn { color: var(--warning-color, #f9a825); font-size: .85em; }
        .target { display: flex; align-items: center; gap: 10px; margin-top: 12px; flex-wrap: wrap; }
        .target select { flex: 1; min-width: 160px; font: inherit; padding: 6px 8px; border-radius: 8px; border: 1px solid var(--divider-color); background: var(--card-background-color); color: var(--primary-text-color); }
        .foot { display: flex; justify-content: flex-end; align-items: center; gap: 8px; margin-top: 14px; }
        .msg { display: flex; align-items: center; gap: 8px; }
        .msg.ok ha-icon { color: var(--success-color, #43a047); }
        .err, .fails { color: var(--error-color, #db4437); margin-top: 10px; }
        .fails { padding-left: 18px; }
      </style>
      <ha-card>
        <div class="head">${esc(this._config?.title || this._t("title"))}</div>
        ${body}
      </ha-card>`;
  }

  _onClick(e) {
    const act = e.target.closest("[data-action]")?.dataset.action;
    if (act === "pick") this.shadowRoot.querySelector('input[type="file"]')?.click();
    else if (act === "reset") { this._state = { phase: "idle", cal: this._state.cal }; this._render(); }
    else if (act === "import") this._import();
  }

  _onChange(e) {
    const t = e.target;
    const st = this._state;
    if (t.type === "file") { if (t.files?.[0]) this._readFile(t.files[0]); return; }
    if (t.hasAttribute("data-cal")) { st.cal = t.value; return; }
    if (t.hasAttribute("data-all")) st.events.forEach((ev) => { ev.selected = t.checked; });
    else if (t.dataset.idx !== undefined) st.events[Number(t.dataset.idx)].selected = t.checked;
    else return;
    const scroll = this.shadowRoot.querySelector(".list")?.scrollTop;
    this._render();
    const list = this.shadowRoot.querySelector(".list");
    if (list && scroll) list.scrollTop = scroll;
  }

  async _readFile(file) {
    const cal = this._state.cal;
    try {
      const events = parseICS(await file.text());
      this._state = events.length
        ? { phase: "preview", fileName: file.name, events, cal }
        : { phase: "idle", error: this._t("no_events"), cal };
    } catch (err) {
      this._state = { phase: "idle", error: this._t("not_ics"), cal };
    }
    this._render();
  }

  _payload(ev) {
    const event = { summary: ev.summary || this._t("untitled") };
    if (ev.allDay) { event.dtstart = ev.start; event.dtend = ev.end; }
    else { event.dtstart = isoOffset(ev.start); event.dtend = isoOffset(ev.end); }
    if (ev.location) event.location = ev.location;
    if (ev.description) event.description = ev.description;
    if (ev.rrule) event.rrule = ev.rrule;
    return event;
  }

  async _import() {
    const st = this._state;
    const todo = st.events.filter((e) => e.selected);
    if (!todo.length || !st.cal) return;
    st.phase = "importing";
    st.error = "";
    const results = [];
    for (const [i, ev] of todo.entries()) {
      st.progress = i + 1;
      this._render();
      try {
        await this._hass.callWS({ type: "calendar/event/create", entity_id: st.cal, event: this._payload(ev) });
        results.push({ ev });
      } catch (err) {
        results.push({ ev, error: err?.message || String(err) });
      }
    }
    this._state = { phase: "done", results, cal: st.cal };
    this._render();
  }
}

if (!customElements.get("ics-import-card")) {
  customElements.define("ics-import-card", IcsImportCard);
}
window.customCards = window.customCards || [];
if (!window.customCards.some((c) => c.type === "ics-import-card")) {
  window.customCards.push({
    type: "ics-import-card",
    name: "ICS Import Card",
    description: "Import the events of an .ics file into an existing Home Assistant calendar.",
    preview: false,
  });
}
console.info(`%c ICS-IMPORT-CARD %c v${CARD_VERSION} `,
  "color:#fff;background:#03a9f4;font-weight:700;border-radius:4px 0 0 4px",
  "color:#03a9f4;background:#e1f5fe;font-weight:700;border-radius:0 4px 4px 0");
