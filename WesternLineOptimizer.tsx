"use client";

/**
 * Mumbai Western Line — Local Train Route Optimizer
 * Single-file Next.js client component.
 *
 * Setup:
 *   npm i xlsx
 *     (SheetJS recommends its own CDN build for current security fixes:
 *      npm i https://cdn.sheetjs.com/xlsx-latest/xlsx-latest.tgz)
 *   Optional: copy western-line-train-template.xlsx into /public so the
 *   "Download template" button serves the formatted file. Without it, the app
 *   generates a simpler template in the browser.
 *
 * Usage:  app/page.tsx  →  export { default } from "./WesternLineOptimizer";
 *
 * Data model: MAIN trains (Andheri→Virar, or direct Bandra→Virar) carry a Train ID.
 * FEEDER trains (Bandra→Andheri) carry a Reference ID naming the main train(s)
 * they connect to. Only linked pairs are ever offered; nothing is guessed.
 */

import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

/* ============================================================================
 * ENGINE START  (pure logic: parsing + ranking, no React)
 * ========================================================================== */

type Speed = "slow" | "fast" | "double-fast";

interface Train {
  uid: string;
  id: string; //                label shown to the user, may be empty
  from: string;
  to: string;
  dep: number; //               minutes since midnight
  arr: number; //               may exceed 1440 for overnight runs
  speed: Speed;
  depPlatform: string | null;
  arrPlatform: string | null;
  lateMin: number; //           average minutes late on arrival
  notes: string;
  estimated: boolean; //        arrival time was estimated (legacy files only)
}

/** A feeder train that the user declared as a way to reach a main train. */
interface Link {
  feederUid: string;
  mainUid: string;
}

interface Settings {
  origin: string;
  destination: string;
  minBufferMin: number; //          minimum change time, same platform
  platformChangeMin: number; //     extra time when the platform differs
}

interface TimetableData {
  trains: Train[];
  links: Link[];
}

interface ParsedData extends TimetableData {
  settings: Settings;
  skipped: string[]; //             problems found while reading (skipped rows and warnings)
  format: "template" | "legacy";
}

interface Journey {
  id: string;
  kind: "direct" | "switch";
  transferStation: string;
  legs: Train[];
  boardMin: number;
  scheduledArrMin: number;
  expectedArrMin: number; //        scheduled + average lateness (when enabled)
  totalMin: number; //              expected arrival − boarding
  waitMin: number; //               time at the transfer station (0 for direct)
  requiredMin: number; //           connection time needed at the transfer station
  platformChange: boolean | null; // null = unknown or not applicable
}

interface SearchResult {
  options: Journey[];
  best: Journey | null;
  queryMin: number;
  travelMin: number;
  readyMin: number;
  atStation: boolean;
  useDelays: boolean;
}

const SEARCH_DELAY_MS = 2000;
const MAX_LIST = 100;
const TEMPLATE_URL = "/western-line-train-template.xlsx";
const TEMPLATE_FILENAME = "western-line-train-template.xlsx";

const DEFAULT_SETTINGS: Settings = {
  origin: "Bandra",
  destination: "Virar",
  minBufferMin: 3,
  platformChangeMin: 4,
};

/** Estimated Bandra → transfer-station ride (min). Only used for old-layout files that lack leg-1 arrival times. */
const DEFAULT_LEG1_RIDE_MIN: Record<Speed, number> = { slow: 15, fast: 11, "double-fast": 10 };
const LEG1_RIDE_MIN: Record<string, Record<Speed, number>> = {
  andheri: { slow: 14, fast: 10, "double-fast": 9 },
  dadar: { slow: 13, fast: 9, "double-fast": 8 },
};

/** Built-in sample in the original layout: [station, main departure, main arrival, ...feeder departures]. */
const LEGACY_SAMPLE: unknown[][] = [
  ["Andheri", "*6:10 PM", "6:59 PM", "5:27 PM", "5:30 PM", "5:34 PM", "5:37 PM", "*5:46 PM"],
  ["Andheri", "**6:27 PM", "7:20 PM", "5:37 PM", "5:40 PM", "*5:46 PM", "5:50 PM", "6:00 PM", "6:03 PM"],
  ["Bandra", "6:06 PM", "7:07 PM"],
  ["Bandra", "6:24 PM", "7:36 PM"],
  ["Bandra", "6:33 PM", "7:45 PM"],
  ["Andheri", "*6:52 PM", "7:49 PM", "6:06 PM", "6:13 PM", "6:15 PM", "*6:16 PM", "6:18 PM", "*6:24 PM"],
  ["Andheri", "*7:20 PM", "8:15 PM"],
  ["Dadar", "*7:23 PM", "8:34 PM", "**7:05 PM"],
];

/* ---------- small helpers ---------- */

const norm = (s: unknown): string => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
const stationKey = (s: string): string => s.trim().toLowerCase().replace(/\s+/g, " ");
const idKey = (s: string): string => s.trim().toLowerCase();
const isBlank = (v: unknown): boolean => v === null || v === undefined || String(v).trim() === "";

function speedFromStars(stars: number): Speed {
  return stars >= 2 ? "double-fast" : stars === 1 ? "fast" : "slow";
}

/** Accepts Excel time serials (0.79166…), "6:10 PM", "18:10", "18:10:00", and legacy "*6:10 PM" / "**6:27 PM". */
function parseTimeCell(cell: unknown): { minutes: number; stars: number } | null {
  if (isBlank(cell)) return null;
  if (typeof cell === "number") {
    if (!Number.isFinite(cell) || cell < 0) return null;
    const frac = cell - Math.floor(cell);
    if (cell >= 1 && frac === 0) return null; // a bare integer is not a time
    return { minutes: Math.round(frac * 1440) % 1440, stars: 0 };
  }
  const m = /^\s*(\*{0,2})\s*(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(AM|PM)?\s*$/i.exec(String(cell));
  if (!m) return null;
  let h = Number(m[2]);
  const min = Number(m[3]);
  if (min > 59) return null;
  const mer = m[4]?.toUpperCase();
  if (mer) {
    if (h < 1 || h > 12) return null;
    h = (h % 12) + (mer === "PM" ? 12 : 0);
  } else if (h > 23) {
    return null;
  }
  return { minutes: h * 60 + min, stars: m[1].length };
}

function parseSpeed(v: unknown): Speed | null {
  const s = String(v ?? "").trim().toLowerCase();
  if (s === "") return "slow";
  const c = s.replace(/[^a-z*]/g, "");
  if (c === "**" || c.includes("double") || c === "df" || c.includes("dfast") || c.includes("superfast")) return "double-fast";
  if (c === "*" || c === "f" || c.includes("fast")) return "fast";
  if (c === "s" || c.includes("slow") || c.includes("local")) return "slow";
  return null;
}

function cleanPlatform(v: unknown): string | null {
  if (isBlank(v)) return null;
  const s = String(v).trim().replace(/^(platform|pf)\s*/i, "");
  return s === "" ? null : s.toUpperCase();
}

/* ---------- header mapping for the template layout ---------- */

const HEADER_ALIASES: Record<string, string[]> = {
  id: ["trainid", "trainno", "trainnumber", "train", "service"],
  ref: ["referenceid", "refid", "reference", "parentid", "maintrainid", "maintrain", "connectsto", "feederof"],
  from: ["fromstation", "from", "origin", "departurestation", "startstation"],
  to: ["tostation", "to", "destination", "arrivalstation", "endstation"],
  dep: ["departuretime", "departure", "deptime", "starttime", "depart"],
  arr: ["arrivaltime", "arrival", "arrtime", "endtime", "arrive"],
  speed: ["speed", "traintype", "type"],
  depPlatform: ["departureplatform", "depplatform", "fromplatform", "boardingplatform", "platform"],
  arrPlatform: ["arrivalplatform", "arrplatform", "toplatform"],
  late: ["avglatemin", "avglate", "averagelate", "averagelatemin", "avgdelaymin", "avgdelay", "averagedelay", "delay", "late"],
  notes: ["traininfonotes", "traininfo", "notes", "info", "remarks", "note"],
};

function mapHeader(row: unknown[]): Record<string, number> {
  const map: Record<string, number> = {};
  const cells = row.map(norm);
  for (const [field, aliases] of Object.entries(HEADER_ALIASES)) {
    const idx = cells.findIndex((c) => aliases.includes(c));
    if (idx >= 0) map[field] = idx;
  }
  return map;
}

function sheetToRows(ws: XLSX.WorkSheet | undefined): { rows: unknown[][]; offset: number } {
  if (!ws || !ws["!ref"]) return { rows: [], offset: 0 };
  const offset = XLSX.utils.decode_range(ws["!ref"]).s.r;
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: null, blankrows: true });
  return { rows, offset };
}

function parseSettings(rows: unknown[][]): Settings {
  const s: Settings = { ...DEFAULT_SETTINGS };
  for (const r of rows) {
    const label = norm(r[0]);
    const value = r[1];
    if (isBlank(value)) continue;
    const num = Number(value);
    if (label.includes("origin")) s.origin = String(value).trim();
    else if (label.includes("destination")) s.destination = String(value).trim();
    else if (label.includes("buffer") && Number.isFinite(num)) s.minBufferMin = Math.max(0, num);
    else if (label.includes("platform") && Number.isFinite(num)) s.platformChangeMin = Math.max(0, num);
  }
  return s;
}

/* ---------- template layout: Train ID + Reference ID ---------- */

interface Cand {
  rowNo: number;
  refs: string[];
  train: Train;
  dropped: boolean;
}

function parseTemplateRows(rows: unknown[][], offset: number, s: Settings): TimetableData & { skipped: string[] } {
  let headerIdx = -1;
  let map: Record<string, number> = {};
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const m = mapHeader(rows[i]);
    if (m.from !== undefined && m.to !== undefined && m.dep !== undefined && m.arr !== undefined) {
      headerIdx = i;
      map = m;
      break;
    }
  }
  if (headerIdx < 0) {
    throw new Error(
      "Couldn't find the expected columns (From Station, To Station, Departure Time, Arrival Time). Download the template and paste your data into it."
    );
  }

  const skipped: string[] = [];
  const cell = (r: unknown[], f: string): unknown => (map[f] === undefined ? null : r[map[f]]);
  const cands: Cand[] = [];

  /* pass 1: read and validate each row on its own */
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.every(isBlank)) continue;
    const rowNo = offset + i + 1;
    const id = String(cell(r, "id") ?? "").trim();
    const refs = String(cell(r, "ref") ?? "")
      .split(/[,;/\n]+/)
      .map((x) => x.trim())
      .filter(Boolean);
    if (id.toLowerCase().startsWith("example") || refs.some((x) => x.toLowerCase().startsWith("example"))) continue;

    const from = String(cell(r, "from") ?? "").trim();
    const to = String(cell(r, "to") ?? "").trim();
    if (!from || !to) { skipped.push(`Row ${rowNo}: From Station and To Station are both required`); continue; }
    if (stationKey(from) === stationKey(to)) { skipped.push(`Row ${rowNo}: From and To are the same station`); continue; }

    const dep = parseTimeCell(cell(r, "dep"));
    const arr = parseTimeCell(cell(r, "arr"));
    if (!dep) { skipped.push(`Row ${rowNo}: Departure Time is missing or not a time`); continue; }
    if (!arr) { skipped.push(`Row ${rowNo}: Arrival Time is missing or not a time`); continue; }
    let arrMin = arr.minutes;
    if (arrMin < dep.minutes) arrMin += 1440; // overnight run
    if (arrMin === dep.minutes || arrMin - dep.minutes > 240) {
      skipped.push(`Row ${rowNo}: Arrival is ${arrMin === dep.minutes ? "the same as" : "over 4 hours after"} departure; check AM/PM`);
      continue;
    }

    const speedCell = cell(r, "speed");
    const speed = isBlank(speedCell) ? speedFromStars(dep.stars) : parseSpeed(speedCell);
    if (!speed) { skipped.push(`Row ${rowNo}: Speed "${String(speedCell)}" is not Slow, Fast or Double Fast`); continue; }

    const lateCell = cell(r, "late");
    const late = isBlank(lateCell) ? 0 : Number(lateCell);
    if (!Number.isFinite(late)) { skipped.push(`Row ${rowNo}: Avg Late must be a number of minutes`); continue; }

    cands.push({
      rowNo,
      refs,
      dropped: false,
      train: {
        uid: `t${rowNo}`,
        id,
        from,
        to,
        dep: dep.minutes,
        arr: arrMin,
        speed,
        depPlatform: cleanPlatform(cell(r, "depPlatform")),
        arrPlatform: cleanPlatform(cell(r, "arrPlatform")),
        lateMin: Math.max(0, late),
        notes: String(cell(r, "notes") ?? "").trim(),
        estimated: false,
      },
    });
  }

  /* pass 2: unique IDs, then resolve Reference IDs */
  const o = stationKey(s.origin);
  const d = stationKey(s.destination);
  const seen = new Set<string>();
  const mainById = new Map<string, Cand>();
  for (const c of cands) {
    if (!c.train.id) continue;
    const k = idKey(c.train.id);
    if (seen.has(k)) {
      skipped.push(`Row ${c.rowNo}: Train ID "${c.train.id}" is already used; IDs must be unique`);
      c.dropped = true;
      continue;
    }
    seen.add(k);
    if (c.refs.length === 0) mainById.set(k, c);
  }

  const mains = cands.filter((c) => !c.dropped && c.refs.length === 0);
  const feeders = cands.filter((c) => !c.dropped && c.refs.length > 0);

  const isDirect = (t: Train) => stationKey(t.from) === o && stationKey(t.to) === d;
  for (const m of mains) {
    if (isDirect(m.train)) continue;
    if (stationKey(m.train.to) !== d) {
      skipped.push(`Row ${m.rowNo}: main train must end at ${s.destination} (it ends at ${m.train.to})`);
      m.dropped = true;
    }
  }

  const links: Link[] = [];
  const hasFeeder = new Set<string>();
  const feederKept = new Set<string>();
  for (const f of feeders) {
    if (stationKey(f.train.from) !== o) {
      skipped.push(`Row ${f.rowNo}: feeder trains must start at ${s.origin} (this starts at ${f.train.from})`);
      continue;
    }
    for (const ref of f.refs) {
      const m = mainById.get(idKey(ref));
      if (!m) { skipped.push(`Row ${f.rowNo}: Reference ID "${ref}" doesn't match any main train`); continue; }
      if (m.dropped) { skipped.push(`Row ${f.rowNo}: main train "${ref}" was skipped, so this link is ignored`); continue; }
      if (isDirect(m.train)) { skipped.push(`Row ${f.rowNo}: "${ref}" is a direct train and needs no feeder`); continue; }
      if (stationKey(f.train.to) !== stationKey(m.train.from)) {
        skipped.push(`Row ${f.rowNo}: this train goes to ${f.train.to}, but main train "${ref}" leaves from ${m.train.from}`);
        continue;
      }
      links.push({ feederUid: f.train.uid, mainUid: m.train.uid });
      hasFeeder.add(m.train.uid);
      feederKept.add(f.train.uid);
    }
  }

  const trains: Train[] = [];
  for (const c of cands) {
    if (c.dropped) continue;
    const t = c.train;
    if (c.refs.length === 0) {
      if (isDirect(t) || hasFeeder.has(t.uid)) trains.push(t);
      else skipped.push(`Row ${c.rowNo}: main train ${t.id || `(${t.from} to ${t.to})`} has no feeder trains linked to it, so it can't be used`);
    } else if (feederKept.has(t.uid)) {
      trains.push(t);
    }
  }
  const rowOf = (m: string) => Number(/^Row (\d+)/.exec(m)?.[1] ?? 0);
  skipped.sort((a, b) => rowOf(a) - rowOf(b));
  return { trains, links, skipped };
}

/* ---------- old layout: [station, main dep, main arr, ...feeder times] — each row already links its feeders ---------- */

function legacyToData(rows: unknown[][], offset: number, s: Settings): TimetableData & { skipped: string[] } {
  const trains: Train[] = [];
  const links: Link[] = [];
  const skipped: string[] = [];

  rows.forEach((r, i) => {
    if (r.every(isBlank)) return;
    const rowNo = offset + i + 1;
    const station = String(r[0] ?? "").trim();
    const first = parseTimeCell(r[1]);
    const last = parseTimeCell(r[2]);
    if (!station || !first || !last || last.minutes <= first.minutes) {
      skipped.push(`Row ${rowNo}: missing or invalid times`);
      return;
    }
    const base = { id: "", depPlatform: null, arrPlatform: null, lateMin: 0, notes: "" };

    if (stationKey(station) === stationKey(s.origin)) {
      trains.push({ ...base, uid: `m${rowNo}`, from: s.origin, to: s.destination, dep: first.minutes, arr: last.minutes, speed: speedFromStars(first.stars), estimated: false });
      return;
    }
    const feederTimes = r.slice(3).map(parseTimeCell).filter((p): p is { minutes: number; stars: number } => p !== null);
    if (feederTimes.length === 0) {
      skipped.push(`Row ${rowNo}: no feeder trains listed, so it can't be used`);
      return;
    }
    const mainUid = `m${rowNo}`;
    trains.push({ ...base, uid: mainUid, from: station, to: s.destination, dep: first.minutes, arr: last.minutes, speed: speedFromStars(first.stars), estimated: false });
    feederTimes.forEach((p, j) => {
      const speed = speedFromStars(p.stars);
      const ride = (LEG1_RIDE_MIN[stationKey(station)] ?? DEFAULT_LEG1_RIDE_MIN)[speed];
      const uid = `f${rowNo}-${j}`;
      trains.push({ ...base, uid, from: s.origin, to: station, dep: p.minutes, arr: p.minutes + ride, speed, estimated: true });
      links.push({ feederUid: uid, mainUid });
    });
  });
  return { trains, links, skipped };
}

/** Reads an uploaded .xlsx/.xls/.csv. Throws an Error with a user-readable message. */
function parseWorkbook(data: ArrayBuffer | Uint8Array): ParsedData {
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(data, { type: "array" });
  } catch {
    throw new Error("That file couldn't be read. Upload an .xlsx, .xls or .csv file.");
  }
  const names = wb.SheetNames;
  if (names.length === 0) throw new Error("That file has no sheets.");

  const settingsName = names.find((n) => norm(n) === "settings");
  const settings = settingsName ? parseSettings(sheetToRows(wb.Sheets[settingsName]).rows) : { ...DEFAULT_SETTINGS };

  const trainsName = names.find((n) => norm(n) === "trains") ?? names[0];
  const { rows, offset } = sheetToRows(wb.Sheets[trainsName]);

  const isLegacy = rows.length > 0 && norm(rows[0][0]) === "train" && norm(rows[0][1]).startsWith("starttime");
  if (isLegacy) {
    return { ...legacyToData(rows.slice(1), offset + 1, settings), settings, format: "legacy" };
  }
  return { ...parseTemplateRows(rows, offset, settings), settings, format: "template" };
}

/* ---------- ranking engine ---------- */

/** Builds every journey the user declared: direct trains, plus each feeder → main link that leaves enough time to change. */
function buildJourneys(data: TimetableData, s: Settings, useDelays: boolean): Journey[] {
  const o = stationKey(s.origin);
  const d = stationKey(s.destination);
  const late = (t: Train) => (useDelays ? t.lateMin : 0);
  const byUid = new Map(data.trains.map((t) => [t.uid, t] as const));
  const out: Journey[] = [];

  for (const t of data.trains) {
    if (stationKey(t.from) === o && stationKey(t.to) === d) {
      out.push({
        id: t.uid,
        kind: "direct",
        transferStation: s.origin,
        legs: [t],
        boardMin: t.dep,
        scheduledArrMin: t.arr,
        expectedArrMin: t.arr + late(t),
        totalMin: t.arr + late(t) - t.dep,
        waitMin: 0,
        requiredMin: 0,
        platformChange: null,
      });
    }
  }

  for (const link of data.links) {
    const t1 = byUid.get(link.feederUid);
    const t2 = byUid.get(link.mainUid);
    if (!t1 || !t2) continue;
    const wait = t2.dep - (t1.arr + late(t1));
    const known = t1.arrPlatform !== null && t2.depPlatform !== null;
    const change = known ? t1.arrPlatform !== t2.depPlatform : null;
    const required = s.minBufferMin + (change ? s.platformChangeMin : 0);
    if (wait < required) continue; // can't make the change
    out.push({
      id: `${t1.uid}>${t2.uid}`,
      kind: "switch",
      transferStation: t1.to,
      legs: [t1, t2],
      boardMin: t1.dep,
      scheduledArrMin: t2.arr,
      expectedArrMin: t2.arr + late(t2),
      totalMin: t2.arr + late(t2) - t1.dep,
      waitMin: wait,
      requiredMin: required,
      platformChange: change,
    });
  }
  return out;
}

/** Valid, catchable journeys ranked: earliest arrival → shortest trip → least waiting. */
function findBest(data: TimetableData, s: Settings, queryMin: number, travelMin: number, useDelays: boolean): SearchResult {
  const readyMin = queryMin + travelMin;
  const options = buildJourneys(data, s, useDelays).filter((j) => j.boardMin >= readyMin);
  options.sort(
    (a, b) =>
      a.expectedArrMin - b.expectedArrMin ||
      a.totalMin - b.totalMin ||
      a.waitMin - b.waitMin ||
      a.boardMin - b.boardMin
  );
  return { options, best: options[0] ?? null, queryMin, travelMin, readyMin, atStation: travelMin === 0, useDelays };
}

/* ---------- formatting ---------- */

function fmtTime(min: number): string {
  const h24 = Math.floor(min / 60) % 24;
  const m = min % 60;
  const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
  return `${h12}:${String(m).padStart(2, "0")} ${h24 >= 12 ? "PM" : "AM"}`;
}

function fmtDuration(min: number): string {
  const r = Math.round(min);
  const h = Math.floor(r / 60);
  const m = r % 60;
  if (h === 0) return `${m} min`;
  return m === 0 ? `${h} hr` : `${h} hr ${m} min`;
}

function to24h(hour12: number, minute: number, meridiem: "AM" | "PM"): number {
  return (hour12 % 12) * 60 + minute + (meridiem === "PM" ? 720 : 0);
}

const SPEED_LABEL: Record<Speed, string> = { slow: "Slow", fast: "Fast", "double-fast": "Double fast" };

/** Lower-cased text a search box can match against: ids, stations, times, speeds, platforms, notes. */
function journeyHaystack(j: Journey): string {
  const parts = j.legs.map((l) =>
    [
      l.id,
      l.from,
      l.to,
      fmtTime(l.dep),
      fmtTime(l.arr),
      SPEED_LABEL[l.speed],
      l.depPlatform ? `platform ${l.depPlatform}` : "",
      l.arrPlatform ? `platform ${l.arrPlatform}` : "",
      l.notes,
    ].join(" ")
  );
  return `${parts.join(" ")} ${j.kind === "direct" ? "direct" : `switch via ${j.transferStation}`}`.toLowerCase();
}

/* ---------- template download (fallback when /public file is absent) ---------- */

function timeCell(min: number): XLSX.CellObject {
  return { t: "n", v: min / 1440, z: "h:mm AM/PM" };
}

function buildTemplateWorkbook(): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  const trains = XLSX.utils.aoa_to_sheet([
    ["Train ID", "Reference ID", "From Station", "To Station", "Departure Time", "Arrival Time", "Speed", "Departure Platform", "Arrival Platform", "Avg Late (min)", "Train Info / Notes"],
    ["EXAMPLE-1", null, "Andheri", "Virar", timeCell(18 * 60 + 10), timeCell(18 * 60 + 59), "Fast", 1, 4, 3, "Example MAIN train: delete or overwrite"],
    ["EXAMPLE-2", "EXAMPLE-1", "Bandra", "Andheri", timeCell(17 * 60 + 46), timeCell(17 * 60 + 56), "Fast", 3, 1, 2, "Example FEEDER: reaches Andheri for EXAMPLE-1"],
  ]);
  trains["!cols"] = [16, 16, 16, 16, 16, 14, 14, 18, 16, 14, 40].map((wch) => ({ wch }));

  const settings = XLSX.utils.aoa_to_sheet([
    ["Setting", "Value", "What it does"],
    ["Origin station", DEFAULT_SETTINGS.origin, "Where you start. Feeder trains must start here."],
    ["Destination station", DEFAULT_SETTINGS.destination, "Where you want to go. Main trains must end here."],
    ["Min transfer buffer (min)", DEFAULT_SETTINGS.minBufferMin, "Minimum change time at the same platform."],
    ["Platform change walk time (min)", DEFAULT_SETTINGS.platformChangeMin, "Extra time when the platform differs."],
  ]);
  settings["!cols"] = [{ wch: 34 }, { wch: 14 }, { wch: 60 }];

  const help = XLSX.utils.aoa_to_sheet([
    ["MAIN train: reaches Virar (Andheri to Virar) or runs direct (Bandra to Virar). Give it a unique Train ID; leave Reference ID blank."],
    ["FEEDER train: Bandra to Andheri. Put the main train's Train ID in Reference ID. Several mains: separate IDs with commas."],
    ["The feeder's To Station must equal the main train's From Station. Only linked trains are offered; nothing is guessed."],
    ["Speed: Slow, Fast or Double Fast. Platforms and Avg Late are optional. Rows starting EXAMPLE are ignored."],
  ]);
  help["!cols"] = [{ wch: 120 }];

  XLSX.utils.book_append_sheet(wb, trains, "Trains");
  XLSX.utils.book_append_sheet(wb, settings, "Settings");
  XLSX.utils.book_append_sheet(wb, help, "Instructions");
  return wb;
}

const LEGACY_PARSED = legacyToData(LEGACY_SAMPLE, 1, DEFAULT_SETTINGS);

/* ============================================================================
 * ENGINE END
 * ========================================================================== */

/* ---------- saved files (IndexedDB) ---------- */

interface StoredFile {
  id: string;
  name: string;
  type: string;
  size: number;
  addedAt: number;
  trainCount: number;
  data: ArrayBuffer;
}

const DB_NAME = "western-line-optimizer";
const STORE = "files";
const ACTIVE_KEY = "wl-active-file";

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("IndexedDB unavailable"));
      return;
    }
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: "id" });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbAll(): Promise<StoredFile[]> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction(STORE).objectStore(STORE).getAll();
    req.onsuccess = () => resolve(req.result as StoredFile[]);
    req.onerror = () => reject(req.error);
  });
}

async function idbPut(file: StoredFile): Promise<void> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).put(file);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

async function idbDelete(id: string): Promise<void> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}

const safeGet = (k: string): string | null => {
  try {
    return window.localStorage.getItem(k);
  } catch {
    return null;
  }
};
const safeSet = (k: string, v: string | null): void => {
  try {
    if (v === null) window.localStorage.removeItem(k);
    else window.localStorage.setItem(k, v);
  } catch {
    /* storage blocked: the app still works for this session */
  }
};
const newId = (): string => `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

const issueLabel = (n: number): string => `${n} ${n === 1 ? "issue" : "issues"} found:`;

function fmtSize(bytes: number): string {
  return bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
function fmtAdded(ts: number): string {
  return new Date(ts).toLocaleString(undefined, { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" });
}

/* ---------- UI types ---------- */

type Phase = "idle" | "locating" | "scanning" | "result";
type LocMode = "station" | "onway";
type SortMode = "best" | "soonest" | "earliest" | "shortest";

interface Dataset extends ParsedData {
  name: string;
}

interface Notice {
  tone: "ok" | "error";
  text: string;
  details: string[];
}

type TlItem =
  | { kind: "stop"; key: string; time: number; station: string; platform: string | null }
  | { kind: "ride"; key: string; leg: Train }
  | { kind: "change"; key: string; text: string };

const SAMPLE_DATASET: Dataset = {
  trains: LEGACY_PARSED.trains,
  links: LEGACY_PARSED.links,
  settings: DEFAULT_SETTINGS,
  skipped: LEGACY_PARSED.skipped,
  format: "legacy",
  name: "Sample timetable",
};

const SORT_LABEL: Record<SortMode, string> = {
  best: "Best match",
  soonest: "Leaves soonest",
  earliest: "Arrives first",
  shortest: "Shortest trip",
};

const SCAN_STEPS = ["Reading the timetable", "Checking platform changes", "Ranking arrival times"];
const ETA_PRESETS = [5, 10, 15, 20, 30, 45];
const HOURS = Array.from({ length: 12 }, (_, i) => i + 1);
const MINUTES = Array.from({ length: 60 }, (_, i) => i);

function buildTimeline(j: Journey): TlItem[] {
  const items: TlItem[] = [];
  j.legs.forEach((leg, idx) => {
    if (idx === 0) items.push({ kind: "stop", key: `${leg.uid}-d`, time: leg.dep, station: leg.from, platform: leg.depPlatform });
    items.push({ kind: "ride", key: `${leg.uid}-r`, leg });
    items.push({ kind: "stop", key: `${leg.uid}-a`, time: leg.arr, station: leg.to, platform: leg.arrPlatform });
    if (idx === 0 && j.legs.length > 1) {
      const next = j.legs[1];
      let text = `Change trains: ${j.waitMin} min`;
      if (j.platformChange === true) text += `, platform ${leg.arrPlatform} to ${next.depPlatform} (needs ${j.requiredMin})`;
      if (j.platformChange === false) text += `, same platform ${next.depPlatform}`;
      items.push({ kind: "change", key: "change", text });
      items.push({ kind: "stop", key: `${next.uid}-d`, time: next.dep, station: next.from, platform: next.depPlatform });
    }
  });
  return items;
}

export default function WesternLineOptimizer() {
  const [phase, setPhase] = useState<Phase>("idle");
  const [result, setResult] = useState<SearchResult | null>(null);
  const [scanStep, setScanStep] = useState(0);

  const [dataset, setDataset] = useState<Dataset>(SAMPLE_DATASET);
  const [activeId, setActiveId] = useState<string>("sample");
  const [files, setFiles] = useState<StoredFile[]>([]);
  const [storageOk, setStorageOk] = useState(true);
  const [libOpen, setLibOpen] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const [notice, setNotice] = useState<Notice | null>(null);
  const [useDelays, setUseDelays] = useState(true);

  const [locMode, setLocMode] = useState<LocMode>("station");
  const [etaMin, setEtaMin] = useState(15);

  const [listOpen, setListOpen] = useState(false);
  const [sortMode, setSortMode] = useState<SortMode>("best");
  const [query, setQuery] = useState("");
  const [pickedId, setPickedId] = useState<string | null>(null);

  const [simOn, setSimOn] = useState(false);
  const [simHour, setSimHour] = useState(5);
  const [simMinute, setSimMinute] = useState(0);
  const [simMeridiem, setSimMeridiem] = useState<"AM" | "PM">("PM");

  const [now, setNow] = useState<Date | null>(null);
  const timers = useRef<number[]>([]);
  const cardRef = useRef<HTMLElement | null>(null);
  const sheetRef = useRef<HTMLDivElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);
  const libRef = useRef<HTMLDivElement | null>(null);
  const fileRef = useRef<HTMLInputElement | null>(null);

  const hasDelays = useMemo(() => dataset.trains.some((t) => t.lateMin > 0), [dataset]);
  const { origin, destination } = dataset.settings;

  /* ----- lifecycle ----- */

  useEffect(() => {
    setNow(new Date());
    const id = window.setInterval(() => setNow(new Date()), 15_000);
    return () => window.clearInterval(id);
  }, []);

  useEffect(() => {
    const pending = timers.current;
    return () => pending.forEach((t) => window.clearTimeout(t));
  }, []);

  useEffect(() => {
    if (phase === "result" && !listOpen) cardRef.current?.focus();
    if (phase === "locating") sheetRef.current?.focus();
  }, [phase, listOpen]);
  useEffect(() => {
    if (listOpen) listRef.current?.focus();
  }, [listOpen]);
  useEffect(() => {
    if (libOpen) libRef.current?.focus();
  }, [libOpen]);

  useEffect(() => {
    if (phase !== "locating" && !listOpen && !libOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape") return;
      setListOpen(false);
      setLibOpen(false);
      setConfirmDeleteId(null);
      if (phase === "locating") setPhase("idle");
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [phase, listOpen, libOpen]);

  const getQueryMinutes = useCallback((): number => {
    if (simOn) return to24h(simHour, simMinute, simMeridiem);
    const d = new Date();
    return d.getHours() * 60 + d.getMinutes();
  }, [simOn, simHour, simMinute, simMeridiem]);

  /* ----- saved files ----- */

  /** Switches the timetable in use. `file === null` means the built-in sample. */
  const activate = useCallback((file: StoredFile | null, announce: boolean) => {
    if (!file) {
      setDataset(SAMPLE_DATASET);
      setActiveId("sample");
      safeSet(ACTIVE_KEY, null);
      if (announce) setNotice(null);
      return;
    }
    try {
      const parsed = parseWorkbook(file.data);
      setDataset({ ...parsed, name: file.name });
      setActiveId(file.id);
      safeSet(ACTIVE_KEY, file.id);
      if (announce) {
        const extra = parsed.format === "legacy" ? " Old layout: arrival times at transfer stations are estimated." : "";
        setNotice({
          tone: "ok",
          text: `Now using ${file.name} (${parsed.trains.length} trains).${extra}`,
          details: parsed.skipped.length ? [issueLabel(parsed.skipped.length), ...parsed.skipped] : [],
        });
      }
    } catch (err) {
      setNotice({ tone: "error", text: err instanceof Error ? err.message : "That file couldn't be read.", details: [] });
    }
  }, []);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const stored = (await idbAll()).sort((a, b) => b.addedAt - a.addedAt);
        if (cancelled) return;
        setFiles(stored);
        const saved = stored.find((f) => f.id === safeGet(ACTIVE_KEY));
        if (saved) activate(saved, false);
      } catch {
        if (!cancelled) setStorageOk(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [activate]);

  const handleFile = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      e.target.value = ""; // lets the same file be chosen again
      if (!file) return;
      try {
        const bytes = await file.arrayBuffer();
        const parsed = parseWorkbook(bytes);
        if (parsed.trains.length === 0) {
          setNotice({ tone: "error", text: `No usable trains found in ${file.name}.`, details: parsed.skipped });
          return;
        }
        const stored: StoredFile = {
          id: newId(),
          name: file.name,
          type: file.type,
          size: file.size,
          addedAt: Date.now(),
          trainCount: parsed.trains.length,
          data: bytes,
        };
        setFiles((prev) => [stored, ...prev]);
        try {
          await idbPut(stored);
        } catch {
          setStorageOk(false);
        }
        setDataset({ ...parsed, name: file.name });
        setActiveId(stored.id);
        safeSet(ACTIVE_KEY, stored.id);
        const extra = parsed.format === "legacy" ? " Old layout detected; arrival times at transfer stations are estimated." : "";
        setNotice({
          tone: "ok",
          text: `Loaded ${parsed.trains.length} trains from ${file.name}.${extra}`,
          details: parsed.skipped.length ? [issueLabel(parsed.skipped.length), ...parsed.skipped] : [],
        });
      } catch (err) {
        setNotice({ tone: "error", text: err instanceof Error ? err.message : "That file couldn't be read.", details: [] });
      }
    },
    []
  );

  const downloadFile = useCallback((f: StoredFile) => {
    const url = URL.createObjectURL(new Blob([f.data], { type: f.type || "application/octet-stream" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = f.name;
    a.click();
    URL.revokeObjectURL(url);
  }, []);

  const deleteFile = useCallback(
    async (f: StoredFile) => {
      try {
        await idbDelete(f.id);
      } catch {
        setStorageOk(false);
      }
      setFiles((prev) => prev.filter((x) => x.id !== f.id));
      setConfirmDeleteId(null);
      if (activeId === f.id) activate(null, true);
    },
    [activeId, activate]
  );

  const handleTemplate = useCallback(async () => {
    try {
      const res = await fetch(TEMPLATE_URL);
      const type = res.headers.get("content-type") ?? "";
      if (res.ok && !type.includes("text/html")) {
        const url = URL.createObjectURL(await res.blob());
        const a = document.createElement("a");
        a.href = url;
        a.download = TEMPLATE_FILENAME;
        a.click();
        URL.revokeObjectURL(url);
        return;
      }
    } catch {
      /* fall through to the generated template */
    }
    XLSX.writeFile(buildTemplateWorkbook(), TEMPLATE_FILENAME);
  }, []);

  /* ----- search flow ----- */

  const handleFind = useCallback(() => {
    if (phase === "idle") setPhase("locating");
  }, [phase]);

  const handleCancel = useCallback(() => setPhase("idle"), []);

  const adjustEta = useCallback((delta: number) => setEtaMin((v) => Math.min(180, Math.max(1, v + delta))), []);

  const handleConfirm = useCallback(() => {
    if (phase !== "locating") return;
    const queryMin = getQueryMinutes();
    const travelMin = locMode === "station" ? 0 : etaMin;
    const delays = useDelays && hasDelays;
    setScanStep(0);
    setPickedId(null);
    setQuery("");
    setSortMode("best");
    setPhase("scanning");

    const t = (fn: () => void, ms: number) => timers.current.push(window.setTimeout(fn, ms));
    t(() => setScanStep(1), SEARCH_DELAY_MS * 0.35);
    t(() => setScanStep(2), SEARCH_DELAY_MS * 0.7);
    t(() => {
      setResult(findBest(dataset, dataset.settings, queryMin, travelMin, delays));
      setPhase("result");
    }, SEARCH_DELAY_MS);
  }, [phase, getQueryMinutes, locMode, etaMin, useDelays, hasDelays, dataset]);

  const handleReset = useCallback(() => {
    setResult(null);
    setPickedId(null);
    setListOpen(false);
    setPhase("idle");
  }, []);

  /* ----- derived ----- */

  const shown: Journey | null = result ? (result.options.find((o) => o.id === pickedId) ?? result.best) : null;
  const isBest = !!shown && shown.id === result?.best?.id;
  const waitAtStation = shown && result ? shown.boardMin - result.readyMin : 0;
  const previewReadyMin = phase === "locating" ? getQueryMinutes() + (locMode === "station" ? 0 : etaMin) : 0;
  const liveMinutes = now ? now.getHours() * 60 + now.getMinutes() : null;
  const hasEstimate = !!shown && shown.legs.some((l) => l.estimated);
  const timeline = useMemo(() => (shown ? buildTimeline(shown) : []), [shown]);

  const haystacks = useMemo(() => (result ? result.options.map(journeyHaystack) : []), [result]);

  const listItems = useMemo(() => {
    if (!result) return { items: [] as { o: Journey; rank: number }[], matches: 0 };
    const tokens = query.toLowerCase().split(/\s+/).filter(Boolean);
    let items = result.options.map((o, rank) => ({ o, rank }));
    if (tokens.length) items = items.filter(({ rank }) => tokens.every((t) => haystacks[rank].includes(t)));
    const matches = items.length;
    if (sortMode === "soonest") items.sort((a, b) => a.o.boardMin - b.o.boardMin || a.rank - b.rank);
    else if (sortMode === "earliest") items.sort((a, b) => a.o.expectedArrMin - b.o.expectedArrMin || a.rank - b.rank);
    else if (sortMode === "shortest") items.sort((a, b) => a.o.totalMin - b.o.totalMin || a.rank - b.rank);
    return { items: items.slice(0, MAX_LIST), matches };
  }, [result, haystacks, query, sortMode]);

  const dataLabel = `${dataset.name}, ${dataset.trains.length} trains`;

  return (
    <div className="lt-root">
      <style>{CSS}</style>

      <main className="lt-main">
        {phase !== "result" && (
          <div className="lt-hero">
            <h1 className="lt-title">
              {origin} to {destination}
            </h1>
            <p className="lt-sub">Western Line, earliest arrival</p>

            <button type="button" className="lt-cta" onClick={handleFind} disabled={phase === "scanning"} aria-label="Find the best train">
              <TrainIcon />
              <span className="lt-cta-text">FIND BEST</span>
            </button>

            <div className="lt-data">
              <p className="lt-data-status">{dataLabel}</p>
              <div className="lt-data-actions">
                <button type="button" className="lt-ghost" onClick={() => setLibOpen(true)}>
                  Timetables ({files.length + 1})
                </button>
                <button type="button" className="lt-ghost" onClick={() => fileRef.current?.click()}>
                  Upload Excel
                </button>
                <button type="button" className="lt-ghost" onClick={handleTemplate}>
                  Download template
                </button>
              </div>
              {notice && (
                <div className="lt-notice" data-tone={notice.tone} role="status">
                  <p>{notice.text}</p>
                  {notice.details.length > 0 && (
                    <details>
                      <summary>Details</summary>
                      <ul>
                        {notice.details.slice(0, 40).map((d, i) => (
                          <li key={i}>{d}</li>
                        ))}
                        {notice.details.length > 40 && <li>and {notice.details.length - 40} more</li>}
                      </ul>
                    </details>
                  )}
                </div>
              )}
            </div>
          </div>
        )}

        {/* Shared hidden file input (used by the home screen and the library) */}
        <input
          ref={fileRef}
          type="file"
          accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv"
          onChange={handleFile}
          hidden
        />

        {phase === "result" && result && (
          <div className="lt-results">
            {shown ? (
              <>
                {/* Card 1: the answer */}
                <section className="lt-card lt-card-hero lt-reveal" style={{ animationDelay: "0ms" }} ref={cardRef} tabIndex={-1} aria-label="Recommended train">
                  <div className="lt-card-head">
                    <span className="lt-pill" data-kind={shown.kind}>
                      {shown.kind === "direct" ? "Direct train" : `Switch at ${shown.transferStation}`}
                    </span>
                    <span className="lt-muted">{isBest ? `Best of ${result.options.length}` : "Your pick"}</span>
                  </div>

                  <div className="lt-hero-grid">
                    <div className="lt-hero-col">
                      <span className="lt-label">Board at {origin}</span>
                      <span className="lt-big">{fmtTime(shown.boardMin)}</span>
                      {shown.legs[0].depPlatform && <span className="lt-plat">Platform {shown.legs[0].depPlatform}</span>}
                    </div>
                    <div className="lt-hero-col lt-hero-right">
                      <span className="lt-label">Arrive at {destination}</span>
                      <span className="lt-big lt-big-em">{fmtTime(shown.expectedArrMin)}</span>
                      {shown.expectedArrMin !== shown.scheduledArrMin && (
                        <span className="lt-muted">Scheduled {fmtTime(shown.scheduledArrMin)}</span>
                      )}
                    </div>
                  </div>
                  <p className="lt-total-line">
                    Total trip <strong>{fmtDuration(shown.totalMin)}</strong>
                  </p>
                </section>

                {/* Card 2: what to do right now */}
                <section className="lt-card lt-reveal" style={{ animationDelay: "90ms" }} aria-label="Your timing">
                  <h2 className="lt-card-title">Your timing</h2>
                  <div className="lt-tiles" data-count={result.atStation ? 2 : 3}>
                    {!result.atStation && (
                      <div className="lt-tile">
                        <span className="lt-tile-label">Reach {origin}</span>
                        <span className="lt-tile-val">{fmtTime(result.readyMin)}</span>
                        <span className="lt-tile-sub">in {fmtDuration(result.travelMin)}</span>
                      </div>
                    )}
                    <div className="lt-tile">
                      <span className="lt-tile-label">Train leaves in</span>
                      <span className="lt-tile-val">{fmtDuration(shown.boardMin - result.queryMin)}</span>
                      <span className="lt-tile-sub">at {fmtTime(shown.boardMin)}</span>
                    </div>
                    <div className="lt-tile lt-tile-key">
                      <span className="lt-tile-label">Wait at station</span>
                      <span className="lt-tile-val">{fmtDuration(waitAtStation)}</span>
                      <span className="lt-tile-sub">{waitAtStation <= 2 ? "Go straight to the platform" : "on the platform"}</span>
                    </div>
                  </div>
                </section>

                {/* Card 3: the route */}
                <section className="lt-card lt-reveal" style={{ animationDelay: "180ms" }} aria-label="Journey details">
                  <h2 className="lt-card-title">Your journey</h2>
                  <ol className="lt-tl">
                    {timeline.map((it, idx) => {
                      const pos = idx === 0 ? "first" : idx === timeline.length - 1 ? "last" : "mid";
                      if (it.kind === "stop") {
                        return (
                          <li key={it.key} className="lt-tl-item lt-tl-stop" data-pos={pos}>
                            <span className="lt-tl-time">{fmtTime(it.time)}</span>
                            <span className="lt-tl-name">{it.station}</span>
                            {it.platform && <span className="lt-chip-plat">Platform {it.platform}</span>}
                          </li>
                        );
                      }
                      if (it.kind === "ride") {
                        const leg = it.leg;
                        return (
                          <li key={it.key} className="lt-tl-item lt-tl-ride" data-pos={pos}>
                            <span className="lt-badge" data-speed={leg.speed}>
                              {SPEED_LABEL[leg.speed]}
                            </span>
                            <span className="lt-muted">{fmtDuration(leg.arr - leg.dep)}</span>
                            {leg.id && <span className="lt-muted">Train {leg.id}</span>}
                            {result.useDelays && leg.lateMin > 0 && <span className="lt-late">Averages {fmtDuration(leg.lateMin)} late</span>}
                            {leg.notes && <span className="lt-note-line">{leg.notes}</span>}
                          </li>
                        );
                      }
                      return (
                        <li key={it.key} className="lt-tl-item lt-tl-change" data-pos={pos}>
                          <span className="lt-change-chip">{it.text}</span>
                        </li>
                      );
                    })}
                  </ol>
                  {hasEstimate && (
                    <p className="lt-note lt-estimate">
                      Arrival at {shown.transferStation} is estimated. Use the template with real arrival times for exact transfers.
                    </p>
                  )}
                </section>

                {!isBest && (
                  <button type="button" className="lt-link" onClick={() => setPickedId(null)}>
                    Back to best match
                  </button>
                )}
                <div className="lt-actions">
                  <button type="button" className="lt-again" onClick={() => setListOpen(true)}>
                    See all {result.options.length} options
                  </button>
                  <button type="button" className="lt-again" onClick={handleReset}>
                    Search again
                  </button>
                </div>
              </>
            ) : (
              <section className="lt-card lt-reveal" ref={cardRef} tabIndex={-1} aria-label="No train available">
                <div className="lt-empty">
                  <h2>No train available</h2>
                  <p>
                    No linked train leaves {origin} at or after {fmtTime(result.readyMin)}
                    {result.atStation ? "" : ", when you reach the station"}, with enough time to change. Try an earlier time in the
                    simulator or add more trains to your Excel file.
                  </p>
                  <button type="button" className="lt-again" onClick={handleReset}>
                    Search again
                  </button>
                </div>
              </section>
            )}
          </div>
        )}
      </main>

      {/* All options: search, sort, manual pick */}
      {listOpen && result && (
        <div className="lt-backdrop" onClick={() => setListOpen(false)}>
          <div
            className="lt-sheet lt-sheet-tall"
            role="dialog"
            aria-modal="true"
            aria-labelledby="lt-list-title"
            tabIndex={-1}
            ref={listRef}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="lt-sheet-head">
              <h2 id="lt-list-title" className="lt-sheet-title">
                All {result.options.length} options
              </h2>
              <p className="lt-sheet-note lt-tight">Tap a train to make it your recommendation.</p>

              <input
                type="search"
                className="lt-search"
                placeholder="Search train, station, time or platform"
                aria-label="Search options"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
              />

              <div className="lt-chips lt-sort" role="group" aria-label="Sort options">
                {(Object.keys(SORT_LABEL) as SortMode[]).map((m) => (
                  <button key={m} type="button" className="lt-chip" aria-pressed={sortMode === m} onClick={() => setSortMode(m)}>
                    {SORT_LABEL[m]}
                  </button>
                ))}
              </div>
              {query.trim() !== "" && (
                <p className="lt-count" role="status">
                  {listItems.matches} of {result.options.length} match
                </p>
              )}
            </div>

            <div className="lt-list-scroll">
              {listItems.items.length === 0 ? (
                <p className="lt-sheet-note">No options match. Try fewer words, such as a station or a time.</p>
              ) : (
                <ul className="lt-list">
                  {listItems.items.map(({ o, rank }) => {
                    const active = shown?.id === o.id;
                    const plat = o.legs[0].depPlatform;
                    return (
                      <li key={o.id}>
                        <button
                          type="button"
                          className="lt-opt"
                          aria-pressed={active}
                          data-active={active}
                          onClick={() => {
                            setPickedId(o.id === result.best?.id ? null : o.id);
                            setListOpen(false);
                          }}
                        >
                          <span className="lt-opt-top">
                            <span className="lt-opt-time">{fmtTime(o.boardMin)}</span>
                            <span className="lt-opt-arrive">arrives {fmtTime(o.expectedArrMin)}</span>
                          </span>
                          <span className="lt-opt-meta">
                            <span>{o.kind === "direct" ? "Direct" : `via ${o.transferStation}`}</span>
                            <span>{o.legs.map((l) => SPEED_LABEL[l.speed]).join(" then ")}</span>
                            <span>{fmtDuration(o.totalMin)}</span>
                            {o.kind === "switch" && <span>change {o.waitMin} min</span>}
                            {plat && <span>platform {plat}</span>}
                            {o.platformChange === true && <span>platform change</span>}
                            {o.legs.some((l) => l.id) && <span>{o.legs.map((l) => l.id).filter(Boolean).join(" + ")}</span>}
                          </span>
                          {(rank === 0 || active) && (
                            <span className="lt-tags">
                              {rank === 0 && <span className="lt-tag">Best match</span>}
                              {active && <span className="lt-tag lt-tag-solid">Selected</span>}
                            </span>
                          )}
                        </button>
                      </li>
                    );
                  })}
                </ul>
              )}
              {listItems.matches > MAX_LIST && <p className="lt-sheet-note">Showing the first {MAX_LIST} of {listItems.matches}.</p>}
            </div>

            <button type="button" className="lt-primary lt-full" onClick={() => setListOpen(false)}>
              Done
            </button>
          </div>
        </div>
      )}

      {/* Saved timetables */}
      {libOpen && (
        <div className="lt-backdrop" onClick={() => { setLibOpen(false); setConfirmDeleteId(null); }}>
          <div
            className="lt-sheet lt-sheet-tall"
            role="dialog"
            aria-modal="true"
            aria-labelledby="lt-lib-title"
            tabIndex={-1}
            ref={libRef}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="lt-sheet-head">
              <h2 id="lt-lib-title" className="lt-sheet-title">
                Your timetables
              </h2>
              <p className="lt-sheet-note lt-tight">Every file you upload is kept here. Pick one to use it.</p>
              {!storageOk && (
                <p className="lt-warn">This browser can't save files, so they will be lost when you close the tab.</p>
              )}
            </div>

            <div className="lt-list-scroll">
              <ul className="lt-list">
                <li className="lt-file" data-active={activeId === "sample"}>
                  <button type="button" className="lt-file-main" aria-pressed={activeId === "sample"} onClick={() => activate(null, true)}>
                    <span className="lt-file-name">Sample timetable</span>
                    <span className="lt-file-meta">Built in, {LEGACY_PARSED.trains.length} trains</span>
                    {activeId === "sample" && <span className="lt-tag lt-tag-solid">In use</span>}
                  </button>
                </li>
                {files.map((f) => (
                  <li key={f.id} className="lt-file" data-active={activeId === f.id}>
                    <button type="button" className="lt-file-main" aria-pressed={activeId === f.id} onClick={() => activate(f, true)}>
                      <span className="lt-file-name">{f.name}</span>
                      <span className="lt-file-meta">
                        {f.trainCount} trains, {fmtSize(f.size)}, added {fmtAdded(f.addedAt)}
                      </span>
                      {activeId === f.id && <span className="lt-tag lt-tag-solid">In use</span>}
                    </button>
                    {confirmDeleteId === f.id ? (
                      <div className="lt-file-actions lt-confirm" role="alert">
                        <span>Delete this file?</span>
                        <button type="button" className="lt-mini" onClick={() => setConfirmDeleteId(null)}>
                          Cancel
                        </button>
                        <button type="button" className="lt-mini lt-mini-danger" onClick={() => deleteFile(f)}>
                          Delete
                        </button>
                      </div>
                    ) : (
                      <div className="lt-file-actions">
                        <button type="button" className="lt-mini" onClick={() => downloadFile(f)} aria-label={`Download ${f.name}`}>
                          Download
                        </button>
                        <button type="button" className="lt-mini lt-mini-danger" onClick={() => setConfirmDeleteId(f.id)} aria-label={`Delete ${f.name}`}>
                          Delete
                        </button>
                      </div>
                    )}
                  </li>
                ))}
              </ul>
              {files.length === 0 && <p className="lt-sheet-note">No uploaded files yet.</p>}
            </div>

            <div className="lt-sheet-actions">
              <button type="button" className="lt-again" onClick={() => fileRef.current?.click()}>
                Upload Excel
              </button>
              <button type="button" className="lt-primary" onClick={() => { setLibOpen(false); setConfirmDeleteId(null); }}>
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Where are you? */}
      {phase === "locating" && (
        <div className="lt-backdrop" onClick={handleCancel}>
          <div
            className="lt-sheet"
            role="dialog"
            aria-modal="true"
            aria-labelledby="lt-loc-title"
            tabIndex={-1}
            ref={sheetRef}
            onClick={(e) => e.stopPropagation()}
          >
            <h2 id="lt-loc-title" className="lt-sheet-title">
              Where are you right now?
            </h2>

            <div className="lt-options" role="radiogroup" aria-label="Your location">
              <label className="lt-option" data-active={locMode === "station"}>
                <input type="radio" name="lt-loc" value="station" checked={locMode === "station"} onChange={() => setLocMode("station")} />
                <span className="lt-option-title">At {origin} station</span>
                <span className="lt-option-sub">Ready to board</span>
              </label>
              <label className="lt-option" data-active={locMode === "onway"}>
                <input type="radio" name="lt-loc" value="onway" checked={locMode === "onway"} onChange={() => setLocMode("onway")} />
                <span className="lt-option-title">On the way</span>
                <span className="lt-option-sub">Still travelling to {origin}</span>
              </label>
            </div>

            {locMode === "onway" && (
              <div className="lt-eta">
                <span className="lt-label">How long to reach {origin} station?</span>
                <div className="lt-stepper">
                  <button type="button" onClick={() => adjustEta(-1)} aria-label="One minute less" disabled={etaMin <= 1}>
                    −
                  </button>
                  <output aria-live="polite">{fmtDuration(etaMin)}</output>
                  <button type="button" onClick={() => adjustEta(1)} aria-label="One minute more" disabled={etaMin >= 180}>
                    +
                  </button>
                </div>
                <div className="lt-chips">
                  {ETA_PRESETS.map((m) => (
                    <button key={m} type="button" className="lt-chip" aria-pressed={etaMin === m} onClick={() => setEtaMin(m)}>
                      {m} min
                    </button>
                  ))}
                </div>
              </div>
            )}

            {hasDelays && (
              <label className="lt-switch lt-delay-switch">
                <input type="checkbox" role="switch" checked={useDelays} onChange={(e) => setUseDelays(e.target.checked)} />
                <span className="lt-track" aria-hidden="true" />
                <span className="lt-switch-text">Include average delays</span>
              </label>
            )}

            <p className="lt-sheet-note">
              Looking for trains leaving {origin} at {fmtTime(previewReadyMin)} or later.
            </p>

            <div className="lt-sheet-actions">
              <button type="button" className="lt-again" onClick={handleCancel}>
                Cancel
              </button>
              <button type="button" className="lt-primary" onClick={handleConfirm}>
                Find best train
              </button>
            </div>
          </div>
        </div>
      )}

      {phase === "scanning" && (
        <div className="lt-overlay" role="status" aria-live="polite" aria-label="Searching timetable">
          <div className="lt-radar" aria-hidden="true">
            <span className="lt-radar-dot" />
          </div>
          <p className="lt-scan-text" key={scanStep}>
            {SCAN_STEPS[scanStep]}
          </p>
          <div className="lt-progress" aria-hidden="true">
            <span style={{ animationDuration: `${SEARCH_DELAY_MS}ms` }} />
          </div>
        </div>
      )}

      <footer className="lt-dev">
        <label className="lt-switch">
          <input type="checkbox" role="switch" checked={simOn} onChange={(e) => setSimOn(e.target.checked)} aria-label="Dev time simulator" />
          <span className="lt-track" aria-hidden="true" />
          <span className="lt-switch-text">Dev time simulator</span>
        </label>

        {simOn ? (
          <div className="lt-time-picker" role="group" aria-label="Simulated time">
            <select value={simHour} onChange={(e) => setSimHour(Number(e.target.value))} aria-label="Hour">
              {HOURS.map((h) => (
                <option key={h} value={h}>
                  {h}
                </option>
              ))}
            </select>
            <span className="lt-colon">:</span>
            <select value={simMinute} onChange={(e) => setSimMinute(Number(e.target.value))} aria-label="Minute">
              {MINUTES.map((m) => (
                <option key={m} value={m}>
                  {String(m).padStart(2, "0")}
                </option>
              ))}
            </select>
            <select value={simMeridiem} onChange={(e) => setSimMeridiem(e.target.value as "AM" | "PM")} aria-label="AM or PM">
              <option value="AM">AM</option>
              <option value="PM">PM</option>
            </select>
          </div>
        ) : (
          <span className="lt-muted lt-live">{liveMinutes !== null ? `Live time ${fmtTime(liveMinutes)}` : "\u00a0"}</span>
        )}
      </footer>
    </div>
  );
}

function TrainIcon() {
  return (
    <svg
      className="lt-cta-icon"
      viewBox="0 0 48 48"
      width="52"
      height="52"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <rect x="10" y="5" width="28" height="30" rx="7" />
      <path d="M10 22h28" />
      <rect x="15" y="10" width="18" height="8" rx="2" />
      <circle cx="17" cy="29" r="1.6" fill="currentColor" stroke="none" />
      <circle cx="31" cy="29" r="1.6" fill="currentColor" stroke="none" />
      <path d="M16 35l-5 8M32 35l5 8M14 43h20" />
    </svg>
  );
}

const CSS = `
.lt-root{
  --bg:#0b0e0d; --em:#10b981; --em-hi:#34d399; --ink:#f3f8f6; --muted:#8fa19a;
  --glass:rgba(255,255,255,.045); --glass-line:rgba(255,255,255,.10);
  min-height:100vh; min-height:100dvh; background:var(--bg); color:var(--ink);
  font-family:ui-sans-serif,system-ui,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",sans-serif;
  font-variant-numeric:tabular-nums; position:relative; overflow-x:hidden;
  display:flex; flex-direction:column; -webkit-font-smoothing:antialiased;
}
.lt-root *,.lt-root *::before,.lt-root *::after{box-sizing:border-box}
.lt-root::before{
  content:""; position:fixed; inset:0; pointer-events:none;
  background:
    radial-gradient(ellipse 60% 45% at 50% 42%, rgba(16,185,129,.16), transparent 70%),
    radial-gradient(ellipse 40% 30% at 85% 5%, rgba(16,185,129,.06), transparent 70%);
}
.lt-main{flex:1; display:grid; place-items:center; padding:40px 20px 150px; position:relative; z-index:1}
.lt-hero{display:flex; flex-direction:column; align-items:center; text-align:center; width:100%; max-width:520px}
.lt-title{margin:0; font-size:clamp(26px,5vw,34px); font-weight:650; letter-spacing:-.02em}
.lt-sub{margin:8px 0 44px; color:var(--muted); font-size:15px}

.lt-cta{
  position:relative; width:clamp(176px,56vw,250px); aspect-ratio:1; border-radius:50%;
  border:1px solid rgba(110,231,183,.6); cursor:pointer; color:#04130d;
  display:flex; flex-direction:column; align-items:center; justify-content:center; gap:12px;
  background:radial-gradient(circle at 30% 22%, #4be3b0, var(--em) 46%, #08785a 100%);
  box-shadow:0 0 0 10px rgba(16,185,129,.07), 0 24px 90px rgba(16,185,129,.35), inset 0 2px 0 rgba(255,255,255,.35);
  transition:transform .18s ease, box-shadow .18s ease;
}
.lt-cta::after{
  content:""; position:absolute; inset:-2px; border-radius:50%;
  border:2px solid rgba(16,185,129,.5); animation:lt-pulse 2.8s ease-out infinite; pointer-events:none;
}
.lt-cta:hover{transform:scale(1.03)}
.lt-cta:active{transform:scale(.97)}
.lt-cta:focus-visible{outline:3px solid #fff; outline-offset:6px}
.lt-cta:disabled{cursor:progress}
.lt-cta-text{font-size:clamp(17px,4.4vw,21px); font-weight:800; letter-spacing:.1em}

.lt-data{margin-top:44px; display:flex; flex-direction:column; align-items:center; gap:12px; width:100%}
.lt-data-status{margin:0; font-size:14px; color:var(--muted); overflow-wrap:anywhere}
.lt-data-actions{display:flex; flex-wrap:wrap; justify-content:center; gap:8px}
.lt-ghost{
  padding:9px 16px; border-radius:999px; font:inherit; font-size:14px; font-weight:550; cursor:pointer; color:var(--ink);
  background:var(--glass); border:1px solid var(--glass-line); transition:background .15s ease, border-color .15s ease;
}
.lt-ghost:hover{background:rgba(255,255,255,.09); border-color:rgba(16,185,129,.5)}
.lt-ghost:focus-visible{outline:2px solid var(--em-hi); outline-offset:2px}
.lt-notice{
  width:100%; padding:12px 14px; border-radius:14px; text-align:left; font-size:14px; line-height:1.5;
  background:rgba(16,185,129,.09); border:1px solid rgba(16,185,129,.35);
}
.lt-notice[data-tone="error"]{background:rgba(248,113,113,.09); border-color:rgba(248,113,113,.45)}
.lt-notice p{margin:0; overflow-wrap:anywhere}
.lt-notice details{margin-top:8px; color:var(--muted)}
.lt-notice summary{cursor:pointer; color:var(--ink)}
.lt-notice ul{margin:8px 0 0; padding-left:18px; max-height:160px; overflow:auto}

.lt-overlay{
  position:fixed; inset:0; z-index:50; display:flex; flex-direction:column; align-items:center; justify-content:center; gap:28px;
  background:rgba(11,14,13,.78); -webkit-backdrop-filter:blur(16px); backdrop-filter:blur(16px);
  animation:lt-fade .2s ease-out;
}
.lt-radar{
  position:relative; width:clamp(170px,50vw,220px); aspect-ratio:1; border-radius:50%;
  border:1px solid rgba(16,185,129,.45); overflow:hidden;
  background:repeating-radial-gradient(circle, transparent 0 26px, rgba(16,185,129,.2) 26px 27px);
  box-shadow:0 0 70px rgba(16,185,129,.22), inset 0 0 40px rgba(16,185,129,.12);
}
.lt-radar::before{
  content:""; position:absolute; inset:0; border-radius:50%;
  background:conic-gradient(from 0deg, transparent 0 250deg, rgba(16,185,129,.7) 360deg);
  animation:lt-spin 1s linear infinite;
}
.lt-radar::after{content:""; position:absolute; left:50%; top:0; bottom:0; width:1px; background:rgba(16,185,129,.18)}
.lt-radar-dot{position:absolute; left:50%; top:50%; width:10px; height:10px; margin:-5px; border-radius:50%; background:var(--em-hi); box-shadow:0 0 18px var(--em-hi); z-index:1}
.lt-scan-text{margin:0; font-size:17px; font-weight:550; animation:lt-fade .3s ease-out}
.lt-progress{width:min(240px,60vw); height:3px; border-radius:3px; background:rgba(255,255,255,.1); overflow:hidden}
.lt-progress span{display:block; height:100%; width:100%; background:var(--em); transform-origin:left; animation:lt-fill linear forwards}

/* ---------- result: three stacked cards ---------- */
.lt-results{width:100%; max-width:440px; display:flex; flex-direction:column; gap:14px}
.lt-card{
  width:100%; padding:22px; border-radius:26px; outline:none;
  background:linear-gradient(180deg, rgba(255,255,255,.07), rgba(255,255,255,.028));
  border:1px solid var(--glass-line); -webkit-backdrop-filter:blur(22px); backdrop-filter:blur(22px);
  box-shadow:0 24px 60px rgba(0,0,0,.45), inset 0 1px 0 rgba(255,255,255,.08);
}
.lt-card-hero{
  padding:24px;
  background:linear-gradient(180deg, rgba(16,185,129,.16), rgba(255,255,255,.03));
  border-color:rgba(16,185,129,.35);
}
.lt-reveal{animation:lt-in .4s cubic-bezier(.2,.8,.2,1) both}
.lt-card-title{margin:0 0 14px; font-size:16px; font-weight:650; letter-spacing:-.01em}
.lt-card-head{display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap}
.lt-pill{
  padding:6px 12px; border-radius:999px; font-size:13px; font-weight:600;
  border:1px solid rgba(16,185,129,.5); color:var(--em-hi); background:rgba(16,185,129,.1);
}
.lt-pill[data-kind="direct"]{background:var(--em); color:#04130d; border-color:var(--em)}
.lt-muted{color:var(--muted); font-size:13px}
.lt-label{color:var(--muted); font-size:13px}

.lt-hero-grid{display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:22px}
.lt-hero-col{display:flex; flex-direction:column; align-items:flex-start; gap:4px; min-width:0}
.lt-hero-right{align-items:flex-end; text-align:right}
.lt-big{font-size:clamp(30px,9vw,40px); line-height:1.05; font-weight:750; letter-spacing:-.035em; color:#fff; white-space:nowrap}
.lt-big-em{color:var(--em-hi)}
.lt-plat{margin-top:6px; padding:4px 12px; border-radius:999px; font-size:13px; font-weight:700; color:#04130d; background:var(--em-hi)}
.lt-total-line{margin:20px 0 0; padding-top:16px; border-top:1px solid var(--glass-line); font-size:15px; color:var(--muted); display:flex; justify-content:space-between; gap:12px}
.lt-total-line strong{color:#fff; font-size:18px; font-weight:700}

.lt-tiles{display:grid; gap:10px; grid-template-columns:repeat(2,1fr)}
.lt-tiles[data-count="3"]{grid-template-columns:repeat(3,1fr)}
.lt-tile{display:flex; flex-direction:column; gap:4px; padding:14px 12px; border-radius:16px; background:rgba(255,255,255,.045); border:1px solid rgba(255,255,255,.07); min-width:0}
.lt-tile-label{font-size:12px; color:var(--muted); line-height:1.25}
.lt-tile-val{font-size:clamp(18px,5.2vw,22px); font-weight:700; letter-spacing:-.02em; color:#fff; line-height:1.15}
.lt-tile-sub{font-size:12px; color:var(--muted); line-height:1.25}
.lt-tile-key{background:rgba(16,185,129,.13); border-color:rgba(16,185,129,.4)}
.lt-tile-key .lt-tile-val{color:var(--em-hi)}

/* journey timeline */
.lt-tl{list-style:none; margin:0; padding:0}
.lt-tl-item{position:relative; padding:0 0 0 30px; display:flex; flex-wrap:wrap; align-items:center; gap:4px 10px; min-height:30px}
.lt-tl-item::before{content:""; position:absolute; left:7px; top:0; bottom:0; width:2px; background:rgba(16,185,129,.45)}
.lt-tl-item[data-pos="first"]::before{top:14px}
.lt-tl-item[data-pos="last"]::before{bottom:auto; height:14px}
.lt-tl-stop{padding-top:2px; padding-bottom:2px}
.lt-tl-stop::after{content:""; position:absolute; left:1px; top:8px; width:14px; height:14px; border-radius:50%; background:var(--bg); border:3px solid var(--em-hi)}
.lt-tl-time{font-size:17px; font-weight:700; color:#fff; min-width:78px}
.lt-tl-name{font-size:15px; color:var(--ink)}
.lt-chip-plat{padding:2px 10px; border-radius:999px; font-size:12px; font-weight:700; color:var(--em-hi); border:1px solid rgba(52,211,153,.5)}
.lt-tl-ride{padding-top:12px; padding-bottom:12px}
.lt-late{font-size:13px; color:#fbbf24}
.lt-note-line{flex-basis:100%; font-size:13px; color:var(--muted)}
.lt-tl-change{padding-top:8px; padding-bottom:8px}
.lt-change-chip{padding:6px 12px; border-radius:12px; font-size:13px; line-height:1.4; color:var(--em-hi); background:rgba(16,185,129,.1); border:1px dashed rgba(52,211,153,.55)}
.lt-note{margin:14px 0 0; font-size:13px; color:var(--muted); line-height:1.45}

.lt-badge{padding:4px 10px; border-radius:999px; font-size:12px; font-weight:700; border:1px solid transparent; white-space:nowrap}
.lt-badge[data-speed="slow"]{color:#c3cfca; border-color:rgba(255,255,255,.22); background:rgba(255,255,255,.05)}
.lt-badge[data-speed="fast"]{color:var(--em-hi); border-color:rgba(52,211,153,.55); background:rgba(16,185,129,.12)}
.lt-badge[data-speed="double-fast"]{color:#04130d; background:var(--em-hi); border-color:var(--em-hi)}

.lt-actions{display:grid; grid-template-columns:1fr 1fr; gap:10px}
.lt-again{
  width:100%; padding:14px; border-radius:14px; font:inherit; font-weight:600; font-size:15px; cursor:pointer;
  color:var(--ink); background:rgba(255,255,255,.06); border:1px solid var(--glass-line); transition:background .15s ease;
}
.lt-again:hover{background:rgba(255,255,255,.1)}
.lt-again:focus-visible,.lt-link:focus-visible{outline:2px solid var(--em-hi); outline-offset:3px}
.lt-link{align-self:center; padding:4px 0; background:none; border:0; font:inherit; font-size:14px; font-weight:600; color:var(--em-hi); cursor:pointer; text-decoration:underline; text-underline-offset:3px}
.lt-empty h2{margin:6px 0 8px; font-size:24px; letter-spacing:-.02em}
.lt-empty p{margin:0 0 20px; color:var(--muted); line-height:1.55; font-size:15px}

/* ---------- sheets ---------- */
.lt-backdrop{
  position:fixed; inset:0; z-index:40; display:flex; align-items:flex-end; justify-content:center; padding:16px;
  background:rgba(4,7,6,.6); -webkit-backdrop-filter:blur(8px); backdrop-filter:blur(8px); animation:lt-fade .2s ease-out;
}
@media (min-width:640px){.lt-backdrop{align-items:center}}
.lt-sheet{
  width:100%; max-width:460px; max-height:calc(100dvh - 32px); overflow-y:auto; padding:24px; border-radius:28px; outline:none;
  background:linear-gradient(180deg, rgba(24,30,28,.97), rgba(14,18,17,.97));
  border:1px solid var(--glass-line); box-shadow:0 30px 80px rgba(0,0,0,.6), inset 0 1px 0 rgba(255,255,255,.08);
  animation:lt-in .3s cubic-bezier(.2,.8,.2,1);
}
.lt-sheet-tall{max-width:520px; height:calc(100dvh - 32px); max-height:720px; display:flex; flex-direction:column; overflow:hidden}
.lt-sheet-head{flex:none}
.lt-list-scroll{flex:1; min-height:0; overflow-y:auto; margin:0 -6px; padding:2px 6px}
.lt-sheet-tall > .lt-primary,.lt-sheet-tall > .lt-sheet-actions{flex:none}
.lt-sheet-title{margin:0; font-size:22px; font-weight:650; letter-spacing:-.02em}
.lt-options{display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-top:18px}
@media (max-width:420px){.lt-options{grid-template-columns:1fr}}
.lt-option{
  position:relative; display:flex; flex-direction:column; gap:3px; padding:16px 14px 16px 46px; border-radius:16px; cursor:pointer;
  background:rgba(255,255,255,.04); border:1px solid var(--glass-line); transition:border-color .15s ease, background .15s ease;
}
.lt-option input{position:absolute; inset:0; width:100%; height:100%; margin:0; opacity:0; cursor:pointer}
.lt-option::before{content:""; position:absolute; left:16px; top:18px; width:18px; height:18px; border-radius:50%; border:2px solid rgba(255,255,255,.3)}
.lt-option[data-active="true"]{border-color:var(--em); background:rgba(16,185,129,.1)}
.lt-option[data-active="true"]::before{border-color:var(--em-hi); background:radial-gradient(circle, var(--em-hi) 0 4px, transparent 5px)}
.lt-option:focus-within{outline:2px solid var(--em-hi); outline-offset:2px}
.lt-option-title{font-size:15px; font-weight:650}
.lt-option-sub{font-size:13px; color:var(--muted)}
.lt-eta{margin-top:20px; display:flex; flex-direction:column; gap:12px}
.lt-stepper{display:flex; align-items:center; justify-content:space-between; gap:12px}
.lt-stepper output{font-size:28px; font-weight:700; letter-spacing:-.02em}
.lt-stepper button{
  width:46px; height:46px; border-radius:50%; font:inherit; font-size:22px; line-height:1; cursor:pointer; color:var(--ink);
  background:rgba(255,255,255,.07); border:1px solid var(--glass-line);
}
.lt-stepper button:disabled{opacity:.35; cursor:default}
.lt-stepper button:focus-visible,.lt-chip:focus-visible,.lt-primary:focus-visible,.lt-opt:focus-visible,.lt-file-main:focus-visible,.lt-mini:focus-visible,.lt-search:focus-visible{outline:2px solid var(--em-hi); outline-offset:2px}
.lt-chips{display:flex; flex-wrap:wrap; gap:8px}
.lt-chip{
  padding:8px 14px; border-radius:999px; font:inherit; font-size:14px; cursor:pointer; color:var(--ink);
  background:rgba(255,255,255,.05); border:1px solid var(--glass-line);
}
.lt-chip[aria-pressed="true"]{background:rgba(16,185,129,.18); border-color:var(--em); color:var(--em-hi)}
.lt-delay-switch{margin-top:20px}
.lt-sheet-note{margin:20px 0 0; font-size:14px; color:var(--muted)}
.lt-tight{margin:6px 0 14px}
.lt-warn{margin:0 0 12px; padding:10px 12px; border-radius:12px; font-size:13px; color:#fbbf24; background:rgba(251,191,36,.08); border:1px solid rgba(251,191,36,.35)}
.lt-sheet-actions{display:grid; grid-template-columns:1fr 1.6fr; gap:10px; margin-top:18px}
.lt-primary{
  padding:14px; border-radius:14px; font:inherit; font-weight:700; font-size:15px; cursor:pointer; color:#04130d;
  background:var(--em); border:1px solid var(--em-hi); transition:filter .15s ease;
}
.lt-primary:hover{filter:brightness(1.1)}
.lt-full{width:100%; margin-top:16px}

.lt-search{
  width:100%; padding:12px 14px; border-radius:14px; font:inherit; font-size:16px; color:var(--ink); color-scheme:dark;
  background:rgba(255,255,255,.07); border:1px solid var(--glass-line);
}
.lt-search::placeholder{color:var(--muted)}
.lt-sort{margin:12px 0 14px}
.lt-count{margin:-4px 0 12px; font-size:13px; color:var(--muted)}
.lt-list{list-style:none; margin:0; padding:0; display:flex; flex-direction:column; gap:8px}
.lt-opt{
  position:relative; width:100%; display:flex; flex-direction:column; gap:8px; padding:14px; border-radius:16px; text-align:left; font:inherit; cursor:pointer; color:var(--ink);
  background:rgba(255,255,255,.04); border:1px solid var(--glass-line); transition:border-color .15s ease, background .15s ease;
}
.lt-opt:hover{background:rgba(255,255,255,.08)}
.lt-opt[data-active="true"]{border-color:var(--em); background:rgba(16,185,129,.12)}
.lt-opt-top{display:flex; justify-content:space-between; align-items:baseline; gap:12px}
.lt-opt-time{font-size:22px; font-weight:700; letter-spacing:-.02em}
.lt-opt-arrive{font-size:14px; font-weight:600; color:var(--em-hi)}
.lt-opt-meta{display:flex; flex-wrap:wrap; gap:4px 12px; font-size:13px; color:var(--muted)}
.lt-tags{display:flex; gap:6px}
.lt-tag{padding:3px 10px; border-radius:999px; font-size:12px; font-weight:700; color:var(--em-hi); border:1px solid rgba(52,211,153,.5)}
.lt-tag-solid{background:var(--em-hi); color:#04130d; border-color:var(--em-hi)}

.lt-file{display:flex; flex-direction:column; gap:8px; padding:12px; border-radius:16px; background:rgba(255,255,255,.04); border:1px solid var(--glass-line)}
.lt-file[data-active="true"]{border-color:var(--em); background:rgba(16,185,129,.1)}
.lt-file-main{display:flex; flex-direction:column; align-items:flex-start; gap:4px; width:100%; padding:2px; text-align:left; font:inherit; color:var(--ink); background:none; border:0; cursor:pointer; border-radius:10px}
.lt-file-name{font-size:15px; font-weight:650; overflow-wrap:anywhere}
.lt-file-meta{font-size:13px; color:var(--muted)}
.lt-file-actions{display:flex; justify-content:flex-end; align-items:center; gap:8px; padding-top:8px; border-top:1px solid rgba(255,255,255,.07)}
.lt-confirm{justify-content:space-between; font-size:14px}
.lt-mini{
  padding:7px 14px; border-radius:999px; font:inherit; font-size:13px; font-weight:600; cursor:pointer; color:var(--ink);
  background:rgba(255,255,255,.06); border:1px solid var(--glass-line);
}
.lt-mini:hover{background:rgba(255,255,255,.11)}
.lt-mini-danger{color:#fca5a5; border-color:rgba(248,113,113,.4)}
.lt-confirm .lt-mini-danger{background:rgba(248,113,113,.18)}

.lt-dev{
  position:fixed; left:0; right:0; bottom:0; z-index:20;
  display:flex; align-items:center; justify-content:center; flex-wrap:wrap; gap:10px 18px;
  padding:12px 16px calc(12px + env(safe-area-inset-bottom));
  background:rgba(11,14,13,.72); -webkit-backdrop-filter:blur(16px); backdrop-filter:blur(16px);
  border-top:1px solid rgba(255,255,255,.08);
}
.lt-switch{display:inline-flex; align-items:center; gap:10px; cursor:pointer; font-size:14px; position:relative}
.lt-switch input{position:absolute; opacity:0; width:42px; height:24px; margin:0; cursor:pointer}
.lt-track{position:relative; width:42px; height:24px; border-radius:999px; background:rgba(255,255,255,.14); transition:background .2s ease; flex:none}
.lt-track::after{content:""; position:absolute; top:3px; left:3px; width:18px; height:18px; border-radius:50%; background:#fff; transition:transform .2s ease}
.lt-switch input:checked + .lt-track{background:var(--em)}
.lt-switch input:checked + .lt-track::after{transform:translateX(18px)}
.lt-switch input:focus-visible + .lt-track{outline:2px solid var(--em-hi); outline-offset:3px}
.lt-time-picker{display:flex; align-items:center; gap:6px}
.lt-colon{color:var(--muted); font-weight:700}
.lt-time-picker select{
  appearance:none; -webkit-appearance:none; color-scheme:dark; cursor:pointer;
  padding:9px 12px; border-radius:10px; font:inherit; font-size:16px; color:var(--ink);
  background:rgba(255,255,255,.07); border:1px solid var(--glass-line);
}
.lt-time-picker select:focus-visible{outline:2px solid var(--em-hi); outline-offset:2px}
.lt-time-picker option{background:#0f1412; color:var(--ink)}
.lt-live{min-height:20px}

@keyframes lt-pulse{0%{transform:scale(1); opacity:.7} 100%{transform:scale(1.28); opacity:0}}
@keyframes lt-spin{to{transform:rotate(360deg)}}
@keyframes lt-fade{from{opacity:0} to{opacity:1}}
@keyframes lt-in{from{opacity:0; transform:translateY(12px) scale(.985)} to{opacity:1; transform:none}}
@keyframes lt-fill{from{transform:scaleX(0)} to{transform:scaleX(1)}}

@media (max-width:420px){
  .lt-card{padding:18px; border-radius:22px}
  .lt-card-hero{padding:20px}
  .lt-sheet{padding:20px}
  .lt-tile{padding:12px 10px}
}
@media (prefers-reduced-motion:reduce){
  .lt-cta::after,.lt-radar::before{animation:none}
  .lt-reveal,.lt-overlay,.lt-scan-text,.lt-backdrop,.lt-sheet{animation:none}
  .lt-cta,.lt-again,.lt-track,.lt-track::after,.lt-option,.lt-primary,.lt-opt,.lt-ghost{transition:none}
}
`;
