import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import {
  Home, Radio, BatteryCharging, Clock, Upload, Search, Menu, X,
  AlertTriangle, PowerOff, Antenna, RotateCcw, ChevronDown, ChevronUp,
  FileSpreadsheet, Plus, Trash2, Settings, ArrowLeft, History,
  Lock, Unlock, Save, Check, Battery, BarChart3,
} from "lucide-react";
import {
  ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";

/* ----------------------------------------------------------------------
   공통 유틸
------------------------------------------------------------------------- */
const norm = (s) => String(s ?? "").replace(/\s+/g, "").replace(/[()]/g, "").toLowerCase();

/* --------------------------- Supabase (실서비스 배포용 저장소) --------------------------- */
// Vercel에 배포할 때 프로젝트 설정 > Environment Variables 에 아래 두 값을 넣어주세요.
// VITE_SUPABASE_URL, VITE_SUPABASE_ANON_KEY (Supabase 대시보드 > Project Settings > API 에서 확인)
// .trim()으로 앞뒤 공백/줄바꿈을 제거하고, 혹시 URL에 /rest/v1 같은 경로가 실수로 같이 들어가 있으면
// (라이브러리가 또 /rest/v1을 붙이면서 "v1/rest/v1"처럼 중복돼 404가 난다) 프로젝트 루트 주소만 남긴다.
const SUPABASE_URL = String(import.meta.env.VITE_SUPABASE_URL || "")
  .trim()
  .replace(/\/(rest|auth|storage|realtime|functions)\/v\d+.*$/i, "") // 실수로 딸려 들어간 API 경로 제거
  .replace(/\/+$/, ""); // 끝 슬래시 제거
const SUPABASE_ANON_KEY = String(import.meta.env.VITE_SUPABASE_ANON_KEY || "").trim();
if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
  console.error("[supabase] VITE_SUPABASE_URL 또는 VITE_SUPABASE_ANON_KEY가 비어있습니다. Vercel 환경변수를 확인하세요.");
} else if (!/^https:\/\/[a-z0-9-]+\.supabase\.co$/i.test(SUPABASE_URL)) {
  console.error(`[supabase] VITE_SUPABASE_URL 형식이 예상과 달라요: "${SUPABASE_URL}" — https://프로젝트ref.supabase.co 형태여야 합니다.`);
}
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

/** window.storage와 동일한 모양({key, value})으로 맞춘 Supabase 기반 key-value 저장소.
 *  Claude 아티팩트 밖(정식 배포)에서는 window.storage가 없으므로 이걸 대신 쓴다.
 *  에러는 콘솔에 그대로 찍고 위로 던진다 — 조용히 삼키면 "저장 안 됐는데 성공한 것처럼 보이는" 문제가 생긴다. */
const kv = {
  async get(key) {
    const { data, error } = await supabase.from("app_kv").select("value").eq("key", key).maybeSingle();
    if (error) { console.error("[kv.get]", key, error); throw error; }
    if (!data) return null;
    return { key, value: data.value };
  },
  async set(key, value) {
    const { error } = await supabase.from("app_kv").upsert({ key, value, updated_at: new Date().toISOString() });
    if (error) { console.error("[kv.set]", key, error); throw error; }
    return { key, value };
  },
  async delete(key) {
    const { error } = await supabase.from("app_kv").delete().eq("key", key);
    if (error) { console.error("[kv.delete]", key, error); throw error; }
    return { key, deleted: true };
  },
};

/** 통합시설코드 전용 정규화. 엑셀이 "0071" 같은 코드를 숫자로 읽어 71로 바꿔버리는 경우가 있어,
 *  숫자로만 이루어진 코드는 앞자리 0을 제거하고 비교해 이런 파일 간 표기 차이에도 매칭되게 한다. */
function normCode(s) {
  const v = norm(s);
  if (/^\d+$/.test(v)) return v.replace(/^0+/, "") || "0";
  return v;
}

/* ----------------------------------------------------------------------------------------
   기지국·축전지 기본정보 전용 엑셀 — 고정 열(컬럼) 매핑
   양식이 항상 동일하다는 전제로, 헤더 텍스트가 아니라 "열 위치"로 직접 읽는다.
   (정확한 열 매핑은 아래 BASE_COLUMNS 정의부 주석 참고)
----------------------------------------------------------------------------------------- */
function colLetterToIndex(letter) {
  let idx = 0;
  for (const ch of letter.toUpperCase()) idx = idx * 26 + (ch.charCodeAt(0) - 64);
  return idx - 1; // 0-based
}

function indexToColLetter(idx) {
  let n = idx + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 실제 시트에서 읽힌 원본 값을 열 문자와 함께 그대로 보여주기 위한 진단용 스냅샷. */
function buildRawPreview(headerRow, sampleRows) {
  const width = Math.max(headerRow?.length || 0, ...sampleRows.map((r) => r?.length || 0), 1);
  return Array.from({ length: width }, (_, idx) => ({
    letter: indexToColLetter(idx),
    header: headerRow?.[idx] ?? "",
    samples: sampleRows.map((r) => r?.[idx] ?? ""),
  }));
}
/** 기본적으로 첫 번째 시트를 그대로 사용한다(기존 동작 유지). 첫 시트가 사실상 비어있을 때만
 *  다른 시트 중 값이 가장 많이 채워진 시트를 찾아 대신 사용한다. */
function pickBestSheetAoa(wb) {
  const readAoa = (name) => {
    const sheet = wb.Sheets[name];
    return sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) : [];
  };
  const scoreOf = (aoa) => aoa.filter((row) => Array.isArray(row) && row.some((c) => c !== "" && c !== undefined && c !== null)).length;

  const firstName = wb.SheetNames[0];
  const firstAoa = firstName ? readAoa(firstName) : [];
  if (scoreOf(firstAoa) > 3) return firstAoa; // 첫 시트에 데이터가 충분히 있으면 그대로 사용

  let best = firstAoa;
  let bestScore = scoreOf(firstAoa);
  wb.SheetNames.forEach((name) => {
    if (name === firstName) return;
    const aoa = readAoa(name);
    const score = scoreOf(aoa);
    if (score > bestScore) { bestScore = score; best = aoa; }
  });
  return best;
}

/** 병합된 셀(같은 값이 첫 행에만 있고 아래 행은 비어있는 경우)을 대비해, 지정한 열들의 빈 칸을 바로 위 값으로 채운다. */
function forwardFillColumns(aoa, letters) {
  const idxs = letters.map(colLetterToIndex);
  const last = {};
  return aoa.map((row) => {
    const newRow = Array.isArray(row) ? [...row] : [];
    idxs.forEach((idx) => {
      const v = newRow[idx];
      if (v === "" || v === undefined || v === null) {
        if (last[idx] !== undefined) newRow[idx] = last[idx];
      } else {
        last[idx] = v;
      }
    });
    return newRow;
  });
}

/** 지정한 열에 그 열의 라벨 텍스트 자체가 값으로 들어있는 동안은 헤더 행으로 보고 건너뛴다.
 *  (예: SITE별 파일처럼 헤더가 2줄이라 "5G"/"4G"/"3G" 같은 라벨이 데이터 자리까지 내려와 있는 경우 대응) */
function countHeaderRows(aoaRaw, checks) {
  let i = 0;
  while (i < aoaRaw.length) {
    const row = aoaRaw[i] || [];
    const isHeaderish = checks.some(({ letter, labels }) => {
      const v = String(row[colLetterToIndex(letter)] ?? "").trim().toLowerCase();
      return v && labels.some((l) => v === l.toLowerCase());
    });
    if (!isHeaderish) break;
    i++;
  }
  return Math.max(i, 1); // 최소 1행은 헤더로 간주
}

/** 헤더가 2줄 이상인 파일에서, 컬럼 라벨 텍스트 자체가 데이터 행에 또 섞여 들어온 경우를 걸러낸다.
 *  예: 국소명 칸에 "국소명"/"공용대표시설명", 주소 칸에 "주소"가 그대로 들어있으면 실제 데이터가 아니라 헤더 잔재로 본다. */
function isHeaderEchoRow(row) {
  const echoChecks = [
    { field: "국소명", labels: ["국소명", "공용대표시설명"] },
    { field: "주소", labels: ["주소", "주소2"] },
    { field: "팀", labels: ["팀", "현장운용팀"] },
    { field: "정류기모델", labels: ["정류기모델", "정류기모델명", "모델명"] },
  ];
  let matches = 0;
  for (const { field, labels } of echoChecks) {
    const v = norm(row[field]);
    if (v && labels.some((l) => v === norm(l))) matches += 1;
  }
  return matches >= 2; // 두 개 이상 열이 동시에 라벨 텍스트면 헤더 잔재 행으로 확정
}

/* ----------------------------------------------------------------------------------------
   기지국·축전지 기본정보 — 하나로 합쳐진 엑셀, 고정 열(컬럼) 매핑
   1행은 목차라서 무시.
   F=통합시설코드 G=국소명 I=본부 J=SKT팀 M=팀(현장운용팀) N=주소 S=국사형태
   CC=5G CD=4G CE=3G(계는 자동 합산)
   CH=5G ARRU(식) CI=RU(식) CJ=5G L9TU(식) CK=링MUX RT 수용 RU/L9TU CM=중계기
   ACTA 내부저항측정 이력(정류기별):
   X=정류기번호 Z=정류기모델 CB=서비스(DU/W) AK=부하전류 AX=축전지번호 BC=규격(Ah, 재고현황과 무관) BD=전압구분(2V/12V)
   BN=축전지상태 BO=양호 BP=열화 BQ=열화2 BR=불량 BT=내부저항측정일시 BW=대개체여부
   BV=불량여부(대시보드 집계 전용). 값이 "대상X" 또는 "폐국"인 행은 모든 집계에서 제외한다(전제조건).
----------------------------------------------------------------------------------------- */
const BASE_COLUMNS = {
  통합시설코드: "F", 국소명: "G", 본부: "I", SKT팀: "J", 팀: "M", 주소: "N", 국사형태: "S",
  "5G": "CC", "4G": "CD", "3G": "CE",
  ARRU_5G: "CH", RU: "CI", L9TU_5G: "CJ", 링MUX_RT_RU_L9TU: "CK", 중계기: "CM",
  정류기번호: "X", 정류기모델: "Z", 서비스: "CB", 부하전류: "AK", 축전지번호: "AX", 규격: "BC", 전압: "BD",
  축전지상태: "BN", 양호: "BO", 열화: "BP", 열화2: "BQ", 불량: "BR", 측정일시: "BT",
  대개체여부: "BW", 불량여부: "BV",
};

/* ----------------------------------------------------------------------------------------
   Backup 시간 산출 — 용량산출계수 C 표 (셀 방전종료전압 1.7V 기준)
   출처: 사내 축전지 용량산정 기준 (용량환산시간K ÷ 보수율80% × 방전전류증가율1.11)
   1시간~48시간은 선형 외삽 추정값 포함.
----------------------------------------------------------------------------------------- */
const BACKUP_HOURS_AXIS = [1, 1.5, 2, 2.5, 3, 4, 6, 10, 12, 24, 36, 48];
const CAPACITY_FACTOR_C = {
  "2V": [2.94, 3.55, 4.10, 4.78, 4.87, 6.94, 9.35, 13.60, 16.23, 30.73, 45.22, 59.72],
  "12V": [3.84, 4.36, 4.88, 5.40, 5.55, 6.93, 9.43, 13.97, 14.43, 27.75, 40.17, 52.62],
};

/** 전압타입 문자열을 "2V"/"12V" 표 키로 정규화한다. */
/** 셀 값에 공백/줄바꿈이 섞여있거나 대소문자가 달라도("RT " , " rt" 등) 같은 값으로 취급하기 위한 비교 함수.
 *  서비스(CB열)·전압(BD열) 등 정확히 일치해야 하는 코드성 값 비교에 쓴다. */
function eqLoose(value, target) {
  return String(value ?? "").trim().toUpperCase() === String(target ?? "").trim().toUpperCase();
}

function normVoltageKey(v) {
  const s = String(v ?? "").trim();
  if (s.includes("12")) return "12V";
  if (s.includes("2")) return "2V";
  return null;
}

/** 특정 백업시간(h)에서의 용량산출계수 C를 표에서 선형보간(구간 밖은 양끝값 고정)한다. */
function interpolateC(voltageKey, hours) {
  const table = CAPACITY_FACTOR_C[voltageKey];
  if (!table) return null;
  const axis = BACKUP_HOURS_AXIS;
  if (hours <= axis[0]) return table[0];
  if (hours >= axis[axis.length - 1]) return table[axis.length - 1];
  for (let i = 0; i < axis.length - 1; i++) {
    if (hours >= axis[i] && hours <= axis[i + 1]) {
      const frac = (hours - axis[i]) / (axis[i + 1] - axis[i]);
      return table[i] + frac * (table[i + 1] - table[i]);
    }
  }
  return null;
}

/** 백업예상시간 = [총용량(Ah) × 백업기준적용시간(h)] ÷ [부하전류(A) × 용량산출계수C]
 *  C가 백업기준적용시간 자체에 의존하므로 수렴할 때까지 반복 계산한다. */
function calcBackupHours(capacityAh, loadA, voltage) {
  const d = calcBackupDetail(capacityAh, loadA, voltage);
  return d ? d.hours : null;
}

/** calcBackupHours와 같은 반복계산을 수행하되, 마지막에 수렴된 C값·전압키까지 함께 돌려준다.
 *  (검증용 상세보기에서 "어떤 공식으로 계산됐는지" 그대로 보여주기 위함) */
function calcBackupDetail(capacityAh, loadA, voltage) {
  const cap = Number(capacityAh);
  const load = Number(loadA);
  const voltageKey = normVoltageKey(voltage);
  if (!voltageKey || !cap || !load || cap <= 0 || load <= 0) return null;
  let t = 10; // 초기 추정치
  let c = null;
  for (let i = 0; i < 60; i++) {
    c = interpolateC(voltageKey, t);
    if (!c) return null;
    const next = (cap * t) / (load * c);
    if (Math.abs(next - t) < 0.01) { t = next; break; }
    t = next;
  }
  return { hours: Math.round(t * 10) / 10, c: Math.round(c * 100) / 100, voltageKey, cap, load };
}

const STATION_LEVEL_FIELDS = ["국소명", "본부", "SKT팀", "팀", "주소", "국사형태", "5G", "4G", "3G", "계", "ARRU_5G", "RU", "L9TU_5G", "링MUX_RT_RU_L9TU", "중계기"];

/** 고정 열 위치로 값을 읽어 행(정류기/축전지 1개) 단위의 평평한 배열을 만든다. */
function parseFixedColumnSheet(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const aoaRaw = pickBestSheetAoa(wb);
        const rawPreview = buildRawPreview(aoaRaw[0], aoaRaw.slice(1, 4));
        // 1행은 목차이므로 항상 1행만 건너뛴다. (혹시 라벨이 더 남아있으면 자동으로 추가 감지)
        const headerCount = countHeaderRows(aoaRaw, [
          { letter: "F", labels: ["통합시설코드", "공용대표시설코드", "공용대표"] },
          { letter: "G", labels: ["국소명", "공용대표시설명"] },
        ]);
        const dataRows = forwardFillColumns(aoaRaw.slice(headerCount), ["F", "G", "I", "J", "M", "N", "S"]);
        const get = (row, letter) => {
          const v = row[colLetterToIndex(letter)];
          return v === undefined || v === null ? "" : v;
        };
        const nonEmptyRows = dataRows.filter((row) => Array.isArray(row) && row.some((c) => c !== "" && c !== undefined && c !== null));
        const flat = nonEmptyRows.map((row) => {
          const out = {};
          Object.entries(BASE_COLUMNS).forEach(([field, letter]) => { out[field] = get(row, letter); });
          const g5 = Number(out["5G"]); const g4 = Number(out["4G"]); const g3 = Number(out["3G"]);
          const nums = [g5, g4, g3].filter((n) => !Number.isNaN(n));
          out["계"] = nums.length ? nums.reduce((s, n) => s + n, 0) : "";
          return out;
        }).filter((out) => !isHeaderEchoRow(out)); // 헤더가 여러 줄이라 라벨 텍스트가 데이터 자리에 또 들어온 행 제외
        const withCode = flat.filter((r) => r.통합시설코드 || r.국소명).length;
        resolve({ rows: flat, total: nonEmptyRows.length, withCode, rawPreview });
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/** 통합시설코드 기준으로 국소 단위 레코드로 묶고, 정류기(축전지) 측정 행들을 rectifiers 배열로 모은다. */
function groupStationRows(flatRows) {
  const map = new Map();
  (flatRows || []).forEach((r) => {
    const code = normCode(r.통합시설코드) || norm(r.국소명);
    if (!code) return;
    if (!map.has(code)) {
      const init = { rectifiers: [] };
      STATION_LEVEL_FIELDS.forEach((f) => { init[f] = r[f]; });
      init.통합시설코드 = r.통합시설코드;
      map.set(code, init);
    }
    const station = map.get(code);
    STATION_LEVEL_FIELDS.forEach((f) => {
      if (!station[f] && r[f]) station[f] = r[f];
    });
    if (r.정류기번호) {
      station.rectifiers.push({
        번호: r.정류기번호, 정류기모델: r.정류기모델, 서비스: r.서비스, 부하전류: r.부하전류, 축전지번호: r.축전지번호, 규격: r.규격, 전압: r.전압,
        축전지상태: r.축전지상태, 양호: r.양호, 열화: r.열화, 열화2: r.열화2, 불량: r.불량, 측정일시: r.측정일시,
        대개체여부: r.대개체여부, 불량여부: r.불량여부,
      });
    }
  });
  return [...map.values()];
}

/** 여러 업로드 파일(국소 단위 배열)을 통합시설코드 기준으로 합친다. 같은 국소면 정류기 목록도 중복 없이 함께 합쳐지고,
 *  그 외의 필드도 값이 있는 쪽으로 자동 채워진다. */
function combineStationLists(lists) {
  const map = new Map();
  lists.flat().forEach((station) => {
    if (!station) return;
    const code = normCode(station.통합시설코드) || norm(station.국소명);
    if (!code) return;
    if (!map.has(code)) {
      map.set(code, { ...station, rectifiers: [...(station.rectifiers || [])] });
      return;
    }
    const existing = map.get(code);
    Object.keys(station).forEach((f) => {
      if (f === "rectifiers") return;
      if (!existing[f] && station[f]) existing[f] = station[f];
    });
    const seen = new Set(existing.rectifiers.map((r) => `${r.번호}_${r.측정일시}`));
    (station.rectifiers || []).forEach((r) => {
      const key = `${r.번호}_${r.측정일시}`;
      if (!seen.has(key)) { existing.rectifiers.push(r); seen.add(key); }
    });
  });
  return [...map.values()];
}

const RECT_SUB = ["번호", "주요부하", "제원", "부하전류", "축전지번호", "양호", "열화1", "열화2", "불량", "측정일시"];
const STOCK_SUB = ["사업물자", "사용", "잔여"];
const ROUNDS = ["1차", "2차", "3차", "4차"];
const STOCK_ROWS = ["12V,100AH", "12V,200AH", "2V,300AH", "2V,600AH", "2V,1000AH"];

function buildAliasMap() {
  const map = {};
  const add = (canonical, aliases = []) => {
    map[norm(canonical)] = canonical;
    aliases.forEach((a) => (map[norm(a)] = canonical));
  };
  add("본부"); add("팀"); add("SKT팀"); add("국사형태"); add("통합시설코드");
  add("국소명", ["국소"]); add("IBN명칭", ["IBN"]); add("주소"); add("상태");
  add("불량유형", ["불량구분", "불량타입"]);
  add("Backup시간", ["백업시간"]);
  add("5G기지국"); add("4G기지국"); add("3G기지국"); add("W+RT");
  add("해당국사LTE모국"); add("5G_PON_COT", ["5G PON COT"]);
  add("링MUX_COT_Line", ["링MUX COT Line"]); add("SI대전용전");
  add("링MUX_RT식", ["링MUX RT(식)"]); add("3G축전지");
  add("5G_ARRU식", ["5G ARRU(식)"]); add("5G_LSH식", ["5G LSH(식)"]);
  add("4G_RRU식", ["4G RRU(식)"]); add("4G_L9TU식", ["4G L9TU(식)"]);
  add("수용된통합RO"); add("5+5G중계기", ["5+5G중계기(COT수용)"]);
  add("RT수용", ["RT수용(RU+L9TU)"]);
  for (let i = 1; i <= 4; i++) {
    RECT_SUB.forEach((sub) => add(`정류기${i}_${sub}`, [`정류기${i}${sub}`]));
  }
  ["2V", "12V"].forEach((v) => {
    ROUNDS.forEach((r) => {
      STOCK_SUB.forEach((sub) => add(`${v}_${r}_${sub}`, [`${v}${r}${sub}`]));
    });
  });
  add("품질개선팀"); add("공용대표명"); add("재고주소", ["주소2"]);
  add("W재고", ["W"]); add("DU재고", ["DU"]); add("5G재고");
  add("작업예정");
  return map;
}
const ALIAS_MAP = buildAliasMap();

function parseSheet(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const rows = raw.map((r) => {
          const out = {};
          Object.keys(r).forEach((h) => {
            const canon = ALIAS_MAP[norm(h)];
            out[canon || h] = r[h];
          });
          return out;
        });
        resolve(rows);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 여러 개의 엑셀(각각 행 배열)을 하나로 병합한다.
 * 1순위 키: 통합시설코드 / 2순위 키(통합시설코드가 없는 행): 국소명
 * 같은 키를 가진 행이 여러 파일에 나눠져 있으면 값이 채워진 컬럼끼리 서로 덮어써서 하나로 합친다.
 */
function mergeDatasets(datasets) {
  const byCode = new Map();
  const byName = new Map();
  const noKey = [];

  const mergeInto = (map, key, row) => {
    const prev = map.get(key);
    if (!prev) { map.set(key, { ...row }); return; }
    const merged = { ...prev };
    Object.keys(row).forEach((k) => {
      const v = row[k];
      if (v !== "" && v !== undefined && v !== null) merged[k] = v;
    });
    map.set(key, merged);
  };

  datasets.forEach((rows) => {
    (rows || []).forEach((row) => {
      const code = norm(row["통합시설코드"]);
      const name = norm(row["국소명"]);
      if (code) mergeInto(byCode, code, row);
      else if (name) mergeInto(byName, name, row);
      else noKey.push(row);
    });
  });

  return [...byCode.values(), ...byName.values(), ...noKey];
}

/* ------------------------- ACTA 내부저항측정 이력(변동 누적) ------------------------- */
const GRADE_PRIORITY = ["불량", "열화2", "열화", "양호"];
function deriveGrade(rect) {
  for (const g of GRADE_PRIORITY) {
    const v = rect[g];
    if (v !== "" && v !== undefined && v !== null) return g;
  }
  return null;
}
/** 모든 집계(대시보드/팀별현황/기지국현황 상태)의 전제조건: 불량여부(BV열)가 "대상X" 또는 "폐국"인 정류기, 그리고
 *  사용자가 지정해 제외한 정류기 모델명(Z열)은 집계에서 제외한다. 빈칸이면 제외 대상이 아니므로 그대로 집계에 포함되고,
 *  등급 정보가 없으면 stationStatus에서 자동으로 "양호"로 판정된다. */
function isCountableRect(rect, excludedModels) {
  const v = rect?.["불량여부"];
  if (eqLoose(v, "대상X") || eqLoose(v, "폐국")) return false;
  if (excludedModels && excludedModels.length && excludedModels.includes(rect?.["정류기모델"])) return false;
  return true;
}
/** 대시보드 수량 클릭 → 기지국 현황 드릴다운에 쓰이는 프리셋 조건과 국소 매칭 여부 판정.
 *  12V "전체"(서비스 미지정) 드릴다운은 대시보드 집계와 동일하게 CB열이 W/DU/5G인 것만 매칭한다. */
function stationMatchesPreset(station, preset, excludedModels) {
  if (!preset) return true;
  return (station.rectifiers || []).some((r) => {
    if (!isCountableRect(r, excludedModels)) return false;
    if (preset.voltage && !eqLoose(r["전압"], preset.voltage)) return false;
    if (preset.voltage === "12V" && !preset.service) {
      if (!["W", "DU", "5G"].some((s) => eqLoose(r["서비스"], s))) return false;
    }
    if (preset.badOnly && !eqLoose(r["불량여부"], "불량")) return false;
    if (preset.service && !eqLoose(r["서비스"], preset.service)) return false;
    return true;
  });
}
function actaKey(record, rectNo) {
  const base = normCode(record["통합시설코드"]) || norm(record["국소명"]) || "unknown";
  return `${base}__${norm(rectNo)}`;
}
/**
 * 정류기별 현재 등급(양호/열화/열화2/불량)을 이전 이력의 마지막 기록과 비교해,
 * 값이 달라졌을 때만 새 이력을 누적한다(변동이 없으면 추가하지 않음 → 중복 없이 변동 시점만 쌓임).
 */
function buildActaHistory(rows, prevHistory) {
  const history = { ...prevHistory };
  let changed = false;
  (rows || []).forEach((record) => {
    (record.rectifiers || []).forEach((rect) => {
      if (!rect.번호) return;
      const grade = deriveGrade(rect);
      if (!grade) return;
      const key = actaKey(record, rect.번호);
      const date = rect.측정일시 || new Date().toLocaleDateString("ko-KR");
      const list = history[key] ? [...history[key]] : [];
      const last = list[list.length - 1];
      if (!last || last.grade !== grade) {
        list.push({ date, grade, 축전지번호: rect.축전지번호, 부하전류: rect.부하전류 });
        history[key] = list;
        changed = true;
      }
    });
  });
  return { history, changed };
}

/** 엑셀(Blob)을 파일로 내려받는다. 아티팩트처럼 iframe 안에서 실행될 때는 <a>가 실제 DOM에
 *  붙어있어야 브라우저가 "정상적인 사용자 다운로드"로 인식해서 팝업/다운로드 차단에 걸리지 않는다. */
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
  // revoke를 곧바로 하면 일부 브라우저에서 다운로드가 시작되기 전에 URL이 무효화될 수 있어 살짝 늦춘다.
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function downloadTemplate() {
  const headers = [
    "본부", "팀", "SKT팀", "국사형태", "통합시설코드", "국소명", "IBN명칭", "주소", "상태", "불량유형", "Backup시간",
    "5G기지국", "4G기지국", "3G기지국", "W+RT",
    "해당국사LTE모국", "5G_PON_COT", "링MUX_COT_Line", "SI대전용전", "링MUX_RT식", "3G축전지",
    "5G_ARRU식", "5G_LSH식", "4G_RRU식", "4G_L9TU식", "수용된통합RO", "5+5G중계기", "RT수용",
  ];
  for (let i = 1; i <= 4; i++) RECT_SUB.forEach((s) => headers.push(`정류기${i}_${s}`));
  ["2V", "12V"].forEach((v) => ROUNDS.forEach((r) => STOCK_SUB.forEach((s) => headers.push(`${v}_${r}_${s}`))));
  headers.push("품질개선팀", "공용대표명", "재고주소", "W재고", "DU재고", "5G재고", "작업예정");

  const example = {
    본부: "충청Access", 팀: "대전", SKT팀: "충남", 국사형태: "실내형", 통합시설코드: "",
    국소명: "대전대후문2", IBN명칭: "대전동구-0071", 주소: "대전 동구 용운동 710",
    상태: "정상", Backup시간: "4.2",
    "5G기지국": 0, "4G기지국": 0, "3G기지국": 2, "W+RT": 2,
    해당국사LTE모국: "SI대전용전", "5G_PON_COT": 0, 링MUX_COT_Line: 5, SI대전용전: "SI대전용전",
    링MUX_RT식: "-", "3G축전지": "분리_정전시 3G Down",
    "5G_ARRU식": 0, "5G_LSH식": 0, "4G_RRU식": 0, "4G_L9TU식": 0,
    수용된통합RO: 45, "5+5G중계기": 0, RT수용: 117,
    정류기1_번호: "대전대후문2-1번정류기", 정류기1_주요부하: "W,WT", 정류기1_제원: "CRS-2400", 정류기1_부하전류: 123,
    정류기2_번호: "대전대후문2-2번정류기", 정류기2_주요부하: "PTS", 정류기2_제원: "SDPS-48N",
    정류기3_번호: "대전대후문2-3번정류기", 정류기3_주요부하: "Ring MUX", 정류기3_제원: "SKR-N48V-30AW",
  };
  const row = headers.map((h) => example[h] ?? "");
  const ws = XLSX.utils.aoa_to_sheet([headers, row]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "국소데이터");
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([buf], { type: "application/octet-stream" });
  downloadBlob(blob, "기지국_업로드_템플릿.xlsx");
}

/* ---------------------------- 기본(샘플) 데이터 ---------------------------- */
const SAMPLE_ROW = {
  통합시설코드: "", 국소명: "대전대후문2", 본부: "충청Access", SKT팀: "충남", 팀: "대전",
  주소: "대전 동구 용운동 710", 국사형태: "실내형",
  rectifiers: [
    { 번호: "대전대후문2-1번정류기", 정류기모델: "CRS-2400", 서비스: "DU", 부하전류: 123, 축전지번호: "B-001", 축전지상태: "양호", 양호: "20", 열화: "", 열화2: "", 불량: "", 측정일시: "2026-07-31", 대개체여부: "", 불량여부: "" },
    { 번호: "대전대후문2-2번정류기", 정류기모델: "SDPS-48N", 서비스: "W", 부하전류: "", 축전지번호: "B-002", 축전지상태: "", 양호: "", 열화: "3", 열화2: "", 불량: "", 측정일시: "2026-07-31", 대개체여부: "", 불량여부: "" },
  ],
};

const TEAM_CHART_SAMPLE = [
  { team: "대전품질개선팀", total: 320, bad: 18 },
  { team: "천안품질개선팀", total: 280, bad: 20 },
  { team: "서산품질개선팀", total: 150, bad: 9 },
  { team: "세종품질개선팀", total: 180, bad: 12 },
  { team: "충주품질개선팀", total: 100, bad: 6 },
  { team: "서청주품질개선팀", total: 60, bad: 4 },
  { team: "동청주품질개선팀", total: 60, bad: 6 },
].map((d) => ({ ...d, rate: +((d.bad / d.total) * 100).toFixed(1) }));

const NAV = [
  { key: "home", label: "홈", icon: Home },
  { key: "stations", label: "기지국 현황", icon: Antenna },
  { key: "battery", label: "축전지 재고 현황", icon: BatteryCharging },
  { key: "backup", label: "Backup 시간 확인", icon: Clock },
  { key: "admin", label: "관리자", icon: Settings },
];

/* ------------------------------- 작은 컴포넌트 ------------------------------- */
function StatCard({ icon: Icon, label, value, unit, tone }) {
  const tones = {
    blue: "bg-blue-50 text-blue-600",
    green: "bg-emerald-50 text-emerald-600",
    amber: "bg-amber-50 text-amber-600",
    red: "bg-red-50 text-red-600",
    rose: "bg-rose-50 text-rose-600",
  };
  return (
    <div className="flex items-center gap-3 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
      <div className={`flex h-11 w-11 shrink-0 items-center justify-center rounded-xl ${tones[tone]}`}>
        <Icon size={20} />
      </div>
      <div className="min-w-0">
        <p className="truncate text-xs text-slate-500">{label}</p>
        <p className="text-xl font-bold text-slate-800">
          {value.toLocaleString()}
          {unit && <span className="ml-0.5 text-sm font-medium text-slate-400">{unit}</span>}
        </p>
      </div>
    </div>
  );
}

function StatusBadge({ status }) {
  const map = {
    정상: "bg-emerald-50 text-emerald-600",
    양호: "bg-emerald-50 text-emerald-600",
    점검필요: "bg-amber-50 text-amber-600",
    열화: "bg-amber-50 text-amber-600",
    불량: "bg-red-50 text-red-600",
    조불량: "bg-red-50 text-red-600",
    다운: "bg-red-50 text-red-600",
    미측정: "bg-slate-100 text-slate-500",
  };
  return (
    <span className={`inline-block whitespace-nowrap rounded-full px-2.5 py-0.5 text-xs font-medium ${map[status] || "bg-slate-100 text-slate-500"}`}>
      {status || "미확인"}
    </span>
  );
}

function GradeBadge({ grade }) {
  const map = {
    양호: "bg-emerald-50 text-emerald-600",
    열화: "bg-amber-50 text-amber-600",
    열화2: "bg-orange-50 text-orange-600",
    불량: "bg-red-50 text-red-600",
  };
  return (
    <span className={`inline-block rounded-full px-2.5 py-0.5 text-xs font-medium ${map[grade] || "bg-slate-100 text-slate-500"}`}>
      {grade || "미확인"}
    </span>
  );
}

const GRID_COLS_SM = { 3: "sm:grid-cols-3", 4: "sm:grid-cols-4", 5: "sm:grid-cols-5", 6: "sm:grid-cols-6" };
function LabelValueRow({ pairs, cols }) {
  const gridClass = `grid-cols-2 ${GRID_COLS_SM[cols] || "sm:grid-cols-4"}`;
  return (
    <div className={`grid ${gridClass}`}>
      {pairs.map(([label, value], i) => (
        <div key={i} className="border border-slate-200 -ml-px -mt-px">
          <div className="bg-slate-50 px-2 py-1.5 text-[11px] font-medium text-slate-500">{label}</div>
          <div className="px-2 py-1.5 text-sm text-slate-800 break-words">{value === "" || value === undefined ? "-" : String(value)}</div>
        </div>
      ))}
    </div>
  );
}

/* --------------------------- 국소 상세 현황 (검색형) --------------------------- */
function LocationDetail({ rows, query, setQuery, actaHistory }) {
  const [historyView, setHistoryView] = useState(null); // { rectNo, label } | null
  const [selectedCode, setSelectedCode] = useState(null);
  const [inputText, setInputText] = useState(query); // 입력창은 따로 두고, "검색" 버튼을 눌러야 실제 검색어(query)에 반영된다.

  const trimmed = query.trim();
  const codeOf = (r) => normCode(r["통합시설코드"]) || norm(r["국소명"]);

  const runSearch = () => setQuery(inputText);
  const resetSearch = () => { setInputText(""); setQuery(""); };

  // F열(공용대표시설명=국소명) 기준으로 2글자 이상 포함되는 국소를 모두 후보로 보여준다.
  const matches = useMemo(() => {
    if (!rows.length) return [];
    if (!trimmed) return rows.slice(0, 1);
    if (trimmed.length < 2) return [];
    return rows.filter((r) => norm(r["국소명"]).includes(norm(trimmed)));
  }, [rows, trimmed]);

  useEffect(() => {
    if (!matches.length) { setSelectedCode(null); return; }
    const stillValid = matches.some((m) => codeOf(m) === selectedCode);
    if (!stillValid) setSelectedCode(codeOf(matches[0]));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [matches]);

  const record = matches.find((m) => codeOf(m) === selectedCode) || matches[0] || null;

  useEffect(() => { setHistoryView(null); }, [record]);

  const rectifiers = record?.rectifiers || [];
  const backupBankMap = useMemo(() => {
    const map = new Map();
    groupRectifierBanks(rectifiers).forEach((b) => {
      map.set(b.번호, calcBackupHours(b.규격합, b.부하전류, b.전압));
    });
    return map;
  }, [rectifiers]);
  const historyEntries = record && historyView ? (actaHistory?.[actaKey(record, historyView.rectNo)] || []) : [];

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      <div className="mb-2 flex flex-col gap-2 sm:flex-row sm:items-center">
        <div className="relative flex-1">
          <Search size={16} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            value={inputText}
            onChange={(e) => setInputText(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") runSearch(); }}
            placeholder="국소명 검색 (2글자 이상)"
            className="w-full rounded-lg border border-slate-200 bg-slate-50 py-2 pl-9 pr-3 text-sm outline-none focus:border-blue-400 focus:bg-white"
          />
        </div>
        <button
          onClick={runSearch}
          className="flex items-center justify-center gap-1 rounded-lg bg-blue-600 px-3 py-2 text-sm font-medium text-white hover:bg-blue-700"
        >
          <Search size={14} /> 검색
        </button>
        <button
          onClick={resetSearch}
          className="flex items-center justify-center gap-1 rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-500 hover:bg-slate-50"
        >
          <RotateCcw size={14} /> 초기화
        </button>
      </div>

      {trimmed.length === 1 && (
        <p className="mb-3 text-xs text-amber-600">2글자 이상 입력하면 검색됩니다.</p>
      )}

      {matches.length > 1 && (
        <div className="mb-3 flex flex-wrap gap-1.5">
          {matches.map((m) => {
            const code = codeOf(m);
            const active = code === selectedCode;
            return (
              <button key={code} onClick={() => setSelectedCode(code)}
                className={`rounded-full px-3 py-1 text-xs font-medium transition-colors ${
                  active ? "bg-blue-600 text-white" : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                }`}>
                {m["국소명"]}
              </button>
            );
          })}
        </div>
      )}

      {!record ? (
        <p className="py-10 text-center text-sm text-slate-400">
          {trimmed.length >= 2 ? "검색 결과가 없습니다. 국소명을 확인해주세요." : "국소명을 검색해주세요."}
        </p>
      ) : (
        <div className="space-y-4 overflow-x-auto">
          <LabelValueRow
            pairs={[
              ["본부", record["본부"]], ["팀", record["팀"]], ["SKT팀", record["SKT팀"]], ["국사형태", record["국사형태"]],
              ["통합시설코드", record["통합시설코드"]], ["국소명", record["국소명"]], ["주소", record["주소"]],
            ]}
          />
          <div>
            <p className="mb-1 text-xs font-semibold text-slate-500">시설구분</p>
            <LabelValueRow pairs={[["5G", record["5G"]], ["4G", record["4G"]], ["3G", record["3G"]], ["계", record["계"]]]} />
          </div>
          <div>
            <p className="mb-1 text-xs font-semibold text-slate-500">중계기</p>
            <LabelValueRow
              cols={5}
              pairs={[
                ["5G ARRU(식)", record["ARRU_5G"]], ["RU(식)", record["RU"]], ["5G L9TU(식)", record["L9TU_5G"]],
                ["링MUX RT 수용 RU/L9TU", record["링MUX_RT_RU_L9TU"]], ["중계기", record["중계기"]],
              ]}
            />
          </div>
          <div>
            {!historyView ? (
              <>
                <p className="mb-1 text-xs font-semibold text-slate-500">ACTA 내부저항측정 이력</p>
                <div className="overflow-x-auto rounded-lg border border-slate-200">
                  <table className="w-full min-w-[960px] whitespace-nowrap text-left text-xs">
                    <thead className="bg-slate-50 text-slate-500">
                      <tr>
                        <th className="px-2 py-2 font-medium">정류기번호</th>
                        <th className="px-2 py-2 font-medium">정류기모델</th>
                        <th className="px-2 py-2 font-medium">서비스</th>
                        <th className="px-2 py-2 font-medium">부하전류(A)</th>
                        <th className="px-2 py-2 font-medium">축전지번호</th>
                        <th className="px-2 py-2 font-medium">규격(Ah)</th>
                        <th className="px-2 py-2 font-medium">예상Backup(h)</th>
                        <th className="px-2 py-2 font-medium">축전지상태</th>
                        <th className="px-2 py-2 font-medium">양호</th>
                        <th className="px-2 py-2 font-medium">열화</th>
                        <th className="px-2 py-2 font-medium">열화2</th>
                        <th className="px-2 py-2 font-medium">불량</th>
                        <th className="px-2 py-2 font-medium">내부저항측정일시</th>
                        <th className="px-2 py-2 font-medium">대개체여부</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rectifiers.length === 0 && (
                        <tr><td colSpan={14} className="px-2 py-4 text-center text-slate-400">등록된 정류기 정보가 없습니다.</td></tr>
                      )}
                      {rectifiers.map((r, idx) => (
                        <tr key={`${r.번호}_${idx}`} className="border-t border-slate-100">
                          <td className="px-2 py-2">
                            <button
                              onClick={() => setHistoryView({ rectNo: r.번호, label: r.번호 })}
                              className="font-medium text-blue-600 hover:underline"
                              title="클릭하면 변동 이력을 확인할 수 있습니다"
                            >
                              {r.번호}
                            </button>
                          </td>
                          <td className="px-2 py-2">{r.정류기모델 || "-"}</td>
                          <td className="px-2 py-2">{r.서비스 || "-"}</td>
                          <td className="px-2 py-2">{r.부하전류 || "-"}</td>
                          <td className="px-2 py-2">{r.축전지번호 || "-"}</td>
                          <td className="px-2 py-2">{r.규격 || "-"}</td>
                          <td className="px-2 py-2 font-medium text-blue-600">
                            {(() => { const h = backupBankMap.get(r.번호); return h ? `${h}h` : "-"; })()}
                          </td>
                          <td className="px-2 py-2">{r.축전지상태 || "-"}</td>
                          <td className="px-2 py-2">{r.양호 || "-"}</td>
                          <td className="px-2 py-2">{r.열화 || "-"}</td>
                          <td className="px-2 py-2">{r.열화2 || "-"}</td>
                          <td className="px-2 py-2">{r.불량 || "-"}</td>
                          <td className="px-2 py-2">{r.측정일시 || "-"}</td>
                          <td className="px-2 py-2">{r.대개체여부 || "-"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            ) : (
              <div className="rounded-lg border border-slate-200">
                <div className="flex items-center justify-between border-b border-slate-100 px-3 py-2.5">
                  <button onClick={() => setHistoryView(null)} className="flex items-center gap-1 text-xs font-medium text-slate-500 hover:text-slate-700">
                    <ArrowLeft size={14} /> 목록으로
                  </button>
                  <p className="flex items-center gap-1.5 text-sm font-semibold text-slate-700">
                    <History size={14} className="text-blue-500" /> {historyView.label} 변동 이력
                  </p>
                  <span className="text-xs text-slate-400">{historyEntries.length}건</span>
                </div>
                {historyEntries.length === 0 ? (
                  <p className="px-3 py-8 text-center text-sm text-slate-400">
                    아직 기록된 변동 이력이 없습니다. 관리자 메뉴에서 기지국 기본정보를 새로 업로드하면 값이 바뀔 때마다 이력이 자동으로 쌓입니다.
                  </p>
                ) : (
                  <ul className="divide-y divide-slate-100">
                    {[...historyEntries].reverse().map((h, i) => (
                      <li key={i} className="flex items-center justify-between gap-3 px-3 py-2.5 text-sm">
                        <span className="text-slate-500">{h.date}</span>
                        <GradeBadge grade={h.grade} />
                        <span className="ml-auto text-xs text-slate-400">
                          {h.축전지번호 ? `축전지 ${h.축전지번호}` : ""}{h.부하전류 ? ` · 부하전류 ${h.부하전류}A` : ""}
                        </span>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* --------------------------------- 팀별 현황 --------------------------------- */
const VALID_TEAMS = [
  "대전품질개선팀", "천안품질개선팀", "서산품질개선팀", "세종품질개선팀",
  "충주품질개선팀", "서청주품질개선팀", "동청주품질개선팀",
];

function TeamStatus({ rows, excludedModels }) {
  const data = useMemo(() => {
    if (!rows.length) return TEAM_CHART_SAMPLE;
    const byTeam = {};
    rows.forEach((station) => {
      const t = station["팀"];
      if (!VALID_TEAMS.includes(t)) return; // 강릉/삼척 등 유효하지 않은 팀은 제외
      byTeam[t] = byTeam[t] || { team: t, total: 0, bad: 0 };
      (station.rectifiers || []).forEach((rect) => {
        if (!isCountableRect(rect, excludedModels)) return;
        // 홈 대시보드와 동일한 분류 기준: RT(1차) → 2V(제한 없음) → 12V(서비스가 W/DU/5G인 것만)
        // 이 세 조건에 해당하지 않는 행(전압 없음, 12V인데 서비스가 다른 경우 등)은 대시보드처럼 집계에서 제외한다.
        const isRt = eqLoose(rect["서비스"], "RT");
        const is2v = eqLoose(rect["전압"], "2V");
        const is12vCounted = eqLoose(rect["전압"], "12V") && ["W", "DU", "5G"].some((s) => eqLoose(rect["서비스"], s));
        if (!isRt && !is2v && !is12vCounted) return;
        byTeam[t].total += 1;
        if (eqLoose(rect["불량여부"], "불량")) byTeam[t].bad += 1;
      });
    });
    return Object.values(byTeam).map((d) => ({ ...d, rate: d.total ? +((d.bad / d.total) * 100).toFixed(1) : 0 }));
  }, [rows, excludedModels]);

  const totalCnt = data.reduce((s, d) => s + d.total, 0);
  const badCnt = data.reduce((s, d) => s + d.bad, 0);
  const overallRate = totalCnt ? ((badCnt / totalCnt) * 100).toFixed(1) : "0.0";

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <h3 className="text-sm font-semibold text-slate-700">팀별 현황</h3>
        <div className="flex gap-4 text-right text-xs">
          <div><p className="text-slate-400">전체 수량</p><p className="font-bold text-slate-700">{totalCnt.toLocaleString()} 조</p></div>
          <div><p className="text-slate-400">조불량 수량</p><p className="font-bold text-red-500">{badCnt.toLocaleString()} 조</p></div>
          <div><p className="text-slate-400">전체 불량률</p><p className="font-bold text-blue-600">{overallRate}%</p></div>
        </div>
      </div>
      <div className="h-72 w-full">
        <ResponsiveContainer width="100%" height="100%">
          <ComposedChart data={data} margin={{ top: 5, right: 10, bottom: 0, left: -10 }}>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f7" />
            <XAxis dataKey="team" tick={{ fontSize: 12, fill: "#64748b" }} />
            <YAxis yAxisId="left" tick={{ fontSize: 11, fill: "#94a3b8" }} />
            <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: "#94a3b8" }} unit="%" />
            <Tooltip />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar yAxisId="left" dataKey="total" name="전체 수량(조)" fill="#3b82f6" radius={[6, 6, 0, 0]} barSize={26} />
            <Bar yAxisId="left" dataKey="bad" name="조불량 수량(조)" fill="#f87171" radius={[6, 6, 0, 0]} barSize={26} />
            <Line yAxisId="right" type="monotone" dataKey="rate" name="불량률(%)" stroke="#10b981" strokeWidth={2} dot={{ r: 4 }} />
          </ComposedChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

/* -------------------------------- 기지국 표 -------------------------------- */
function stationStatus(station, excludedModels) {
  const rects = (station.rectifiers || []).filter((r) => isCountableRect(r, excludedModels));
  if (!rects.length) return "미측정";
  if (rects.some((r) => eqLoose(r["불량여부"], "불량"))) return "조불량";
  if (rects.some((r) => { const g = deriveGrade(r); return g === "열화" || g === "열화2"; })) return "열화";
  return "양호";
}

/** 국소의 정류기들에서 서로 다른 모델명(Z열)만 모아 콤마로 이어붙인다. (엑셀 내보내기용) */
function stationModels(station) {
  const models = [...new Set((station.rectifiers || []).map((r) => r.정류기모델).filter(Boolean))];
  return models.length ? models.join(", ") : "-";
}

/** 국소 1개를 정류기 모델명(Z열)별로 쪼개서, 모델마다 별도의 행으로 만든다.
 *  같은 국소에 sars-375, SDPS-48N-10A 두 모델이 있으면 이 국소는 2개 행이 된다.
 *  각 행의 rectifiers는 그 모델에 해당하는 정류기만 담아서, 상태·조치여부도 모델 단위로 판정되게 한다. */
function expandStationsByModel(stations) {
  const expanded = [];
  (stations || []).forEach((station) => {
    const rects = station.rectifiers || [];
    const models = [...new Set(rects.map((r) => r.정류기모델).filter(Boolean))];
    if (!models.length) {
      expanded.push({ ...station, rectifiers: rects, _model: null });
      return;
    }
    models.forEach((model) => {
      expanded.push({ ...station, rectifiers: rects.filter((r) => r.정류기모델 === model), _model: model });
    });
  });
  return expanded;
}



/** 국소의 정류기들에서 조치여부(BW열: 대개체필요/불용철거필요 등)를 모아 콤마로 이어붙인다. */
function stationAction(station) {
  const actions = [...new Set((station.rectifiers || []).map((r) => r.대개체여부).filter(Boolean))];
  return actions.length ? actions.join(", ") : "";
}

/** 국소(또는 정류기 모델별로 나뉜 행)의 정류기들에서 내부저항측정일시(BT열) 중 가장 최근 값 하나만 돌려준다. */
function stationMeasuredAt(station) {
  const dates = (station.rectifiers || []).map((r) => r.측정일시).filter(Boolean);
  if (!dates.length) return "";
  return dates.reduce((latest, d) => {
    const dLen = new Date(d).getTime();
    const latestLen = new Date(latest).getTime();
    if (Number.isNaN(dLen) || Number.isNaN(latestLen)) return String(d) > String(latest) ? d : latest; // 날짜로 못 읽으면 문자열 비교로 대체
    return dLen > latestLen ? d : latest;
  });
}

function ActionBadge({ text }) {
  if (!text) return <span className="text-slate-300">-</span>;
  const tone = text.includes("불용철거") ? "bg-violet-50 text-violet-600" : text.includes("대개체") ? "bg-amber-50 text-amber-600" : "bg-slate-100 text-slate-500";
  return <span className={`inline-block whitespace-nowrap rounded-full px-2.5 py-0.5 text-xs font-medium ${tone}`}>{text}</span>;
}

/** 기지국 현황 표를 엑셀(xlsx)로 내려받는다. */
/** 기지국 현황에서 화면에 보이는 대상(정류기 모델별로 나뉜 행) 기준으로, 처음 업로드한 엑셀과
 *  같은 항목 구성(통합시설코드~불량여부)으로 정류기 1개당 1행씩 엑셀을 만든다. */
const EXPORT_FIELDS = [
  ["통합시설코드", "통합시설코드"], ["국소명", "국소명"], ["본부", "본부"], ["SKT팀", "SKT팀"], ["팀", "팀(현장운용팀)"],
  ["주소", "주소"], ["국사형태", "국사형태"],
  ["5G", "5G"], ["4G", "4G"], ["3G", "3G"],
  ["ARRU_5G", "5G ARRU(식)"], ["RU", "RU(식)"], ["L9TU_5G", "5G L9TU(식)"],
  ["링MUX_RT_RU_L9TU", "링MUX RT 수용 RU/L9TU"], ["중계기", "중계기"],
];
const EXPORT_RECT_FIELDS = [
  ["번호", "정류기번호"], ["정류기모델", "정류기모델"], ["서비스", "서비스"], ["부하전류", "부하전류"],
  ["축전지번호", "축전지번호"], ["규격", "규격(Ah)"], ["전압", "전압"], ["축전지상태", "축전지상태"],
  ["양호", "양호"], ["열화", "열화"], ["열화2", "열화2"], ["불량", "불량"],
  ["측정일시", "내부저항측정일시"], ["대개체여부", "대개체여부"], ["불량여부", "불량여부"],
];

function exportStationsToExcel(stations, label, excludedModels) {
  const headers = [...EXPORT_FIELDS.map(([, h]) => h), ...EXPORT_RECT_FIELDS.map(([, h]) => h), "상태", "조치여부"];
  const aoa = [headers];
  stations.forEach((s) => {
    const stationCells = EXPORT_FIELDS.map(([f]) => s[f] ?? "");
    const status = stationStatus(s, excludedModels);
    const action = stationAction(s);
    const rects = s.rectifiers && s.rectifiers.length ? s.rectifiers : [{}];
    rects.forEach((r) => {
      const rectCells = EXPORT_RECT_FIELDS.map(([f]) => r[f] ?? "");
      aoa.push([...stationCells, ...rectCells, status, action]);
    });
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "기지국현황");
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([buf], { type: "application/octet-stream" });
  const safeLabel = (label || "기지국현황").replace(/[\\/:*?"<>|]/g, "_");
  downloadBlob(blob, `${safeLabel}_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/** 축전지 재고 현황(재고표 + 대상목록)을 엑셀로 내려받는다. */
function exportStockToExcel(stock, targets) {
  const stockUsage = (type) => (targets || []).filter((t) => t.종류 === type).length;
  const stockTotal = (type) => ROUNDS.reduce((s, round) => s + (Number(stock[`${round}_${type}`]) || 0), 0);

  const stockAoa = [
    ["구분", ...ROUNDS],
    ...STOCK_ROWS.map((type) => [type, ...ROUNDS.map((round) => stock[`${round}_${type}`] ?? "")]),
  ];
  // 합계/사용/잔여는 종류(열)별 값이라 위 표와 축이 달라 별도 표로 붙인다.
  const summaryAoa = [
    ["구분", ...STOCK_ROWS],
    ["합계", ...STOCK_ROWS.map((type) => stockTotal(type))],
    ["사용", ...STOCK_ROWS.map((type) => stockUsage(type))],
    ["잔여", ...STOCK_ROWS.map((type) => stockTotal(type) - stockUsage(type))],
  ];

  const targetHeaders = ["종류", "국소명", "주소", "현장운용팀", "W", "DU", "5G", "작업 예정(완료)일", "완료여부", "비고"];
  const targetAoa = [targetHeaders, ...(targets || []).map((t) => [
    t.종류 || "", t.국소명 || "", t.주소 || "", t.현장운용팀 || "", t.W || "", t.DU || "", t["5G"] || "", t.작업예정일 || "", t.완료여부 || "대기", t.비고 || "",
  ])];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(stockAoa), "재고(차수별)");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryAoa), "합계·사용·잔여");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(targetAoa), "대상목록");
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([buf], { type: "application/octet-stream" });
  downloadBlob(blob, `축전지재고현황_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function StationTable({ rows, compact, presetFilter, onClearPreset, excludedModels }) {
  const [q, setQ] = useState("");
  const [teamFilter, setTeamFilter] = useState("전체");
  const [statusFilter, setStatusFilter] = useState("전체");
  const [page, setPage] = useState(1);
  const perPage = compact ? 5 : 10;

  useEffect(() => { setPage(1); }, [presetFilter]);

  const teams = useMemo(() => ["전체", ...new Set(rows.map((r) => r["팀"]).filter(Boolean))], [rows]);

  // 국소 1개를 정류기 모델명별로 쪼갠 뒤 필터링한다 — 그래야 상태·조치여부도 모델 단위로 정확히 표시된다.
  const expandedRows = useMemo(() => expandStationsByModel(rows), [rows]);

  const filtered = useMemo(() => {
    return expandedRows.filter((r) => {
      const okQ = !q || norm(r["국소명"]).includes(norm(q)) || norm(r["주소"]).includes(norm(q));
      const okTeam = teamFilter === "전체" || r["팀"] === teamFilter;
      const okStatus = statusFilter === "전체" || stationStatus(r, excludedModels) === statusFilter;
      const okPreset = stationMatchesPreset(r, presetFilter, excludedModels);
      return okQ && okTeam && okStatus && okPreset;
    });
  }, [expandedRows, q, teamFilter, statusFilter, presetFilter, excludedModels]);

  const shown = compact ? filtered.slice(0, perPage) : filtered.slice((page - 1) * perPage, page * perPage);
  const totalPages = Math.max(1, Math.ceil(filtered.length / perPage));

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      {!compact && presetFilter && (
        <div className="mb-3 flex items-center justify-between gap-2 rounded-lg bg-blue-50 px-3 py-2 text-xs text-blue-700">
          <span>대시보드에서 선택한 대상만 보고 있어요: <b>{presetFilter.label}</b> · {filtered.length}건</span>
          <button onClick={onClearPreset} className="flex items-center gap-1 rounded-md bg-white px-2 py-1 font-medium text-blue-600 hover:bg-blue-100">
            <X size={12} /> 필터 해제
          </button>
        </div>
      )}
      <div className="mb-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <h3 className="text-sm font-semibold text-slate-700">기지국 현황 <span className="font-normal text-slate-400">({filtered.length}건 · 정류기 모델별로 행이 나뉩니다)</span></h3>
        {!compact && (
          <div className="flex flex-wrap gap-2">
            <div className="relative">
              <Search size={14} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input value={q} onChange={(e) => { setQ(e.target.value); setPage(1); }} placeholder="국소명 검색"
                className="rounded-lg border border-slate-200 bg-slate-50 py-1.5 pl-8 pr-2 text-xs outline-none focus:border-blue-400 focus:bg-white" />
            </div>
            <select value={teamFilter} onChange={(e) => { setTeamFilter(e.target.value); setPage(1); }}
              className="rounded-lg border border-slate-200 bg-slate-50 px-2 py-1.5 text-xs outline-none">
              {teams.map((t) => <option key={t}>{t}</option>)}
            </select>
            <select value={statusFilter} onChange={(e) => { setStatusFilter(e.target.value); setPage(1); }}
              className="rounded-lg border border-slate-200 bg-slate-50 px-2 py-1.5 text-xs outline-none">
              {["전체", "양호", "열화", "조불량", "미측정"].map((s) => <option key={s}>{s}</option>)}
            </select>
            <button onClick={() => exportStationsToExcel(filtered, presetFilter?.label, excludedModels)}
              className="flex items-center gap-1 rounded-lg border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 text-xs font-medium text-emerald-700 hover:bg-emerald-100">
              <FileSpreadsheet size={14} /> 엑셀 다운로드
            </button>
          </div>
        )}
      </div>
      <div className="overflow-x-auto">
        <table className="w-full min-w-[720px] text-left text-xs">
          <thead className="bg-slate-50 text-slate-500">
            <tr>
              <th className="px-2 py-2 font-medium">국소명</th>
              <th className="px-2 py-2 font-medium">주소</th>
              <th className="px-2 py-2 font-medium">현장운용팀</th>
              <th className="px-2 py-2 font-medium">정류기 모델명</th>
              <th className="px-2 py-2 font-medium">상태</th>
              <th className="px-2 py-2 font-medium">조치여부</th>
              <th className="px-2 py-2 font-medium">내부저항측정일시</th>
            </tr>
          </thead>
          <tbody>
            {shown.length === 0 && (
              <tr><td colSpan={7} className="px-2 py-6 text-center text-slate-400">표시할 데이터가 없습니다.</td></tr>
            )}
            {shown.map((r, i) => (
              <tr key={i} className="border-t border-slate-100 hover:bg-slate-50">
                <td className="px-2 py-2 font-medium text-slate-700">{r["국소명"] || "-"}</td>
                <td className="px-2 py-2 text-slate-500">{r["주소"] || "-"}</td>
                <td className="px-2 py-2 text-slate-500">{r["팀"] || "-"}</td>
                <td className="px-2 py-2 text-slate-500 whitespace-nowrap">{r._model || "-"}</td>
                <td className="px-2 py-2"><StatusBadge status={stationStatus(r, excludedModels)} /></td>
                <td className="px-2 py-2"><ActionBadge text={stationAction(r)} /></td>
                <td className="px-2 py-2 text-slate-500 whitespace-nowrap">{stationMeasuredAt(r) || "-"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {!compact && totalPages > 1 && (
        <div className="mt-3 flex items-center justify-center gap-2 text-xs">
          <button onClick={() => setPage(1)} disabled={page === 1}
            className="rounded-lg border border-slate-200 px-2 py-1.5 text-slate-500 hover:bg-slate-50 disabled:opacity-30 disabled:hover:bg-transparent">
            처음
          </button>
          <button onClick={() => setPage((p) => Math.max(1, p - 1))} disabled={page === 1}
            className="rounded-lg border border-slate-200 px-2 py-1.5 text-slate-500 hover:bg-slate-50 disabled:opacity-30 disabled:hover:bg-transparent">
            이전
          </button>
          <span className="px-2 font-medium text-slate-600">{page} / {totalPages}</span>
          <button onClick={() => setPage((p) => Math.min(totalPages, p + 1))} disabled={page === totalPages}
            className="rounded-lg border border-slate-200 px-2 py-1.5 text-slate-500 hover:bg-slate-50 disabled:opacity-30 disabled:hover:bg-transparent">
            다음
          </button>
          <button onClick={() => setPage(totalPages)} disabled={page === totalPages}
            className="rounded-lg border border-slate-200 px-2 py-1.5 text-slate-500 hover:bg-slate-50 disabled:opacity-30 disabled:hover:bg-transparent">
            마지막
          </button>
        </div>
      )}
    </div>
  );
}

/* -------------------------------- 홈 화면 -------------------------------- */
/* --------------------------------- 축전지 현황 대시보드 --------------------------------- */
function BatteryOverviewDashboard({ rows, lastUpdatedAt, onDrill, excludedModels }) {
  const dash = useMemo(() => {
    const makeBucket = () => ({ total: 0, bad: 0, du: 0, w: 0 });
    if (!rows.length) {
      return { v2: { total: 850, bad: 60, du: 42, w: 18 }, v12: { total: 398, bad: 82, du: 0, w: 27 }, rt: { total: 50, bad: 15 } };
    }
    const v2 = makeBucket();
    const v12 = makeBucket();
    const rt = { total: 0, bad: 0 };
    rows.forEach((station) => {
      (station.rectifiers || []).forEach((rect) => {
        if (!isCountableRect(rect, excludedModels)) return; // 전제조건: 대상X/폐국/제외 모델 제외

        // 1차 필터: CB열이 'RT'인 대상은 전압(BD열)과 무관하게 먼저 전부 RT 축전지로 집계한다.
        if (eqLoose(rect["서비스"], "RT")) {
          rt.total += 1;
          if (eqLoose(rect["불량여부"], "불량")) rt.bad += 1;
          return;
        }

        // 2V 축전지: BD열이 '2V'인 것만 (CB열 제한 없음)
        if (eqLoose(rect["전압"], "2V")) {
          v2.total += 1;
          if (eqLoose(rect["불량여부"], "불량")) {
            v2.bad += 1;
            if (eqLoose(rect["서비스"], "DU")) v2.du += 1;
            if (eqLoose(rect["서비스"], "W")) v2.w += 1;
          }
          return;
        }

        // 12V 축전지: 1차 BD열 '12V' → 2차 CB열이 'W'·'DU'·'5G' 중 하나인 것만
        if (eqLoose(rect["전압"], "12V") && ["W", "DU", "5G"].some((s) => eqLoose(rect["서비스"], s))) {
          v12.total += 1;
          if (eqLoose(rect["불량여부"], "불량")) {
            v12.bad += 1;
            if (eqLoose(rect["서비스"], "DU")) v12.du += 1;
            if (eqLoose(rect["서비스"], "W")) v12.w += 1;
          }
          return;
        }
        // 그 외(전압구분 없음, 12V인데 CB열이 W/DU/5G가 아닌 경우 등)는 어느 집계에도 포함하지 않는다.
      });
    });
    return { v2, v12, rt };
  }, [rows, excludedModels]);

  const { v2, v12, rt } = dash;
  const total = v2.total + v12.total + rt.total;
  const totalBad = v2.bad + v12.bad + rt.bad;

  const fmt = (n) => n.toLocaleString();
  const drill = (preset) => onDrill?.(preset);
  const [open, setOpen] = useState(true);

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-center gap-2">
          <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-blue-50 text-blue-600"><Battery size={18} /></div>
          <h2 className="text-base font-bold text-slate-800">축전지 현황</h2>
        </div>
        <div className="flex items-center gap-2 text-xs text-slate-400">
          <span className="hidden sm:inline">기준일 : {lastUpdatedAt || "-"}</span>
          <RotateCcw size={13} className="hidden sm:inline" />
          <button onClick={() => setOpen((v) => !v)}
            className="flex items-center gap-1 rounded-lg border border-slate-200 px-2 py-1 text-slate-500 hover:bg-slate-50">
            {open ? <><ChevronUp size={13} /> 접기</> : <><ChevronDown size={13} /> 펼치기</>}
          </button>
        </div>
      </div>

      {!open ? (
        <button onClick={() => setOpen(true)} className="flex w-full flex-wrap items-center gap-x-4 gap-y-1 rounded-lg bg-slate-50 px-3 py-2.5 text-left text-xs hover:bg-slate-100">
          <span className="text-slate-500">전체 <b className="text-blue-600">{fmt(total)}조</b></span>
          <span className="text-slate-500">불량 <b className="text-red-600">{fmt(totalBad)}조</b></span>
          <span className="text-slate-400">(2V {fmt(v2.total)}조 · 12V {fmt(v12.total)}조) — 눌러서 펼치기</span>
        </button>
      ) : (
      <>
      <p className="mb-3 text-[11px] text-slate-400">숫자를 클릭하면 기지국 현황에서 해당 대상들을 바로 확인할 수 있어요.</p>

      {/* 전체 / 불량 요약 카드 */}
      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
        <div className="overflow-hidden rounded-xl border border-blue-100 bg-gradient-to-b from-blue-50/60 to-white">
          <div className="flex items-center gap-4 p-4">
            <div className="flex h-14 w-14 shrink-0 items-center justify-center rounded-full bg-blue-100 text-blue-600"><Battery size={26} /></div>
            <div>
              <p className="text-sm font-semibold text-blue-600">전체 축전지</p>
              <button onClick={() => drill({ label: "전체 축전지" })}
                className="text-3xl font-extrabold text-slate-800 hover:text-blue-600 hover:underline">
                {fmt(total)}<span className="ml-1 text-lg font-bold text-slate-400">조</span>
              </button>
              <p className="mt-0.5 text-xs text-slate-400">
                (2V:{" "}
                <button onClick={() => drill({ voltage: "2V", label: "2V 전체" })} className="font-medium text-slate-500 hover:text-blue-600 hover:underline">{fmt(v2.total)}조</button>
                {" | "}12V:{" "}
                <button onClick={() => drill({ voltage: "12V", label: "12V 전체" })} className="font-medium text-slate-500 hover:text-blue-600 hover:underline">{fmt(v12.total)}조</button>
                {" | "}RT:{" "}
                <button onClick={() => drill({ service: "RT", label: "RT 전체" })} className="font-medium text-slate-500 hover:text-blue-600 hover:underline">{fmt(rt.total)}조</button>)
              </p>
            </div>
          </div>
          <div className="h-1 bg-blue-500" />
        </div>
        <div className="overflow-hidden rounded-xl border border-red-100 bg-gradient-to-b from-red-50/60 to-white">
          <div className="flex items-center gap-4 p-4">
            <div className="flex h-14 w-14 shrink-0 items-center justify-center rounded-full bg-red-100 text-red-600"><AlertTriangle size={26} /></div>
            <div>
              <p className="text-sm font-semibold text-red-600">불량 축전지</p>
              <button onClick={() => drill({ badOnly: true, label: "불량 축전지 전체" })}
                className="text-3xl font-extrabold text-slate-800 hover:text-red-600 hover:underline">
                {fmt(totalBad)}<span className="ml-1 text-lg font-bold text-slate-400">조</span>
              </button>
              <p className="mt-0.5 text-xs text-slate-400">
                (2V:{" "}
                <button onClick={() => drill({ voltage: "2V", badOnly: true, label: "2V 불량" })} className="font-medium text-slate-500 hover:text-red-600 hover:underline">{fmt(v2.bad)}조</button>
                {" | "}12V:{" "}
                <button onClick={() => drill({ voltage: "12V", badOnly: true, label: "12V 불량" })} className="font-medium text-slate-500 hover:text-red-600 hover:underline">{fmt(v12.bad)}조</button>
                {" | "}RT:{" "}
                <button onClick={() => drill({ service: "RT", badOnly: true, label: "RT 불량" })} className="font-medium text-slate-500 hover:text-red-600 hover:underline">{fmt(rt.bad)}조</button>)
              </p>
            </div>
          </div>
          <div className="h-1 bg-red-500" />
        </div>
      </div>

      {/* 불량 축전지 상세 현황 */}
      <div className="mt-5">
        <div className="mb-2 flex items-center gap-1.5">
          <BarChart3 size={15} className="text-blue-500" />
          <p className="text-sm font-semibold text-slate-700">불량 축전지 상세 현황</p>
        </div>
        <div className="space-y-3">
          <div className="rounded-xl border border-emerald-100 bg-emerald-50/40 p-3">
            <div className="mb-2 flex items-center gap-2">
              <span className="flex h-6 w-9 items-center justify-center rounded-full bg-emerald-600 text-[10px] font-bold text-white">2V</span>
              <p className="text-sm font-semibold text-emerald-700">2V 축전지</p>
            </div>
            <div className="grid grid-cols-3 divide-x divide-emerald-100 rounded-lg bg-white/70">
              <button onClick={() => drill({ voltage: "2V", label: "2V 전체" })} className="flex flex-col items-center gap-1 py-3 hover:bg-emerald-50">
                <p className="text-xs text-slate-400">전체</p>
                <p className="text-xl font-extrabold text-emerald-600">{fmt(v2.total)}조</p>
                <Battery size={20} className="text-emerald-500" />
              </button>
              <button onClick={() => drill({ voltage: "2V", badOnly: true, service: "DU", label: "2V DU 불량" })} className="flex flex-col items-center gap-1 py-3 hover:bg-red-50">
                <p className="text-xs text-slate-400">DU 불량</p>
                <p className="text-xl font-extrabold text-red-500">{fmt(v2.du)}조</p>
                <AlertTriangle size={20} className="text-red-500" />
              </button>
              <button onClick={() => drill({ voltage: "2V", badOnly: true, service: "W", label: "2V W 불량" })} className="flex flex-col items-center gap-1 py-3 hover:bg-orange-50">
                <p className="text-xs text-slate-400">W 불량</p>
                <p className="text-xl font-extrabold text-orange-500">{fmt(v2.w)}조</p>
                <AlertTriangle size={20} className="text-orange-500" />
              </button>
            </div>
          </div>

          <div className="rounded-xl border border-amber-100 bg-amber-50/40 p-3">
            <div className="mb-2 flex items-center gap-2">
              <span className="flex h-6 w-11 items-center justify-center rounded-full bg-amber-500 text-[10px] font-bold text-white">12V</span>
              <p className="text-sm font-semibold text-amber-700">12V 축전지</p>
            </div>
            <div className="grid grid-cols-3 divide-x divide-amber-100 rounded-lg bg-white/70">
              <button onClick={() => drill({ voltage: "12V", label: "12V 전체" })} className="flex flex-col items-center gap-1 py-3 hover:bg-amber-50">
                <p className="text-xs text-slate-400">전체</p>
                <p className="text-xl font-extrabold text-amber-600">{fmt(v12.total)}조</p>
                <Battery size={20} className="text-amber-500" />
              </button>
              <button onClick={() => drill({ voltage: "12V", badOnly: true, service: "DU", label: "12V DU 불량" })} className="flex flex-col items-center gap-1 py-3 hover:bg-red-50">
                <p className="text-xs text-slate-400">DU 불량</p>
                <p className="text-xl font-extrabold text-red-500">{fmt(v12.du)}조</p>
                <AlertTriangle size={20} className="text-red-500" />
              </button>
              <button onClick={() => drill({ voltage: "12V", badOnly: true, service: "W", label: "12V W 불량" })} className="flex flex-col items-center gap-1 py-3 hover:bg-orange-50">
                <p className="text-xs text-slate-400">W 불량</p>
                <p className="text-xl font-extrabold text-orange-500">{fmt(v12.w)}조</p>
                <AlertTriangle size={20} className="text-orange-500" />
              </button>
            </div>
          </div>

          <div className="rounded-xl border border-violet-100 bg-violet-50/40 p-3">
            <div className="mb-2 flex items-center gap-2">
              <span className="flex h-6 w-9 items-center justify-center rounded-full bg-violet-500 text-[10px] font-bold text-white">RT</span>
              <p className="text-sm font-semibold text-violet-700">RT 축전지</p>
            </div>
            <div className="grid grid-cols-2 divide-x divide-violet-100 rounded-lg bg-white/70">
              <button onClick={() => drill({ service: "RT", label: "RT 전체" })} className="flex flex-col items-center gap-1 py-3 hover:bg-violet-50">
                <p className="text-xs text-slate-400">전체</p>
                <p className="text-xl font-extrabold text-violet-600">{fmt(rt.total)}조</p>
                <Battery size={20} className="text-violet-500" />
              </button>
              <button onClick={() => drill({ service: "RT", badOnly: true, label: "RT 불량" })} className="flex flex-col items-center gap-1 py-3 hover:bg-red-50">
                <p className="text-xs text-slate-400">불량</p>
                <p className="text-xl font-extrabold text-red-500">{fmt(rt.bad)}조</p>
                <AlertTriangle size={20} className="text-red-500" />
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* 요약 표 */}
      <div className="mt-5">
        <div className="mb-2 flex items-center gap-1.5">
          <FileSpreadsheet size={14} className="text-slate-500" />
          <p className="text-sm font-semibold text-slate-700">요약 표</p>
        </div>
        <div className="overflow-x-auto overflow-hidden rounded-xl border border-slate-200">
          <table className="w-full min-w-[480px] text-center text-xs">
            <thead>
              <tr className="bg-slate-800 text-white">
                <th className="px-3 py-2 font-medium">구분</th>
                <th className="px-3 py-2 font-medium">전체 축전지</th>
                <th className="px-3 py-2 font-medium">불량 축전지</th>
                <th className="px-3 py-2 font-medium">DU 불량</th>
                <th className="px-3 py-2 font-medium">W 불량</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              <tr>
                <td className="px-3 py-2 font-semibold text-emerald-600">2V</td>
                <td className="px-3 py-2"><button onClick={() => drill({ voltage: "2V", label: "2V 전체" })} className="text-slate-700 hover:text-blue-600 hover:underline">{fmt(v2.total)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ voltage: "2V", badOnly: true, label: "2V 불량" })} className="text-red-500 hover:underline">{fmt(v2.bad)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ voltage: "2V", badOnly: true, service: "DU", label: "2V DU 불량" })} className="text-red-500 hover:underline">{fmt(v2.du)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ voltage: "2V", badOnly: true, service: "W", label: "2V W 불량" })} className="text-orange-500 hover:underline">{fmt(v2.w)}조</button></td>
              </tr>
              <tr>
                <td className="px-3 py-2 font-semibold text-amber-600">12V</td>
                <td className="px-3 py-2"><button onClick={() => drill({ voltage: "12V", label: "12V 전체" })} className="text-slate-700 hover:text-blue-600 hover:underline">{fmt(v12.total)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ voltage: "12V", badOnly: true, label: "12V 불량" })} className="text-red-500 hover:underline">{fmt(v12.bad)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ voltage: "12V", badOnly: true, service: "DU", label: "12V DU 불량" })} className="text-red-500 hover:underline">{fmt(v12.du)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ voltage: "12V", badOnly: true, service: "W", label: "12V W 불량" })} className="text-orange-500 hover:underline">{fmt(v12.w)}조</button></td>
              </tr>
              <tr>
                <td className="px-3 py-2 font-semibold text-violet-600">RT</td>
                <td className="px-3 py-2"><button onClick={() => drill({ service: "RT", label: "RT 전체" })} className="text-slate-700 hover:text-blue-600 hover:underline">{fmt(rt.total)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ service: "RT", badOnly: true, label: "RT 불량" })} className="text-red-500 hover:underline">{fmt(rt.bad)}조</button></td>
                <td className="px-3 py-2 text-slate-300">-</td>
                <td className="px-3 py-2 text-slate-300">-</td>
              </tr>
              <tr className="bg-slate-50">
                <td className="px-3 py-2 font-bold text-blue-600">합계</td>
                <td className="px-3 py-2"><button onClick={() => drill({ label: "전체 축전지" })} className="font-bold text-blue-600 hover:underline">{fmt(total)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ badOnly: true, label: "불량 축전지 전체" })} className="font-bold text-blue-600 hover:underline">{fmt(totalBad)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ badOnly: true, service: "DU", label: "DU 불량 전체" })} className="font-bold text-blue-600 hover:underline">{fmt(v2.du + v12.du)}조</button></td>
                <td className="px-3 py-2"><button onClick={() => drill({ badOnly: true, service: "W", label: "W 불량 전체" })} className="font-bold text-blue-600 hover:underline">{fmt(v2.w + v12.w)}조</button></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
      </>
      )}
    </div>
  );
}

function HomePage({ rows, query, setQuery, actaHistory, lastUpdatedAt, onDrill, excludedModels }) {
  return (
    <div className="space-y-6">
      <section>
        <h2 className="mb-3 text-sm font-semibold text-slate-700">1. 대시보드</h2>
        <BatteryOverviewDashboard rows={rows} lastUpdatedAt={lastUpdatedAt} onDrill={onDrill} excludedModels={excludedModels} />
      </section>

      <section>
        <h2 className="mb-3 text-sm font-semibold text-slate-700">2. 국소 상세 현황</h2>
        <LocationDetail rows={rows.length ? rows : [SAMPLE_ROW]} query={query} setQuery={setQuery} actaHistory={actaHistory} />
      </section>

      <section>
        <h2 className="mb-3 text-sm font-semibold text-slate-700">3. 팀별 현황</h2>
        <TeamStatus rows={rows} excludedModels={excludedModels} />
      </section>
    </div>
  );
}

/* ---------------------------- 축전지 재고 현황 ---------------------------- */
/* ---------------------------- 축전지 재고 색상 팔레트 (차분한 톤) ---------------------------- */
const TYPE_STYLE = {
  "2V": { chip: "bg-sky-50 text-sky-600 ring-1 ring-inset ring-sky-100", dot: "bg-sky-400", tabActive: "bg-white text-sky-700 shadow-sm" },
  "12V": { chip: "bg-emerald-50 text-emerald-600 ring-1 ring-inset ring-emerald-100", dot: "bg-emerald-400", tabActive: "bg-white text-emerald-700 shadow-sm" },
  "전체": { tabActive: "bg-white text-slate-700 shadow-sm" },
};
function chipStyleFor(type) {
  return String(type).startsWith("12V") ? TYPE_STYLE["12V"].chip : TYPE_STYLE["2V"].chip;
}

function CombinedStockTable({ values, onChange, targets, isAdmin }) {
  const usageByType = useMemo(() => {
    const map = {};
    STOCK_ROWS.forEach((type) => { map[type] = (targets || []).filter((t) => t.종류 === type).length; });
    return map;
  }, [targets]);

  const totalByType = useMemo(() => {
    const map = {};
    STOCK_ROWS.forEach((type) => {
      map[type] = ROUNDS.reduce((s, round) => s + (Number(values[`${round}_${type}`]) || 0), 0);
    });
    return map;
  }, [values]);

  return (
    <div className="overflow-x-auto overflow-hidden rounded-lg border border-slate-200 shadow-sm">
      <table className="w-full min-w-[620px] text-center text-xs">
        <thead>
          <tr className="bg-slate-50">
            <th className="border border-slate-200 px-3 py-2 font-semibold text-slate-500"></th>
            {STOCK_ROWS.map((type) => <th key={type} className="whitespace-nowrap border border-slate-200 px-3 py-2 font-medium text-slate-500">{type}</th>)}
          </tr>
        </thead>
        <tbody>
          {ROUNDS.map((round) => (
            <tr key={round}>
              <td className="whitespace-nowrap border border-slate-200 bg-slate-50/70 px-3 py-2 text-left font-medium text-slate-600">{round}</td>
              {STOCK_ROWS.map((type) => (
                <td key={type} className="border border-slate-200 p-0">
                  {isAdmin ? (
                    <input
                      value={values[`${round}_${type}`] ?? ""}
                      onChange={(e) => onChange(round, type, e.target.value)}
                      className="w-full bg-transparent px-2 py-2 text-center outline-none focus:bg-slate-50"
                      placeholder="-"
                    />
                  ) : (
                    <span className="block px-2 py-2 text-slate-700">{values[`${round}_${type}`] || "-"}</span>
                  )}
                </td>
              ))}
            </tr>
          ))}
          <tr className="bg-slate-100/70">
            <td className="whitespace-nowrap border border-slate-200 px-3 py-2 text-left font-semibold text-slate-700">합계</td>
            {STOCK_ROWS.map((type) => (
              <td key={type} className="border border-slate-200 px-3 py-2 font-semibold text-slate-700">{totalByType[type]}</td>
            ))}
          </tr>
          <tr className="bg-blue-50/50">
            <td className="whitespace-nowrap border border-slate-200 px-3 py-2 text-left font-semibold text-blue-600" title="아래 '대상'에 등록된 개수를 자동으로 센 값입니다">사용</td>
            {STOCK_ROWS.map((type) => (
              <td key={type} className="border border-slate-200 px-3 py-2 font-semibold text-blue-600">{usageByType[type]}</td>
            ))}
          </tr>
          <tr className="bg-emerald-50/50">
            <td className="whitespace-nowrap border border-slate-200 px-3 py-2 text-left font-semibold text-emerald-600" title="합계에서 사용 수량을 뺀 값입니다">잔여</td>
            {STOCK_ROWS.map((type) => (
              <td key={type} className="border border-slate-200 px-3 py-2 font-semibold text-emerald-600">{totalByType[type] - usageByType[type]}</td>
            ))}
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function normalizeTargetRecord(r) {
  return {
    id: r.id ?? (Date.now() + Math.random()),
    종류: r.종류 || STOCK_ROWS[0],
    국소명: r.국소명 ?? r.공용대표명 ?? r.국소검색 ?? "",
    주소: r.주소 ?? "",
    현장운용팀: r.현장운용팀 ?? "",
    W: r.W ?? "",
    DU: r.DU ?? "",
    "5G": r["5G"] ?? "",
    작업예정일: r.작업예정일 ?? r.작업예정 ?? "",
    완료여부: r.완료여부 || "대기",
    비고: r.비고 ?? "",
  };
}

function TargetTable({ records, setRecords, storageKey, isAdmin }) {
  const [open, setOpen] = useState(true);
  const [tab, setTab] = useState("전체");
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved
  const [editing, setEditing] = useState(false); // 관리자 모드에서도, 저장을 누르면 이 값이 꺼지면서 화면이 일반 모드처럼 읽기 전용으로 바뀐다.

  useEffect(() => {
    (async () => {
      if (records.length > 0 || !storageKey) return;
      try {
        const res = await kv.get(storageKey);
        if (res?.value) { setRecords(JSON.parse(res.value).map(normalizeTargetRecord)); return; }
      } catch (e) { /* 저장된 대상 없음 */ }
      // 이전 버전(2V/12V 분리 저장, 구 컬럼 구성)과의 하위 호환: 있으면 불러와서 합쳐준다
      try {
        const [legacy2V, legacy12V] = await Promise.all([
          kv.get("battery-stock-targets-2v")?.catch(() => null),
          kv.get("battery-stock-targets-12v")?.catch(() => null),
        ]);
        const merged = [
          ...(legacy2V?.value ? JSON.parse(legacy2V.value).map((r) => normalizeTargetRecord({ ...r, 종류: "2V" })) : []),
          ...(legacy12V?.value ? JSON.parse(legacy12V.value).map((r) => normalizeTargetRecord({ ...r, 종류: "12V" })) : []),
        ];
        if (merged.length) setRecords(merged);
      } catch (e) { /* 이전 데이터 없음 */ }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [storageKey]);

  const [teamFilter, setTeamFilter] = useState("전체");
  const filtered = records
    .filter((r) => tab === "전체" || String(r.종류).startsWith(tab))
    .filter((r) => teamFilter === "전체" || r.현장운용팀 === teamFilter);

  const addRow = () =>
    setRecords((p) => [...p, normalizeTargetRecord({
      id: Date.now(),
      종류: tab === "전체" ? STOCK_ROWS[0] : (STOCK_ROWS.find((s) => s.startsWith(tab)) || STOCK_ROWS[0]),
    })]);
  const removeRow = (id) => setRecords((p) => p.filter((r) => r.id !== id));
  const editCell = (id, field, val) => setRecords((p) => p.map((r) => (r.id === id ? { ...r, [field]: val } : r)));

  const handleSave = async () => {
    if (!storageKey || !isAdmin) return;
    setSaveState("saving");
    try {
      await kv.set(storageKey, JSON.stringify(records));
      setSaveState("saved");
      setEditing(false); // 저장되면 수정 화면(입력칸)을 닫고 일반 모드처럼 읽기 전용으로 보여준다.
    } catch (e) {
      setSaveState("idle");
    } finally {
      setTimeout(() => setSaveState("idle"), 1800);
    }
  };

  const canEdit = isAdmin && editing;
  const HEADERS = ["종류", "국소명", "주소", "현장운용팀", "W", "DU", "5G", "작업 예정(완료)일", "완료여부", "비고"];
  const colCount = HEADERS.length + (canEdit ? 1 : 0);

  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 shadow-sm">
      <button
        onClick={() => setOpen((v) => !v)}
        className="flex w-full items-center justify-between bg-slate-50 px-4 py-3 text-left"
      >
        <div>
          <h3 className="text-sm font-semibold text-slate-700">대상</h3>
          <p className="mt-0.5 text-[11px] text-slate-400">관리자 모드에서 국소명 · 주소 · 현장운용팀 · W · DU · 5G를 직접 입력합니다</p>
        </div>
        <div className="flex items-center gap-2">
          <span className="rounded-full bg-white px-2 py-0.5 text-[11px] font-medium text-slate-500 ring-1 ring-inset ring-slate-200">{records.length}건</span>
          {open ? <ChevronUp size={16} className="text-slate-400" /> : <ChevronDown size={16} className="text-slate-400" />}
        </div>
      </button>

      {open && (
        <div className="border-t border-slate-100">
          <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-2">
            <div className="flex flex-wrap items-center gap-2">
              <div className="flex items-center gap-1 rounded-lg bg-slate-100 p-1">
                {["전체", "2V", "12V"].map((t) => (
                  <button key={t} onClick={() => setTab(t)}
                    className={`rounded-md px-3 py-1.5 text-xs font-medium transition-colors ${
                      tab === t ? TYPE_STYLE[t].tabActive : "text-slate-500 hover:text-slate-700"
                    }`}>
                    {t === "전체" ? "전체" : `${t} 축전지`}
                  </button>
                ))}
              </div>
              <select value={teamFilter} onChange={(e) => setTeamFilter(e.target.value)}
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-600 outline-none focus:border-slate-400">
                <option value="전체">현장운용팀 전체</option>
                {VALID_TEAMS.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <div className="flex items-center gap-2">
              {!isAdmin ? (
                <span className="flex items-center gap-1 text-[11px] text-slate-400">
                  <Lock size={11} /> 대상 추가·수정·저장은 관리자 모드에서만 가능합니다
                </span>
              ) : !editing ? (
                <button onClick={() => setEditing(true)}
                  className="flex items-center gap-1 rounded-lg bg-slate-700 px-3 py-1.5 text-xs font-medium text-white shadow-sm hover:bg-slate-800">
                  <Plus size={13} className="rotate-45" /> 수정
                </button>
              ) : (
                <>
                  <button onClick={handleSave}
                    className={`flex items-center gap-1 rounded-lg border px-3 py-1.5 text-xs font-medium shadow-sm transition-colors ${
                      saveState === "saved" ? "border-emerald-200 bg-emerald-50 text-emerald-600" : "border-slate-200 text-slate-600 hover:bg-slate-50"
                    }`}>
                    {saveState === "saved" ? <><Check size={13} /> 저장됨</> : <><Save size={13} /> {saveState === "saving" ? "저장 중..." : "저장"}</>}
                  </button>
                  <button onClick={addRow} className="flex items-center gap-1 rounded-lg bg-slate-700 px-3 py-1.5 text-xs font-medium text-white shadow-sm hover:bg-slate-800">
                    <Plus size={13} /> 대상 추가
                  </button>
                </>
              )}
            </div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full min-w-[820px] text-center text-xs">
              <thead>
                <tr className="bg-slate-50 text-[11px] uppercase tracking-wide text-slate-500">
                  {HEADERS.map((h) => (
                    <th key={h} className="px-3 py-2.5 font-semibold">{h}</th>
                  ))}
                  {canEdit && <th className="px-2 py-2.5" />}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filtered.length === 0 && (
                  <tr><td colSpan={colCount} className="px-3 py-6 text-slate-400">
                    {canEdit ? "대상을 추가하고 정보를 입력해주세요." : "등록된 대상이 없습니다."}
                  </td></tr>
                )}
                {filtered.map((row, idx) => (
                  <tr key={row.id} className={`${idx % 2 === 1 ? "bg-slate-50/60" : "bg-white"} transition-colors hover:bg-slate-50`}>
                    <td className="px-2 py-1.5">
                      {canEdit ? (
                        <select value={row.종류 || STOCK_ROWS[0]} onChange={(e) => editCell(row.id, "종류", e.target.value)}
                          className={`rounded-full border-0 px-2 py-1 text-[11px] font-medium outline-none ${chipStyleFor(row.종류)}`}>
                          {STOCK_ROWS.map((type) => <option key={type} value={type}>{type}</option>)}
                        </select>
                      ) : (
                        <span className={`inline-block whitespace-nowrap rounded-full px-2 py-0.5 text-[11px] font-medium ${chipStyleFor(row.종류)}`}>{row.종류}</span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <input value={row.국소명} onChange={(e) => editCell(row.id, "국소명", e.target.value)} placeholder="국소명"
                          className="w-full rounded-md border border-slate-200 bg-white px-2 py-1.5 text-center outline-none focus:border-slate-400" />
                      ) : (
                        <span className="font-medium text-slate-700">{row.국소명 || <span className="text-slate-300">-</span>}</span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <input value={row.주소} onChange={(e) => editCell(row.id, "주소", e.target.value)} placeholder="주소"
                          className="w-full rounded-md border border-slate-200 bg-white px-2 py-1.5 text-center outline-none focus:border-slate-400" />
                      ) : (
                        <span className="text-slate-500">{row.주소 || <span className="text-slate-300">-</span>}</span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <select value={row.현장운용팀 || ""} onChange={(e) => editCell(row.id, "현장운용팀", e.target.value)}
                          className="w-full rounded-md border border-slate-200 bg-white px-2 py-1.5 text-center outline-none focus:border-slate-400">
                          <option value="">선택</option>
                          {VALID_TEAMS.map((t) => <option key={t} value={t}>{t}</option>)}
                        </select>
                      ) : (
                        <span className="text-slate-500">{row.현장운용팀 || <span className="text-slate-300">-</span>}</span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <input value={row.W} onChange={(e) => editCell(row.id, "W", e.target.value)} placeholder="-"
                          className="w-full rounded-md border border-slate-200 bg-white px-1 py-1.5 text-center outline-none focus:border-slate-400" />
                      ) : (
                        <span className="inline-flex min-w-[1.75rem] justify-center rounded-full bg-slate-100 px-2 py-0.5 text-slate-600">
                          {row.W === "" || row.W === undefined ? "-" : row.W}
                        </span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <input value={row.DU} onChange={(e) => editCell(row.id, "DU", e.target.value)} placeholder="-"
                          className="w-full rounded-md border border-slate-200 bg-white px-1 py-1.5 text-center outline-none focus:border-slate-400" />
                      ) : (
                        <span className="inline-flex min-w-[1.75rem] justify-center rounded-full bg-slate-100 px-2 py-0.5 text-slate-600">
                          {row.DU === "" || row.DU === undefined ? "-" : row.DU}
                        </span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <input value={row["5G"]} onChange={(e) => editCell(row.id, "5G", e.target.value)} placeholder="-"
                          className="w-full rounded-md border border-slate-200 bg-white px-1 py-1.5 text-center outline-none focus:border-slate-400" />
                      ) : (
                        <span className="inline-flex min-w-[1.75rem] justify-center rounded-full bg-slate-100 px-2 py-0.5 text-slate-600">
                          {row["5G"] === "" || row["5G"] === undefined ? "-" : row["5G"]}
                        </span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <input type="date" value={row.작업예정일} onChange={(e) => editCell(row.id, "작업예정일", e.target.value)}
                          className="w-full rounded-md border border-slate-200 bg-white px-2 py-1.5 text-center text-[11px] outline-none focus:border-slate-400" />
                      ) : (
                        <span className="flex items-center justify-center gap-1 text-slate-500">
                          {row.작업예정일 || <span className="text-slate-300">-</span>}
                          <Lock size={10} className="text-slate-300" />
                        </span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <select value={row.완료여부 || "대기"} onChange={(e) => editCell(row.id, "완료여부", e.target.value)}
                          className={`rounded-full border-0 px-2 py-1 text-[11px] font-medium outline-none ${
                            row.완료여부 === "완료" ? "bg-emerald-50 text-emerald-600" : "bg-amber-50 text-amber-600"
                          }`}>
                          <option value="대기">대기</option>
                          <option value="완료">완료</option>
                        </select>
                      ) : (
                        <span className={`inline-block whitespace-nowrap rounded-full px-2 py-0.5 text-[11px] font-medium ${
                          row.완료여부 === "완료" ? "bg-emerald-50 text-emerald-600" : "bg-amber-50 text-amber-600"
                        }`}>{row.완료여부 || "대기"}</span>
                      )}
                    </td>
                    <td className="px-1 py-1.5">
                      {canEdit ? (
                        <input value={row.비고} onChange={(e) => editCell(row.id, "비고", e.target.value)}
                          placeholder="비고"
                          className="w-full rounded-md border border-transparent bg-transparent px-2 py-1.5 text-center outline-none focus:border-slate-300 focus:bg-white" />
                      ) : (
                        <span className="text-slate-500">{row.비고 || <span className="text-slate-300">-</span>}</span>
                      )}
                    </td>
                    {canEdit && (
                      <td className="px-2 py-1.5">
                        <button onClick={() => removeRow(row.id)}
                          className="rounded-md p-1.5 text-slate-300 transition-colors hover:bg-red-50 hover:text-red-500">
                          <Trash2 size={14} />
                        </button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function BatteryStockPage({ isAdmin }) {
  const [stock, setStock] = useState({});
  const [targets, setTargets] = useState([]);
  const [stockSaveState, setStockSaveState] = useState("idle"); // idle | saving | saved

  useEffect(() => {
    (async () => {
      try {
        const res = await kv.get("battery-stock-quantities");
        if (res?.value) setStock(JSON.parse(res.value) || {});
      } catch (e) { console.error("[kv load failed] battery-stock-quantities", e); }
    })();
  }, []);

  const updateStock = (round, label, val) => setStock((p) => ({ ...p, [`${round}_${label}`]: val }));

  const handleStockSave = async () => {
    if (!isAdmin) return;
    setStockSaveState("saving");
    try {
      await kv.set("battery-stock-quantities", JSON.stringify(stock));
      setStockSaveState("saved");
    } catch (e) {
      setStockSaveState("idle");
    } finally {
      setTimeout(() => setStockSaveState("idle"), 1800);
    }
  };

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-6">
        <div className="mb-5 flex flex-wrap items-center justify-between gap-2">
          <div className="rounded-md border border-slate-300 px-6 py-1.5 text-sm font-semibold text-slate-700">축전지 재고</div>
          <div className="flex items-center gap-2">
            {isAdmin ? (
              <button onClick={handleStockSave}
                className={`flex items-center gap-1 rounded-lg border px-3 py-1.5 text-xs font-medium shadow-sm transition-colors ${
                  stockSaveState === "saved" ? "border-emerald-200 bg-emerald-50 text-emerald-600" : "border-slate-200 text-slate-600 hover:bg-slate-50"
                }`}>
                {stockSaveState === "saved" ? <><Check size={13} /> 저장됨</> : <><Save size={13} /> {stockSaveState === "saving" ? "저장 중..." : "저장"}</>}
              </button>
            ) : (
              <span className="flex items-center gap-1 text-[11px] text-slate-400">
                <Lock size={11} /> 재고 수량 수정·저장은 관리자 모드에서만 가능합니다
              </span>
            )}
            <button onClick={() => exportStockToExcel(stock, targets)}
              className="flex items-center gap-1 rounded-lg border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 text-xs font-medium text-emerald-700 hover:bg-emerald-100">
              <FileSpreadsheet size={14} /> 엑셀 다운로드
            </button>
          </div>
        </div>

        <div className="space-y-4">
          <CombinedStockTable values={stock} onChange={updateStock} targets={targets} isAdmin={isAdmin} />
          <TargetTable records={targets} setRecords={setTargets}
            storageKey="battery-stock-targets" isAdmin={isAdmin} />
        </div>
      </div>
    </div>
  );
}

/* -------------------------------- Backup 시간 -------------------------------- */
/** 같은 정류기번호로 묶인 행들은 같은 정류기에 병렬 연결된 축전지 조(組)다.
 *  규격(Ah)은 조 수만큼 합산하고, 부하전류·전압은 정류기 단위로 1번만 반영한다. */
function groupRectifierBanks(rects) {
  const map = new Map();
  rects.forEach((r) => {
    const key = r.번호 || `_unk_${r.축전지번호 || Math.random()}`;
    if (!map.has(key)) {
      map.set(key, {
        번호: r.번호, 정류기모델: r.정류기모델, 서비스: r.서비스, 전압: r.전압, 부하전류: r.부하전류,
        축전지번호목록: [], 규격합: 0, 조수: 0, 축전지상태: r.축전지상태,
      });
    }
    const bank = map.get(key);
    const cap = Number(r.규격);
    if (!Number.isNaN(cap) && cap > 0) { bank.규격합 += cap; bank.조수 += 1; }
    if (r.축전지번호) bank.축전지번호목록.push(r.축전지번호);
    // 부하전류·전압은 정류기당 값이 반복 기재되므로 비어있을 때만 채운다.
    if (!bank.부하전류 && r.부하전류) bank.부하전류 = r.부하전류;
    if (!bank.전압 && r.전압) bank.전압 = r.전압;
  });
  return [...map.values()];
}

/** 국소 1개(rectifiers 배열)에서 정류기(축전지 뱅크)별 예상 Backup 시간을 계산하고, 국소 대표값(최소값)을 함께 돌려준다.
 *  대개체필요/불용철거필요 등은 계산에서 제외하지 않지만, 규격·부하전류·전압 중 하나라도 없으면 해당 뱅크는 계산 대상에서 빠진다. */
function computeStationBackup(station) {
  const rects = (station.rectifiers || []).filter((r) => r.불량여부 !== "대상X" && r.불량여부 !== "폐국");
  const banks = groupRectifierBanks(rects);
  const results = banks
    .map((b) => {
      const d = calcBackupDetail(b.규격합, b.부하전류, b.전압);
      return d ? { ...b, backupH: d.hours, usedC: d.c, voltageKey: d.voltageKey } : null;
    })
    .filter(Boolean);
  const repH = results.length ? Math.min(...results.map((r) => r.backupH)) : null;
  return { results, repH, missingCount: banks.length - results.length };
}

/** 예상 Backup 시간에 따른 시각적 등급(위험/양호). 절대 기준이 아니라 한눈에 보기 위한 참고용 표시다.
 *  1.5시간을 넘으면 양호로 본다. */
function backupTier(h) {
  if (h == null) return { tone: "slate", label: "확인불가" };
  if (h <= 1.5) return { tone: "red", label: "위험" };
  return { tone: "emerald", label: "양호" };
}

const BACKUP_TONE_CLASSES = {
  red: { border: "border-red-100", bg: "bg-red-50/40", chip: "bg-red-50 text-red-600", big: "text-red-600", badge: "bg-red-100 text-red-600" },
  emerald: { border: "border-emerald-100", bg: "bg-emerald-50/40", chip: "bg-emerald-50 text-emerald-600", big: "text-emerald-600", badge: "bg-emerald-100 text-emerald-600" },
  slate: { border: "border-slate-200", bg: "bg-slate-50", chip: "bg-slate-100 text-slate-500", big: "text-slate-400", badge: "bg-slate-100 text-slate-500" },
};

function StationBackupCard({ station, results, repH, missingCount, large }) {
  const [openKey, setOpenKey] = useState(null);
  const tier = backupTier(repH);
  const tone = BACKUP_TONE_CLASSES[tier.tone];

  // 전압(2V/12V) 계열별로 묶어서, 어떤 축전지 조합이 계산에 쓰였는지 나눠 보여준다.
  const groups = useMemo(() => {
    const map = new Map();
    results.forEach((r) => {
      const key = r.voltageKey || r.전압 || "기타";
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(r);
    });
    return [...map.entries()];
  }, [results]);

  return (
    <div className={`rounded-xl border ${tone.border} ${tone.bg} ${large ? "p-5" : "p-3.5"}`}>
      <div className="mb-2 flex items-start justify-between gap-2">
        <div className="min-w-0">
          <p className={`truncate font-semibold text-slate-800 ${large ? "text-lg" : "text-sm"}`}>{station["국소명"] || "-"}</p>
          <p className={`truncate text-slate-400 ${large ? "text-sm" : "text-[11px]"}`}>{station["주소"] || "-"} · {station["팀"] || "-"}</p>
        </div>
        <span className={`inline-block shrink-0 rounded-full font-medium ${tone.badge} ${large ? "px-3 py-1 text-xs" : "px-2 py-0.5 text-[10px]"}`}>{tier.label}</span>
      </div>

      <div className={`mb-3 flex items-end gap-1.5 rounded-lg bg-white/70 ${large ? "px-4 py-4" : "px-3 py-2"}`}>
        <Clock size={large ? 24 : 16} className={tone.big} />
        <p className={`font-extrabold leading-none ${tone.big} ${large ? "text-4xl" : "text-2xl"}`}>{repH != null ? repH : "-"}</p>
        <p className={`pb-0.5 font-medium text-slate-400 ${large ? "text-sm" : "text-xs"}`}>{repH != null ? "시간 예상 (뱅크 중 최소값 기준)" : "계산 불가"}</p>
      </div>


      {groups.length > 0 && (
        <div className="space-y-2">
          {groups.map(([voltageKey, banks], gi) => (
            <div key={gi}>
              <p className="mb-1 text-[10px] font-semibold text-slate-400">{voltageKey} 계열 · 축전지 뱅크 {banks.length}개</p>
              <div className="space-y-1">
                {banks.map((res, j) => {
                  const rt = backupTier(res.backupH);
                  const rtone = BACKUP_TONE_CLASSES[rt.tone];
                  const key = `${gi}-${j}`;
                  const isOpen = openKey === key;
                  return (
                    <div key={j} className="overflow-hidden rounded-lg bg-white/70">
                      <button type="button" onClick={() => setOpenKey(isOpen ? null : key)}
                        className="flex w-full items-center justify-between gap-2 px-2 py-1.5 text-left">
                        <span className={`inline-flex flex-wrap items-center gap-1 rounded-full px-2 py-1 text-[10px] font-medium ${rtone.chip}`}>
                          <span className="font-semibold">{res.전압 || "-"}</span>
                          <span className="text-slate-400">·</span>
                          <span>{res.규격합 || "-"}Ah{res.조수 > 1 ? `(${res.조수}조)` : ""}</span>
                          <span className="text-slate-400">·</span>
                          <span>{res.부하전류 || "-"}A</span>
                          <span className="text-slate-400">→</span>
                          <span className="font-semibold">{res.backupH}h</span>
                        </span>
                        <ChevronDown size={12} className={`shrink-0 text-slate-400 transition-transform ${isOpen ? "rotate-180" : ""}`} />
                      </button>
                      {isOpen && (
                        <div className="border-t border-slate-100 px-2 py-1.5 text-[10px] leading-relaxed text-slate-500">
                          <p>정류기 {res.번호 || "-"} · 축전지번호 {res.축전지번호목록.join(", ") || "-"}</p>
                          <p className="mt-1 break-all font-mono text-slate-600">
                            {res.backupH}h = ({res.규격합}Ah × {res.backupH}h) ÷ ({res.부하전류}A × C<sub>{res.전압}</sub>({res.backupH}h)≈{res.usedC})
                          </p>
                          <p className="mt-1 text-slate-400">※ 위 C값은 용량산출계수 표를 해당 시간축에서 선형보간해, 값이 수렴할 때까지 반복계산한 결과예요.</p>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
      )}
      {missingCount > 0 && (
        <p className="mt-1.5 text-[10px] text-slate-400">규격·부하전류 정보 누락으로 계산 제외된 축전지 뱅크 {missingCount}개</p>
      )}
    </div>
  );
}

function BackupPage({ rows }) {
  const [q, setQ] = useState("");
  const data = rows.length ? rows : [SAMPLE_ROW];
  const query = q.trim();
  // 국소 상세현황(홈)과 동일하게, 2글자 이상 입력됐을 때 국소명에 포함되는 국소를 모두 후보로 보여준다.
  const filtered = useMemo(() => {
    if (!data.length || query.length < 2) return [];
    return data.filter((r) => norm(r["국소명"]).includes(norm(query)));
  }, [data, query]);

  const computed = useMemo(
    () => filtered.map((r) => ({ station: r, ...computeStationBackup(r) })),
    [filtered]
  );
  const SHOWN_LIMIT = 24;
  const shown = computed.slice(0, SHOWN_LIMIT);

  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      <div className="mb-3">
        <h3 className="text-sm font-semibold text-slate-700">Backup 시간 확인</h3>
        <p className="mt-0.5 text-[11px] text-slate-400">
          축전지 규격(Ah) · 부하전류(A) · 전압(2V/12V) 기반 자동 산출 · 용량산출계수C 반복계산
        </p>
      </div>
      <div className="relative mb-4">
        <Search size={16} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input
          value={q} onChange={(e) => setQ(e.target.value)} placeholder="국소명 2글자 이상 입력"
          className="w-full rounded-xl border border-slate-200 bg-slate-50 py-2.5 pl-9 pr-3 text-sm outline-none focus:border-blue-400 focus:bg-white"
        />
      </div>

      {(!query || query.length < 2) && (
        <div className="rounded-xl border border-dashed border-slate-200 px-4 py-10 text-center">
          <Search size={22} className="mx-auto mb-2 text-slate-300" />
          <p className="text-sm text-slate-400">국소명을 2글자 이상 입력하면 예상 Backup 시간이 카드로 표시됩니다.</p>
        </div>
      )}

      {query.length >= 2 && computed.length === 0 && (
        <p className="rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-600">'{query}'와(과) 일치하는 국소가 없습니다.</p>
      )}

      {query.length >= 2 && computed.length > 0 && (
        <>
          <div className={shown.length === 1 ? "grid grid-cols-1" : "grid grid-cols-1 gap-3 sm:grid-cols-2"}>
            {shown.map((c, i) => (
              <StationBackupCard key={i} station={c.station} results={c.results} repH={c.repH} missingCount={c.missingCount} large={shown.length === 1} />
            ))}
          </div>
          {computed.length > SHOWN_LIMIT && (
            <p className="mt-3 text-center text-[11px] text-slate-400">
              {computed.length}건 중 {SHOWN_LIMIT}건 표시 중 · 검색어를 더 구체적으로 입력해보세요.
            </p>
          )}
        </>
      )}
    </div>
  );
}

/* ----------------------------------- 관리자 ----------------------------------- */
function AdminUploadSlot({ title, desc, count, uploadedAt, loading, onUpload, onReset }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
        <div>
          <p className="text-sm font-semibold text-slate-700">{title}</p>
          <p className="mt-1 text-[11px] leading-relaxed text-slate-400">{desc}</p>
          <p className="mt-2 text-[11px] text-slate-500">
            {count > 0 ? <>현재 <b className="text-blue-600">{count}건</b> 연동됨 · 최종 업로드 {uploadedAt || "-"}</> : "업로드된 엑셀 없음"}
          </p>
        </div>
        <div className="flex shrink-0 items-center gap-2">
          <label title="여러 파일을 한 번에 선택할 수 있습니다. 통합시설코드가 같으면 하나로 합쳐집니다(같은 통합시설코드 값이 있는지 정확히 대조)."
            className="flex cursor-pointer items-center gap-1 rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-blue-700">
            <Upload size={14} /> {loading ? "업로드 중..." : "엑셀 업로드"}
            <input type="file" accept=".xlsx,.xls,.csv" multiple onChange={onUpload} className="hidden" />
          </label>
          {count > 0 && (
            <button onClick={onReset} className="flex items-center rounded-lg border border-slate-200 p-1.5 text-slate-400 hover:bg-slate-50">
              <RotateCcw size={14} />
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function RawColumnPreview({ rawPreview, highlightLetters }) {
  const [open, setOpen] = useState(false);
  if (!rawPreview || !rawPreview.length) {
    return <p className="mt-2 text-slate-400">(새로고침 후에는 원본 미리보기가 사라져요. 다시 확인하려면 파일을 다시 업로드해주세요.)</p>;
  }
  return (
    <div className="mt-2">
      <button onClick={() => setOpen((v) => !v)} className="flex items-center gap-1 text-blue-600 hover:underline">
        {open ? <ChevronUp size={12} /> : <ChevronDown size={12} />} 원본 열 전체보기 (A열부터 실제 값 그대로) — 이 열 문자와 실제 엑셀을 나란히 비교해보세요
      </button>
      {open && (
        <div className="mt-2 max-h-64 overflow-auto rounded border border-slate-200">
          <table className="w-full min-w-[600px] text-left">
            <thead className="sticky top-0 bg-slate-100 text-slate-500">
              <tr>
                <th className="px-2 py-1">열</th>
                <th className="px-2 py-1">헤더행 값</th>
                <th className="px-2 py-1">1행</th>
                <th className="px-2 py-1">2행</th>
                <th className="px-2 py-1">3행</th>
              </tr>
            </thead>
            <tbody>
              {rawPreview.map((col) => {
                const marked = highlightLetters?.includes(col.letter);
                return (
                  <tr key={col.letter} className={`border-t border-slate-100 ${marked ? "bg-blue-50" : ""}`}>
                    <td className={`px-2 py-1 font-medium ${marked ? "text-blue-600" : "text-slate-500"}`}>{col.letter}</td>
                    <td className="px-2 py-1 text-slate-400">{String(col.header ?? "") || "-"}</td>
                    {col.samples.map((v, i) => (
                      <td key={i} className="px-2 py-1 text-slate-700">{String(v ?? "") || "-"}</td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function AdminMultiUploadSlot({ title, desc, files, loading, onUpload, onRemove, countLabel, getCount, totalLabel, totalCount, guide, renderPreview }) {
  return (
    <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
      <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
        <div>
          <p className="text-sm font-semibold text-slate-700">{title}</p>
          <p className="mt-1 text-[11px] leading-relaxed text-slate-400">{desc}</p>
          <p className="mt-2 text-[11px] text-slate-500">
            {files.length > 0 ? <>파일 <b className="text-blue-600">{files.length}개</b> · 합산 {totalLabel} <b className="text-blue-600">{totalCount}개</b> 연동됨</> : "업로드된 엑셀 없음"}
          </p>
        </div>
        <label className="flex shrink-0 cursor-pointer items-center gap-1 rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-blue-700">
          <Upload size={14} /> {loading ? "업로드 중..." : "엑셀 업로드 추가"}
          <input type="file" accept=".xlsx,.xls,.csv" multiple onChange={onUpload} className="hidden" />
        </label>
      </div>

      {files.length > 0 && (
        <div className="mt-3 space-y-2">
          {files.map((f) => (
            <div key={f.id} className="rounded-lg bg-slate-50 px-3 py-2 text-[11px]">
              <div className="flex items-center justify-between gap-3">
                <span className="min-w-0 flex-1 truncate text-slate-600">
                  <FileSpreadsheet size={12} className="mr-1 inline text-slate-400" />
                  {f.name} · {getCount(f)}개 {countLabel} · {f.uploadedAt}
                </span>
                <button onClick={() => onRemove(f.id)} className="shrink-0 rounded-md p-1 text-slate-300 hover:bg-red-50 hover:text-red-500">
                  <Trash2 size={13} />
                </button>
              </div>
              {renderPreview && (
                <div className="mt-2 overflow-x-auto border-t border-slate-200 pt-2">
                  {renderPreview(f)}
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      {guide && (
        <div className="mt-3 overflow-x-auto rounded-lg border border-slate-100 bg-slate-50/70 p-3">
          {guide}
        </div>
      )}
    </div>
  );
}

const BASE_COLUMN_GUIDE = (
  <>
    <p className="mb-1.5 text-[11px] font-semibold text-slate-500">열 매핑 기준 (양식 고정 · 1행은 목차라 자동으로 건너뜀)</p>
    <div className="grid grid-cols-2 gap-x-4 gap-y-1 text-[11px] text-slate-500 sm:grid-cols-4">
      <span>F열 → 통합시설코드</span><span>G열 → 국소명</span><span>I열 → 본부</span><span>J열 → SKT팀</span>
      <span>M열 → 팀(현장운용팀)</span><span>N열 → 주소</span><span>S열 → 국사형태</span>
      <span>CC열 → 5G</span><span>CD열 → 4G</span><span>CE열 → 3G (계는 자동 합산)</span>
      <span>CH열 → 5G ARRU(식)</span><span>CI열 → RU(식)</span><span>CJ열 → 5G L9TU(식)</span>
      <span>CK열 → 링MUX RT 수용 RU/L9TU</span><span>CM열 → 중계기</span>
      <span>X열 → 정류기번호</span><span>Z열 → 정류기모델</span><span>CB열 → 서비스(DU/W)</span><span>AK열 → 부하전류</span><span>BC열 → 규격(Ah)</span>
      <span>AX열 → 축전지번호</span><span>BD열 → 전압구분(2V/12V)</span><span>BN열 → 축전지상태</span><span>BO열 → 양호</span>
      <span>BP열 → 열화</span><span>BQ열 → 열화2</span><span>BR열 → 불량</span><span>BT열 → 내부저항측정일시</span>
      <span>BW열 → 대개체여부</span><span>BV열 → 불량여부(대시보드 집계용)</span>
    </div>
    <p className="mt-2 text-[11px] text-amber-600">
      전제조건: BV열(불량여부) 값이 "대상X" 또는 "폐국"인 행은 대시보드·팀별현황·기지국현황 등 모든 집계에서 제외됩니다.
    </p>
  </>
);

function AdminPage({
  baseFiles, baseLoading, onBaseUpload, onRemoveBaseFile, onBaseReset, totalStations,
  actaHistoryCount, isAdmin, onChangePin,
  rows, excludedModels, onAddExcludedModel, onRemoveExcludedModel,
}) {
  const [pin1, setPin1] = useState("");
  const [pin2, setPin2] = useState("");
  const [pinMsg, setPinMsg] = useState("");
  const [modelInput, setModelInput] = useState("");

  const allModels = useMemo(() => {
    const set = new Set();
    (rows || []).forEach((s) => (s.rectifiers || []).forEach((r) => { if (r.정류기모델) set.add(r.정류기모델); }));
    return [...set].sort();
  }, [rows]);
  const availableModels = allModels.filter((m) => !(excludedModels || []).includes(m));

  // 진단용: 전압(BD열)별 원본 행 수 / 제외된 행 수 / 서비스(CB열)별 분포를 그대로 보여준다 (숫자가 왜 이렇게 나오는지 확인용). 관리자 전용.
  const diag = useMemo(() => {
    const build = () => ({ raw: 0, excluded: 0, byService: {}, byBadValue: {} });
    const byVoltage = { "2V": build(), "12V": build(), "(전압없음)": build() };
    (rows || []).forEach((station) => {
      (station.rectifiers || []).forEach((rect) => {
        const v = rect["전압"] || "(전압없음)";
        const bucket = byVoltage[v] || (byVoltage[v] = build());
        bucket.raw += 1;
        const bv = rect["불량여부"] === "" || rect["불량여부"] == null ? "(빈값)" : String(rect["불량여부"]);
        bucket.byBadValue[bv] = (bucket.byBadValue[bv] || 0) + 1;
        if (!isCountableRect(rect, excludedModels)) { bucket.excluded += 1; return; }
        const svc = rect["서비스"] || "(빈값)";
        bucket.byService[svc] = (bucket.byService[svc] || 0) + 1;
      });
    });
    return byVoltage;
  }, [rows, excludedModels]);

  const submitPin = () => {
    if (!isAdmin) { setPinMsg("PIN을 변경하려면 먼저 관리자 모드로 로그인하세요."); return; }
    if (pin1.length < 4) { setPinMsg("PIN은 4자리 이상으로 설정해주세요."); return; }
    if (pin1 !== pin2) { setPinMsg("입력한 PIN이 서로 다릅니다."); return; }
    onChangePin(pin1);
    setPin1(""); setPin2("");
    setPinMsg("관리자 PIN이 변경되었습니다.");
  };

  return (
    <div className="space-y-4">
      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
        <p className="text-sm font-semibold text-slate-700">관리자 · 엑셀 업로드</p>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-400">
          엑셀 파일이 여러 개라면, 대상(용도)에 맞는 칸에 각각 업로드하세요. 기지국·축전지 기본정보는 파일을 여러 개 추가할 수 있고, 추가한 파일들은 <b className="text-slate-500">통합시설코드</b>가
          정확히 일치하는 국소끼리만 하나로 합쳐져 함께 표시됩니다.
        </p>
      </div>

      {!isAdmin ? (
        <div className="rounded-2xl border border-dashed border-slate-200 bg-white p-6 text-center">
          <Lock size={20} className="mx-auto mb-2 text-slate-300" />
          <p className="text-sm font-medium text-slate-500">엑셀 업로드·삭제·초기화는 관리자만 할 수 있어요.</p>
          <p className="mt-1 text-[11px] text-slate-400">우측 상단 "일반 모드" 버튼을 눌러 PIN을 입력하면 관리자 모드로 전환됩니다.</p>
        </div>
      ) : (
        <>
          {baseFiles.length > 0 && (
            <div className="flex justify-end">
              <button onClick={onBaseReset} className="flex items-center gap-1 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs text-slate-500 hover:bg-slate-50">
                <RotateCcw size={13} /> 전체 초기화
              </button>
            </div>
          )}

          <AdminMultiUploadSlot
            title="기지국 · 축전지 기본정보"
            desc="홈 대시보드, 기지국 현황, 국소 상세 현황, 팀별 현황, Backup 시간에서 사용됩니다. 시설구분·중계기 정보와 정류기 ACTA 이력이 모두 이 한 파일에서 채워집니다. 정류기 등급이 바뀌면 ACTA 이력이 자동으로 누적됩니다. 파일을 여러 개 올리면 모두 통합시설코드 기준으로 합쳐져 함께 표시됩니다."
            files={baseFiles}
            loading={baseLoading}
            onUpload={onBaseUpload}
            onRemove={onRemoveBaseFile}
            countLabel="국소"
            getCount={(f) => f.stations.length}
            totalLabel="국소"
            totalCount={totalStations}
            guide={BASE_COLUMN_GUIDE}
            renderPreview={(f) => (
              <div>
                <p className="mb-1 text-slate-400">실제로 읽힌 값 미리보기 (처음 3건) — 열 위치가 맞는지 확인해보세요</p>
                {f.stations.length === 0 ? (
                  <p className="text-amber-600">이 파일에서 국소를 하나도 찾지 못했습니다. F열(통합시설코드) 또는 G열(국소명)이 비어있거나 열 위치가 다를 수 있습니다.</p>
                ) : (
                  <table className="w-full min-w-[560px] text-left">
                    <thead className="text-slate-400">
                      <tr>
                        <th className="pr-3 py-1">국소명(G)</th><th className="pr-3 py-1">통합시설코드(F)</th>
                        <th className="pr-3 py-1">본부(I)</th><th className="pr-3 py-1">SKT팀(J)</th>
                        <th className="pr-3 py-1">5G/4G/3G</th><th className="pr-3 py-1">정류기수</th>
                      </tr>
                    </thead>
                    <tbody>
                      {f.stations.slice(0, 3).map((s, i) => (
                        <tr key={i} className="border-t border-slate-200">
                          <td className="pr-3 py-1 text-slate-700">{s.국소명 || <span className="text-red-400">(없음)</span>}</td>
                          <td className="pr-3 py-1 text-slate-600">{s.통합시설코드 || <span className="text-red-400">(없음)</span>}</td>
                          <td className="pr-3 py-1 text-slate-500">{s.본부 || "-"}</td>
                          <td className="pr-3 py-1 text-slate-500">{s.SKT팀 || "-"}</td>
                          <td className="pr-3 py-1 text-slate-500">{s["5G"] || "-"}/{s["4G"] || "-"}/{s["3G"] || "-"}</td>
                          <td className="pr-3 py-1 text-slate-500">{s.rectifiers.length}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
                <RawColumnPreview rawPreview={f.rawPreview} highlightLetters={["F", "G", "I", "J", "M", "N", "S", "CC", "CD", "CE", "CH", "CI", "CJ", "CK", "CM", "X", "Z", "CB", "AK", "AX", "BC", "BD", "BN", "BO", "BP", "BQ", "BR", "BT", "BV", "BW"]} />
              </div>
            )}
          />
        </>
      )}

      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
        <p className="text-sm font-semibold text-slate-700">축전지 재고 대상 정보</p>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-400">
          더 이상 엑셀 업로드가 아니라, 축전지 재고 현황 화면에서 관리자가 대상을 직접 추가·입력합니다. 관리자 모드로 전환한 뒤 "축전지 재고 현황" 메뉴에서 편집해주세요.
        </p>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
        <div className="flex items-center gap-2">
          <History size={16} className="text-blue-500" />
          <p className="text-sm font-semibold text-slate-700">ACTA 내부저항측정 변동 이력</p>
        </div>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-400">
          기본정보가 바뀔 때마다 정류기별 등급(양호/열화/열화2/불량)을 이전 기록과 비교합니다. 값이 달라진 시점만 자동으로 쌓이며,
          홈 화면 &gt; 국소 상세 현황 &gt; ACTA 내부저항측정 이력에서 정류기번호를 클릭하면 확인할 수 있습니다.
        </p>
        <p className="mt-2 text-[11px] text-slate-500">현재 <b className="text-blue-600">{actaHistoryCount}개</b> 정류기의 변동 이력이 기록되어 있습니다.</p>
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
        <div className="flex items-center gap-2">
          <X size={16} className="text-blue-500" />
          <p className="text-sm font-semibold text-slate-700">집계에서 제외할 정류기 모델명</p>
        </div>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-400">
          여기에 추가한 모델명(Z열)을 가진 정류기는 대시보드·팀별현황·기지국현황 등 모든 집계와 상태 판정에서 빠집니다(표시는 그대로 됩니다).
        </p>
        {!isAdmin && (
          <p className="mt-2 flex items-center gap-1 text-[11px] text-slate-400">
            <Lock size={11} /> 추가·삭제는 관리자만 가능합니다. 목록은 아래에서 확인만 가능해요.
          </p>
        )}
        {isAdmin && (
          <div className="mt-3 flex flex-wrap items-center gap-2">
            <input
              list="admin-model-suggestions"
              value={modelInput}
              onChange={(e) => setModelInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && modelInput.trim()) { onAddExcludedModel(modelInput); setModelInput(""); } }}
              placeholder="정류기 모델명 입력 후 Enter"
              className="min-w-[180px] flex-1 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs outline-none focus:border-blue-400"
            />
            <datalist id="admin-model-suggestions">
              {availableModels.map((m) => <option key={m} value={m} />)}
            </datalist>
            <button
              onClick={() => { if (modelInput.trim()) { onAddExcludedModel(modelInput); setModelInput(""); } }}
              className="flex items-center gap-1 rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-blue-700"
            >
              <Plus size={13} /> 추가
            </button>
          </div>
        )}
        {isAdmin && availableModels.length > 0 && (
          <p className="mt-2 text-[11px] text-slate-400">
            현재 데이터에 있는 모델명: {availableModels.slice(0, 12).join(", ")}{availableModels.length > 12 ? " 외" : ""}
          </p>
        )}
        {(excludedModels || []).length > 0 ? (
          <div className="mt-3 flex flex-wrap gap-1.5">
            {excludedModels.map((m) => (
              <span key={m} className="flex items-center gap-1 rounded-full bg-red-50 px-2.5 py-1 text-xs font-medium text-red-600">
                {m}
                {isAdmin && (
                  <button onClick={() => onRemoveExcludedModel(m)} className="rounded-full p-0.5 hover:bg-red-100">
                    <X size={11} />
                  </button>
                )}
              </span>
            ))}
          </div>
        ) : (
          <p className="mt-3 text-[11px] text-slate-400">제외 중인 모델명이 없습니다.</p>
        )}
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
        <div className="flex items-center gap-2">
          <BarChart3 size={16} className="text-blue-500" />
          <p className="text-sm font-semibold text-slate-700">집계 근거 확인 (관리자 전용)</p>
        </div>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-400">
          홈 대시보드 숫자가 왜 그렇게 나오는지 확인용입니다. 전압(BD열)별 원본 행 수 · 제외된 건수(대상X·폐국·제외모델) · 서비스(CB열)별 분포를 그대로 보여줍니다.
        </p>
        {!isAdmin ? (
          <p className="mt-3 flex items-center gap-1 text-xs text-slate-400">
            <Lock size={12} /> 관리자 모드로 로그인해야 볼 수 있습니다.
          </p>
        ) : (
          <div className="mt-3 overflow-x-auto rounded-lg border border-slate-200 bg-slate-50/60 p-3 text-[11px]">
            <table className="w-full min-w-[520px] text-left">
              <thead className="text-slate-400">
                <tr>
                  <th className="pr-4 py-1">전압(BD열)</th>
                  <th className="pr-4 py-1">원본 행 수</th>
                  <th className="pr-4 py-1">제외됨(대상X·폐국·제외모델)</th>
                  <th className="pr-4 py-1">서비스(CB열)별 분포</th>
                  <th className="pr-4 py-1">불량여부(BV열) 실제 값 분포</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(diag).map(([voltage, d]) => (
                  <tr key={voltage} className="border-t border-slate-200">
                    <td className="pr-4 py-1.5 font-semibold text-slate-600">{voltage}</td>
                    <td className="pr-4 py-1.5 text-slate-700">{d.raw}건</td>
                    <td className="pr-4 py-1.5 text-slate-500">{d.excluded}건</td>
                    <td className="pr-4 py-1.5 text-slate-500">
                      {Object.keys(d.byService).length
                        ? Object.entries(d.byService).map(([svc, cnt]) => `${svc}: ${cnt}건`).join(" · ")
                        : "-"}
                    </td>
                    <td className="pr-4 py-1.5 text-slate-500">
                      {Object.keys(d.byBadValue).length
                        ? Object.entries(d.byBadValue).map(([bv, cnt]) => `"${bv}": ${cnt}건`).join(" · ")
                        : "-"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <p className="mt-2 text-slate-400">
              예) 12V 원본 행 수에서 "제외됨" 건수를 빼고, 서비스별 분포 중 W(또는 DU)만 더한 값이 홈 화면 12V 패널의 "전체"입니다. RT로 잡힌 만큼은 12V 패널이 아니라 RT 패널 쪽에 들어갑니다.
              불량 카운트는 "불량여부(BV열) 실제 값 분포"에서 정확히 <b>"불량"</b>이라고 찍힌 값만 셉니다 — 여기에 "불량 "(공백 포함)이나 다른 표기가 따옴표 안에 별도로 보이면 그게 원인입니다.
            </p>
          </div>
        )}
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
        <div className="flex items-center gap-2">
          <Lock size={16} className="text-blue-500" />
          <p className="text-sm font-semibold text-slate-700">관리자 PIN 변경</p>
        </div>
        <p className="mt-1 text-[11px] leading-relaxed text-slate-400">
          관리자 모드로 로그인해야 축전지 재고의 대상 추가·수정·저장, 작업예정일 입력이 가능합니다. 여기서 관리자 PIN을 바꿀 수 있어요(기본값 0000).
        </p>
        <div className="mt-3 flex flex-wrap items-center gap-2">
          <input type="password" value={pin1} onChange={(e) => setPin1(e.target.value.replace(/\D/g, ""))} placeholder="새 PIN"
            className="w-28 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs outline-none focus:border-blue-400" />
          <input type="password" value={pin2} onChange={(e) => setPin2(e.target.value.replace(/\D/g, ""))} placeholder="새 PIN 확인"
            className="w-28 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs outline-none focus:border-blue-400" />
          <button onClick={submitPin} className="rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-blue-700">
            변경
          </button>
        </div>
        {pinMsg && <p className="mt-2 text-[11px] text-slate-500">{pinMsg}</p>}
      </div>
    </div>
  );
}

/* ----------------------------------- App ----------------------------------- */
export default function App() {
  const [page, setPage] = useState("home");
  const [menuOpen, setMenuOpen] = useState(false);
  const [presetFilter, setPresetFilter] = useState(null);
  const [excludedModels, setExcludedModels] = useState([]);

  const goToStations = (preset) => {
    setPresetFilter(preset || null);
    setPage("stations");
  };

  const addExcludedModel = async (name) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    setExcludedModels((prev) => {
      if (prev.includes(trimmed)) return prev;
      const next = [...prev, trimmed];
      kv.set("excluded-rectifier-models", JSON.stringify(next)).catch((err) => console.error("[kv persist failed]", err));
      return next;
    });
  };
  const removeExcludedModel = (name) => {
    setExcludedModels((prev) => {
      const next = prev.filter((m) => m !== name);
      kv.set("excluded-rectifier-models", JSON.stringify(next)).catch((err) => console.error("[kv persist failed]", err));
      return next;
    });
  };

  const [baseFiles, setBaseFiles] = useState([]); // [{id, name, uploadedAt, stations}]
  const [baseLoading, setBaseLoading] = useState(false);

  const rows = useMemo(
    () => combineStationLists(baseFiles.map((f) => f.stations)),
    [baseFiles]
  );

  const [actaHistory, setActaHistory] = useState({});

  const [isAdmin, setIsAdmin] = useState(false);
  const [adminPin, setAdminPin] = useState("0000");

  // 관리자 모드가 꺼지면(로그아웃/시작 시 기본값) 관리자 화면에 남아있지 않도록 홈으로 이동시킨다.
  useEffect(() => {
    if (!isAdmin && page === "admin") setPage("home");
  }, [isAdmin, page]);

  const [searchQuery, setSearchQuery] = useState("");
  const [toast, setToast] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const res = await kv.get("battery-base-files");
        if (res?.value) setBaseFiles(JSON.parse(res.value) || []);
      } catch (e) { console.error("[kv load failed] battery-base-files", e); }
      try {
        const res3 = await kv.get("acta-history-log");
        if (res3?.value) setActaHistory(JSON.parse(res3.value) || {});
      } catch (e) { console.error("[kv load failed] acta-history-log", e); }
      try {
        const res4 = await kv.get("admin-pin");
        if (res4?.value) setAdminPin(res4.value);
      } catch (e) { console.error("[kv load failed] admin-pin", e); }
      try {
        const res5 = await kv.get("excluded-rectifier-models");
        if (res5?.value) setExcludedModels(JSON.parse(res5.value) || []);
      } catch (e) { console.error("[kv load failed] excluded-rectifier-models", e); }
    })();
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(""), 5000);
    return () => clearTimeout(t);
  }, [toast]);

  const [pinModalOpen, setPinModalOpen] = useState(false);
  const [pinInput, setPinInput] = useState("");
  const [pinError, setPinError] = useState("");

  const toggleAdmin = () => {
    if (isAdmin) { setIsAdmin(false); setToast("관리자 모드를 종료했습니다."); return; }
    setPinInput(""); setPinError(""); setPinModalOpen(true);
  };

  const submitPinModal = () => {
    if (pinInput === adminPin) {
      setIsAdmin(true);
      setPinModalOpen(false);
      setToast("관리자 모드로 전환되었습니다.");
    } else {
      setPinError("PIN이 올바르지 않습니다.");
    }
  };

  const changePin = async (newPin) => {
    setAdminPin(newPin);
    try { await kv.set("admin-pin", newPin); } catch (e) {}
  };

  // ① 기지국·축전지 기본정보 — 여러 개의 엑셀 파일을 각각 슬롯으로 추가할 수 있고, 모두 합쳐서(통합시설코드 기준) 화면에 함께 표시된다.
  //    같은 파일을 다시 올리는 게 아니라 "새 파일을 추가"하는 방식이라 누적 병합 걱정 없이 여러 개를 그대로 더할 수 있다.
  const handleBaseUpload = async (e) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    setBaseLoading(true);
    try {
      const newEntries = [];
      let totalRaw = 0, totalWithCode = 0;
      for (const file of files) {
        const { rows: flat, total, withCode, rawPreview } = await parseFixedColumnSheet(file);
        totalRaw += total; totalWithCode += withCode;
        const stations = groupStationRows(flat);
        newEntries.push({
          id: `${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
          name: file.name,
          uploadedAt: new Date().toLocaleString("ko-KR"),
          stations,
          rawPreview,
        });
      }
      // setState의 함수형 업데이터는 React가 나중에(비동기로) 실행할 수도 있어서, 그 안에서 만든 값을
      // 바로 다음 줄에서 쓰면 아직 비어있을 수 있다(→ 그대로 저장하면 DB에 value가 NULL로 들어감).
      // 그래서 next/toPersist는 지금 갖고 있는 baseFiles로 미리 계산해서 쓴다.
      const next = [...baseFiles, ...newEntries];
      const toPersist = next.map(({ rawPreview, ...rest }) => rest); // 미리보기는 저장 용량 아끼려고 세션에만 둔다
      setBaseFiles(next);
      const totalStations = newEntries.reduce((s, f) => s + f.stations.length, 0);
      try {
        await kv.set("battery-base-files", JSON.stringify(toPersist));
        if (totalWithCode === 0 && totalRaw > 0) {
          setToast(`⚠ ${totalRaw}행을 읽었지만 통합시설코드/국소명이 있는 행이 0건입니다. E열·F열 위치를 다시 확인해주세요.`);
        } else {
          setToast(`엑셀 업로드 완료 · 파일 ${newEntries.length}개, 국소 ${totalStations}건 추가됨 (읽은 행 ${totalRaw}건 중 ${totalWithCode}건 매칭)`);
        }
      } catch (persistErr) {
        setToast(`⚠ 화면에는 반영됐지만 서버 저장에 실패했습니다 — 새로고침하면 사라져요. (${persistErr?.message || persistErr})`);
      }
    } catch (err) {
      setToast(`엑셀 파일을 읽는 중 오류가 발생했습니다: ${err?.message || err}`);
    } finally {
      setBaseLoading(false);
      e.target.value = "";
    }
  };

  const removeBaseFile = async (id) => {
    const next = baseFiles.filter((f) => f.id !== id);
    setBaseFiles(next);
    try {
      await kv.set("battery-base-files", JSON.stringify(next));
    } catch (err) {
      setToast(`⚠ 삭제가 서버에 저장되지 않았습니다: ${err?.message || err}`);
    }
  };

  const resetBase = async () => {
    setBaseFiles([]);
    try {
      await kv.delete("battery-base-files");
      setToast("기본정보 데이터를 초기화했습니다.");
    } catch (err) {
      setToast(`⚠ 초기화가 서버에 반영되지 않았습니다: ${err?.message || err}`);
    }
  };

  // 합쳐진 국소 데이터(rows)가 바뀔 때마다 정류기 등급 변동을 감지해 ACTA 이력에 자동으로 쌓는다.
  useEffect(() => {
    if (!rows.length) return;
    setActaHistory((prev) => {
      const { history, changed } = buildActaHistory(rows, prev);
      if (changed) {
        kv.set("acta-history-log", JSON.stringify(history)).catch((err) => console.error("[kv persist failed]", err));
      }
      return changed ? history : prev;
    });
  }, [rows]);

  return (
    <div className="min-h-screen w-full bg-slate-50 text-slate-800">
      {/* Header */}
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto flex max-w-7xl items-center gap-3 px-4 py-3">
          <button className="md:hidden" onClick={() => setMenuOpen((v) => !v)}>
            {menuOpen ? <X size={20} /> : <Menu size={20} />}
          </button>
          <div className="flex items-center gap-2">
            <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-blue-600 text-white"><Radio size={16} /></div>
            <span className="text-sm font-bold text-slate-800 sm:text-base">기지국 · 축전지 통합 관리 시스템</span>
          </div>
          <div className="ml-auto flex items-center gap-2">
            <span className="hidden text-xs text-slate-400 lg:block">
              {baseFiles.length ? `국소 ${rows.length}건 · 파일 ${baseFiles.length}개` : "샘플 데이터 표시 중"}
            </span>
            {isAdmin && (
              <button onClick={() => setPage("admin")}
                title="관리자 메뉴에서 엑셀을 업로드하세요"
                className={`flex items-center gap-1 rounded-lg border px-2.5 py-1.5 text-xs font-medium hover:bg-slate-50 ${
                  page === "admin" ? "border-blue-300 text-blue-600 bg-blue-50" : "border-slate-200 text-slate-500"
                }`}>
                <Settings size={14} /> 관리자
              </button>
            )}
            <button onClick={toggleAdmin}
              title={isAdmin ? "클릭하면 관리자 모드가 종료됩니다" : "PIN을 입력하면 관리자 모드로 전환됩니다"}
              className={`flex items-center gap-1 rounded-lg border px-2.5 py-1.5 text-xs font-medium ${
                isAdmin ? "border-emerald-300 bg-emerald-50 text-emerald-600" : "border-slate-200 text-slate-500 hover:bg-slate-50"
              }`}>
              {isAdmin ? <Unlock size={14} /> : <Lock size={14} />} {isAdmin ? "관리자 모드" : "일반 모드"}
            </button>
          </div>
        </div>
        <nav className="hidden border-t border-slate-100 md:block">
          <div className="mx-auto flex max-w-7xl gap-1 px-4">
            {NAV.filter((n) => n.key !== "admin" || isAdmin).map(({ key, label, icon: Icon }) => (
              <button key={key} onClick={() => { setPage(key); setPresetFilter(null); }}
                className={`flex items-center gap-1.5 border-b-2 px-3 py-2.5 text-sm font-medium transition-colors ${
                  page === key ? "border-blue-600 text-blue-600" : "border-transparent text-slate-500 hover:text-slate-700"
                }`}>
                <Icon size={15} /> {label}
              </button>
            ))}
          </div>
        </nav>
        {menuOpen && (
          <div className="border-t border-slate-100 bg-white md:hidden">
            {NAV.filter((n) => n.key !== "admin" || isAdmin).map(({ key, label, icon: Icon }) => (
              <button key={key} onClick={() => { setPage(key); setPresetFilter(null); setMenuOpen(false); }}
                className={`flex w-full items-center gap-2 px-4 py-3 text-sm ${page === key ? "bg-blue-50 text-blue-600" : "text-slate-600"}`}>
                <Icon size={16} /> {label}
              </button>
            ))}
          </div>
        )}
      </header>

      {/* 관리자 PIN 모달 */}
      {pinModalOpen && (
        <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/40 px-4">
          <div className="w-full max-w-xs rounded-2xl bg-white p-5 shadow-xl">
            <div className="mb-3 flex items-center gap-2">
              <div className="flex h-9 w-9 items-center justify-center rounded-full bg-blue-50 text-blue-600"><Lock size={16} /></div>
              <div>
                <p className="text-sm font-semibold text-slate-700">관리자 모드</p>
                <p className="text-[11px] text-slate-400">PIN을 입력하세요</p>
              </div>
            </div>
            <input
              type="password"
              value={pinInput}
              autoFocus
              onChange={(e) => { setPinInput(e.target.value.replace(/\D/g, "")); setPinError(""); }}
              onKeyDown={(e) => { if (e.key === "Enter") submitPinModal(); if (e.key === "Escape") setPinModalOpen(false); }}
              placeholder="PIN 입력 (기본 0000)"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-center text-sm tracking-widest outline-none focus:border-blue-400"
            />
            {pinError && <p className="mt-2 text-center text-[11px] text-red-500">{pinError}</p>}
            <div className="mt-4 flex gap-2">
              <button onClick={() => setPinModalOpen(false)}
                className="flex-1 rounded-lg border border-slate-200 py-2 text-xs font-medium text-slate-500 hover:bg-slate-50">
                취소
              </button>
              <button onClick={submitPinModal}
                className="flex-1 rounded-lg bg-blue-600 py-2 text-xs font-medium text-white hover:bg-blue-700">
                확인
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toast && (
        <div className="fixed left-1/2 top-16 z-30 max-w-[92vw] -translate-x-1/2 rounded-lg bg-slate-800 px-4 py-2 text-center text-xs text-white shadow-lg sm:max-w-md">
          {toast}
        </div>
      )}

      {/* Content */}
      <main className="mx-auto max-w-7xl px-4 py-5 sm:py-6">
        {page === "home" && (
          <HomePage
            rows={rows} query={searchQuery} setQuery={setSearchQuery} actaHistory={actaHistory}
            lastUpdatedAt={baseFiles.length ? baseFiles[baseFiles.length - 1].uploadedAt : null}
            onDrill={goToStations} excludedModels={excludedModels}
          />
        )}
        {page === "stations" && (
          <StationTable rows={rows.length ? rows : [SAMPLE_ROW]} presetFilter={presetFilter} onClearPreset={() => setPresetFilter(null)} excludedModels={excludedModels} />
        )}
        {page === "battery" && (
          <BatteryStockPage isAdmin={isAdmin} />
        )}
        {page === "backup" && <BackupPage rows={rows} />}
        {page === "admin" && (
          <AdminPage
            baseFiles={baseFiles} baseLoading={baseLoading} totalStations={rows.length}
            onBaseUpload={handleBaseUpload} onRemoveBaseFile={removeBaseFile} onBaseReset={resetBase}
            actaHistoryCount={Object.keys(actaHistory).length}
            isAdmin={isAdmin} onChangePin={changePin}
            rows={rows} excludedModels={excludedModels} onAddExcludedModel={addExcludedModel} onRemoveExcludedModel={removeExcludedModel}
          />
        )}
      </main>
    </div>
  );
}
