const WEEKDAYS: string[] = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

/** Wandelt Date, 'YYYY-MM-DD' oder 'YYYY-MM-DDTHH:mm' in ein lokales Datum (00:00 Uhr) um. */
export function toLocalDate(x: unknown): Date {
  if (x instanceof Date) {
    return new Date(x.getFullYear(), x.getMonth(), x.getDate());
  }
  if (typeof x === 'string') {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(x);
    if (m) {
      return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    }
  }
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

export function today(): Date {
  return toLocalDate(new Date());
}

export function addDays(d: unknown, n: number): Date {
  const x = toLocalDate(d);
  x.setDate(x.getDate() + Number(n || 0));
  return x;
}

/** Addiert Werktage (Mo–Fr). Feiertage werden nicht berücksichtigt. */
export function addWorkdays(d: unknown, n: number): Date {
  const x = toLocalDate(d);
  let k = Number(n || 0);
  const step = k < 0 ? -1 : 1;
  k = Math.abs(k);
  while (k > 0) {
    x.setDate(x.getDate() + step);
    const w = x.getDay();
    if (w !== 0 && w !== 6) {
      k--;
    }
  }
  return x;
}

export function endOfMonth(d: unknown): Date {
  const x = toLocalDate(d);
  return new Date(x.getFullYear(), x.getMonth() + 1, 0);
}

export function toIsoDate(d: unknown): string {
  const x = toLocalDate(d);
  return `${x.getFullYear()}-${pad2(x.getMonth() + 1)}-${pad2(x.getDate())}`;
}

export function formatDate(d: unknown): string {
  const x = toLocalDate(d);
  return `${pad2(x.getDate())}.${pad2(x.getMonth() + 1)}.${x.getFullYear()}`;
}

export function formatDateWithWeekday(d: unknown): string {
  const x = toLocalDate(d);
  return `${WEEKDAYS[x.getDay()]}, ${formatDate(x)}`;
}

/** 'TT.MM.JJJJ' oder 'T.M.JJ' → Date; ungültig → null. */
export function parseSwissDate(s: string): Date | null {
  const m = /^\s*(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})\s*$/.exec(s || '');
  if (!m) {
    return null;
  }
  let year = Number(m[3]);
  if (year < 100) {
    year += 2000;
  }
  const d = new Date(year, Number(m[2]) - 1, Number(m[1]));
  return isNaN(d.getTime()) ? null : d;
}

/** Datumshilfen, die im Formular-Skript als form.date verfügbar sind. */
export const scriptDateHelpers = {
  today,
  addDays,
  addWorkdays,
  endOfMonth,
  format: formatDate,
  iso: toIsoDate
};

export function pad(n: number): string {
  return pad2(n);
}
