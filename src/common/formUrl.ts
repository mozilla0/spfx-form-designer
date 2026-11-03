import { FormMode } from './model';

/** Ein erkannter Aufruf eines Listenformulars. */
export interface IFormRequest {
  mode: FormMode;
  url: string;
  /** Listen-ID in Kleinbuchstaben ohne Klammern, falls in der URL enthalten. */
  listId?: string;
  /** Serverrelativer Pfad der Liste in Kleinbuchstaben, falls aus dem Pfad ableitbar. */
  listPath?: string;
  itemId?: number;
  contentTypeId?: string;
  source?: string;
}

const PAGE_TYPES: Record<string, FormMode> = { '4': 'display', '6': 'edit', '8': 'new' };
const CLASSIC_FORMS: Record<string, FormMode> = { 'dispform.aspx': 'display', 'editform.aspx': 'edit', 'newform.aspx': 'new' };

export function normalizeGuid(value: string | null | undefined): string | undefined {
  if (!value) {
    return undefined;
  }
  const v = decodeURIComponent(value).replace(/[{}]/g, '').trim().toLowerCase();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(v) ? v : undefined;
}

/** Liest einen Query-Parameter ohne Rücksicht auf Gross-/Kleinschreibung (bereits dekodiert). */
function param(url: URL, name: string): string | null {
  const wanted = name.toLowerCase();
  let found: string | null = null;
  url.searchParams.forEach((value, key) => {
    if (found === null && key.toLowerCase() === wanted) {
      found = value;
    }
  });
  return found;
}

/**
 * Erkennt URLs von SharePoint-Listenformularen:
 * /_layouts/15/SPListForm.aspx, /_layouts/15/listform.aspx sowie …/NewForm.aspx, EditForm.aspx, DispForm.aspx.
 */
export function parseFormUrl(raw: string | URL | null | undefined, base?: string): IFormRequest | undefined {
  if (!raw) {
    return undefined;
  }
  let url: URL;
  try {
    url = new URL(String(raw), base || (typeof window !== 'undefined' ? window.location.href : 'https://localhost/'));
  } catch {
    return undefined;
  }
  const path = url.pathname.toLowerCase();
  const file = path.slice(path.lastIndexOf('/') + 1);
  const idText = param(url, 'ID');
  const itemId = idText && /^\d+$/.test(idText) ? Number(idText) : undefined;
  const contentTypeId = param(url, 'ContentTypeId') || undefined;
  const source = param(url, 'Source') || undefined;

  if (file === 'splistform.aspx' || file === 'listform.aspx') {
    const mode = PAGE_TYPES[param(url, 'PageType') || ''];
    if (!mode) {
      return undefined;
    }
    return {
      mode,
      url: url.href,
      listId: normalizeGuid(param(url, 'List') || param(url, 'ListId')),
      itemId,
      contentTypeId,
      source
    };
  }

  const classic = CLASSIC_FORMS[file];
  if (classic) {
    return { mode: classic, url: url.href, listPath: decodeURIComponent(path.slice(0, path.lastIndexOf('/'))), itemId, contentTypeId, source };
  }
  return undefined;
}
