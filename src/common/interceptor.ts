import { IFormRequest, parseFormUrl } from './formUrl';

/**
 * Fängt das Öffnen von Listenformularen auf der Seite ab und leitet es an eine registrierte Liste weiter.
 * Drei Wege werden beobachtet: Navigation API (Navigationen per Skript oder Link),
 * Klicks auf Links (Capture-Phase) und window.open.
 *
 * Diagnose (gilt für die Browser-Sitzung):
 *   ?fdDebug=log  – nur protokollieren, nichts umleiten
 *   ?fdDebug=on   – protokollieren und umleiten
 *   ?fdDebug=off  – Diagnose beenden (Standard: umleiten ohne Protokoll)
 */

export type DebugMode = 'off' | 'log' | 'on';

/** Liefert true, wenn die Liste das Formular übernimmt. */
export type FormRequestHandler = (req: IFormRequest, via: string) => boolean;

interface IRegistration {
  listId: string;
  listPath: string;
  handler: FormRequestHandler;
}

const STORAGE_KEY = 'formDesignerDebug';
const SKIP_KEY = 'formDesignerSkipUrl';
const registrations: IRegistration[] = [];
let installed = false;
let debugMode: DebugMode = 'off';

export function getDebugMode(): DebugMode {
  return debugMode;
}

function readDebugMode(): DebugMode {
  try {
    const m = /[?&]fdDebug=(log|on|off)\b/i.exec(window.location.search);
    if (m) {
      const v = m[1].toLowerCase() as DebugMode;
      if (v === 'off') {
        window.sessionStorage.removeItem(STORAGE_KEY);
      } else {
        window.sessionStorage.setItem(STORAGE_KEY, v);
      }
      return v;
    }
    const stored = window.sessionStorage.getItem(STORAGE_KEY);
    return stored === 'log' || stored === 'on' ? stored : 'off';
  } catch {
    return 'off';
  }
}

export function fdLog(...args: unknown[]): void {
  if (debugMode !== 'off') {
    console.log('%c[Form Designer]', 'color:#0f6cbd;font-weight:bold', ...args);
  }
}

/** Lässt die nächste Navigation zu dieser URL unbehelligt (z. B. Rückfall auf die ganze Seite). */
export function allowNextNavigation(url: string): void {
  try {
    window.sessionStorage.setItem(SKIP_KEY, url);
  } catch {
    // ignorieren
  }
}

function consumeSkip(url: string): boolean {
  try {
    if (window.sessionStorage.getItem(SKIP_KEY) === url) {
      window.sessionStorage.removeItem(SKIP_KEY);
      return true;
    }
  } catch {
    // ignorieren
  }
  return false;
}

/** Letzte Entscheidung, um doppelt gemeldete Navigationen zur selben URL nur einmal zu behandeln. */
let lastDecision: { url: string; at: number; handled: boolean } | undefined;
const DUPLICATE_WINDOW_MS = 1500;

function dispatch(req: IFormRequest | undefined, via: string): boolean {
  if (!req) {
    return false;
  }
  const now = Date.now();
  if (lastDecision && lastDecision.url === req.url && now - lastDecision.at < DUPLICATE_WINDOW_MS) {
    fdLog(`${via}: doppelte Meldung derselben Navigation – ${lastDecision.handled ? 'ebenfalls abgefangen' : 'ignoriert'}`);
    return lastDecision.handled;
  }
  const handled = decide(req, via);
  lastDecision = { url: req.url, at: now, handled };
  return handled;
}

function decide(req: IFormRequest, via: string): boolean {
  if (consumeSkip(req.url)) {
    fdLog(`${via}: übersprungen (Rückfall auf ganze Seite)`, req);
    return false;
  }
  const reg = registrations.filter(
    (r) => (req.listId && req.listId === r.listId) || (!req.listId && req.listPath && req.listPath === r.listPath)
  )[0];
  if (!reg) {
    fdLog(`${via}: Formular erkannt, aber nicht für eine registrierte Liste`, req);
    return false;
  }
  if (debugMode === 'log') {
    fdLog(`${via}: würde übernommen (Modus «log», keine Umleitung)`, req);
    return false;
  }
  const handled = reg.handler(req, via);
  fdLog(`${via}: ${handled ? 'übernommen' : 'an SharePoint weitergegeben'}`, req);
  return handled;
}

function describeTarget(el: Element | null): string {
  const parts: string[] = [];
  let node: Element | null = el;
  let depth = 0;
  while (node && depth < 6) {
    const auto = node.getAttribute('data-automationid');
    const key = node.getAttribute('data-automation-key');
    const label = node.getAttribute('aria-label') || node.getAttribute('name');
    if (auto || key || label) {
      parts.push(`${node.tagName.toLowerCase()}${auto ? `[automationid=${auto}]` : ''}${key ? `[key=${key}]` : ''}${label ? `[«${label}»]` : ''}`);
    }
    node = node.parentElement;
    depth++;
  }
  return parts.join(' ← ') || (el ? el.tagName.toLowerCase() : '?');
}

function install(): void {
  if (installed) {
    return;
  }
  installed = true;

  // 1. Navigation API (Chromium, neuere Firefox-Versionen)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const nav = (window as any).navigation;
  if (nav && typeof nav.addEventListener === 'function') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    nav.addEventListener('navigate', (e: any) => {
      if (!e.cancelable || e.hashChange || e.downloadRequest) {
        return;
      }
      if (dispatch(parseFormUrl(e.destination && e.destination.url), 'navigate')) {
        e.preventDefault();
      }
    });
    fdLog('Navigation API aktiv');
  } else {
    fdLog('Navigation API nicht verfügbar – nur Klicks und window.open werden abgefangen');
  }

  // 2. Klicks auf Links
  document.addEventListener(
    'click',
    (e: MouseEvent) => {
      const target = e.target as Element | null;
      if (debugMode !== 'off') {
        fdLog('Klick:', describeTarget(target));
      }
      if (e.defaultPrevented || e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) {
        return;
      }
      const anchor = target && target.closest ? (target.closest('a[href]') as HTMLAnchorElement | null) : null;
      if (anchor && dispatch(parseFormUrl(anchor.href), 'click')) {
        e.preventDefault();
        e.stopPropagation();
      }
    },
    true
  );

  // 3. window.open (z. B. wenn SharePoint Formulare in einem neuen Tab öffnet)
  const originalOpen = window.open;
  window.open = function (url?: string | URL, target?: string, features?: string): Window | null {
    if (dispatch(parseFormUrl(url), 'window.open')) {
      return null;
    }
    return originalOpen.call(window, url, target, features);
  } as typeof window.open;
}

/** Registriert eine Liste. Mehrere Listen auf einer Seite (z. B. Listen-Webparts) sind möglich. */
export function registerList(listId: string, listServerRelativeUrl: string, handler: FormRequestHandler): void {
  debugMode = readDebugMode();
  const id = listId.replace(/[{}]/g, '').toLowerCase();
  const path = decodeURIComponent(listServerRelativeUrl).toLowerCase().replace(/\/$/, '');
  for (let i = registrations.length - 1; i >= 0; i--) {
    if (registrations[i].listId === id) {
      registrations.splice(i, 1);
    }
  }
  registrations.push({ listId: id, listPath: path, handler });
  install();
  fdLog(`Liste registriert: ${id} (${path}), Modus «${debugMode === 'off' ? 'normal' : debugMode}»`);
}

export function unregisterList(listId: string): void {
  const id = listId.replace(/[{}]/g, '').toLowerCase();
  for (let i = registrations.length - 1; i >= 0; i--) {
    if (registrations[i].listId === id) {
      registrations.splice(i, 1);
    }
  }
}
