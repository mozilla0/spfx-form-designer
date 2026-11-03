import { IFormLayout, PanelSize } from './model';
import { t } from './i18n';

/**
 * Experimentell: verbreitert das SharePoint-Panel, in dem die Runtime gerendert wird.
 * Microsoft bietet dafür keine API. Die Anpassung sucht den Panel-Container im DOM und
 * überschreibt dessen Breite. Wird kein Panel gefunden, bleibt alles unverändert.
 */

export const PANEL_SIZES: Array<{ key: PanelSize; text: () => string; width?: number }> = [
  { key: 'default', text: () => t('Standard') },
  { key: 'small', text: () => t('Schmal (480 px)'), width: 480 },
  { key: 'medium', text: () => t('Mittel (640 px)'), width: 640 },
  { key: 'large', text: () => t('Breit (880 px)'), width: 880 },
  { key: 'xlarge', text: () => t('Sehr breit (1100 px)'), width: 1100 },
  { key: 'wide', text: () => t('Fast ganze Breite (90 %)') },
  { key: 'custom', text: () => t('Eigene Breite') }
];

/** CSS-Breite für das Panel oder undefined für die Standardbreite. */
export function panelCssWidth(layout: IFormLayout | undefined): string | undefined {
  if (!layout || layout.panelSize === 'default') {
    return undefined;
  }
  if (layout.panelSize === 'wide') {
    return '90vw';
  }
  const preset = PANEL_SIZES.filter((s) => s.key === layout.panelSize)[0];
  const px = layout.panelSize === 'custom' ? layout.panelWidth : preset ? preset.width : undefined;
  if (!px || px < 320) {
    return undefined;
  }
  return `${Math.round(px)}px`;
}

/** Breite in Pixeln für die Darstellung im Designer (Vorschau der Formularbreite). */
export function panelPreviewWidth(layout: IFormLayout | undefined): number | undefined {
  const css = panelCssWidth(layout);
  if (!css) {
    return undefined;
  }
  return css.indexOf('px') > 0 ? parseInt(css, 10) : undefined;
}

/** Sucht den Container des Seitenpanels, in dem das Element liegt. */
export function findPanelContainer(el: HTMLElement): HTMLElement | undefined {
  const fluent = el.closest('.ms-Panel-main') as HTMLElement | null;
  if (fluent) {
    return fluent;
  }
  // Rückfall: nächster fixierter oder absolut positionierter Vorfahre, der am rechten Rand andockt.
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  let node: HTMLElement | null = el.parentElement;
  while (node && node !== document.body) {
    const style = window.getComputedStyle(node);
    if (style.position === 'fixed' || style.position === 'absolute') {
      const r = node.getBoundingClientRect();
      const docksRight = Math.abs(vw - r.right) <= 4;
      const tall = r.height >= vh * 0.6;
      const notFullPage = r.width < vw * 0.95;
      if (docksRight && tall && notFullPage) {
        return node;
      }
    }
    node = node.parentElement;
  }
  return undefined;
}

const PROPS: string[] = ['width', 'max-width', 'min-width'];

/**
 * Setzt die Panelbreite und liefert eine Funktion zum Zurücksetzen.
 * Versucht es mehrmals, weil SharePoint das Panel animiert einblendet.
 */
export function applyPanelWidth(el: HTMLElement, cssWidth: string | undefined): () => void {
  if (!cssWidth) {
    return () => undefined;
  }
  let target: HTMLElement | undefined;
  const previous: Record<string, [string, string]> = {};
  const timers: number[] = [];

  const apply = (): boolean => {
    if (target) {
      return true;
    }
    const found = findPanelContainer(el);
    if (!found) {
      return false;
    }
    target = found;
    PROPS.forEach((p) => {
      previous[p] = [found.style.getPropertyValue(p), found.style.getPropertyPriority(p)];
    });
    found.style.setProperty('width', `min(${cssWidth}, 100vw)`, 'important');
    found.style.setProperty('max-width', '100vw', 'important');
    found.style.setProperty('min-width', '0', 'important');
    found.setAttribute('data-form-designer-width', cssWidth);
    return true;
  };

  [0, 150, 500, 1200].forEach((delay) => {
    timers.push(window.setTimeout(apply, delay));
  });

  return () => {
    timers.forEach((t) => window.clearTimeout(t));
    if (target) {
      const t = target;
      PROPS.forEach((p) => {
        const [value, priority] = previous[p] || ['', ''];
        if (value) {
          t.style.setProperty(p, value, priority);
        } else {
          t.style.removeProperty(p);
        }
      });
      t.removeAttribute('data-form-designer-width');
    }
  };
}
