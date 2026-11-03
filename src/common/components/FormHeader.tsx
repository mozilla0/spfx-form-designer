import * as React from 'react';
import { Icon, Text } from '@fluentui/react';
import { IFormHeader } from '../model';

export const DEFAULT_ACCENT: string = '#0f6cbd';

export interface IFormHeaderProps {
  title: string;
  subtitle?: string;
  header: IFormHeader | undefined;
  /** Kompakte Darstellung im Designer. */
  compact?: boolean;
}

/** Relative Helligkeit nach WCAG, um die Textfarbe auf farbigem Grund zu wählen. */
function luminance(hex: string): number {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) {
    return 0;
  }
  const n = parseInt(m[1], 16);
  const channel = (c: number): number => {
    const v = c / 255;
    return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel((n >> 16) & 255) + 0.7152 * channel((n >> 8) & 255) + 0.0722 * channel(n & 255);
}

export function textColorOn(background: string): string {
  // Schwarz oder Weiss – je nachdem, was mehr Kontrast ergibt.
  const l = luminance(background);
  return (l + 0.05) / 0.05 > 1.05 / (l + 0.05) ? '#1b1a19' : '#ffffff';
}

export function isHexColor(v: string | undefined): boolean {
  return !!v && /^#[0-9a-f]{6}$/i.test(v.trim());
}

/** Kopfbereich mit optionalem Icon, Logo und Untertitel. */
export const FormHeader: React.FC<IFormHeaderProps> = ({ title, subtitle, header, compact }) => {
  const h = header || {};
  const accent = isHexColor(h.accentColor) ? (h.accentColor as string).trim() : DEFAULT_ACCENT;
  const style = h.style || 'plain';
  const band = style === 'band';
  const fg = band ? textColorOn(accent) : undefined;
  const iconSize = compact ? 22 : 28;
  const imageHeight = compact ? 32 : 44;
  const image = h.imageUrl ? (
    <img src={h.imageUrl} alt="" style={{ height: imageHeight, maxWidth: compact ? 120 : 180, objectFit: 'contain', flex: 'none', marginLeft: h.imagePosition === 'right' ? 'auto' : undefined }} />
  ) : null;

  return (
    <header
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: 12,
        padding: band ? (compact ? '10px 14px' : '14px 18px') : style === 'line' ? '0 0 10px' : 0,
        background: band ? accent : undefined,
        color: fg,
        borderRadius: band ? 4 : undefined,
        borderBottom: style === 'line' ? `3px solid ${accent}` : undefined,
        minWidth: 0
      }}
    >
      {image && h.imagePosition !== 'right' && image}
      {h.icon && <Icon iconName={h.icon} aria-hidden="true" style={{ fontSize: iconSize, color: band ? fg : accent, flex: 'none', lineHeight: 1 }} />}
      <div style={{ minWidth: 0 }}>
        <Text as="h1" variant={compact ? 'large' : 'xLarge'} styles={{ root: { margin: 0, fontWeight: 600, color: fg, overflowWrap: 'anywhere', display: 'block' } }}>
          {title}
        </Text>
        {subtitle && (
          <Text variant="medium" styles={{ root: { color: band ? fg : '#605e5c', display: 'block', marginTop: 2, opacity: band ? 0.9 : 1 } }}>
            {subtitle}
          </Text>
        )}
      </div>
      {image && h.imagePosition === 'right' && image}
    </header>
  );
};
