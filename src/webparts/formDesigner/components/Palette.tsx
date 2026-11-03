import * as React from 'react';
import { useDraggable, useDroppable } from '@dnd-kit/core';
import { DefaultButton, IconButton } from '@fluentui/react';
import { fieldKindLabel, ISpField } from '../../../common/spTypes';
import { GRIP } from './Canvas';
import styles from './FormDesigner.module.scss';
import { t } from '../../../common/i18n';

export interface IPaletteProps {
  unplaced: ISpField[];
  dragActiveType: string | undefined;
  overId: string | undefined;
  onAdd: (name: string) => void;
  onAddSection: () => void;
  onAddTab: () => void;
}

const PaletteItem: React.FC<{ field: ISpField; onAdd: (name: string) => void }> = ({ field, onAdd }) => {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: `palette:${field.internalName}`,
    data: { type: 'palette', name: field.internalName }
  });
  return (
    <div ref={setNodeRef} className={styles.paletteItem} style={{ opacity: isDragging ? 0.4 : 1 }} {...attributes} {...listeners} aria-label={t('Spalte {0} ins Formular ziehen', field.title)}>
      <span style={{ color: '#8a8886', display: 'flex' }}>{GRIP}</span>
      <span className={styles.paletteItemText}>
        <strong>{field.title}</strong>
        <span className={styles.typeLabel}>{fieldKindLabel(field.kind)}</span>
      </span>
      <IconButton
        iconProps={{ iconName: 'Add' }}
        title={t('{0} einfügen', field.title)}
        ariaLabel={t('{0} einfügen', field.title)}
        onPointerDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => e.stopPropagation()}
        onClick={() => onAdd(field.internalName)}
      />
    </div>
  );
};

/** Werkzeugkasten mit den noch nicht platzierten Spalten. Dient auch als Ablage zum Entfernen. */
export const Palette: React.FC<IPaletteProps> = (props) => {
  const { setNodeRef } = useDroppable({ id: 'paletteZone', data: { type: 'paletteZone' } });
  const removing = props.dragActiveType === 'field' && props.overId === 'paletteZone';
  return (
    <aside className={styles.palette} aria-label={t('Werkzeugkasten')}>
      <div ref={setNodeRef} className={`${styles.panel} ${styles.paletteZone} ${removing ? styles.paletteZoneOver : ''}`}>
        <p className={styles.panelTitle}>{t('Spalten der Liste')}</p>
        {props.unplaced.length === 0 && <span className={styles.hint}>{t('Alle Spalten sind im Formular platziert.')}</span>}
        {props.unplaced.map((f) => (
          <PaletteItem key={f.internalName} field={f} onAdd={props.onAdd} />
        ))}
        <span className={styles.hint}>{t('Spalten ins Formular ziehen oder mit + in den gewählten Abschnitt einfügen. Ein Feld hierher zurückziehen entfernt es aus dem Formular.')}</span>
      </div>
      <div className={`${styles.panel} ${styles.paletteZone}`}>
        <p className={styles.panelTitle}>{t('Layout')}</p>
        <DefaultButton text={t('Abschnitt hinzufügen')} iconProps={{ iconName: 'GridViewSmall' }} onClick={props.onAddSection} />
        <DefaultButton text={t('Tab hinzufügen')} iconProps={{ iconName: 'Tab' }} onClick={props.onAddTab} />
      </div>
    </aside>
  );
};
