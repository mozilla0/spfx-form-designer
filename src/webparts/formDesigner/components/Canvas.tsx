import * as React from 'react';
import { useDroppable } from '@dnd-kit/core';
import { rectSortingStrategy, SortableContext, useSortable, verticalListSortingStrategy } from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { DefaultButton } from '@fluentui/react';
import { FormMode, IFormDefinition, ISection } from '../../../common/model';
import { FIELD_KIND_LABELS, ISpField } from '../../../common/spTypes';
import { hasValidationRules } from '../../../common/validation';
import { panelPreviewWidth } from '../../../common/panelSize';
import { FormHeader } from '../../../common/components/FormHeader';
import { Selection } from './designerState';
import styles from './FormDesigner.module.scss';

export interface ICanvasProps {
  definition: IFormDefinition;
  fieldsByName: Record<string, ISpField>;
  formMode: FormMode;
  activeTab: string;
  selection: Selection;
  overId: string | undefined;
  onSelect: (s: Selection) => void;
  onTabChange: (tabId: string) => void;
}

const GRIP = (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
    <circle cx="9" cy="6" r="1.6" />
    <circle cx="15" cy="6" r="1.6" />
    <circle cx="9" cy="12" r="1.6" />
    <circle cx="15" cy="12" r="1.6" />
    <circle cx="9" cy="18" r="1.6" />
    <circle cx="15" cy="18" r="1.6" />
  </svg>
);

export { GRIP };

function cx(...names: Array<string | false | undefined>): string {
  return names.filter(Boolean).join(' ');
}

interface IFieldCardProps {
  name: string;
  section: ISection;
  props: ICanvasProps;
}

const FieldCard: React.FC<IFieldCardProps> = ({ name, section, props }) => {
  const { definition, fieldsByName, selection, overId, onSelect } = props;
  const field = fieldsByName[name];
  const settings = definition.fields[name];
  const id = `field:${name}`;
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({
    id,
    data: { type: 'field', name, sectionId: section.id }
  });
  const selected = !!selection && selection.kind === 'field' && selection.name === name;
  const label = (settings && settings.label) || (field ? field.title : name);
  const kindLabel = field ? FIELD_KIND_LABELS[field.kind] : 'Spalte fehlt in der Liste';
  const required = !!settings && (settings.required || (!!field && field.required));
  const scripted = definition.script.indexOf(`'${name}'`) >= 0 || definition.script.indexOf(`"${name}"`) >= 0;

  return (
    <div
      ref={setNodeRef}
      className={cx(styles.fieldCard, selected && styles.fieldCardSelected, overId === id && !isDragging && styles.fieldCardOver)}
      style={{
        transform: CSS.Transform.toString(transform),
        transition,
        opacity: isDragging ? 0.4 : 1,
        gridColumn: section.columns === 2 && settings && settings.width === 'full' ? '1 / -1' : undefined
      }}
      {...attributes}
      {...listeners}
      aria-label={`Feld ${label}. Zum Verschieben Leertaste drücken.`}
      onClick={() => onSelect({ kind: 'field', name })}
      onFocus={() => onSelect({ kind: 'field', name })}
    >
      <div className={styles.fieldHead}>
        <span className={styles.fieldLabel}>{label}</span>
        {required && <span className={styles.required}>*</span>}
        <span className={styles.typeLabel}>{kindLabel}</span>
        {scripted && <span className={cx(styles.badge, styles.badgeScript)}>JS</span>}
        {settings && hasValidationRules(settings) && <span className={cx(styles.badge, styles.badgeRule)}>Prüfung</span>}
        {settings && settings.defaultValue.mode !== 'none' && <span className={cx(styles.badge, styles.badgeDefault)}>Standardwert</span>}
        {settings && settings.visibility.operator !== 'always' && <span className={cx(styles.badge, styles.badgeCond)}>Bedingt sichtbar</span>}
        {settings && settings.readOnly && <span className={styles.badge}>Schreibgeschützt</span>}
      </div>
      {field && field.kind !== 'boolean' && <div className={cx(styles.mockInput, field.kind === 'note' && styles.mockInputTall)} />}
    </div>
  );
};

interface ISectionProps {
  section: ISection;
  props: ICanvasProps;
}

const SectionBlock: React.FC<ISectionProps> = ({ section, props }) => {
  const id = `section:${section.id}`;
  const { attributes, listeners, setNodeRef, setActivatorNodeRef, transform, transition, isDragging } = useSortable({
    id,
    data: { type: 'section', id: section.id, tabId: section.tabId }
  });
  const bodyId = `sectionBody:${section.id}`;
  const { setNodeRef: setBodyRef } = useDroppable({ id: bodyId, data: { type: 'sectionBody', id: section.id } });
  const selected = !!props.selection && props.selection.kind === 'section' && props.selection.id === section.id;

  return (
    <section
      ref={setNodeRef}
      className={cx(styles.section, selected && styles.sectionSelected, props.overId === id && styles.sectionOver)}
      style={{ transform: CSS.Transform.toString(transform), transition, opacity: isDragging ? 0.5 : 1 }}
      aria-label={section.title || 'Abschnitt'}
    >
      <div className={styles.sectionHeader}>
        <h3 className={styles.sectionTitle}>
          <button type="button" ref={setActivatorNodeRef} className={styles.handle} aria-label={`Abschnitt ${section.title} verschieben`} {...attributes} {...listeners}>
            {GRIP}
          </button>
          {section.title || 'Abschnitt ohne Titel'}
        </h3>
        <DefaultButton text="Abschnitt bearbeiten" onClick={() => props.onSelect({ kind: 'section', id: section.id })} styles={{ root: { minWidth: 0, height: 28, padding: '0 8px' } }} />
      </div>
      <SortableContext items={section.fields.map((f) => `field:${f}`)} strategy={rectSortingStrategy}>
        <div
          ref={setBodyRef}
          className={cx(styles.sectionBody, props.overId === bodyId && styles.sectionBodyOver)}
          style={{ gridTemplateColumns: section.columns === 2 ? 'repeat(2, minmax(0, 1fr))' : 'minmax(0, 1fr)' }}
        >
          {section.fields.length === 0 && <div className={styles.empty}>Leerer Abschnitt – Spalte hierher ziehen oder links mit + einfügen.</div>}
          {section.fields.map((name) => (
            <FieldCard key={name} name={name} section={section} props={props} />
          ))}
        </div>
      </SortableContext>
    </section>
  );
};

const TabButton: React.FC<{ id: string; label: string; active: boolean; over: boolean; onClick: () => void }> = (p) => {
  const { setNodeRef } = useDroppable({ id: `tab:${p.id}`, data: { type: 'tab', id: p.id } });
  return (
    <button ref={setNodeRef} type="button" className={cx(styles.tab, p.active && styles.tabActive, p.over && styles.tabOver)} aria-pressed={p.active} onClick={p.onClick}>
      {p.label}
    </button>
  );
};

/** Die Formularfläche im Entwurfsmodus. */
export const Canvas: React.FC<ICanvasProps> = (props) => {
  const { definition, activeTab, formMode } = props;
  const sections = definition.sections.filter((s) => s.tabId === activeTab);
  const title = definition.titles[formMode] || '';
  return (
    <div className={styles.canvas}>
      <div className={styles.formSurface} style={{ maxWidth: panelPreviewWidth(definition.layout) || undefined }}>
        <div className={styles.formHeader}>
          <div style={{ flex: '1 1 auto', minWidth: 0 }}>
            <FormHeader compact title={title || '(Standardtitel)'} subtitle={definition.header ? definition.header.subtitle : undefined} header={definition.header} />
          </div>
          <DefaultButton text="Titel und Tabs" onClick={() => props.onSelect({ kind: 'form' })} styles={{ root: { flex: 'none' } }} />
        </div>
        <nav className={styles.tabs} aria-label="Tabs des Formulars">
          {definition.tabs.map((t) => (
            <TabButton key={t.id} id={t.id} label={t.label} active={t.id === activeTab} over={props.overId === `tab:${t.id}`} onClick={() => props.onTabChange(t.id)} />
          ))}
        </nav>
        <SortableContext items={sections.map((s) => `section:${s.id}`)} strategy={verticalListSortingStrategy}>
          <div className={styles.sections}>
            {sections.length === 0 && <div className={styles.empty}>Dieser Tab hat noch keinen Abschnitt.</div>}
            {sections.map((s) => (
              <SectionBlock key={s.id} section={s} props={props} />
            ))}
          </div>
        </SortableContext>
      </div>
    </div>
  );
};
