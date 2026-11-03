import * as React from 'react';
import {
  closestCenter,
  CollisionDetection,
  DndContext,
  DragEndEvent,
  DragOverEvent,
  DragOverlay,
  DragStartEvent,
  KeyboardSensor,
  PointerSensor,
  pointerWithin,
  useSensor,
  useSensors
} from '@dnd-kit/core';
import { sortableKeyboardCoordinates } from '@dnd-kit/sortable';
import {
  ChoiceGroup,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  TextField
} from '@fluentui/react';
import { RUNTIME_COMPONENT_ID } from '../../../common/constants';
import { createDefinition, FormMode, IFormDefinition, IRuntimeProperties, normalizeDefinition } from '../../../common/model';
import { IContentTypeInfo, IListInfo, SpService } from '../../../common/spService';
import { FieldValues, ICurrentUser, ISpField } from '../../../common/spTypes';
import { FormHost } from '../../../common/components/FormHost';
import { panelPreviewWidth } from '../../../common/panelSize';
import { Canvas } from './Canvas';
import { Palette } from './Palette';
import { PropertiesPanel } from './PropertiesPanel';
import { ScriptEditor } from './ScriptEditor';
import { DeployDialog } from './DeployDialog';
import { TranslationsEditor } from './TranslationsEditor';
import { LinkBuilder } from './LinkBuilder';
import { readUrlParams } from '../../../common/urlPrefill';
import { ContentLanguage } from '../../../common/model';
import { getLanguage } from '../../../common/i18n';
import * as ds from './designerState';
import styles from './FormDesigner.module.scss';
import { t } from '../../../common/i18n';

export interface IFormDesignerProps {
  sp: SpService;
  currentUser: ICurrentUser;
}

type View = 'design' | 'script' | 'translations' | 'preview';

interface IStatus {
  kind: MessageBarType;
  text: string;
}

const PRIORITY: Record<string, number> = { field: 0, tab: 1, paletteZone: 2, sectionBody: 3, section: 4 };

/** Kollisionserkennung: das spezifischste Ziel unter dem Mauszeiger gewinnt. */
const collision: CollisionDetection = (args) => {
  const type = args.active.data.current ? args.active.data.current.type : undefined;
  const allowed = type === 'section' ? ['section', 'tab'] : type === 'palette' ? ['field', 'sectionBody', 'tab'] : ['field', 'sectionBody', 'tab', 'paletteZone'];
  const containers = args.droppableContainers.filter((c) => c.id !== args.active.id && c.data.current && allowed.indexOf(c.data.current.type) >= 0);
  const filtered = { ...args, droppableContainers: containers };
  const hits = args.pointerCoordinates ? pointerWithin(filtered) : closestCenter(filtered);
  return hits.slice().sort((a, b) => {
    const ta = a.data && a.data.droppableContainer ? a.data.droppableContainer.data.current.type : '';
    const tb = b.data && b.data.droppableContainer ? b.data.droppableContainer.data.current.type : '';
    return (PRIORITY[ta] ?? 9) - (PRIORITY[tb] ?? 9);
  });
};

function definitionUrlOf(ct: IContentTypeInfo): string | undefined {
  const pairs: Array<[string | null, string | null]> = [
    [ct.NewFormClientSideComponentId, ct.NewFormClientSideComponentProperties],
    [ct.EditFormClientSideComponentId, ct.EditFormClientSideComponentProperties],
    [ct.DisplayFormClientSideComponentId, ct.DisplayFormClientSideComponentProperties]
  ];
  for (const [id, props] of pairs) {
    if (id && id.replace(/[{}]/g, '').toLowerCase() === RUNTIME_COMPONENT_ID && props) {
      try {
        const parsed = JSON.parse(props) as IRuntimeProperties;
        if (parsed.definitionUrl) {
          return parsed.definitionUrl;
        }
      } catch {
        // ignorieren
      }
    }
  }
  return undefined;
}

/** WYSIWYG-Designer für SharePoint-Listenformulare. */
const FormDesigner: React.FC<IFormDesignerProps> = ({ sp, currentUser }) => {
  const [lists, setLists] = React.useState<IListInfo[]>([]);
  const [listId, setListId] = React.useState<string>('');
  const [contentTypes, setContentTypes] = React.useState<IContentTypeInfo[]>([]);
  const [ctId, setCtId] = React.useState<string>('');
  const [fields, setFields] = React.useState<ISpField[]>([]);
  const [definition, setDefinition] = React.useState<IFormDefinition | undefined>(undefined);
  const [view, setView] = React.useState<View>('design');
  const [formMode, setFormMode] = React.useState<FormMode>('new');
  const [selection, setSelection] = React.useState<ds.Selection>(undefined);
  const [activeTab, setActiveTab] = React.useState<string>('');
  const [loading, setLoading] = React.useState<boolean>(false);
  const [status, setStatus] = React.useState<IStatus | undefined>(undefined);
  const [dirty, setDirty] = React.useState<boolean>(false);
  const [deployOpen, setDeployOpen] = React.useState<boolean>(false);
  const [previewKey, setPreviewKey] = React.useState<number>(0);
  const [previewItemId, setPreviewItemId] = React.useState<string>('');
  const [previewValues, setPreviewValues] = React.useState<FieldValues>({});
  const [linkBuilderOpen, setLinkBuilderOpen] = React.useState<boolean>(false);
  const [previewQuery, setPreviewQuery] = React.useState<string>('');
  const [previewLanguage, setPreviewLanguage] = React.useState<ContentLanguage>(getLanguage());
  const [drag, setDrag] = React.useState<{ id: string; type: string; label: string } | undefined>(undefined);
  const [overId, setOverId] = React.useState<string | undefined>(undefined);

  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 6 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
  );

  const fieldsByName = React.useMemo(() => {
    const map: Record<string, ISpField> = {};
    fields.forEach((f) => {
      map[f.internalName] = f;
    });
    return map;
  }, [fields]);

  const fail = (e: unknown): void => setStatus({ kind: MessageBarType.error, text: (e as Error).message });

  // Listen laden
  React.useEffect(() => {
    sp.getLists().then(setLists).catch(fail);
  }, []);

  // Inhaltstypen laden
  React.useEffect(() => {
    if (!listId) {
      return;
    }
    setDefinition(undefined);
    setContentTypes([]);
    setCtId('');
    sp.getContentTypes(listId)
      .then((cts) => {
        setContentTypes(cts);
        if (cts.length) {
          setCtId(cts[0].id);
        }
      })
      .catch(fail);
  }, [listId]);

  const loadDefinition = async (): Promise<void> => {
    const ct = contentTypes.filter((c) => c.id === ctId)[0];
    if (!listId || !ct) {
      return;
    }
    setLoading(true);
    setStatus(undefined);
    try {
      const loadedFields = await sp.getFields(listId, ct.id);
      let def = await sp.readDefinition(listId, ct.id, 'draft');
      if (!def) {
        const url = definitionUrlOf(ct);
        def = url ? await sp.readJson<IFormDefinition>(url) : undefined;
      }
      const usable = loadedFields.filter((f) => f.kind !== 'unsupported');
      const normalized = def ? normalizeDefinition(def) : createDefinition(usable.map((f) => f.internalName));
      setFields(loadedFields);
      setDefinition(normalized);
      setActiveTab(normalized.tabs[0] ? normalized.tabs[0].id : '');
      setSelection(undefined);
      setDirty(false);
      if (!def) {
        setStatus({ kind: MessageBarType.info, text: t('Für diesen Inhaltstyp gibt es noch kein Formular. Alle unterstützten Spalten wurden in einen Abschnitt gelegt.') });
      }
    } catch (e) {
      fail(e);
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => {
    loadDefinition().catch(fail);
  }, [ctId]);

  const update = (def: IFormDefinition, change: ds.SelectionChange = 'keep'): void => {
    setDefinition(def);
    setDirty(true);
    if (change === 'clear') {
      setSelection(undefined);
    } else if (change !== 'keep') {
      setSelection(change);
    }
  };

  const saveDraft = async (): Promise<void> => {
    if (!definition) {
      return;
    }
    try {
      await sp.saveDefinition(listId, ctId, definition, 'draft');
      setDirty(false);
      setStatus({ kind: MessageBarType.success, text: t('Entwurf gespeichert. Erst «Bereitstellen» macht die Änderungen in der Liste sichtbar.') });
    } catch (e) {
      fail(e);
    }
  };

  const targetSectionId = (): string | undefined => {
    if (!definition) {
      return undefined;
    }
    if (selection && selection.kind === 'section') {
      return selection.id;
    }
    if (selection && selection.kind === 'field') {
      const s = ds.sectionOfField(definition, selection.name);
      if (s) {
        return s.id;
      }
    }
    const first = definition.sections.filter((s) => s.tabId === activeTab)[0];
    return first ? first.id : undefined;
  };

  // ---------- Drag-and-drop ----------

  const onDragStart = (e: DragStartEvent): void => {
    const data = e.active.data.current || {};
    let label = '';
    if (data.type === 'section' && definition) {
      const s = definition.sections.filter((x) => x.id === data.id)[0];
      label = s ? s.title : t('Abschnitt');
    } else if (data.name) {
      const f = fieldsByName[data.name];
      const settings = definition ? definition.fields[data.name] : undefined;
      label = (settings && settings.label) || (f ? f.title : data.name);
    }
    setDrag({ id: String(e.active.id), type: data.type, label });
  };

  const onDragOver = (e: DragOverEvent): void => {
    setOverId(e.over ? String(e.over.id) : undefined);
  };

  const onDragEnd = (e: DragEndEvent): void => {
    setDrag(undefined);
    setOverId(undefined);
    if (!definition || !e.over) {
      return;
    }
    const a = e.active.data.current || {};
    const o = e.over.data.current || {};

    if (a.type === 'section') {
      if (o.type === 'section' && o.id !== a.id) {
        update(ds.moveSection(definition, a.id, o.id, o.tabId));
      } else if (o.type === 'tab') {
        update(ds.moveSection(definition, a.id, undefined, o.id));
        setActiveTab(o.id);
      }
      return;
    }

    const name: string = a.name;
    if (o.type === 'paletteZone' && a.type === 'field') {
      update(ds.removeField(definition, name), 'clear');
      return;
    }
    if (o.type === 'tab') {
      const r = ds.placeFieldOnTab(definition, name, o.id);
      update(r.def, { kind: 'field', name });
      setActiveTab(o.id);
      return;
    }
    if (o.type === 'field') {
      if (o.name === name) {
        return;
      }
      if (a.type === 'field' && a.sectionId === o.sectionId) {
        const sec = definition.sections.filter((s) => s.id === o.sectionId)[0];
        update(ds.reorderInSection(definition, sec.id, sec.fields.indexOf(name), sec.fields.indexOf(o.name)), { kind: 'field', name });
      } else {
        update(ds.placeField(definition, name, o.sectionId, o.name), { kind: 'field', name });
      }
      return;
    }
    if (o.type === 'sectionBody') {
      update(ds.placeField(definition, name, o.id), { kind: 'field', name });
    }
  };

  // ---------- Vorschau ----------

  const openPreview = (): void => {
    setPreviewKey((k) => k + 1);
    setView('preview');
  };

  const loadPreviewItem = async (): Promise<void> => {
    const id = Number(previewItemId);
    if (!id) {
      setPreviewValues({});
      setPreviewKey((k) => k + 1);
      return;
    }
    try {
      setPreviewValues(await sp.getItem(listId, id, fields));
      setPreviewKey((k) => k + 1);
    } catch (e) {
      fail(e);
    }
  };

  // ---------- Darstellung ----------

  const ct = contentTypes.filter((c) => c.id === ctId)[0];
  const list = lists.filter((l) => l.id === listId)[0];
  const placed = definition ? ds.placedFields(definition) : [];
  const unplaced = fields.filter((f) => f.kind !== 'unsupported' && placed.indexOf(f.internalName) < 0);

  const listOptions: IDropdownOption[] = lists.map((l) => ({ key: l.id, text: l.title }));
  const ctOptions: IDropdownOption[] = contentTypes.map((c) => ({ key: c.id, text: c.name }));

  return (
    <div className={styles.designer}>
      <div className={styles.toolbar}>
        <Dropdown label={t('Liste')} placeholder={t('Liste wählen')} options={listOptions} selectedKey={listId || null} onChange={(_, o) => o && setListId(String(o.key))} styles={{ root: { minWidth: 220 } }} />
        <Dropdown label={t('Inhaltstyp')} placeholder={t('Inhaltstyp wählen')} options={ctOptions} selectedKey={ctId || null} disabled={!contentTypes.length} onChange={(_, o) => o && setCtId(String(o.key))} styles={{ root: { minWidth: 180 } }} />
        <ChoiceGroup
          label={t('Formulartyp')}
          selectedKey={formMode}
          options={[
            { key: 'new', text: t('Neu') },
            { key: 'edit', text: t('Bearbeiten') },
            { key: 'display', text: t('Anzeigen') }
          ]}
          styles={{ flexContainer: { display: 'flex', gap: 12 } }}
          onChange={(_, o) => {
            if (o) {
              setFormMode(o.key as FormMode);
              setPreviewKey((k) => k + 1);
            }
          }}
        />
        <div className={styles.toolbarGrow} />
        <DefaultButton text={dirty ? t('Entwurf speichern') + ' *' : t('Entwurf speichern')} iconProps={{ iconName: 'Save' }} disabled={!definition} onClick={() => saveDraft().catch(fail)} />
        <PrimaryButton text={t('Bereitstellen')} iconProps={{ iconName: 'CloudUpload' }} disabled={!definition || !ct} onClick={() => setDeployOpen(true)} />
      </div>

      {status && (
        <MessageBar messageBarType={status.kind} onDismiss={() => setStatus(undefined)} dismissButtonAriaLabel={t('Schliessen')}>
          {status.text}
        </MessageBar>
      )}

      {definition && (
        <Pivot selectedKey={view} onLinkClick={(item) => item && (item.props.itemKey === 'preview' ? openPreview() : setView(item.props.itemKey as View))} styles={{ root: { padding: '0 16px', background: '#fff' } }}>
          <PivotItem itemKey="design" headerText={t('Entwurf')} itemIcon="Design" />
          <PivotItem itemKey="script" headerText={t('JavaScript')} itemIcon="Code" />
          <PivotItem itemKey="translations" headerText={t('Übersetzungen')} itemIcon="Translate" />
          <PivotItem itemKey="preview" headerText={t('Vorschau')} itemIcon="Preview" />
        </Pivot>
      )}

      <div className={styles.body}>
        {loading && <Spinner size={SpinnerSize.large} label={t('Formular wird geladen …')} />}
        {!loading && !definition && <span className={styles.hint}>{t('Liste und Inhaltstyp wählen, um dessen Formular zu gestalten.')}</span>}

        {!loading && definition && view === 'design' && (
          <DndContext sensors={sensors} collisionDetection={collision} onDragStart={onDragStart} onDragOver={onDragOver} onDragEnd={onDragEnd} onDragCancel={() => { setDrag(undefined); setOverId(undefined); }}>
            <Palette
              unplaced={unplaced}
              dragActiveType={drag ? drag.type : undefined}
              overId={overId}
              onAdd={(name) => {
                const target = targetSectionId();
                if (target) {
                  update(ds.placeField(definition, name, target), { kind: 'field', name });
                }
              }}
              onAddSection={() => {
                const r = ds.addSection(definition, activeTab);
                update(r.def, { kind: 'section', id: r.sectionId });
              }}
              onAddTab={() => {
                const r = ds.addTab(definition);
                update(r.def, { kind: 'section', id: r.sectionId });
                setActiveTab(r.tabId);
              }}
            />
            <Canvas
              definition={definition}
              fieldsByName={fieldsByName}
              formMode={formMode}
              activeTab={activeTab}
              selection={selection}
              overId={overId}
              onSelect={setSelection}
              onTabChange={setActiveTab}
            />
            <PropertiesPanel definition={definition} fieldsByName={fieldsByName} selection={selection} formMode={formMode} currentUser={currentUser} onChange={(d, s) => update(d, s)} onUploadImage={(file) => sp.uploadImage(file)} onOpenLinkBuilder={() => setLinkBuilderOpen(true)} />
            <DragOverlay>{drag ? <div className={styles.overlayCard}>{drag.label}</div> : null}</DragOverlay>
          </DndContext>
        )}

        {!loading && definition && view === 'script' && (
          <ScriptEditor
            script={definition.script}
            fields={fields.filter((f) => f.kind !== 'unsupported')}
            onChange={(script) => update({ ...definition, script })}
            onTest={openPreview}
          />
        )}

        {!loading && definition && view === 'translations' && (
          <TranslationsEditor definition={definition} fieldsByName={fieldsByName} onChange={(d) => update(d)} />
        )}

        {!loading && definition && view === 'preview' && (
          <div className={styles.preview} style={{ maxWidth: panelPreviewWidth(definition.layout) ? (panelPreviewWidth(definition.layout) as number) + 2 : undefined, margin: '0 auto' }}>
            <div style={{ display: 'flex', gap: 8, alignItems: 'flex-end', padding: '12px 24px 0', flexWrap: 'wrap' }}>
              <TextField label={t('Element-ID für die Vorschau (optional)')} type="number" value={previewItemId} onChange={(_, v) => setPreviewItemId(v || '')} styles={{ root: { width: 240 } }} />
              <DefaultButton text={t('Laden und neu starten')} onClick={() => loadPreviewItem().catch(fail)} />
              <TextField
                label={t('URL-Parameter für die Vorschau')}
                placeholder="fd_Title=Test&fd_Status=Offen"
                value={previewQuery}
                onChange={(_, v) => setPreviewQuery(v || '')}
                onBlur={() => setPreviewKey((k) => k + 1)}
                styles={{ root: { width: 320 } }}
              />
              <Dropdown
                label={t('Sprache der Vorschau')}
                selectedKey={previewLanguage}
                options={[
                  { key: 'de', text: t('Deutsch') },
                  { key: 'en', text: t('Englisch') }
                ]}
                styles={{ root: { width: 160 } }}
                onChange={(_, o) => o && setPreviewLanguage(o.key as ContentLanguage)}
              />
            </div>
            <FormHost
              key={`${previewKey}-${formMode}-${previewLanguage}`}
              contentLanguage={previewLanguage}
              urlParams={readUrlParams(`?${previewQuery.replace(/^\?/, '')}`)}
              definition={definition}
              fields={fields}
              mode={formMode}
              initialValues={formMode === 'new' ? {} : previewValues}
              currentUser={currentUser}
              sp={sp}
              preview
              onSubmit={async () => undefined}
              onCancel={() => setPreviewKey((k) => k + 1)}
              onRequestEdit={() => setFormMode('edit')}
            />
          </div>
        )}
      </div>

      {linkBuilderOpen && definition && ct && list && (
        <LinkBuilder
          webUrl={sp.webUrl}
          listId={listId}
          listUrl={list.url}
          contentTypeId={ct.id}
          definition={definition}
          fieldsByName={fieldsByName}
          onDismiss={() => setLinkBuilderOpen(false)}
        />
      )}

      {deployOpen && definition && ct && list && (
        <DeployDialog
          sp={sp}
          listId={listId}
          listTitle={list.title}
          contentType={ct}
          definition={definition}
          onDismiss={() => setDeployOpen(false)}
          onChanged={() => {
            setDirty(false);
            sp.getContentTypes(listId).then(setContentTypes).catch(fail);
          }}
        />
      )}
    </div>
  );
};

export default FormDesigner;
