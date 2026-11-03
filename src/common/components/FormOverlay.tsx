import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { ContextualMenu, DefaultButton, IconButton, MessageBar, MessageBarType, Modal, Panel, PanelType, Spinner, SpinnerSize, Stack } from '@fluentui/react';
import { DisplayMode, FormMode, IFormDefinition } from '../model';
import { IFormRequest } from '../formUrl';
import { allowNextNavigation, fdLog } from '../interceptor';
import { SpService } from '../spService';
import { FieldValues, ICurrentUser, ISpField } from '../spTypes';
import { FormHost } from './FormHost';

/** Formularkonfiguration eines Inhaltstyps der Liste. */
export interface IListFormConfig {
  contentTypeId: string;
  definition: IFormDefinition;
}

export interface IFormOverlayProps {
  sp: SpService;
  listId: string;
  request: IFormRequest;
  configs: IListFormConfig[];
  currentUser: ICurrentUser;
  displayMode: Exclude<DisplayMode, 'page'>;
  cssWidth: string | undefined;
  onClose: (saved: boolean) => void;
}

interface ILoaded {
  config: IListFormConfig;
  fields: ISpField[];
  values: FieldValues;
}

function sameCt(a: string, b: string): boolean {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  return x === y || x.indexOf(y) === 0 || y.indexOf(x) === 0;
}

function fallbackToPage(url: string): void {
  allowNextNavigation(url);
  window.location.href = url;
}

const DRAG_OPTIONS = { moveMenuItemText: 'Verschieben', closeMenuItemText: 'Schliessen', menu: ContextualMenu, dragHandleSelector: '.fd-dialog-handle' };

/** Lädt und zeigt ein Formular im Panel oder Dialog über der Listenansicht. */
export const FormOverlay: React.FC<IFormOverlayProps> = (p) => {
  const [loaded, setLoaded] = React.useState<ILoaded | undefined>(undefined);
  const [error, setError] = React.useState<string>('');
  const [mode, setMode] = React.useState<FormMode>(p.request.mode);

  React.useEffect(() => {
    let alive = true;
    const load = async (): Promise<void> => {
      let config: IListFormConfig | undefined;
      const req = p.request;
      if (req.contentTypeId) {
        config = p.configs.filter((c) => sameCt(c.contentTypeId, req.contentTypeId as string))[0];
      } else if (p.configs.length === 1) {
        config = p.configs[0];
      } else if (req.itemId) {
        const item = await p.sp.get(`/_api/web/lists(guid'${p.listId}')/items(${req.itemId})?$select=ContentTypeId`);
        config = p.configs.filter((c) => sameCt(c.contentTypeId, String(item.ContentTypeId)))[0];
      } else {
        config = p.configs[0];
      }
      if (!config) {
        fdLog('Inhaltstyp ohne Form-Designer-Formular – Rückfall auf SharePoint', req);
        fallbackToPage(req.url);
        return;
      }
      const fields = await p.sp.getFields(p.listId, config.contentTypeId);
      const values = req.itemId && req.mode !== 'new' ? await p.sp.getItem(p.listId, req.itemId, fields) : {};
      if (alive) {
        setLoaded({ config, fields, values });
      }
    };
    load().catch((e: Error) => alive && setError(e.message));
    return () => {
      alive = false;
    };
  }, []);

  const body = (
    <div style={{ minHeight: 200 }}>
      {!loaded && !error && <Spinner size={SpinnerSize.large} label="Formular wird geladen …" styles={{ root: { padding: 40 } }} />}
      {error && (
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 20 } }}>
          <MessageBar messageBarType={MessageBarType.error}>Das Formular konnte nicht geladen werden. {error}</MessageBar>
          <DefaultButton text="Als ganze Seite öffnen" onClick={() => fallbackToPage(p.request.url)} styles={{ root: { alignSelf: 'flex-start' } }} />
        </Stack>
      )}
      {loaded && (
        <FormHost
          key={mode}
          definition={loaded.config.definition}
          fields={loaded.fields}
          mode={mode}
          initialValues={loaded.values}
          currentUser={p.currentUser}
          sp={p.sp}
          onSubmit={async (values) => {
            await p.sp.saveItem(p.listId, mode === 'new' ? undefined : p.request.itemId, loaded.fields, values);
            p.onClose(true);
          }}
          onCancel={() => p.onClose(false)}
          onRequestEdit={p.request.itemId ? () => setMode('edit') : undefined}
        />
      )}
    </div>
  );

  if (p.displayMode === 'dialog') {
    return (
      <Modal
        isOpen
        onDismiss={() => p.onClose(false)}
        dragOptions={DRAG_OPTIONS}
        isBlocking
        styles={{
          main: {
            width: `min(${p.cssWidth || '720px'}, 95vw)`,
            maxWidth: '95vw',
            maxHeight: '90vh',
            minWidth: 320,
            resize: 'both',
            overflow: 'auto'
          },
          scrollableContent: { overflow: 'visible', maxHeight: 'none' }
        }}
      >
        <div className="fd-dialog-handle" style={{ display: 'flex', justifyContent: 'flex-end', padding: '4px 4px 0', cursor: 'move' }}>
          <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Schliessen" title="Schliessen" onClick={() => p.onClose(false)} />
        </div>
        {body}
      </Modal>
    );
  }

  return (
    <Panel
      isOpen
      type={PanelType.custom}
      customWidth={p.cssWidth ? `min(${p.cssWidth}, 100vw)` : '640px'}
      isLightDismiss={false}
      closeButtonAriaLabel="Schliessen"
      onDismiss={() => p.onClose(false)}
      styles={{ content: { padding: 0 }, scrollableContent: { overflowY: 'auto' } }}
    >
      {body}
    </Panel>
  );
};

/** Öffnet ein Formular über der Seite. Nach dem Speichern wird die Listenansicht neu geladen. */
let overlayOpen = false;

/** true, solange ein Formular im Panel oder Dialog offen ist. */
export function isOverlayOpen(): boolean {
  return overlayOpen;
}

export function openFormOverlay(props: Omit<IFormOverlayProps, 'onClose'>): void {
  if (overlayOpen) {
    return;
  }
  overlayOpen = true;
  const host = document.createElement('div');
  host.setAttribute('data-form-designer-overlay', 'true');
  document.body.appendChild(host);
  const close = (saved: boolean): void => {
    overlayOpen = false;
    ReactDOM.unmountComponentAtNode(host);
    if (host.parentNode) {
      host.parentNode.removeChild(host);
    }
    if (saved) {
      window.location.reload();
    }
  };
  ReactDOM.render(<FormOverlay {...props} onClose={close} />, host);
}
