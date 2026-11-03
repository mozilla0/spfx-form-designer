import * as React from 'react';
import { Checkbox, DefaultButton, Dialog, DialogFooter, DialogType, Icon, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';
import { PANEL_COMPONENT_ID, RUNTIME_COMPONENT_ID } from '../../../common/constants';
import { displayModeOf, IFormBackup, IFormDefinition, IRuntimeProperties } from '../../../common/model';
import { checkScriptSyntax } from '../../../common/scriptEngine';
import { ContentTypeFormUpdate, IContentTypeInfo, SpService } from '../../../common/spService';

export interface IDeployDialogProps {
  sp: SpService;
  listId: string;
  listTitle: string;
  contentType: IContentTypeInfo;
  definition: IFormDefinition;
  onDismiss: () => void;
  onChanged: () => void;
}

type StepState = 'pending' | 'running' | 'done' | 'error';

interface IStep {
  label: string;
  state: StepState;
}

function sameId(a: string | null | undefined, b: string): boolean {
  return !!a && a.replace(/[{}]/g, '').toLowerCase() === b.toLowerCase();
}

function isForeign(id: string | null | undefined): boolean {
  return !!id && !sameId(id, RUNTIME_COMPONENT_ID);
}

const STEP_ICONS: Record<StepState, string> = { pending: 'CircleRing', running: 'Sync', done: 'CompletedSolid', error: 'StatusErrorFull' };
const STEP_COLORS: Record<StepState, string> = { pending: '#8a8886', running: '#0f6cbd', done: '#107c10', error: '#a4262c' };

/** Stellt die Formulardefinition bereit und verknüpft die Runtime mit dem Inhaltstyp. */
export const DeployDialog: React.FC<IDeployDialogProps> = (p) => {
  const ct = p.contentType;
  const [forms, setForms] = React.useState({ new: true, edit: true, display: true });
  const [replaceForeign, setReplaceForeign] = React.useState(false);
  const [steps, setSteps] = React.useState<IStep[]>([]);
  const [phase, setPhase] = React.useState<'idle' | 'running' | 'done' | 'restored'>('idle');
  const [error, setError] = React.useState('');

  const foreign = isForeign(ct.NewFormClientSideComponentId) || isForeign(ct.EditFormClientSideComponentId) || isForeign(ct.DisplayFormClientSideComponentId);
  const ours = sameId(ct.NewFormClientSideComponentId, RUNTIME_COMPONENT_ID) || sameId(ct.EditFormClientSideComponentId, RUNTIME_COMPONENT_ID) || sameId(ct.DisplayFormClientSideComponentId, RUNTIME_COMPONENT_ID);
  const noneSelected = !forms.new && !forms.edit && !forms.display;
  const usesOverlay = (['new', 'edit', 'display'] as const).some((m) => displayModeOf(p.definition.layout, m) !== 'page');
  const canStart = phase === 'idle' && !noneSelected && (!foreign || replaceForeign);

  const run = async (): Promise<void> => {
    const labels = ['Formular-Skript prüfen', 'Bisherige Formularkonfiguration sichern', 'Formulardefinition veröffentlichen', 'Runtime mit dem Inhaltstyp verknüpfen', usesOverlay ? 'Panel-Erweiterung auf der Liste registrieren' : 'Panel-Erweiterung von der Liste entfernen'];
    let current: IStep[] = labels.map((label) => ({ label, state: 'pending' as StepState }));
    const mark = (i: number, state: StepState): void => {
      current = current.map((s, idx) => (idx === i ? { ...s, state } : s));
      setSteps(current);
    };
    setPhase('running');
    setError('');
    setSteps(current);
    let i = 0;
    try {
      mark(i, 'running');
      const syntax = checkScriptSyntax(p.definition.script);
      if (syntax) {
        throw new Error(`Das Formular-Skript enthält einen Syntaxfehler: ${syntax}`);
      }
      mark(i++, 'done');

      mark(i, 'running');
      if (!ours) {
        const backup: IFormBackup = {
          savedAt: new Date().toISOString(),
          NewFormClientSideComponentId: ct.NewFormClientSideComponentId,
          NewFormClientSideComponentProperties: ct.NewFormClientSideComponentProperties,
          EditFormClientSideComponentId: ct.EditFormClientSideComponentId,
          EditFormClientSideComponentProperties: ct.EditFormClientSideComponentProperties,
          DisplayFormClientSideComponentId: ct.DisplayFormClientSideComponentId,
          DisplayFormClientSideComponentProperties: ct.DisplayFormClientSideComponentProperties
        };
        await p.sp.saveBackup(p.listId, ct.id, backup);
      }
      mark(i++, 'done');

      mark(i, 'running');
      const url = await p.sp.saveDefinition(p.listId, ct.id, p.definition, 'live');
      await p.sp.saveDefinition(p.listId, ct.id, p.definition, 'draft');
      mark(i++, 'done');

      mark(i, 'running');
      const props: IRuntimeProperties = { definitionUrl: url };
      const json = JSON.stringify(props);
      const update: ContentTypeFormUpdate = {};
      if (forms.new) {
        update.NewFormClientSideComponentId = RUNTIME_COMPONENT_ID;
        update.NewFormClientSideComponentProperties = json;
      }
      if (forms.edit) {
        update.EditFormClientSideComponentId = RUNTIME_COMPONENT_ID;
        update.EditFormClientSideComponentProperties = json;
      }
      if (forms.display) {
        update.DisplayFormClientSideComponentId = RUNTIME_COMPONENT_ID;
        update.DisplayFormClientSideComponentProperties = json;
      }
      await p.sp.updateContentTypeForms(p.listId, ct.id, update);
      mark(i++, 'done');

      mark(i, 'running');
      if (usesOverlay) {
        await p.sp.ensureListCommandSet(p.listId, PANEL_COMPONENT_ID, 'Form Designer Panel');
      } else {
        await p.sp.removeListCommandSet(p.listId, PANEL_COMPONENT_ID);
      }
      mark(i++, 'done');
      setPhase('done');
      p.onChanged();
    } catch (e) {
      mark(i, 'error');
      setError((e as Error).message);
      setPhase('idle');
    }
  };

  const restore = async (): Promise<void> => {
    setError('');
    try {
      const backup = await p.sp.readBackup(p.listId, ct.id);
      const update: ContentTypeFormUpdate = backup
        ? {
            NewFormClientSideComponentId: backup.NewFormClientSideComponentId || '',
            NewFormClientSideComponentProperties: backup.NewFormClientSideComponentProperties || '',
            EditFormClientSideComponentId: backup.EditFormClientSideComponentId || '',
            EditFormClientSideComponentProperties: backup.EditFormClientSideComponentProperties || '',
            DisplayFormClientSideComponentId: backup.DisplayFormClientSideComponentId || '',
            DisplayFormClientSideComponentProperties: backup.DisplayFormClientSideComponentProperties || ''
          }
        : {
            NewFormClientSideComponentId: '',
            NewFormClientSideComponentProperties: '',
            EditFormClientSideComponentId: '',
            EditFormClientSideComponentProperties: '',
            DisplayFormClientSideComponentId: '',
            DisplayFormClientSideComponentProperties: ''
          };
      await p.sp.updateContentTypeForms(p.listId, ct.id, update);
      await p.sp.removeListCommandSet(p.listId, PANEL_COMPONENT_ID);
      setPhase('restored');
      setSteps([]);
      p.onChanged();
    } catch (e) {
      setError((e as Error).message);
    }
  };

  return (
    <Dialog
      hidden={false}
      onDismiss={phase === 'running' ? undefined : p.onDismiss}
      minWidth={520}
      dialogContentProps={{ type: DialogType.normal, title: 'Als List Form Customizer bereitstellen', subText: `Liste «${p.listTitle}» · Inhaltstyp «${ct.name}»` }}
      modalProps={{ isBlocking: phase === 'running' }}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Formulare ersetzen</Text>
        <Stack horizontal tokens={{ childrenGap: 20 }}>
          <Checkbox label="Neu" checked={forms.new} disabled={phase === 'running'} onChange={(_, c) => setForms({ ...forms, new: !!c })} />
          <Checkbox label="Bearbeiten" checked={forms.edit} disabled={phase === 'running'} onChange={(_, c) => setForms({ ...forms, edit: !!c })} />
          <Checkbox label="Anzeigen" checked={forms.display} disabled={phase === 'running'} onChange={(_, c) => setForms({ ...forms, display: !!c })} />
        </Stack>

        {foreign && phase === 'idle' && (
          <MessageBar messageBarType={MessageBarType.warning}>
            Am Inhaltstyp hängt bereits ein anderer Form Customizer. Er wird vor dem Ersetzen gesichert und lässt sich später wiederherstellen.
            <Checkbox label="Bestehenden Form Customizer ersetzen" checked={replaceForeign} onChange={(_, c) => setReplaceForeign(!!c)} styles={{ root: { marginTop: 8 } }} />
          </MessageBar>
        )}
        {phase === 'idle' && (
          <MessageBar messageBarType={MessageBarType.info}>
            Ist auf der Liste ein Power-Apps-Formular aktiv, in den Listeneinstellungen unter «Formulareinstellungen» auf «Standard-SharePoint-Formular» umstellen. Die Power App bleibt dabei erhalten.
          </MessageBar>
        )}

        {steps.length > 0 && (
          <Stack tokens={{ childrenGap: 6 }} role="list">
            {steps.map((s) => (
              <Stack key={s.label} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} role="listitem">
                {s.state === 'running' ? <Spinner size={SpinnerSize.xSmall} /> : <Icon iconName={STEP_ICONS[s.state]} styles={{ root: { color: STEP_COLORS[s.state] } }} />}
                <Text>{s.label}</Text>
              </Stack>
            ))}
          </Stack>
        )}

        {phase === 'done' && <MessageBar messageBarType={MessageBarType.success}>Bereitgestellt. Die Liste öffnet ab sofort die neuen Formulare.</MessageBar>}
        {phase === 'restored' && <MessageBar messageBarType={MessageBarType.success}>Die vorherige Formularkonfiguration ist wiederhergestellt.</MessageBar>}
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        {noneSelected && phase === 'idle' && <Text styles={{ root: { color: '#a4262c' } }}>Mindestens ein Formular auswählen.</Text>}
      </Stack>
      <DialogFooter>
        {(ours || phase === 'done') && phase !== 'running' && phase !== 'restored' && <DefaultButton text="Zurücksetzen" onClick={() => restore().catch(() => undefined)} />}
        <DefaultButton text={phase === 'done' || phase === 'restored' ? 'Schliessen' : 'Abbrechen'} disabled={phase === 'running'} onClick={p.onDismiss} />
        {phase === 'idle' && <PrimaryButton text="Jetzt bereitstellen" disabled={!canStart} onClick={() => run().catch(() => undefined)} />}
      </DialogFooter>
    </Dialog>
  );
};
