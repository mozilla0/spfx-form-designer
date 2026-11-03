import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { DEFINITION_FOLDER } from './constants';
import { IFormBackup, IFormDefinition } from './model';
import {
  FieldKind,
  FieldValue,
  FieldValues,
  ILookupValue,
  ISpField,
  ITermValue,
  IUrlValue,
  IUserValue
} from './spTypes';
import { toLocalDate, pad } from './dateUtils';
import { getCulture } from './i18n';

/** Begriff aus dem Ausdrucksspeicher mit Pfad zur Unterscheidung gleichnamiger Begriffe. */
export interface ITermOption {
  termGuid: string;
  label: string;
  path: string;
}

const EMPTY_GUID = '00000000-0000-0000-0000-000000000000';

/** Spaltentypen, die per ValidateUpdateListItem geschrieben werden. */
const VALIDATE_KINDS: FieldKind[] = ['taxonomy', 'taxonomymulti', 'url'];

/* eslint-disable @typescript-eslint/no-explicit-any */

const JSON_HEADERS: Record<string, string> = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'odata-version': ''
};

export interface IListInfo {
  id: string;
  title: string;
}

export interface IContentTypeInfo {
  id: string;
  name: string;
  NewFormClientSideComponentId: string | null;
  NewFormClientSideComponentProperties: string | null;
  EditFormClientSideComponentId: string | null;
  EditFormClientSideComponentProperties: string | null;
  DisplayFormClientSideComponentId: string | null;
  DisplayFormClientSideComponentProperties: string | null;
}

export type ContentTypeFormUpdate = Partial<Omit<IContentTypeInfo, 'id' | 'name'>>;

/** Systemspalten, die nie im Formular erscheinen. */
const EXCLUDED_FIELDS: string[] = ['ContentType', 'Attachments', 'FileLeafRef', '_ComplianceTag', '_ComplianceTagWrittenTime', '_ComplianceTagUserId', '_IsRecord'];

function odataEscape(value: string): string {
  return value.replace(/'/g, "''");
}

function mapKind(f: any): FieldKind {
  switch (f.TypeAsString) {
    case 'Text':
      return 'text';
    case 'Note':
      return 'note';
    case 'Number':
    case 'Currency':
      return 'number';
    case 'Choice':
      return 'choice';
    case 'MultiChoice':
      return 'multichoice';
    case 'Boolean':
      return 'boolean';
    case 'DateTime':
      return f.DisplayFormat === 1 ? 'datetime' : 'date';
    case 'User':
    case 'UserMulti':
      return f.AllowMultipleValues || f.TypeAsString === 'UserMulti' ? 'usermulti' : 'user';
    case 'Lookup':
    case 'LookupMulti':
      return f.AllowMultipleValues || f.TypeAsString === 'LookupMulti' ? 'lookupmulti' : 'lookup';
    case 'TaxonomyFieldType':
      return 'taxonomy';
    case 'TaxonomyFieldTypeMulti':
      return 'taxonomymulti';
    case 'URL':
      return 'url';
    case 'Calculated':
      return 'calculated';
    default:
      return 'unsupported';
  }
}

/** Zugriff auf SharePoint über SPHttpClient (REST). */
export class SpService {
  constructor(private readonly http: SPHttpClient, public readonly webUrl: string, public readonly webServerRelativeUrl: string) {}

  public async get<T = any>(path: string): Promise<T> {
    const res = await this.http.get(this.url(path), SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
    return this.json<T>(res);
  }

  public async post<T = any>(path: string, body?: unknown, extraHeaders?: Record<string, string>): Promise<T> {
    const res = await this.http.post(this.url(path), SPHttpClient.configurations.v1, {
      headers: { ...JSON_HEADERS, ...(extraHeaders || {}) },
      body: body === undefined ? undefined : typeof body === 'string' ? body : JSON.stringify(body)
    });
    return this.json<T>(res);
  }

  // ---------- Listen, Inhaltstypen, Spalten ----------

  public async getLists(): Promise<IListInfo[]> {
    const r = await this.get('/_api/web/lists?$select=Id,Title,BaseTemplate,Hidden&$filter=Hidden eq false and BaseType eq 0&$orderby=Title');
    return (r.value as any[]).map((l) => ({ id: l.Id, title: l.Title }));
  }

  public async getContentTypes(listId: string): Promise<IContentTypeInfo[]> {
    const r = await this.get(
      `/_api/web/lists(guid'${listId}')/contenttypes?$select=Id,Name,Hidden,NewFormClientSideComponentId,NewFormClientSideComponentProperties,EditFormClientSideComponentId,EditFormClientSideComponentProperties,DisplayFormClientSideComponentId,DisplayFormClientSideComponentProperties`
    );
    return (r.value as any[])
      .filter((c) => !c.Hidden && String(c.Id.StringValue).indexOf('0x0120') !== 0)
      .map((c) => ({
        id: c.Id.StringValue,
        name: c.Name,
        NewFormClientSideComponentId: c.NewFormClientSideComponentId || null,
        NewFormClientSideComponentProperties: c.NewFormClientSideComponentProperties || null,
        EditFormClientSideComponentId: c.EditFormClientSideComponentId || null,
        EditFormClientSideComponentProperties: c.EditFormClientSideComponentProperties || null,
        DisplayFormClientSideComponentId: c.DisplayFormClientSideComponentId || null,
        DisplayFormClientSideComponentProperties: c.DisplayFormClientSideComponentProperties || null
      }));
  }

  public async getFields(listId: string, contentTypeId: string): Promise<ISpField[]> {
    const r = await this.get(
      `/_api/web/lists(guid'${listId}')/contenttypes('${contentTypeId}')/fields?$filter=Hidden eq false`
    );
    return (r.value as any[])
      .filter((f) => EXCLUDED_FIELDS.indexOf(f.InternalName) < 0)
      .filter((f) => !f.ReadOnlyField || f.InternalName === 'Title' || f.TypeAsString === 'Calculated')
      .map((f) => ({
        id: f.Id,
        internalName: f.InternalName,
        title: f.Title,
        typeAsString: f.TypeAsString,
        kind: mapKind(f),
        required: !!f.Required,
        readOnly: !!f.ReadOnlyField,
        choices: (f.Choices as string[]) || [],
        lookupListId: f.LookupList ? String(f.LookupList).replace(/[{}]/g, '') : undefined,
        lookupField: f.LookupField || 'Title',
        termSetId: f.TermSetId ? String(f.TermSetId) : undefined,
        anchorId: f.AnchorId && String(f.AnchorId) !== EMPTY_GUID ? String(f.AnchorId) : undefined,
        urlFormat: f.TypeAsString === 'URL' ? (f.DisplayFormat === 1 ? 'image' : 'hyperlink') : undefined,
        description: f.Description || ''
      }));
  }

  public async updateContentTypeForms(listId: string, contentTypeId: string, update: ContentTypeFormUpdate): Promise<void> {
    await this.post(`/_api/web/lists(guid'${listId}')/contenttypes('${contentTypeId}')`, update, {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    });
  }

  // ---------- Panel-Erweiterung (Benutzerdefinierte Aktion der Liste) ----------

  public async getListCommandSetActions(listId: string, componentId: string): Promise<string[]> {
    const r = await this.get(`/_api/web/lists(guid'${listId}')/UserCustomActions?$select=Id,ClientSideComponentId`);
    return (r.value as any[])
      .filter((a) => String(a.ClientSideComponentId || '').replace(/[{}]/g, '').toLowerCase() === componentId.toLowerCase())
      .map((a) => String(a.Id));
  }

  /** Registriert das Command Set auf der Liste, falls noch nicht vorhanden. */
  public async ensureListCommandSet(listId: string, componentId: string, title: string): Promise<boolean> {
    const existing = await this.getListCommandSetActions(listId, componentId);
    if (existing.length) {
      return false;
    }
    await this.post(`/_api/web/lists(guid'${listId}')/UserCustomActions`, {
      Title: title,
      Location: 'ClientSideExtension.ListViewCommandSet.CommandBar',
      ClientSideComponentId: componentId,
      ClientSideComponentProperties: '{}'
    });
    return true;
  }

  public async removeListCommandSet(listId: string, componentId: string): Promise<number> {
    const ids = await this.getListCommandSetActions(listId, componentId);
    for (const id of ids) {
      await this.post(`/_api/web/lists(guid'${listId}')/UserCustomActions('${id}')/deleteObject()`);
    }
    return ids.length;
  }

  // ---------- Elemente ----------

  public async getItem(listId: string, itemId: number, fields: ISpField[]): Promise<FieldValues> {
    const select: string[] = ['Id'];
    const expand: string[] = [];
    let needsText = false;
    fields.forEach((f) => {
      if (f.kind === 'user' || f.kind === 'usermulti') {
        select.push(`${f.internalName}/Id`, `${f.internalName}/Title`, `${f.internalName}/EMail`);
        expand.push(f.internalName);
      } else if (f.kind === 'lookup' || f.kind === 'lookupmulti') {
        select.push(`${f.internalName}/Id`, `${f.internalName}/${f.lookupField || 'Title'}`);
        expand.push(f.internalName);
      } else if (f.kind !== 'unsupported') {
        select.push(f.internalName);
      }
      if (f.kind === 'taxonomy' || f.kind === 'taxonomymulti' || f.kind === 'calculated') {
        needsText = true;
      }
    });
    const query = `$select=${select.join(',')}${expand.length ? `&$expand=${expand.join(',')}` : ''}`;
    const item = await this.get(`/_api/web/lists(guid'${listId}')/items(${itemId})?${query}`);
    const texts: Record<string, string> = needsText ? await this.get(`/_api/web/lists(guid'${listId}')/items(${itemId})/FieldValuesAsText`) : {};
    const values: FieldValues = {};
    fields.forEach((f) => {
      const text = texts[f.internalName] !== undefined ? texts[f.internalName] : texts[f.internalName.replace(/_/g, '_x005f_')];
      values[f.internalName] = this.fromSp(f, item[f.internalName], text);
    });
    return values;
  }

  public async saveItem(listId: string, itemId: number | undefined, fields: ISpField[], values: FieldValues, onlyFields?: string[]): Promise<number> {
    const payload: Record<string, unknown> = {};
    const formValues: Array<{ FieldName: string; FieldValue: string }> = [];
    fields.forEach((f) => {
      if (f.readOnly || f.kind === 'unsupported' || f.kind === 'calculated') {
        return;
      }
      if (onlyFields && onlyFields.indexOf(f.internalName) < 0) {
        return;
      }
      const v = values[f.internalName];
      if (VALIDATE_KINDS.indexOf(f.kind) >= 0) {
        formValues.push({ FieldName: f.internalName, FieldValue: this.toFormValue(f, v) });
      } else if (f.kind === 'user' || f.kind === 'lookup') {
        payload[`${f.internalName}Id`] = v && typeof v === 'object' && !Array.isArray(v) && 'id' in v ? v.id : null;
      } else if (f.kind === 'usermulti' || f.kind === 'lookupmulti') {
        payload[`${f.internalName}Id`] = Array.isArray(v) ? (v as Array<IUserValue | ILookupValue>).map((x) => x.id) : [];
      } else {
        payload[f.internalName] = this.toSp(f, v);
      }
    });
    let id = itemId;
    if (id) {
      if (Object.keys(payload).length) {
        await this.post(`/_api/web/lists(guid'${listId}')/items(${id})`, payload, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
      }
    } else {
      const created = await this.post(`/_api/web/lists(guid'${listId}')/items`, payload);
      id = created.Id as number;
    }
    if (formValues.length) {
      const r = await this.post(`/_api/web/lists(guid'${listId}')/items(${id})/ValidateUpdateListItem`, { formValues, bNewDocumentUpdate: false });
      const failed = ((r && r.value) || []).filter((x: any) => x.HasException);
      if (failed.length) {
        throw new Error(failed.map((x: any) => `${x.FieldName}: ${x.ErrorMessage}`).join('; '));
      }
    }
    return id as number;
  }

  // ---------- Verwaltete Metadaten ----------

  private termCache: Record<string, ITermOption[]> = {};

  /** Lädt die Begriffe eines Ausdruckssatzes (optional unterhalb eines Ankerbegriffs). */
  public async getTerms(field: ISpField): Promise<ITermOption[]> {
    if (!field.termSetId) {
      return [];
    }
    const cacheKey = `${field.termSetId}|${field.anchorId || ''}`;
    if (this.termCache[cacheKey]) {
      return this.termCache[cacheKey];
    }
    const lang = getCulture().toLowerCase();
    const labelOf = (term: any): string => {
      const labels: any[] = term.labels || [];
      const exact = labels.filter((l) => String(l.languageTag || '').toLowerCase() === lang)[0];
      const sameLang = labels.filter((l) => String(l.languageTag || '').toLowerCase().split('-')[0] === lang.split('-')[0])[0];
      const def = labels.filter((l) => l.isDefault)[0];
      return ((exact || sameLang || def || labels[0] || {}).name as string) || '';
    };
    const base = `/_api/v2.1/termStore/sets/${field.termSetId}`;
    const fetchAll = async (path: string): Promise<any[]> => {
      const out: any[] = [];
      let next: string | undefined = `${path}${path.indexOf('?') >= 0 ? '&' : '?'}$top=500`;
      while (next) {
        const res = await this.http.get(next.indexOf('http') === 0 ? next : this.url(next), SPHttpClient.configurations.v1, { headers: { Accept: 'application/json' } });
        if (!res.ok) {
          throw new Error(await this.errorText(res));
        }
        const body = await res.json();
        (body.value || []).forEach((x: any) => out.push(x));
        next = body['@odata.nextLink'];
      }
      return out;
    };
    const result: ITermOption[] = [];
    const walk = async (path: string, parentPath: string, depth: number): Promise<void> => {
      const terms = await fetchAll(path);
      for (const term of terms) {
        const label = labelOf(term);
        const fullPath = parentPath ? `${parentPath} › ${label}` : label;
        result.push({ termGuid: String(term.id), label, path: fullPath });
        if (depth < 6 && term.childrenCount !== 0) {
          await walk(`${base}/terms/${term.id}/children`, fullPath, depth + 1);
        }
      }
    };
    await walk(field.anchorId ? `${base}/terms/${field.anchorId}/children` : `${base}/children`, '', 0);
    this.termCache[cacheKey] = result;
    return result;
  }

  // ---------- Personen und Nachschlagewerte ----------

  public async searchUsers(query: string): Promise<IUserValue[]> {
    const r = await this.post('/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser', {
      queryParams: {
        QueryString: query,
        MaximumEntitySuggestions: 10,
        AllowEmailAddresses: false,
        AllowMultipleEntities: false,
        PrincipalSource: 15,
        PrincipalType: 1
      }
    });
    const raw = typeof r.value === 'string' ? JSON.parse(r.value) : [];
    return (raw as any[]).map((e) => ({
      id: 0,
      title: e.DisplayText,
      email: (e.EntityData && e.EntityData.Email) || '',
      loginName: e.Key
    })) as IUserValue[];
  }

  /** Stellt sicher, dass der Benutzer in der Website existiert, und liefert seine ID. */
  public async ensureUser(loginName: string): Promise<IUserValue> {
    const r = await this.post('/_api/web/ensureuser', { logonName: loginName });
    return { id: r.Id, title: r.Title, email: r.Email };
  }

  public async getLookupItems(field: ISpField): Promise<ILookupValue[]> {
    if (!field.lookupListId) {
      return [];
    }
    const col = field.lookupField || 'Title';
    const r = await this.get(`/_api/web/lists(guid'${field.lookupListId}')/items?$select=Id,${col}&$orderby=${col}&$top=500`);
    return (r.value as any[]).map((i) => ({ id: i.Id, title: String(i[col] ?? '') }));
  }

  // ---------- Definitionen und Sicherungen ----------

  public definitionFileName(listId: string, contentTypeId: string, kind: 'live' | 'draft' | 'backup' = 'live'): string {
    const suffix = kind === 'live' ? '' : `.${kind}`;
    return `${listId}_${contentTypeId}${suffix}.json`;
  }

  public async readJson<T>(serverRelativeUrl: string): Promise<T | undefined> {
    const res = await this.http.get(
      this.url(`/_api/web/GetFileByServerRelativePath(decodedurl='${odataEscape(serverRelativeUrl)}')/$value`),
      SPHttpClient.configurations.v1
    );
    if (res.status === 404) {
      return undefined;
    }
    if (!res.ok) {
      throw new Error(await this.errorText(res));
    }
    return (await res.json()) as T;
  }

  public async writeJson(fileName: string, content: unknown): Promise<string> {
    const folder = await this.ensureDefinitionFolder();
    const res = await this.http.post(
      this.url(`/_api/web/GetFolderByServerRelativePath(decodedurl='${odataEscape(folder)}')/Files/AddUsingPath(decodedurl='${odataEscape(fileName)}',overwrite=true)`),
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' }, body: JSON.stringify(content, undefined, 2) }
    );
    if (!res.ok) {
      throw new Error(await this.errorText(res));
    }
    return `${folder}/${fileName}`;
  }

  /** Lädt ein Bild in «Websiteobjekte/FormDesigner/Bilder» und liefert die serverrelative URL. */
  public async uploadImage(file: File): Promise<string> {
    const folder = `${await this.ensureDefinitionFolder()}/Bilder`;
    await this.post(`/_api/web/folders/AddUsingPath(decodedurl='${odataEscape(folder)}',overwrite=false)`).catch(() => undefined);
    const safeName = file.name.replace(/[^A-Za-z0-9._-]+/g, '_');
    const name = `${Date.now()}_${safeName}`;
    const res = await this.http.post(
      this.url(`/_api/web/GetFolderByServerRelativePath(decodedurl='${odataEscape(folder)}')/Files/AddUsingPath(decodedurl='${odataEscape(name)}',overwrite=true)`),
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' }, body: await file.arrayBuffer() }
    );
    if (!res.ok) {
      throw new Error(await this.errorText(res));
    }
    return `${folder}/${name}`;
  }

  /** Speichert die Definition als Entwurf (nur Designer) oder live (wird von der Runtime gelesen). */
  public async saveDefinition(listId: string, contentTypeId: string, def: IFormDefinition, kind: 'live' | 'draft'): Promise<string> {
    return this.writeJson(this.definitionFileName(listId, contentTypeId, kind), def);
  }

  public async readDefinition(listId: string, contentTypeId: string, kind: 'live' | 'draft'): Promise<IFormDefinition | undefined> {
    const folder = await this.ensureDefinitionFolder();
    return this.readJson<IFormDefinition>(`${folder}/${this.definitionFileName(listId, contentTypeId, kind)}`);
  }

  public async saveBackup(listId: string, contentTypeId: string, backup: IFormBackup): Promise<string> {
    return this.writeJson(this.definitionFileName(listId, contentTypeId, 'backup'), backup);
  }

  public async readBackup(listId: string, contentTypeId: string): Promise<IFormBackup | undefined> {
    const folder = await this.ensureDefinitionFolder();
    return this.readJson<IFormBackup>(`${folder}/${this.definitionFileName(listId, contentTypeId, 'backup')}`);
  }

  private _folder: string | undefined;

  private async ensureDefinitionFolder(): Promise<string> {
    if (this._folder) {
      return this._folder;
    }
    const lib = await this.post('/_api/web/lists/EnsureSiteAssetsLibrary');
    const rootFolder = await this.get(`/_api/web/lists(guid'${lib.Id}')/RootFolder?$select=ServerRelativeUrl`);
    const folder = `${rootFolder.ServerRelativeUrl}/${DEFINITION_FOLDER}`;
    await this.post(`/_api/web/folders/AddUsingPath(decodedurl='${odataEscape(folder)}',overwrite=false)`).catch(() => undefined);
    this._folder = folder;
    return folder;
  }

  // ---------- Hilfsfunktionen ----------

  private url(path: string): string {
    return `${this.webUrl.replace(/\/$/, '')}${path}`;
  }

  private async json<T>(res: SPHttpClientResponse): Promise<T> {
    if (!res.ok) {
      throw new Error(await this.errorText(res));
    }
    if (res.status === 204) {
      return undefined as unknown as T;
    }
    const text = await res.text();
    return (text ? JSON.parse(text) : undefined) as T;
  }

  private async errorText(res: SPHttpClientResponse): Promise<string> {
    try {
      const body = await res.json();
      const msg = body?.['odata.error']?.message?.value || body?.error?.message?.value || body?.error?.message;
      return msg ? `${msg} (HTTP ${res.status})` : `HTTP ${res.status}`;
    } catch {
      return `HTTP ${res.status}`;
    }
  }

  private fromSp(f: ISpField, raw: any, text?: string): FieldValue {
    if (f.kind === 'calculated') {
      return text !== undefined ? text : raw === undefined || raw === null ? '' : String(raw);
    }
    if (raw === undefined || raw === null) {
      switch (f.kind) {
        case 'multichoice':
        case 'usermulti':
        case 'lookupmulti':
        case 'taxonomymulti':
          return [];
        case 'boolean':
          return false;
        case 'user':
        case 'lookup':
        case 'taxonomy':
          return null;
        case 'url':
          return { url: '', description: '' };
        default:
          return '';
      }
    }
    const list = (x: any): any[] => (Array.isArray(x) ? x : Array.isArray(x && x.results) ? x.results : []);
    switch (f.kind) {
      case 'number':
        return String(raw);
      case 'boolean':
        return !!raw;
      case 'multichoice':
        return Array.isArray(raw) ? raw : Array.isArray(raw.results) ? raw.results : [];
      case 'date': {
        const d = new Date(raw);
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
      }
      case 'datetime': {
        const d = new Date(raw);
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
      }
      case 'user':
        return raw.Id ? { id: raw.Id, title: raw.Title, email: raw.EMail } : null;
      case 'usermulti':
        return list(raw).filter((u) => u && u.Id).map((u) => ({ id: u.Id, title: u.Title, email: u.EMail }));
      case 'lookup':
        return raw.Id ? { id: raw.Id, title: String(raw[f.lookupField || 'Title'] ?? '') } : null;
      case 'lookupmulti':
        return list(raw).filter((l) => l && l.Id).map((l) => ({ id: l.Id, title: String(l[f.lookupField || 'Title'] ?? '') }));
      case 'taxonomy': {
        if (!raw.TermGuid) {
          return null;
        }
        return { termGuid: String(raw.TermGuid), label: text !== undefined && text !== '' ? text : String(raw.Label || '') } as ITermValue;
      }
      case 'taxonomymulti': {
        const labels = text ? text.split(';').map((x) => x.trim()) : [];
        return list(raw)
          .filter((x) => x && x.TermGuid)
          .map((x, i) => ({ termGuid: String(x.TermGuid), label: labels[i] || String(x.Label || '') })) as ITermValue[];
      }
      case 'url':
        return { url: String(raw.Url || ''), description: String(raw.Description || '') } as IUrlValue;
      default:
        return typeof raw === 'string' ? raw : String(raw);
    }
  }

  /** Wert im Format von ValidateUpdateListItem. */
  private toFormValue(f: ISpField, v: FieldValue): string {
    if (f.kind === 'taxonomy') {
      const term = v && typeof v === 'object' && !Array.isArray(v) && 'termGuid' in v ? (v as ITermValue) : undefined;
      return term ? `${term.label}|${term.termGuid}` : '';
    }
    if (f.kind === 'taxonomymulti') {
      return Array.isArray(v) ? (v as ITermValue[]).map((x) => `${x.label}|${x.termGuid}`).join(';') : '';
    }
    if (f.kind === 'url') {
      const u = v && typeof v === 'object' && !Array.isArray(v) && 'url' in v ? (v as IUrlValue) : undefined;
      if (!u || !u.url) {
        return '';
      }
      return `${u.url.replace(/,/g, ',,')}, ${u.description || u.url}`;
    }
    return '';
  }

  private toSp(f: ISpField, v: FieldValue): unknown {
    switch (f.kind) {
      case 'number': {
        if (v === '' || v === undefined || v === null) {
          return null;
        }
        const n = Number(v);
        return isNaN(n) ? null : n;
      }
      case 'boolean':
        return !!v;
      case 'multichoice':
        return Array.isArray(v) ? v : [];
      case 'date': {
        if (!v || typeof v !== 'string') {
          return null;
        }
        // Mittag UTC, damit reine Datumswerte unabhängig von der Zeitzone auf dem richtigen Tag landen.
        const d = toLocalDate(v);
        return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0)).toISOString();
      }
      case 'datetime': {
        if (!v || typeof v !== 'string') {
          return null;
        }
        const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/.exec(v);
        if (!m) {
          return null;
        }
        return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0)).toISOString();
      }
      default:
        return v === undefined || v === null ? null : v;
    }
  }
}
