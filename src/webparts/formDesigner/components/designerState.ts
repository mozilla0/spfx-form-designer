import { createFieldSettings, IFieldSettings, IFormDefinition, ISection, newId } from '../../../common/model';

/** Reine Hilfsfunktionen, die eine Formulardefinition unverändert lassen und eine neue liefern. */

export type Selection = { kind: 'field'; name: string } | { kind: 'section'; id: string } | { kind: 'form' } | undefined;

/** 'keep' lässt die Auswahl unverändert, 'clear' hebt sie auf. */
export type SelectionChange = Selection | 'keep' | 'clear';

function removeFromSections(sections: ISection[], name: string): ISection[] {
  return sections.map((s) => (s.fields.indexOf(name) >= 0 ? { ...s, fields: s.fields.filter((f) => f !== name) } : s));
}

export function placedFields(def: IFormDefinition): string[] {
  const out: string[] = [];
  def.sections.forEach((s) => s.fields.forEach((f) => out.push(f)));
  return out;
}

export function sectionOfField(def: IFormDefinition, name: string): ISection | undefined {
  return def.sections.filter((s) => s.fields.indexOf(name) >= 0)[0];
}

/** Fügt ein Feld in einen Abschnitt ein – vor «beforeName» oder am Ende. */
export function placeField(def: IFormDefinition, name: string, sectionId: string, beforeName?: string): IFormDefinition {
  const fields = def.fields[name] ? def.fields : { ...def.fields, [name]: createFieldSettings() };
  const sections = removeFromSections(def.sections, name).map((s) => {
    if (s.id !== sectionId) {
      return s;
    }
    const list = s.fields.slice();
    const idx = beforeName ? list.indexOf(beforeName) : -1;
    if (idx < 0) {
      list.push(name);
    } else {
      list.splice(idx, 0, name);
    }
    return { ...s, fields: list };
  });
  return { ...def, fields, sections };
}

export function reorderInSection(def: IFormDefinition, sectionId: string, from: number, to: number): IFormDefinition {
  return {
    ...def,
    sections: def.sections.map((s) => {
      if (s.id !== sectionId) {
        return s;
      }
      const list = s.fields.slice();
      const [moved] = list.splice(from, 1);
      list.splice(to, 0, moved);
      return { ...s, fields: list };
    })
  };
}

export function moveFieldBy(def: IFormDefinition, name: string, delta: number): IFormDefinition {
  const sec = sectionOfField(def, name);
  if (!sec) {
    return def;
  }
  const from = sec.fields.indexOf(name);
  const to = from + delta;
  if (to < 0 || to >= sec.fields.length) {
    return def;
  }
  return reorderInSection(def, sec.id, from, to);
}

export function removeField(def: IFormDefinition, name: string): IFormDefinition {
  return { ...def, sections: removeFromSections(def.sections, name) };
}

/** Legt das Feld im ersten Abschnitt des Tabs ab; fehlt einer, wird er angelegt. */
export function placeFieldOnTab(def: IFormDefinition, name: string, tabId: string): { def: IFormDefinition; sectionId: string } {
  let target = def.sections.filter((s) => s.tabId === tabId)[0];
  let next = def;
  if (!target) {
    target = { id: newId('sec'), tabId, title: 'Neuer Abschnitt', columns: 2, fields: [] };
    next = { ...def, sections: def.sections.concat([target]) };
  }
  return { def: placeField(next, name, target.id), sectionId: target.id };
}

export function addSection(def: IFormDefinition, tabId: string): { def: IFormDefinition; sectionId: string } {
  const section: ISection = { id: newId('sec'), tabId, title: 'Neuer Abschnitt', columns: 2, fields: [] };
  return { def: { ...def, sections: def.sections.concat([section]) }, sectionId: section.id };
}

export function addTab(def: IFormDefinition): { def: IFormDefinition; tabId: string; sectionId: string } {
  const tabId = newId('tab');
  const section: ISection = { id: newId('sec'), tabId, title: 'Neuer Abschnitt', columns: 2, fields: [] };
  return {
    def: { ...def, tabs: def.tabs.concat([{ id: tabId, label: 'Neuer Tab' }]), sections: def.sections.concat([section]) },
    tabId,
    sectionId: section.id
  };
}

export function updateSection(def: IFormDefinition, id: string, patch: Partial<ISection>): IFormDefinition {
  return { ...def, sections: def.sections.map((s) => (s.id === id ? { ...s, ...patch } : s)) };
}

export function deleteSection(def: IFormDefinition, id: string): IFormDefinition {
  return { ...def, sections: def.sections.filter((s) => s.id !== id) };
}

/** Verschiebt einen Abschnitt vor einen anderen oder ans Ende eines Tabs. */
export function moveSection(def: IFormDefinition, id: string, beforeId: string | undefined, tabId: string): IFormDefinition {
  const moving = def.sections.filter((s) => s.id === id)[0];
  if (!moving) {
    return def;
  }
  const rest = def.sections.filter((s) => s.id !== id);
  const moved = { ...moving, tabId };
  const idx = beforeId ? rest.map((s) => s.id).indexOf(beforeId) : -1;
  if (idx < 0) {
    rest.push(moved);
  } else {
    rest.splice(idx, 0, moved);
  }
  return { ...def, sections: rest };
}

export function renameTab(def: IFormDefinition, tabId: string, label: string): IFormDefinition {
  return { ...def, tabs: def.tabs.map((t) => (t.id === tabId ? { ...t, label } : t)) };
}

export function deleteTab(def: IFormDefinition, tabId: string): IFormDefinition {
  return { ...def, tabs: def.tabs.filter((t) => t.id !== tabId), sections: def.sections.filter((s) => s.tabId !== tabId) };
}

export function updateFieldSettings(def: IFormDefinition, name: string, patch: Partial<IFieldSettings>): IFormDefinition {
  const current = def.fields[name] || createFieldSettings();
  return { ...def, fields: { ...def.fields, [name]: { ...current, ...patch } } };
}
