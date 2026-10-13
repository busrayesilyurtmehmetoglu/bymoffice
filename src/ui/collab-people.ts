// Who is in the file: a button in the bar of the room that says how many, and opens the list.
// Names are text of other people: they are only ever set as textContent.

import { initials, type People } from '../collab/people';
import type { UiStrings } from './strings';

export class PeopleButton {
  readonly element: HTMLButtonElement;
  private readonly count: HTMLSpanElement;
  private panel: HTMLDivElement | undefined;
  private people: People = { list: [], more: 0, count: 0 };
  private words: UiStrings['collab'];
  private sheetName: (index: number) => string | undefined = () => undefined;
  private readonly onDocumentDown: (event: Event) => void;
  private readonly onKey: (event: KeyboardEvent) => void;

  constructor(
    private readonly host: HTMLElement,
    words: UiStrings['collab'],
  ) {
    this.words = words;
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'bym-collab-count';
    button.setAttribute('aria-haspopup', 'true');
    button.setAttribute('aria-expanded', 'false');
    button.style.cssText = 'display:inline-flex;align-items:center;gap:4px;flex:0 0 auto;margin:0;padding:1px 8px;font:inherit;color:inherit;background:#fff;border:1px solid #dadce0;border-radius:10px;cursor:pointer';
    // A head and shoulders, drawn: no text of anybody is in it.
    const icon = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    icon.setAttribute('viewBox', '0 0 16 16');
    icon.setAttribute('width', '12');
    icon.setAttribute('height', '12');
    icon.setAttribute('aria-hidden', 'true');
    const head = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    head.setAttribute('cx', '8');
    head.setAttribute('cy', '5');
    head.setAttribute('r', '3');
    const body = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    body.setAttribute('d', 'M2 15c0-3.3 2.7-6 6-6s6 2.7 6 6z');
    for (const shape of [head, body]) {
      shape.setAttribute('fill', 'currentColor');
      icon.appendChild(shape);
    }
    this.count = document.createElement('span');
    this.count.className = 'bym-collab-count-number';
    button.appendChild(icon);
    button.appendChild(this.count);
    button.addEventListener('click', () => (this.panel ? this.close(false) : this.open()));
    this.element = button;
    this.onDocumentDown = (event) => {
      const target = event.target as Node | null;
      if (this.panel && target && !this.panel.contains(target) && !this.element.contains(target)) this.close(false);
    };
    this.onKey = (event) => {
      if (event.key !== 'Escape' || !this.panel) return;
      event.preventDefault();
      event.stopPropagation();
      this.close(true);
    };
    button.addEventListener('keydown', this.onKey);
  }

  get isOpen(): boolean {
    return this.panel !== undefined;
  }

  update(people: People, words: UiStrings['collab'], sheetName: (index: number) => string | undefined): void {
    this.people = people;
    this.words = words;
    this.sheetName = sheetName;
    this.count.textContent = String(people.count);
    this.element.setAttribute('aria-label', words.peopleCount(people.count));
    this.element.title = words.peopleCount(people.count);
    if (this.panel) this.fill(this.panel);
  }

  private open(): void {
    if (this.panel) return;
    const panel = document.createElement('div');
    panel.className = 'bym-collab-people';
    panel.setAttribute('role', 'list');
    panel.setAttribute('aria-label', this.words.peopleCount(this.people.count));
    panel.tabIndex = -1;
    const box = this.element.getBoundingClientRect();
    const width = Math.min(280, Math.max(160, (typeof window !== 'undefined' ? window.innerWidth : 320) - 16));
    const left = Math.max(8, Math.min(box.left, (typeof window !== 'undefined' ? window.innerWidth : 320) - width - 8));
    panel.style.cssText = `position:fixed;z-index:2147483000;left:${left}px;top:${box.bottom + 4}px;width:${width}px;max-height:min(60vh,360px);overflow:auto;box-sizing:border-box;padding:4px 0;font:12px system-ui,sans-serif;color:#3c4043;background:#fff;border:1px solid #dadce0;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.2)`;
    panel.addEventListener('keydown', this.onKey);
    this.fill(panel);
    this.host.appendChild(panel);
    this.panel = panel;
    this.element.setAttribute('aria-expanded', 'true');
    document.addEventListener('pointerdown', this.onDocumentDown, true);
    document.addEventListener('mousedown', this.onDocumentDown, true);
  }

  /** Closes the list; `focus`: the button gets the focus back (it was closed from the keyboard). */
  close(focus: boolean): void {
    const panel = this.panel;
    if (!panel) return;
    this.panel = undefined;
    if (panel.parentNode) panel.parentNode.removeChild(panel);
    this.element.setAttribute('aria-expanded', 'false');
    document.removeEventListener('pointerdown', this.onDocumentDown, true);
    document.removeEventListener('mousedown', this.onDocumentDown, true);
    if (focus) this.element.focus();
  }

  private fill(panel: HTMLDivElement): void {
    while (panel.firstChild) panel.removeChild(panel.firstChild);
    panel.setAttribute('aria-label', this.words.peopleCount(this.people.count));
    for (const who of this.people.list) {
      const row = document.createElement('div');
      row.className = 'bym-collab-person';
      row.setAttribute('role', 'listitem');
      row.setAttribute('data-peer', who.id);
      if (who.self) row.setAttribute('data-self', 'true');
      if (who.readOnly) row.setAttribute('data-read-only', 'true');
      row.style.cssText = 'display:flex;align-items:baseline;gap:6px;padding:3px 10px;min-width:0';
      const dot = document.createElement('span');
      dot.setAttribute('aria-hidden', 'true');
      dot.className = 'bym-collab-person-color';
      dot.style.cssText = `flex:0 0 auto;display:inline-block;width:9px;height:9px;border-radius:50%;background:${who.color}`;
      const name = document.createElement('span');
      name.className = 'bym-collab-person-name';
      name.style.cssText = 'flex:0 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:600;unicode-bidi:isolate';
      name.dir = 'auto';
      name.textContent = who.name;
      row.appendChild(dot);
      row.appendChild(name);
      const notes: string[] = [];
      if (who.self) notes.push(this.words.you);
      if (who.readOnly) notes.push(this.words.viewOnly);
      const sheet = who.sheet !== undefined ? this.sheetName(who.sheet) : undefined;
      if (sheet !== undefined) notes.push(this.words.onSheet(sheet));
      if (notes.length > 0) {
        const note = document.createElement('span');
        note.className = 'bym-collab-person-note';
        note.style.cssText = 'flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#5f6368';
        note.textContent = notes.join(' · ');
        row.appendChild(note);
      }
      panel.appendChild(row);
    }
    if (this.people.more > 0) {
      const more = document.createElement('div');
      more.className = 'bym-collab-people-more';
      more.setAttribute('role', 'listitem');
      more.style.cssText = 'padding:3px 10px;color:#5f6368';
      more.textContent = this.words.morePeople(this.people.more);
      panel.appendChild(more);
    }
  }

  /** A small circle with the letters of a participant, for the bar. */
  static avatar(name: string, color: string, id: string): HTMLSpanElement {
    const badge = document.createElement('span');
    badge.className = 'bym-collab-peer';
    badge.setAttribute('data-peer', id);
    badge.title = name;
    badge.setAttribute('aria-hidden', 'true');
    badge.textContent = initials(name);
    badge.style.cssText = `display:inline-flex;align-items:center;justify-content:center;flex:0 0 auto;box-sizing:border-box;min-width:20px;height:20px;padding:0 3px;border-radius:10px;font-size:10px;font-weight:600;color:#fff;background:${color}`;
    return badge;
  }

  destroy(): void {
    this.close(false);
  }
}
