// Asking for the password of a protected file: the loop mount() opens its file through
// (openAsking), the question of the editor over a DOM that is made here, and the messages of
// the WebView bridge. The encrypted files are made by the engine itself: what is tested here
// is who is asked what and when, and that the password shows up nowhere (the encryption has
// its own tests against the reference implementation).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { toHex, utf16le } from '../src/crypto/bytes';
import { base64Decode, base64Encode, createBridge, type BridgeHost } from '../src/embed/bridge';
import { SaveController } from '../src/embed/save-to';
import { PasswordError, encryptionOf, open, type OfficeFile } from '../src/index';
import { getCell } from '../src/model/sheet/types';
import { setCellInput } from '../src/ops/commands';
import { applyOp } from '../src/ops/ops';
import { PasswordPrompt, type PasswordRequest } from '../src/ui/password-dialog';
import { openAsking } from '../src/ui/password-flow';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { minimalXlsx } from './fixtures';
import { cat, number, workbookStream } from './xls-fixtures';
import { CfbFile } from '../src/cfb';
import { writeCfb } from '../src/cfb-write';

const PASSWORD = 'Gizli-Parola-42';
const isCfb = (bytes: Uint8Array): boolean => toHex(bytes.subarray(0, 8)) === 'd0cf11e0a1b11ae1';

let cached: Uint8Array | undefined;
/** An xlsx encrypted with PASSWORD (once: a hundred thousand rounds take their time). */
async function encrypted(): Promise<Uint8Array> {
  if (!cached) cached = await (await open(await minimalXlsx())).save({ password: PASSWORD });
  return cached;
}

const opener = (bytes: Uint8Array, tried: (string | undefined)[] = []) => (password: string | undefined, signal: { readonly aborted: boolean }): Promise<OfficeFile> => {
  tried.push(password);
  return open(bytes, { password, signal });
};

async function rejectsWith(work: Promise<unknown>, code: string, label: string): Promise<PasswordError> {
  try {
    await work;
  } catch (error) {
    assert.ok(error instanceof PasswordError, `${label}: a PasswordError is thrown, got ${String(error)}`);
    assert.equal(error.code, code, `${label}: ${error.message}`);
    return error;
  }
  throw new assert.AssertionError({ message: `${label}: nothing was thrown` });
}

// ------------------------------------------------------------------ the loop

test('asking: a file that needs no password, and one whose password was given, ask nobody', async () => {
  const asked: PasswordRequest[] = [];
  const onPassword = (request: PasswordRequest) => {
    asked.push(request);
    return null;
  };
  const plain = await openAsking(opener(await minimalXlsx()), { onPassword });
  assert.deepEqual([plain.format, plain.passwordProtected], ['xlsx', false]);
  const file = await openAsking(opener(await encrypted()), { password: PASSWORD, onPassword });
  assert.deepEqual([file.format, file.passwordProtected], ['xlsx', true]);
  assert.deepEqual(asked, []);
});

test('asking: onPassword is asked when no password was given, and again for as long as the password is wrong', async () => {
  const asked: PasswordRequest[] = [];
  const answers = ['yanlış', 'yine yanlış', PASSWORD];
  const tried: (string | undefined)[] = [];
  const file = await openAsking(opener(await encrypted(), tried), {
    onPassword: async (request) => {
      asked.push({ ...request });
      return answers[asked.length - 1];
    },
  });
  assert.equal(file.passwordProtected, true);
  assert.deepEqual(asked, [
    { attempts: 0, wrong: false },
    { attempts: 1, wrong: true },
    { attempts: 2, wrong: true },
  ]);
  assert.deepEqual(tried, [undefined, 'yanlış', 'yine yanlış', PASSWORD]);
});

test('asking: a password of the options that is wrong is said to be wrong the first time', async () => {
  const asked: PasswordRequest[] = [];
  await openAsking(opener(await encrypted()), {
    password: 'eski parola',
    onPassword: (request) => {
      asked.push({ ...request });
      return PASSWORD;
    },
  });
  assert.deepEqual(asked, [{ attempts: 1, wrong: true }]);
});

test('asking: null or nothing from onPassword gives up with "cancelled"; without anyone to ask the error is "required" or "wrong"', async () => {
  const bytes = await encrypted();
  const cancelled = await rejectsWith(openAsking(opener(bytes), { onPassword: () => null }), 'cancelled', 'null');
  assert.equal(cancelled.format, 'ooxml');
  await rejectsWith(openAsking(opener(bytes), { onPassword: async () => undefined }), 'cancelled', 'undefined');
  await rejectsWith(openAsking(opener(bytes), {}), 'required', 'nobody to ask, no password');
  await rejectsWith(openAsking(opener(bytes), { password: 'x' }), 'wrong', 'nobody to ask, a wrong password');
  await assert.rejects(openAsking(opener(bytes), { onPassword: () => 42 as unknown as string }), /onPassword\(\) gives the password as a string/);
  // What onPassword throws is the caller's own error.
  await assert.rejects(
    openAsking(opener(bytes), {
      onPassword: () => {
        throw new Error('the application failed');
      },
    }),
    /the application failed/,
  );
});

test('asking: the empty text is a password like any other, not giving up', async () => {
  const tried: (string | undefined)[] = [];
  const answers: (string | null)[] = ['', null];
  await rejectsWith(openAsking(opener(await encrypted(), tried), { onPassword: () => answers.shift() }), 'cancelled', 'empty, then null');
  assert.deepEqual(tried, [undefined, '']);
});

test('asking: every password is tried once, as it was given: nothing else is tried behind one that is wrong', async () => {
  const calls: unknown[][] = [];
  const answers = ['a', 'b', null];
  await rejectsWith(
    openAsking(
      async (...args: unknown[]) => {
        calls.push(args.slice(0, 1).concat(args.length));
        throw new PasswordError(args[0] === undefined ? 'required' : 'wrong', 'no');
      },
      { onPassword: () => answers.shift() },
    ),
    'cancelled',
    'given up after two passwords',
  );
  // The password and the signal, and nothing that says "try another one as well".
  assert.deepEqual(calls, [
    [undefined, 2],
    ['a', 2],
    ['b', 2],
  ]);
});

test('asking: errors that are not about the password are passed on, and nobody is asked', async () => {
  let asked = 0;
  const onPassword = () => {
    asked++;
    return 'x';
  };
  await rejectsWith(openAsking(async () => Promise.reject(new PasswordError('damaged', 'broken')), { onPassword }), 'damaged', 'a damaged file');
  await rejectsWith(openAsking(async () => Promise.reject(new PasswordError('unsupported', 'unknown')), { onPassword }), 'unsupported', 'an encryption that is not read');
  await assert.rejects(openAsking(async () => Promise.reject(new Error('Not a Word or Excel file')), { onPassword }), /Not a Word or Excel file/);
  assert.equal(asked, 0);
});

test('asking: the question of the editor is made once, asked for every wrong password, and closed at the end, whatever the end is', async () => {
  const log: string[] = [];
  const answers: (string | null)[] = ['wrong', PASSWORD];
  const make = () => ({
    ask: async (request: PasswordRequest) => {
      log.push(`ask ${request.attempts} ${request.wrong}`);
      return answers.shift() as string | null;
    },
    close: () => void log.push('close'),
  });
  const file = await openAsking(opener(await encrypted()), {}, make);
  assert.equal(file.passwordProtected, true);
  assert.deepEqual(log, ['ask 0 false', 'ask 1 true', 'close']);
  log.length = 0;
  answers.push(null);
  await rejectsWith(openAsking(opener(await encrypted()), {}, make), 'cancelled', 'the user gave up');
  assert.deepEqual(log, ['ask 0 false', 'close']);
});

test('asking: giving up while a password is being checked calls the check off', async () => {
  let cancel: (() => void) | undefined;
  let checks = 0;
  const work = openAsking(
    async (password, signal) => {
      checks++;
      if (password === undefined) throw new PasswordError('required', 'needs one');
      // A check that takes its time: the user presses Escape meanwhile.
      await new Promise((resolve) => setTimeout(resolve, 5));
      (cancel as () => void)();
      await new Promise((resolve) => setTimeout(resolve, 5));
      if (signal.aborted) throw new PasswordError('cancelled', 'called off');
      throw new PasswordError('wrong', 'wrong');
    },
    {},
    (onCancel) => {
      cancel = onCancel;
      return { ask: async () => 'something', close: () => undefined };
    },
  );
  await rejectsWith(work, 'cancelled', 'called off during the check');
  assert.equal(checks, 2);
});

// ------------------------------------------------------------------ the dialog

interface FakeEvent {
  key?: string;
  target?: FakeElement;
  isComposing?: boolean;
  stopped?: boolean;
  prevented?: boolean;
  stopPropagation(): void;
  preventDefault(): void;
}

class FakeStyle {
  cssText = '';
  private readonly set = new Map<string, { value: string; priority: string }>();
  setProperty(name: string, value: string, priority = ''): void {
    this.set.set(name, { value, priority });
  }
  getPropertyValue(name: string): string {
    const entry = this.set.get(name);
    return entry ? entry.value : '';
  }
  getPropertyPriority(name: string): string {
    const entry = this.set.get(name);
    return entry ? entry.priority : '';
  }
  removeProperty(name: string): void {
    this.set.delete(name);
  }
}

class FakeElement {
  children: FakeElement[] = [];
  parentNode: FakeElement | null = null;
  readonly attributes = new Map<string, string>();
  readonly style = new FakeStyle();
  readonly listeners = new Map<string, ((event: FakeEvent) => void)[]>();
  className = '';
  id = '';
  textContent = '';
  type = '';
  value = '';
  disabled = false;
  constructor(readonly tag: string) {}
  get ownerDocument(): typeof fakeDocument {
    return fakeDocument;
  }
  get isConnected(): boolean {
    for (let at: FakeElement | null = this; at; at = at.parentNode) if (at === fakeDocument.body) return true;
    return false;
  }
  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    child.parentNode = this;
    return child;
  }
  removeChild(child: FakeElement): FakeElement {
    this.children.splice(this.children.indexOf(child), 1);
    child.parentNode = null;
    if (!fakeDocument.activeElement.isConnected) fakeDocument.activeElement = fakeDocument.body;
    return child;
  }
  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }
  getAttribute(name: string): string | null {
    const value = this.attributes.get(name);
    return value === undefined ? null : value;
  }
  removeAttribute(name: string): void {
    this.attributes.delete(name);
  }
  addEventListener(type: string, listener: (event: FakeEvent) => void): void {
    const list = this.listeners.get(type) || [];
    list.push(listener);
    this.listeners.set(type, list);
  }
  removeEventListener(type: string, listener: (event: FakeEvent) => void): void {
    const list = this.listeners.get(type) || [];
    const at = list.indexOf(listener);
    if (at >= 0) list.splice(at, 1);
  }
  contains(other: FakeElement | null): boolean {
    for (let at = other; at; at = at.parentNode) if (at === this) return true;
    return false;
  }
  getRootNode(): typeof fakeDocument {
    return fakeDocument;
  }
  focus(): void {
    if (this.isConnected && !this.disabled) fakeDocument.activeElement = this;
  }
  /** Sends an event to this element and up through its parents, as the DOM does. */
  fire(type: string, more: Partial<FakeEvent> = {}): FakeEvent {
    const event: FakeEvent = {
      target: this,
      ...more,
      stopPropagation() {
        event.stopped = true;
      },
      preventDefault() {
        event.prevented = true;
      },
    };
    for (let at: FakeElement | null = this; at && !event.stopped; at = at.parentNode) for (const listener of (at.listeners.get(type) || []).slice()) listener(event);
    return event;
  }
  /** Every element under this one, itself included. */
  all(): FakeElement[] {
    return [this as FakeElement].concat(...this.children.map((child) => child.all()));
  }
  find(attribute: string, value: string): FakeElement {
    const found = this.all().find((element) => element.getAttribute(attribute) === value);
    if (!found) throw new Error(`no element with ${attribute}="${value}"`);
    return found;
  }
}

const fakeDocument = {
  body: new FakeElement('body'),
  activeElement: undefined as unknown as FakeElement,
  createElement: (tag: string) => new FakeElement(tag),
};
fakeDocument.activeElement = fakeDocument.body;

/** A container in a page, the globals the dialog uses, and what takes them away again. */
function page(position = 'static'): { container: FakeElement; outside: FakeElement; restore(): void } {
  const scope = globalThis as unknown as { document?: unknown; window?: unknown };
  const before = { document: scope.document, window: scope.window };
  fakeDocument.body.children = [];
  const container = fakeDocument.body.appendChild(new FakeElement('div'));
  const outside = fakeDocument.body.appendChild(new FakeElement('button'));
  fakeDocument.activeElement = outside;
  scope.document = fakeDocument;
  scope.window = { getComputedStyle: () => ({ position }), matchMedia: () => ({ matches: false }) };
  return {
    container,
    outside,
    restore: () => {
      scope.document = before.document;
      scope.window = before.window;
    },
  };
}

const promptIn = (container: FakeElement, strings = STRINGS_EN, onCancel?: () => void): PasswordPrompt =>
  new PasswordPrompt({ container: container as unknown as HTMLElement, strings, onCancel });

test('dialog: the question is a dialog inside the container with a password field that the browser neither fills in nor corrects', () => {
  const { container, restore } = page();
  try {
    const prompt = promptIn(container);
    assert.equal(container.children.length, 0, 'nothing is shown before something is asked');
    void prompt.ask({ attempts: 0, wrong: false });
    assert.equal(container.children.length, 1);
    const panel = container.find('role', 'dialog');
    assert.equal(panel.getAttribute('aria-modal'), 'true');
    const title = container.all().find((element) => element.id === panel.getAttribute('aria-labelledby'));
    assert.equal(title && title.textContent, 'Password required');
    const field = container.find('data-bym-field', 'password');
    // "new-password": a browser does not fill in the password it has stored for the site, and "off" is not honoured for passwords.
    assert.deepEqual([field.tag, field.type, field.getAttribute('autocomplete'), field.getAttribute('spellcheck'), field.getAttribute('autocapitalize')], ['input', 'password', 'new-password', 'false', 'off']);
    // No name and no form: nothing to submit, and nothing a password manager is offered to save.
    assert.equal(field.getAttribute('name'), null);
    assert.deepEqual(container.all().filter((element) => element.tag === 'form'), []);
    const label = container.all().find((element) => element.tag === 'label');
    assert.deepEqual([label && label.getAttribute('for'), label && label.textContent], [field.id, 'Password']);
    assert.equal(fakeDocument.activeElement, field, 'the field has the focus');
    assert.equal(container.style.getPropertyValue('position'), 'relative', 'a container that is not positioned is, while the question is there');
    assert.deepEqual([container.find('data-bym-action', 'open').textContent, container.find('data-bym-action', 'cancel').textContent], ['Open', 'Cancel']);
    prompt.close();
    assert.equal(container.children.length, 0, 'nothing is left of the question');
    assert.equal(container.style.getPropertyValue('position'), '', 'the container is as it was');
  } finally {
    restore();
  }
});

test('dialog: Enter and the button answer with what was typed, and the field is empty by then', async () => {
  const { container, outside, restore } = page('relative');
  try {
    const prompt = promptIn(container);
    const first = prompt.ask({ attempts: 0, wrong: false });
    const field = container.find('data-bym-field', 'password');
    field.value = PASSWORD;
    const event = field.fire('keydown', { key: 'Enter' });
    assert.deepEqual([event.stopped, event.prevented], [true, true], 'the key is the question\'s alone');
    assert.equal(field.value, '', 'the password is taken out of the field at once');
    assert.equal(await first, PASSWORD);
    assert.equal(field.disabled, true, 'nothing is typed while the password is checked');
    const message = container.all().find((element) => element.getAttribute('role') === 'alert') as FakeElement;
    assert.equal(message.textContent, 'Checking the password…');

    // It was wrong: asked again, in the same dialog.
    const second = prompt.ask({ attempts: 1, wrong: true });
    assert.equal(container.children.length, 1, 'the same dialog');
    assert.deepEqual([field.disabled, field.value, field.getAttribute('aria-invalid')], [false, '', 'true']);
    assert.equal(message.textContent, 'Wrong password. Try again.');
    assert.equal(field.getAttribute('aria-describedby'), message.id);
    field.value = 'ikinci deneme';
    container.find('data-bym-action', 'open').fire('click');
    assert.equal(await second, 'ikinci deneme');
    assert.equal(field.value, '');

    void prompt.ask({ attempts: 2, wrong: true });
    assert.equal(message.textContent, 'Wrong password (2 attempts). Try again.');
    field.value = 'typed and never sent';
    prompt.close();
    assert.equal(field.value, '', 'what was typed and not sent goes with the dialog');
    assert.equal(fakeDocument.activeElement, outside, 'the focus is back where it was');
    // No element that was made holds a password in an attribute or in its text.
    for (const element of [field, message, ...container.all()]) {
      const said = `${element.textContent} ${element.value} ${[...element.attributes.values()].join(' ')}`;
      assert.ok(!said.includes(PASSWORD) && !said.includes('ikinci deneme') && !said.includes('never sent'), `nothing typed is left in <${element.tag}>`);
    }
  } finally {
    restore();
  }
});

test('dialog: Escape and Cancel give up, also while a password is being checked', async () => {
  const { container, restore } = page();
  try {
    let cancelled = 0;
    const prompt = promptIn(container, STRINGS_TR, () => cancelled++);
    const first = prompt.ask({ attempts: 0, wrong: false });
    const field = container.find('data-bym-field', 'password');
    assert.equal(container.find('data-bym-action', 'cancel').textContent, 'Vazgeç');
    field.value = 'yarım';
    const escape = field.fire('keydown', { key: 'Escape' });
    assert.deepEqual([escape.stopped, escape.prevented], [true, true]);
    assert.equal(await first, null);
    assert.equal(field.value, '');
    assert.equal(cancelled, 1);
    // Nothing is asked after the user gave up.
    assert.equal(await prompt.ask({ attempts: 0, wrong: false }), null);
    prompt.close();

    // While a password is checked there is no question open: giving up is told through onCancel, and to the next question.
    const other = promptIn(container, STRINGS_TR, () => cancelled++);
    const asked = other.ask({ attempts: 0, wrong: false });
    const field2 = container.find('data-bym-field', 'password');
    field2.value = 'x';
    field2.fire('keydown', { key: 'Enter' });
    assert.equal(await asked, 'x');
    container.find('data-bym-action', 'cancel').fire('click');
    assert.equal(cancelled, 2);
    assert.equal(await other.ask({ attempts: 1, wrong: true }), null);
    other.close();
    other.close();
    assert.equal(container.children.length, 0);
  } finally {
    restore();
  }
});

test('dialog: Enter while an input method composes is not an answer, and keys do not reach the page', async () => {
  const { container, restore } = page();
  try {
    const prompt = promptIn(container);
    let answered = false;
    void prompt.ask({ attempts: 0, wrong: false }).then(() => (answered = true));
    const field = container.find('data-bym-field', 'password');
    let heard = 0;
    container.addEventListener('keydown', () => heard++);
    container.addEventListener('keyup', () => heard++);
    container.addEventListener('input', () => heard++);
    field.value = 'şifre';
    field.fire('keydown', { key: 'Enter', isComposing: true });
    field.fire('keydown', { key: 's' });
    field.fire('keyup', { key: 's' });
    field.fire('input');
    await Promise.resolve();
    assert.deepEqual([answered, field.value, heard], [false, 'şifre', 0]);
    prompt.close();
  } finally {
    restore();
  }
});

// ------------------------------------------------------------------ saving with the controller

test('saveTo: a file that was opened with a password is handed over encrypted by the Save command and by auto-save alike', async () => {
  const file = await openAsking(opener(await encrypted()), { password: PASSWORD });
  const workbook = await file.workbook();
  for (const op of setCellInput(workbook, 0, 0, 0, '1234')) applyOp(workbook, op);
  const saves: { bytes: Uint8Array; auto: boolean }[] = [];
  const controller = new SaveController({
    target: {
      kind: 'function',
      save: async (bytes, info) => {
        saves.push({ bytes, auto: info.auto === true });
      },
    },
    writable: () => true,
    commit: () => true,
    bytes: () => file.save(),
    bytesAsIs: () => file.save(),
    pending: () => false,
    readOnly: () => false,
    autoSave: { delayMs: 250, maxWaitMs: 250, unconditional: true },
    info: () => ({ format: 'xlsx' }),
    strings: () => STRINGS_EN,
    notice: () => undefined,
    etag: null,
  } as unknown as ConstructorParameters<typeof SaveController>[0]);
  controller.changed();
  assert.equal(await controller.saveNow(), true, 'the save of the button');
  controller.changed();
  for (let i = 0; i < 300 && saves.length < 2; i++) await new Promise((resolve) => setTimeout(resolve, 10));
  controller.destroy();
  assert.deepEqual(saves.map((save) => save.auto), [false, true], 'one save of the user, one of auto-save');
  for (const save of saves) {
    assert.ok(isCfb(save.bytes), `the ${save.auto ? 'automatic' : 'user\'s'} save is a compound file`);
    assert.deepEqual(encryptionOf(save.bytes), { format: 'ooxml', encryption: { scheme: 'agile', cipher: 'AES', keyBits: 256, hash: 'SHA512', spinCount: 100000, integrity: true } });
    const again = await open(save.bytes, { password: PASSWORD });
    const cell = getCell((await again.workbook()).sheets[0], 0, 0);
    assert.equal(cell && cell.value, 1234);
  }
});

// ------------------------------------------------------------------ the bridge

interface Fake {
  host: BridgeHost;
  sent: any[];
  states: string[];
  views: number;
  destroyed: number;
  prompts: { asked: PasswordRequest[]; answer?: (value: string | null) => void; closed: number; cancel(): void }[];
  /** What open() of the host does instead of opening, when a test sets it. */
  opener?: (bytes: Uint8Array, options?: { password?: string; signal?: { readonly aborted: boolean } }) => Promise<OfficeFile>;
  /** How often the host was asked to open a file, and with which passwords. */
  opens: (string | undefined)[];
  opened: OfficeFile[];
}

function fake(withPrompt = false): Fake {
  const state: Fake = { host: undefined as unknown as BridgeHost, sent: [], states: [], views: 0, destroyed: 0, prompts: [], opened: [], opens: [] };
  const view = () => {
    state.views++;
    return {
      history: { canUndo: false, canRedo: false },
      activeSheet: 0,
      isEditing: false,
      commitEditing: () => undefined,
      commitPending: () => true,
      localizeError: (error: Error) => error.message,
      showNotice: () => undefined,
      setReadOnly: () => undefined,
      setLocale: () => undefined,
      destroy: () => void state.destroyed++,
    } as any;
  };
  state.host = {
    container: { addEventListener: () => undefined, removeEventListener: () => undefined } as unknown as HTMLElement,
    version: '9.9.9',
    post(json) {
      state.sent.push(JSON.parse(json));
      return true;
    },
    open: async (bytes, options) => {
      state.opens.push(options && options.password);
      const file = state.opener ? await state.opener(bytes, options) : await open(bytes, options);
      state.opened.push(file);
      return file;
    },
    createView: view,
    onState: (name) => void state.states.push(name),
    timer: () => undefined,
    frame: () => undefined,
  };
  if (withPrompt) {
    state.host.passwordPrompt = (_container, _locale, onCancel) => {
      const prompt: Fake['prompts'][number] = { asked: [], closed: 0, cancel: onCancel };
      state.prompts.push(prompt);
      return {
        ask: (request) =>
          new Promise<string | null>((resolve) => {
            prompt.asked.push({ ...request });
            prompt.answer = resolve;
          }),
        close: () => void prompt.closed++,
      };
    };
  }
  return state;
}

/** Waits until the page has sent a message of the type (the check of a password takes a moment). */
async function until(f: Fake, type: string, count = 1): Promise<any> {
  for (let i = 0; i < 3000; i++) {
    const found = f.sent.filter((message) => message.type === type);
    if (found.length >= count) return found[count - 1];
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error(`no "${type}" message (${count}) among: ${f.sent.map((message) => message.type).join(', ')}`);
}

/** Everything that left the page, and everything a test caught, as text: the password must be in none of it. */
function nothingLeaks(f: Fake, passwords: string[], more: string[] = []): void {
  const said = JSON.stringify(f.sent) + more.join(' ');
  for (const password of passwords) {
    for (const form of [password, base64Encode(new TextEncoder().encode(password)), toHex(utf16le(password)), base64Encode(utf16le(password))]) {
      assert.ok(!said.includes(form), `a message holds the password (${form.slice(0, 6)}…)`);
    }
  }
}

test('bridge: a load with the password of the file opens it, and loaded says the file is encrypted', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()), password: PASSWORD });
  const loaded = f.sent.find((message) => message.type === 'loaded');
  assert.deepEqual([loaded && loaded.id, loaded && loaded.format, loaded && loaded.encrypted], ['a', 'xlsx', true]);
  assert.equal(f.sent.some((message) => message.type === 'passwordRequired'), false);
  nothingLeaks(f, [PASSWORD]);

  // What `save` hands back is encrypted with the same password.
  await bridge.receive({ type: 'save', id: 's' });
  const saved = f.sent.find((message) => message.type === 'saved');
  assert.equal(saved.encrypted, true);
  const bytes = base64Decode(saved.data);
  assert.ok(isCfb(bytes), 'the saved bytes are a compound file');
  assert.equal((await open(bytes, { password: PASSWORD })).format, 'xlsx');
  nothingLeaks(f, [PASSWORD]);

  // A file without a password says nothing of the kind.
  await bridge.receive({ type: 'load', id: 'b', data: base64Encode(await minimalXlsx()) });
  const plain = f.sent.filter((message) => message.type === 'loaded')[1];
  assert.equal('encrypted' in plain, false);
  bridge.destroy();
});

test('bridge: without a password the application is asked, told when it was wrong, and the file opens with the right one', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()) });
  const first = await until(f, 'passwordRequired');
  assert.deepEqual(first, { type: 'passwordRequired', wrong: false, attempts: 0, id: 'a', format: 'ooxml' });
  await bridge.receive({ type: 'password', id: 'p1', value: 'yanlış-parola' });
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'done', id: 'p1', command: 'password' });
  const second = await until(f, 'passwordRequired', 2);
  assert.deepEqual([second.wrong, second.attempts], [true, 1]);
  await bridge.receive({ type: 'password', value: PASSWORD });
  await loading;
  const loaded = f.sent.find((message) => message.type === 'loaded');
  assert.deepEqual([loaded && loaded.id, loaded && loaded.encrypted], ['a', true]);
  assert.equal(f.views, 1);
  nothingLeaks(f, [PASSWORD, 'yanlış-parola']);
  bridge.destroy();
});

test('bridge: a password of null gives up: the load fails with "password-cancelled" and nothing is shown', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()), password: 'eski-parola' });
  const asked = await until(f, 'passwordRequired');
  assert.deepEqual([asked.wrong, asked.attempts], [true, 1], 'the password of the load was wrong');
  await bridge.receive({ type: 'password', value: null });
  await loading;
  const error = f.sent.find((message) => message.type === 'error');
  assert.deepEqual([error.code, error.id], ['password-cancelled', 'a']);
  assert.equal(f.views, 0);
  assert.deepEqual(f.states.slice(-1), ['error']);
  nothingLeaks(f, [PASSWORD, 'eski-parola']);
  bridge.destroy();
});

test('bridge: a password message nobody asked for, or of the wrong shape, is refused without repeating it', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'password', id: 'x', value: 'Sızmamalı-1' });
  assert.deepEqual([f.sent[f.sent.length - 1].type, f.sent[f.sent.length - 1].code, f.sent[f.sent.length - 1].id], ['error', 'password-not-expected', 'x']);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()) });
  await until(f, 'passwordRequired');
  await bridge.receive({ type: 'password', id: 'y', value: 12345678 });
  assert.equal(f.sent[f.sent.length - 1].code, 'bad-message');
  await bridge.receive({ type: 'password', id: 'z' });
  assert.equal(f.sent[f.sent.length - 1].code, 'bad-message');
  await bridge.receive({ type: 'password', id: 'w', value: 'Sızmamalı-2'.repeat(200) });
  assert.equal(f.sent[f.sent.length - 1].code, 'bad-message');
  await bridge.receive({ type: 'load', id: 'bad', data: 'AAAA', password: 7 });
  await bridge.receive({ type: 'close' });
  await loading;
  nothingLeaks(f, ['Sızmamalı-1', 'Sızmamalı-2', '12345678']);
  bridge.destroy();
});

test('bridge: closing, or loading another file, ends a load that waits for its password', async () => {
  const f = fake(true);
  const bridge = createBridge(f.host);
  const first = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()) });
  await until(f, 'passwordRequired');
  assert.equal(f.prompts.length, 1);
  await bridge.receive({ type: 'close', id: 'c' });
  await first;
  assert.deepEqual(f.sent.filter((message) => message.id === 'a' && message.type === 'error').map((message) => message.code), ['load-failed']);
  assert.equal(f.prompts[0].closed >= 1, true, 'the question inside the page is taken away');
  // The answer comes too late: nobody is asking.
  await bridge.receive({ type: 'password', id: 'late', value: PASSWORD });
  assert.equal(f.sent[f.sent.length - 1].code, 'password-not-expected');

  const second = bridge.receive({ type: 'load', id: 'b', data: base64Encode(await encrypted()) });
  await until(f, 'passwordRequired', 2);
  const third = bridge.receive({ type: 'load', id: 'c2', data: base64Encode(await minimalXlsx()) });
  await Promise.all([second, third]);
  assert.deepEqual(f.sent.filter((message) => message.type === 'loaded').map((message) => message.id), ['c2']);
  assert.equal(f.views, 1);
  nothingLeaks(f, [PASSWORD]);
  bridge.destroy();
});

test('bridge: the page asks the user as well; the first answer counts, from the user or from the application', async () => {
  const f = fake(true);
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()), locale: 'tr' });
  await until(f, 'passwordRequired');
  assert.deepEqual(f.prompts[0].asked, [{ attempts: 0, wrong: false }]);
  assert.deepEqual(f.states.slice(-2), ['loading', 'password'], 'the page shows the container for the question');
  // The user types a wrong one.
  (f.prompts[0].answer as (value: string | null) => void)('kullanıcının yanlışı');
  await until(f, 'passwordRequired', 2);
  assert.deepEqual(f.prompts[0].asked[1], { attempts: 1, wrong: true });
  // The application answers this time.
  await bridge.receive({ type: 'password', value: PASSWORD });
  await loading;
  assert.equal(f.sent.some((message) => message.type === 'loaded'), true);
  assert.equal(f.prompts.length, 1, 'one question for the whole load');
  assert.equal(f.prompts[0].closed >= 1, true, 'and it is gone when the file is open');
  assert.deepEqual(f.states.slice(-1), ['open']);
  nothingLeaks(f, [PASSWORD, 'kullanıcının yanlışı']);
  bridge.destroy();
});

test('bridge: the user giving up inside the page fails the load; "passwordPrompt: false" asks the application alone', async () => {
  const f = fake(true);
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()) });
  await until(f, 'passwordRequired');
  (f.prompts[0].answer as (value: string | null) => void)(null);
  await loading;
  assert.equal(f.sent.find((message) => message.type === 'error').code, 'password-cancelled');

  const alone = bridge.receive({ type: 'load', id: 'b', data: base64Encode(await encrypted()), passwordPrompt: false });
  await until(f, 'passwordRequired', 2);
  assert.equal(f.prompts.length, 1, 'no second question inside the page');
  await bridge.receive({ type: 'password', value: PASSWORD });
  await alone;
  assert.equal(f.sent.filter((message) => message.type === 'loaded').length, 1);
  bridge.destroy();
});

test('bridge: a file that is given a password is saved encrypted from then on, converted or not; closing it forgets the password', async () => {
  // An xls that is converted to be edited. (A protected xls, which only the reference script can make, is in password-files.test.ts.)
  const stream = workbookStream({ sheets: [{ name: 'S', records: cat(number(0, 0, 5)) }] });
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'a', data: base64Encode(writeCfb([{ path: 'Workbook', data: stream }])), convert: true });
  const loaded = f.sent.find((message) => message.type === 'loaded');
  assert.deepEqual([loaded.format, loaded.convertedFrom, 'encrypted' in loaded], ['xlsx', 'xls', false]);
  f.opened[f.opened.length - 1].setPassword(PASSWORD);
  await bridge.receive({ type: 'save', id: 's' });
  const saved = f.sent.find((message) => message.type === 'saved');
  assert.equal(saved.encrypted, true);
  assert.ok(isCfb(base64Decode(saved.data)), 'encrypted from the moment the file has a password');
  // Closing the file forgets the password.
  const file = f.opened[f.opened.length - 1];
  await bridge.receive({ type: 'close' });
  await assert.rejects(file.save(), /password of this file was forgotten/);
  nothingLeaks(f, [PASSWORD]);
  bridge.destroy();
});

test('bridge: an encryption that is not read is "unsupported", a damaged encrypted file "corrupt", and nobody is asked', async () => {
  const f = fake(true);
  const bridge = createBridge(f.host);
  const extensible = writeCfb([
    { path: 'EncryptionInfo', data: new Uint8Array([4, 0, 3, 0, 0x10, 0, 0, 0, 1, 2, 3, 4]) },
    { path: 'EncryptedPackage', data: new Uint8Array(64) },
  ]);
  await bridge.receive({ type: 'load', id: 'a', data: base64Encode(extensible) });
  assert.deepEqual([f.sent[f.sent.length - 1].type, f.sent[f.sent.length - 1].code], ['error', 'unsupported']);
  const damaged = (await encrypted()).slice();
  damaged[damaged.length - 700] ^= 1;
  damaged[damaged.length - 1300] ^= 1;
  damaged[damaged.length - 2000] ^= 1;
  await bridge.receive({ type: 'load', id: 'b', data: base64Encode(damaged), password: PASSWORD });
  const error = f.sent[f.sent.length - 1];
  assert.deepEqual([error.type, error.id], ['error', 'b']);
  assert.ok(error.code === 'corrupt', `a damaged encrypted file is ${error.code}: ${error.message}`);
  assert.equal(f.prompts.length, 0);
  nothingLeaks(f, [PASSWORD]);
  bridge.destroy();
});

// ------------------------------------------------------------------ the bridge, after the review

const wait = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

test('bridge: every password is tried once: the host opens the file once for each', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()) });
  await until(f, 'passwordRequired');
  await bridge.receive({ type: 'password', value: 'yanlış' });
  await until(f, 'passwordRequired', 2);
  await bridge.receive({ type: 'password', value: PASSWORD });
  await loading;
  assert.deepEqual(f.opens, [undefined, 'yanlış', PASSWORD]);
  bridge.destroy();
});

test('bridge: giving up in the question of the page while a password is being checked calls the check off', async () => {
  const f = fake(true);
  let aborted = false;
  f.opener = async (bytes, options) => {
    if (!options || options.password === undefined) return open(bytes, options);
    // A check that takes its time, as a hundred thousand rounds do on a phone: it ends when it is called off.
    for (let i = 0; i < 400 && !(options.signal && options.signal.aborted); i++) await wait(5);
    aborted = options.signal !== undefined && options.signal.aborted;
    throw new PasswordError(aborted ? 'cancelled' : 'wrong', aborted ? 'called off' : 'wrong');
  };
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()) });
  await until(f, 'passwordRequired');
  (f.prompts[0].answer as (value: string | null) => void)('denetleniyor');
  await wait(20);
  // "Cancel" in the dialog: no question is open, the check is under way.
  f.prompts[0].cancel();
  await loading;
  assert.equal(aborted, true, 'the check was called off');
  const error = f.sent.find((message) => message.type === 'error');
  assert.deepEqual([error.code, error.id], ['password-cancelled', 'a']);
  assert.equal(f.sent.filter((message) => message.type === 'passwordRequired').length, 1, 'nobody is asked again');
  nothingLeaks(f, [PASSWORD, 'denetleniyor']);
  bridge.destroy();
});

test('bridge: the password of the application takes the question of the page away; a wrong one brings a new question', async () => {
  const f = fake(true);
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()) });
  await until(f, 'passwordRequired');
  assert.equal(f.prompts[0].closed, 0, 'the page asks');
  await bridge.receive({ type: 'password', value: 'uygulamanın yanlışı' });
  assert.equal(f.prompts[0].closed >= 1, true, 'the question of the page is gone as soon as the application answered');
  await until(f, 'passwordRequired', 2);
  assert.equal(f.prompts.length, 2, 'a new question for the second try');
  assert.deepEqual(f.prompts[1].asked, [{ attempts: 1, wrong: true }]);
  await bridge.receive({ type: 'password', value: PASSWORD });
  await loading;
  assert.equal(f.sent.some((message) => message.type === 'loaded'), true);
  assert.equal(f.prompts[1].closed >= 1, true);
  nothingLeaks(f, [PASSWORD, 'uygulamanın yanlışı']);
  bridge.destroy();
});

test('bridge: a password that comes while another is being checked is refused with a code of its own, and is not kept', async () => {
  const f = fake(true);
  let release: (() => void) | undefined;
  f.opener = async (bytes, options) => {
    if (options && options.password !== undefined) await new Promise<void>((resolve) => (release = resolve));
    return open(bytes, options);
  };
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()) });
  await until(f, 'passwordRequired');
  await bridge.receive({ type: 'password', id: 'p1', value: 'yanlış' });
  // The check of "yanlış" is under way: a second password has nothing to answer.
  await bridge.receive({ type: 'password', id: 'p2', value: PASSWORD });
  const refused = f.sent.find((message) => message.id === 'p2');
  assert.deepEqual([refused.type, refused.code], ['error', 'password-not-expected']);
  (release as () => void)();
  // It was not kept for the question that comes now: the file is still not open.
  const again = await until(f, 'passwordRequired', 2);
  assert.deepEqual([again.wrong, again.attempts], [true, 1]);
  assert.equal(f.sent.some((message) => message.type === 'loaded'), false);
  await bridge.receive({ type: 'close' });
  await loading;
  nothingLeaks(f, [PASSWORD, 'yanlış']);
  bridge.destroy();
});

test('bridge: a file that was opened and is not shown after all is closed: an overtaken load, a model that cannot be read', async () => {
  const f = fake();
  let release: (() => void) | undefined;
  f.opener = async (bytes, options) => {
    const file = await open(bytes, options);
    await new Promise<void>((resolve) => (release = resolve));
    return file;
  };
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()), password: PASSWORD });
  for (let i = 0; i < 2000 && !release; i++) await wait(5);
  // The file is open, and the application closes before it is shown.
  await bridge.receive({ type: 'close' });
  (release as () => void)();
  await loading;
  assert.equal(f.views, 0);
  assert.equal(f.opened.length, 1);
  await assert.rejects(f.opened[0].save(), /password of this file was forgotten/);

  // A file whose model cannot be read.
  let closed = 0;
  f.opener = async () =>
    ({
      format: 'xlsx',
      kind: 'spreadsheet',
      passwordProtected: true,
      workbook: async () => {
        throw new Error('The model cannot be read');
      },
      save: async () => new Uint8Array(0),
      close: () => void closed++,
    }) as unknown as OfficeFile;
  await bridge.receive({ type: 'load', id: 'b', data: base64Encode(await minimalXlsx()) });
  assert.equal(f.sent[f.sent.length - 1].code, 'corrupt');
  assert.equal(closed, 1, 'the file that could not be shown was closed');
  bridge.destroy();
});

test('bridge: loaded says when anyone can open the file, and when it has no integrity check', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const plain = await open(await minimalXlsx());
  plain.setPassword('VelvetSweatshop');
  await bridge.receive({ type: 'load', id: 'a', data: base64Encode(await plain.save()) });
  const first = f.sent.find((message) => message.type === 'loaded');
  assert.deepEqual([first.encrypted, first.defaultPassword, 'integrity' in first], [true, true, false]);
  assert.equal(f.sent.some((message) => message.type === 'passwordRequired'), false, 'nobody is asked for a password everyone knows');

  // The integrity check taken out of the descriptor, by someone who does not know the password.
  const cfb = CfbFile.read(await encrypted());
  const info = cfb.stream('EncryptionInfo');
  const text = new TextDecoder().decode(info.subarray(8)).replace(/<dataIntegrity [^>]*\/>/, '');
  const body = new TextEncoder().encode(text);
  const stripped = new Uint8Array(8 + body.length);
  stripped.set(info.subarray(0, 8));
  stripped.set(body, 8);
  const file = writeCfb([
    { path: 'EncryptionInfo', data: stripped },
    { path: 'EncryptedPackage', data: cfb.stream('EncryptedPackage') },
  ]);
  await bridge.receive({ type: 'load', id: 'b', data: base64Encode(file), password: PASSWORD });
  const second = f.sent.filter((message) => message.type === 'loaded')[1];
  assert.deepEqual([second.encrypted, second.integrity, 'defaultPassword' in second], [true, false, false]);
  // A file with its check says nothing about it.
  await bridge.receive({ type: 'load', id: 'c', data: base64Encode(await encrypted()), password: PASSWORD });
  assert.equal('integrity' in f.sent.filter((message) => message.type === 'loaded')[2], false);
  bridge.destroy();
});
