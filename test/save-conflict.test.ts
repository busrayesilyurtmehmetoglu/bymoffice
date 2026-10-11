import assert from 'node:assert/strict';
import { test } from 'node:test';
import { SAVE_TIMEOUT_MS, SaveController, SaveError, resolveSaveTarget, type ConflictQuestion, type SaveHost, type SaveInfo, type SaveState, type SaveTarget } from '../src/embed/save-to';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

// What the controller does with a save the server refused because the file there is not the one
// that was opened (HTTP 412): it asks, and only the answer "overwrite" sends the file again.

interface Sent {
  url: string;
  init: RequestInit;
  answer(status: number, headers?: { [name: string]: string }): void;
}

interface Rig {
  controller: SaveController;
  sent: Sent[];
  states: SaveState[];
  notices: [kind: string, message: string][];
  errors: unknown[];
  /** The questions that were asked, with whether each was taken away again. */
  asked: { question: ConflictQuestion; closed: number }[];
  /** What happened, in the order it happened. */
  order: string[];
  infos: SaveInfo[];
  content: number[];
  writable: boolean;
}

const settle = async () => {
  for (let i = 0; i < 40; i++) await Promise.resolve();
};

const request = (more: Partial<Extract<SaveTarget, { kind: 'request' }>> = {}): SaveTarget => ({
  kind: 'request',
  url: 'https://app.example.com/files/7',
  method: 'PUT',
  headers: [],
  credentials: 'same-origin',
  sameAsFile: true,
  timeoutMs: SAVE_TIMEOUT_MS,
  allowConverted: true,
  onConflict: 'ask',
  ...more,
});

function rig(target: SaveTarget, more: Partial<SaveHost> = {}): Rig {
  const state: Rig = { controller: undefined as unknown as SaveController, sent: [], states: [], notices: [], errors: [], asked: [], order: [], infos: [], content: [1, 2, 3], writable: true };
  const fakeFetch = (url: unknown, init?: RequestInit) =>
    new Promise<Response>((resolve, reject) => {
      const signal = init && init.signal;
      if (signal) signal.addEventListener('abort', () => reject(new Error('aborted by the signal')));
      state.order.push('request');
      state.sent.push({
        url: String(url),
        init: init || {},
        answer(status, headers = {}) {
          const lower: { [name: string]: string } = {};
          for (const name of Object.keys(headers)) lower[name.toLowerCase()] = headers[name];
          resolve({ status, headers: { get: (name: string) => (name.toLowerCase() in lower ? lower[name.toLowerCase()] : null) } } as unknown as Response);
        },
      });
    });
  const host: SaveHost = {
    target,
    writable: () => state.writable,
    commit: () => true,
    bytes: async () => new Uint8Array(state.content),
    info: () => ({ format: 'xlsx' }),
    strings: () => STRINGS_EN,
    notice: (message, kind) => {
      state.order.push(`notice ${kind}`);
      state.notices.push([kind, message]);
    },
    onState: (value) => state.states.push(value),
    onError: (error) => {
      state.order.push('onError');
      state.errors.push(error);
    },
    askOverwrite: (question) => {
      const entry = { question, closed: 0 };
      state.order.push('ask');
      state.asked.push(entry);
      return () => void entry.closed++;
    },
    fetch: fakeFetch as typeof fetch,
    now: () => 1000,
    timer: () => () => undefined,
    etag: '"v1"',
    ...more,
  };
  state.controller = new SaveController(host);
  return state;
}

const headersOf = (sent: Sent) => sent.init.headers as { [name: string]: string };
/** The If-Match a request carries, however its name is written. */
const ifMatchOf = (sent: Sent): string | undefined => {
  const headers = headersOf(sent);
  for (const name of Object.keys(headers)) if (name.toLowerCase() === 'if-match') return headers[name];
  return undefined;
};

/** A rig with a change that was sent and refused with 412: the question is open. */
async function conflicted(target: SaveTarget = request(), more: Partial<SaveHost> = {}): Promise<{ r: Rig; first: Promise<boolean> }> {
  const r = rig(target, more);
  r.controller.changed();
  const first = r.controller.saveNow();
  await settle();
  assert.equal(r.sent.length, 1, 'the save was sent');
  assert.notEqual(ifMatchOf(r.sent[0]), undefined, 'the first save names the file it expects on the server');
  r.sent[0].answer(412);
  await settle();
  return { r, first };
}

test('a save that is refused with 412 asks whether to overwrite, after the notice and onError, and sends nothing by itself', async () => {
  const { r, first } = await conflicted();
  assert.equal(r.asked.length, 1, 'the question was asked once');
  assert.equal(r.asked[0].question.detail, '', 'a server says nothing that is shown');
  assert.deepEqual(r.order, ['request', 'notice error', 'onError', 'ask'], 'the notice and onError come before the question');
  assert.equal((r.errors[0] as SaveError).code, 'conflict', 'onError is given the conflict as before');
  assert.equal(r.notices[0][1], STRINGS_EN.saveConflict, 'the notice says what it said before');
  assert.equal(r.sent.length, 1, 'nothing is sent while the question is open');
  assert.equal(r.controller.asking, true, 'the controller says that it asks');
  assert.deepEqual(r.controller.state, { dirty: true, saving: false, error: STRINGS_EN.saveConflict }, 'the file is unsaved and no save is under way while the question is open');
  let answered: boolean | undefined;
  void first.then((ok) => void (answered = ok));
  await settle();
  assert.equal(answered, undefined, 'the save that ran into the conflict is answered when the question is');
  r.asked[0].question.answer(false);
  await settle();
  assert.equal(answered, false, 'cancelled: the save has failed');
});

test('the answer "overwrite" sends the file once more without If-Match and says that it was saved', async () => {
  const { r, first } = await conflicted(request({ headers: [['if-match', '"of the page"'], ['X-Token', 't']] }));
  r.asked[0].question.answer(true);
  await settle();
  assert.equal(r.asked[0].closed, 1, 'the question is taken away when it is answered');
  assert.equal(r.controller.asking, false, 'nothing is asked any more');
  assert.equal(r.sent.length, 2, 'one more request');
  assert.equal(ifMatchOf(r.sent[1]), undefined, 'the overwrite carries no If-Match, not the one of the page either');
  assert.equal(headersOf(r.sent[1])['X-Token'], 't', 'the other headers of the page are sent');
  assert.equal(r.controller.saving, true, 'the overwrite is a save under way');
  r.sent[1].answer(200, { ETag: '"v3"' });
  await settle();
  assert.equal(await first, true, 'the save that ran into the conflict ends as saved');
  assert.equal(r.sent.length, 2, 'and nothing else is sent');
  assert.equal(r.controller.dirty, false, 'nothing is unsaved');
  assert.deepEqual(r.notices[r.notices.length - 1], ['info', STRINGS_EN.saved], 'the user is told that the file was saved');
  assert.equal(r.controller.currentETag, '"v3"', 'later saves are guarded by the new ETag');
  // An answer given twice is one answer.
  r.asked[0].question.answer(true);
  await settle();
  assert.equal(r.sent.length, 2, 'a second answer to the same question sends nothing');
});

test('an overwrite that fails says why, and one the server still answers with 412 does not ask again', async () => {
  const { r, first } = await conflicted();
  r.asked[0].question.answer(true);
  await settle();
  r.sent[1].answer(412);
  await settle();
  assert.equal(await first, false, 'the overwrite failed');
  assert.equal(r.asked.length, 1, 'a save without a condition that is refused is not asked about');
  assert.equal(r.controller.asking, false, 'no question is open');
  assert.equal((r.errors[1] as SaveError).code, 'conflict', 'onError hears of it');

  const other = await conflicted();
  other.r.asked[0].question.answer(true);
  await settle();
  other.r.sent[1].answer(500);
  await settle();
  assert.equal(await other.first, false, 'the overwrite failed');
  assert.equal(other.r.notices[other.r.notices.length - 1][1], STRINGS_EN.saveFailed('HTTP 500'), 'the other failure is what is said');
  assert.equal(other.r.controller.dirty, true, 'the file stays unsaved');
});

test('the answer "cancel" sends nothing, leaves the file unsaved, and the next save asks again after its own 412', async () => {
  const { r, first } = await conflicted();
  r.asked[0].question.answer(false);
  await settle();
  assert.equal(await first, false, 'the save was not made');
  assert.equal(r.asked[0].closed, 1, 'the question is taken away');
  assert.equal(r.sent.length, 1, 'nothing was sent');
  assert.equal(r.controller.dirty, true, 'the file is still unsaved');
  assert.equal(r.controller.saving, false, 'no save is under way');
  const again = r.controller.saveNow();
  await settle();
  assert.equal(r.asked.length, 1, 'the question is not asked before the server has answered');
  assert.equal(r.sent.length, 2, 'the next save is sent');
  assert.equal(headersOf(r.sent[1])['If-Match'], '"v1"', 'with the ETag as it was: nothing was overwritten');
  r.sent[1].answer(412);
  await settle();
  assert.equal(r.asked.length, 2, 'and its 412 asks again');
  r.asked[1].question.answer(false);
  assert.equal(await again, false, 'cancelled again');
});

test('a save asked for while the question is open asks nothing and sends nothing; it ends as the question does', async () => {
  const { r } = await conflicted();
  const second = r.controller.saveNow();
  r.controller.run();
  // A change made meanwhile does not send anything either.
  r.content = [9];
  r.controller.changed();
  const third = r.controller.saveNow();
  await settle();
  assert.equal(r.asked.length, 1, 'no second question');
  assert.equal(r.sent.length, 1, 'no request');
  r.asked[0].question.answer(true);
  await settle();
  assert.equal(r.sent.length, 2, 'the answer sends the one overwrite');
  assert.deepEqual(Array.from(r.sent[1].init.body as unknown as Uint8Array), [9], 'with the file as it is now');
  r.sent[1].answer(204);
  await settle();
  assert.equal(await second, true, 'the call that waited is answered with what came of the overwrite');
  assert.equal(await third, true, 'and so is the one after a change');
  assert.equal(r.sent.length, 2, 'nothing else was sent');

  const other = await conflicted();
  const waiting = other.r.controller.saveNow();
  other.r.asked[0].question.answer(false);
  assert.equal(await waiting, false, 'cancelled: the call that waited is answered false');
  assert.equal(other.r.sent.length, 1, 'and nothing was sent');
});

test('with onConflict "fail" a 412 is told as before and nothing is asked', async () => {
  const { r, first } = await conflicted(request({ onConflict: 'fail' }));
  assert.equal(await first, false, 'the save failed at once');
  assert.equal(r.asked.length, 0, 'no question');
  assert.equal(r.controller.asking, false, 'the controller asks nothing');
  assert.deepEqual(r.order, ['request', 'notice error', 'onError'], 'the notice and onError, as before');
  assert.equal((r.errors[0] as SaveError).code, 'conflict', 'onError is given the conflict');
});

test('a host that cannot ask gets the conflict told as before', async () => {
  const { r, first } = await conflicted(request(), { askOverwrite: undefined });
  assert.equal(await first, false, 'the save failed at once');
  assert.equal(r.controller.asking, false, 'the controller asks nothing');
  assert.equal(r.notices[0][1], STRINGS_EN.saveConflict, 'the notice is the one of a conflict');
});

test('a function of the application that fails is not asked about, whatever it throws', async () => {
  let calls = 0;
  const r = rig({ kind: 'function', save: () => (calls++, Promise.reject(new Error('HTTP 412'))) }, { askOverwrite: undefined });
  r.controller.changed();
  assert.equal(await r.controller.saveNow(), false, 'the save failed');
  assert.equal(r.asked.length, 0, 'no question');
  assert.equal((r.errors[0] as SaveError).code, 'refused', 'a function that rejects has refused');
  assert.equal(calls, 1, 'the function was called once');
});

test('a function is told when the save it is given is an overwrite', async () => {
  const infos: SaveInfo[] = [];
  const r = rig({ kind: 'function', save: (_bytes, info) => void infos.push(info) });
  r.controller.changed();
  assert.equal(await r.controller.saveNow(), true, 'saved');
  r.controller.changed();
  assert.equal(await r.controller.saveNow({ overwrite: true }), true, 'saved over');
  assert.deepEqual(infos, [{ format: 'xlsx' }, { format: 'xlsx', overwrite: true }], 'only the overwrite is marked');
});

test('a function that rejects with a conflict is asked about when the host can ask, and its detail is shown as one short line', async () => {
  const infos: SaveInfo[] = [];
  let fail = true;
  const save = (_bytes: Uint8Array, info: SaveInfo): Promise<void> => {
    infos.push(info);
    return fail ? Promise.reject(new SaveError('conflict', 'The application reported a conflict', undefined, `changed by\n<b>Ada</b>\u0000${'x'.repeat(400)}`)) : Promise.resolve();
  };
  const r = rig({ kind: 'function', save });
  r.controller.changed();
  const first = r.controller.saveNow();
  await settle();
  assert.equal(r.asked.length, 1, 'the question was asked');
  const detail = r.asked[0].question.detail;
  assert.equal(detail.indexOf('changed by <b>Ada</b> xxx'), 0, 'the detail is the text that was given, as text, on one line');
  assert.equal(detail.length, 201, 'cut to 200 characters and an ellipsis');
  fail = false;
  r.asked[0].question.answer(true);
  assert.equal(await first, true, 'the overwrite was taken');
  assert.deepEqual(infos, [{ format: 'xlsx' }, { format: 'xlsx', overwrite: true }], 'the second call is marked as an overwrite');
});

test('a page that overwrites from inside onError is not asked as well, and one request is sent', async () => {
  let own: Promise<boolean> | undefined;
  const r: Rig = rig(request(), {
    onError: (error) => {
      if ((error as SaveError).code === 'conflict' && !own) own = r.controller.saveNow({ overwrite: true });
    },
  });
  r.controller.changed();
  const first = r.controller.saveNow();
  await settle();
  r.sent[0].answer(412);
  await settle();
  assert.equal(r.asked.length, 0, 'the page answered by itself: nothing is asked');
  assert.equal(r.sent.length, 2, 'its overwrite is the one request');
  assert.equal(headersOf(r.sent[1])['If-Match'], undefined, 'without If-Match');
  r.sent[1].answer(200);
  await settle();
  assert.equal(await (own as Promise<boolean>), true, 'the overwrite of the page succeeded');
  assert.equal(await first, false, 'the save that ran into the conflict failed');
  assert.equal(r.sent.length, 2, 'nothing else was sent');
});

test('a page that overwrites while the question is open closes it, and one request is sent', async () => {
  const { r, first } = await conflicted();
  const own = r.controller.saveNow({ overwrite: true });
  await settle();
  assert.equal(r.asked[0].closed, 1, 'the question is taken away');
  assert.equal(r.controller.asking, false, 'nothing is asked any more');
  assert.equal(r.sent.length, 2, 'one overwrite');
  // An answer to the question that is gone changes nothing.
  r.asked[0].question.answer(true);
  r.asked[0].question.answer(false);
  await settle();
  assert.equal(r.sent.length, 2, 'the late answers send nothing');
  r.sent[1].answer(200);
  assert.equal(await own, true, 'the overwrite succeeded');
  assert.equal(await first, true, 'and the save that asked ends with it');
});

test('a question that is answered while it is being asked is closed and overwrites once', async () => {
  let closed = 0;
  const r = rig(request(), {
    askOverwrite: (question) => {
      question.answer(true);
      return () => void closed++;
    },
  });
  r.controller.changed();
  const first = r.controller.saveNow();
  await settle();
  r.sent[0].answer(412);
  await settle();
  assert.equal(closed, 1, 'what the host handed back to take the question away is called');
  assert.equal(r.sent.length, 2, 'one overwrite');
  r.sent[1].answer(200);
  assert.equal(await first, true, 'saved');
});

test('a host that throws while asking leaves the conflict as a failure', async () => {
  const { r, first } = await conflicted(request(), {
    askOverwrite: () => {
      throw new Error('no document');
    },
  });
  assert.equal(await first, false, 'the save failed');
  assert.equal(r.controller.asking, false, 'nothing is asked');
  assert.equal(r.sent.length, 1, 'nothing was sent');
});

test('the question goes away, unanswered, when the file can no longer be written, when it is dismissed and when the controller is destroyed', async () => {
  const locked = await conflicted();
  locked.r.writable = false;
  assert.equal(await locked.r.controller.saveNow(), false, 'a file that cannot be written is not saved');
  assert.equal(locked.r.asked[0].closed, 1, 'and the question is taken away');
  assert.equal(await locked.first, false, 'cancelled');
  locked.r.asked[0].question.answer(true);
  await settle();
  assert.equal(locked.r.sent.length, 1, 'an answer that comes late sends nothing');

  const dismissed = await conflicted();
  dismissed.r.controller.dismissConflict();
  assert.equal(dismissed.r.asked[0].closed, 1, 'dismissed: the question is taken away');
  assert.equal(dismissed.r.controller.asking, false, 'nothing is asked');
  assert.equal(await dismissed.first, false, 'cancelled');
  assert.equal(dismissed.r.controller.dirty, true, 'the file stays unsaved');
  dismissed.r.controller.dismissConflict();
  assert.equal(dismissed.r.asked[0].closed, 1, 'dismissing twice closes once');

  const gone = await conflicted();
  const waiting = gone.r.controller.saveNow();
  gone.r.controller.destroy();
  assert.equal(gone.r.asked[0].closed, 1, 'destroyed: the question is taken away');
  assert.equal(await gone.first, false, 'the save is answered');
  assert.equal(await waiting, false, 'and so is the call that waited for the question');
  gone.r.asked[0].question.answer(true);
  await settle();
  assert.equal(gone.r.sent.length, 1, 'nothing is sent for a view that is gone');
});

test('a save that was asked to overwrite while the refused one was on its way is not asked about', async () => {
  const r = rig(request());
  r.controller.changed();
  const first = r.controller.saveNow();
  await settle();
  const over = r.controller.saveNow({ overwrite: true });
  r.sent[0].answer(412);
  await settle();
  assert.equal(r.asked.length, 0, 'the overwrite was already asked for');
  assert.equal(r.sent.length, 2, 'it follows the refused save');
  r.sent[1].answer(200);
  assert.equal(await over, true, 'the overwrite succeeded');
  assert.equal(await first, false, 'the refused save failed');
});

test('saveTo.onConflict is "ask" unless the page says "fail", and anything else is refused', () => {
  const page = 'https://app.example.com/editor/';
  const of = (saveTo: unknown) => resolveSaveTarget(saveTo, '/files/7', page) as Extract<SaveTarget, { kind: 'request' }>;
  assert.equal(of(true).onConflict, 'ask', 'true asks');
  assert.equal(of('/files/8').onConflict, 'ask', 'an address asks');
  assert.equal(of({ method: 'POST' }).onConflict, 'ask', 'a request asks by default');
  assert.equal(of({ onConflict: 'ask' }).onConflict, 'ask', 'a request that says so asks');
  assert.equal(of({ onConflict: 'fail' }).onConflict, 'fail', 'a request that says "fail" does not');
  assert.throws(() => of({ onConflict: 'overwrite' }), TypeError, 'no value lets a save overwrite without being asked');
  assert.throws(() => of({ onConflict: true }), TypeError, 'a value that is no text is refused');
});

test('the question has its words in both languages, and they say that the file on the server is overwritten', () => {
  for (const strings of [STRINGS_EN, STRINGS_TR]) {
    assert.equal(typeof strings.saveConflictTitle, 'string', 'a title');
    assert.equal(strings.saveConflictTitle.length > 0, true, 'the title is not empty');
    assert.equal(strings.saveConflictQuestion.length > 40, true, 'the question explains');
    assert.equal(strings.saveOverwrite.length > 0 && strings.saveCancel.length > 0, true, 'the two answers have names');
    assert.notEqual(strings.saveOverwrite, strings.saveCancel, 'which differ');
  }
  assert.match(STRINGS_TR.saveConflictQuestion, /üzerine yazılır/, 'Turkish: says that the changes there are written over');
  assert.match(STRINGS_TR.saveConflictQuestion, /başkası değiştirmiş ya da önceki kaydınız yanıtı gelmeden yazılmış olabilir/, 'Turkish: keeps the two ways it can have come about');
  assert.match(STRINGS_EN.saveConflictQuestion, /someone else changed it, or an earlier save of yours was written although its answer did not arrive/, 'English: keeps the two ways it can have come about');
  assert.equal(STRINGS_TR.saveOverwrite, 'Üzerine yaz', 'Turkish: the name of the overwrite');
  assert.equal(STRINGS_TR.saveCancel, 'Vazgeç', 'Turkish: the name of the cancel');
});

test('an overwrite that waited behind a save the server took is not carried on: the next save names the file the server holds now', async () => {
  const r = rig(request());
  r.controller.changed();
  const first = r.controller.saveNow();
  await settle();
  assert.equal(ifMatchOf(r.sent[0]), '"v1"', 'the save under way is guarded');
  // Asked for while that save is on its way, for a conflict that then did not happen.
  const over = r.controller.saveNow({ overwrite: true });
  r.controller.changed();
  r.content = [4, 5];
  r.sent[0].answer(200, { ETag: '"v2"' });
  await settle();
  assert.equal(await first, true, 'the first save was taken');
  assert.equal(r.sent.length, 2, 'what was changed meanwhile is sent');
  assert.equal(ifMatchOf(r.sent[1]), '"v2"', 'guarded by the ETag the server just gave, not sent over whatever it holds');
  r.sent[1].answer(412);
  await settle();
  assert.equal(r.asked.length, 1, 'the user is asked about the conflict of that save');
  assert.equal(r.sent.length, 2, 'and nothing is sent over the file someone changed in between');
  r.asked[0].question.answer(false);
  assert.equal(await over, false, 'the call that asked to overwrite ends as the question does');
  assert.equal(r.sent.length, 2, 'nothing more was sent');
});

test('the answer "overwrite" covers the request it leads to and no later one: a save the page starts from inside commit() is guarded after it', async () => {
  let reenter = false;
  let r: Rig | undefined;
  const made = await conflicted(request(), {
    commit: () => {
      if (reenter && r) {
        reenter = false;
        void r.controller.saveNow();
      }
      return true;
    },
  });
  r = made.r;
  reenter = true;
  r.asked[0].question.answer(true);
  await settle();
  assert.equal(r.sent.length, 2, 'one request is on its way');
  // The page's own save went first, guarded; the server takes it this time and names a new file.
  assert.equal(ifMatchOf(r.sent[1]), '"v1"', 'the save the page asked for is guarded');
  r.controller.changed();
  r.sent[1].answer(200, { ETag: '"v2"' });
  await settle();
  assert.equal(r.sent.length, 3, 'the change made meanwhile is sent');
  assert.equal(ifMatchOf(r.sent[2]), '"v2"', 'with the ETag of the file the server holds now: the answer was to the conflict before it');
  r.sent[2].answer(204);
  await settle();
  assert.equal(await made.first, true, 'the call that led to the question hears that the file is saved');
});
