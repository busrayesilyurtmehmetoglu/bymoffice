import assert from 'node:assert/strict';
import { test } from 'node:test';
import { SaveController, resolveSaveTarget, type ConflictQuestion, type SaveNowOptions, type SaveState } from '../src/embed/save-to';
import { STRINGS_EN } from '../src/ui/strings';

// Random sequences of changes, saves, overwrites and answers of a server that
// keeps an ETag, with a page that calls back into the controller from inside
// each of its callbacks. Whatever the order, the controller keeps to this:
//
// - never two requests on their way at once;
// - "nothing unsaved" only when the last 2xx answer was to the content as it is now;
// - content that is older never goes over content that is newer;
// - "saving" only while a request is on its way, and every saveNow() is answered;
// - a conflict asks at most one question at a time, nothing is on its way while it is open, every
//   question is taken away exactly once, and what the controller says of it is what the host was asked;
// - while the ETag of the file is known, a request goes without If-Match only when saveNow({ overwrite: true })
//   was called, or a question that was open was answered with "overwrite", since the server last took a save.
//
// The same sequences run once more with auto-save on, with time that passes, something that is being
// typed, a view that turns read-only and a page that is hidden among the steps. Everything above holds
// there as well, and with it:
//
// - auto-save never writes the file while something is being typed, while the user is asked about a
//   conflict, from a view that is only read, for a file that cannot be written or while it is off;
// - after a conflict, a refusal or a redirect nothing is sent until the user has acted (a save, an
//   overwrite, an answer): auto-save does not try again by itself;
// - a request auto-save sends always carries If-Match: where the server named no version of the file
//   (its answers come without an ETag, or with a weak one, in these sequences) auto-save sends nothing;
// - auto-save never falls asleep: with changes that no save holds and nothing in its way (no save under
//   way, no question, not paused, a view that can be edited) a timer of its own is waiting, and never two.
//
// The sequences come from a seed: one that fails is named in the message and runs again as it was.

type Reenter = 'none' | 'onState' | 'onError' | 'onSaved' | 'notice' | 'strings' | 'commit' | 'bytes' | 'info' | 'ask';
const MODES: Reenter[] = ['none', 'onState', 'onError', 'onSaved', 'notice', 'strings', 'commit', 'bytes', 'info', 'ask'];
const SEEDS = 120;
const STEPS = 60;

function generator(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

/** How many questions the sequences of a test led to, and how many of them were answered with "overwrite": the test is one of them only if there were some. */
const seen = { questions: 0, overwrites: 0, autoWrites: 0, deferred: 0, unguarded: 0, paused: 0 };
/** The time limit of a request in these runs: no wait of auto-save is ever this long, which tells the two kinds of timer apart. */
const REQUEST_TIMEOUT = 123457;

const settle = async () => {
  for (let i = 0; i < 60; i++) await Promise.resolve();
};

interface Flight {
  version: number;
  ifMatch: string | undefined;
  resolve(response: unknown): void;
  reject(error: unknown): void;
  done: boolean;
}

async function run(seed: number, reenter: Reenter, auto = false): Promise<string[]> {
  const random = generator(seed * 7919 + MODES.indexOf(reenter) + (auto ? 104729 : 0));
  let clock = 5000;
  /** Something is being typed; the view is only read; auto-save is on. */
  let typing = false;
  let readOnly = false;
  let autoOn = auto;
  /** A save ended in a way auto-save must not try again by itself, and the user has not acted since. */
  let halted = false;
  /** The file was last written for a save nobody asked for: the request that follows is auto-save's. */
  let writtenByAuto = false;
  /** The file became writable again and nothing told the controller (nothing does: a change is what looks again). */
  let unwoken = false;
  /** The content of the file now. */
  let version = 0;
  const server = { body: 0, etag: '"e0"', serial: 0, writes: [] as number[] };
  const flights: Flight[] = [];
  const timers: { callback: () => void; off: boolean; at: number; request: boolean }[] = [];
  const problems: string[] = [];
  const promises: { settled: boolean }[] = [];
  let last2xxBody = 0;
  let writable = true;
  let stuck = false;
  let told: SaveState | undefined;
  let depth = 0;
  let controller: SaveController;
  /** The questions about a conflict that were asked: how often each was taken away. */
  const questions: { question: ConflictQuestion; closed: number }[] = [];
  const asked = () => questions.filter((entry) => entry.closed === 0);
  /**
   * How often the file was allowed to go over whatever the server holds (saveNow({ overwrite: true }), or
   * "overwrite" as the answer to a question that was open) since the server last took a save: what was
   * allowed is the overwriting of the file the server held then, not of the one a later save put there.
   */
  let allowed = 0;

  const call = (options?: SaveNowOptions): void => {
    const entry = { settled: false };
    promises.push(entry);
    if (options && options.overwrite === true) allowed++;
    controller.saveNow(options).then(
      () => void (entry.settled = true),
      () => {
        entry.settled = true;
        problems.push('saveNow() rejected');
      },
    );
  };
  /** What a page does from inside a callback: asks for a save, for an overwrite, or changes the file. */
  const back = (from: Reenter): void => {
    if (reenter !== from || depth > 0) return;
    depth++;
    try {
      const pick = random();
      if (pick < 0.5) call();
      else if (pick < 0.7) call({ overwrite: true });
      else if (pick < 0.8) {
        version++;
        controller.changed();
      }
    } finally {
      depth--;
    }
  };
  const open = (): Flight[] => flights.filter((flight) => !flight.done);

  controller = new SaveController({
    target: resolveSaveTarget(true, 'http://files.example/f.xlsx', undefined),
    etag: '"e0"',
    writable: () => writable,
    commit: () => {
      back('commit');
      return !stuck;
    },
    bytes: async () => {
      const taken = version;
      // A save of the user: whatever kept auto-save still is over, and what comes of this save decides anew.
      halted = false;
      writtenByAuto = false;
      back('bytes');
      return new Uint8Array([taken & 255, taken >> 8]);
    },
    // What auto-save writes the file with: every call is a save nobody asked for.
    bytesAsIs: async () => {
      const taken = version;
      seen.autoWrites++;
      writtenByAuto = true;
      if (controller.currentETag === undefined) seen.unguarded++;
      if (!autoOn) problems.push('auto-save wrote the file while it was off');
      if (typing) problems.push('auto-save wrote the file while something was being typed');
      if (readOnly) problems.push('auto-save wrote the file from a view that is only read');
      if (!writable) problems.push('auto-save wrote a file that cannot be written');
      if (asked().length > 0) problems.push('auto-save wrote the file while the user was asked about a conflict');
      if (halted) problems.push('auto-save wrote the file after a conflict or a refusal, and the user had not acted since');
      back('bytes');
      return new Uint8Array([taken & 255, taken >> 8]);
    },
    pending: () => {
      if (typing) seen.deferred++;
      return typing;
    },
    readOnly: () => readOnly,
    autoSave: auto ? { delayMs: 1000, maxWaitMs: 4000 } : undefined,
    now: () => clock,
    timeoutMs: REQUEST_TIMEOUT,
    info: () => {
      back('info');
      return { format: 'xlsx' };
    },
    strings: () => {
      back('strings');
      return STRINGS_EN;
    },
    notice: () => back('notice'),
    onState: (state) => {
      told = state;
      back('onState');
    },
    onError: () => back('onError'),
    onSaved: () => back('onSaved'),
    askOverwrite: (question) => {
      const entry = { question, closed: 0 };
      questions.push(entry);
      seen.questions++;
      if (asked().length > 1) problems.push('two questions about a conflict open at once');
      if (open().length > 0) problems.push('a question was asked while a request was on its way');
      back('ask');
      return () => void entry.closed++;
    },
    timer: (callback, ms) => {
      const entry = { callback, off: false, at: clock + ms, request: ms === REQUEST_TIMEOUT };
      timers.push(entry);
      return () => void (entry.off = true);
    },
    fetch: ((_url: string, init: { body: Uint8Array; headers: { [name: string]: string }; signal?: AbortSignal }) =>
      new Promise((resolve, reject) => {
        const flight: Flight = { version: init.body[0] + init.body[1] * 256, ifMatch: init.headers['If-Match'], resolve, reject, done: false };
        if (init.signal) {
          init.signal.addEventListener('abort', () => {
            flight.done = true;
            reject(new Error('aborted'));
          });
        }
        flights.push(flight);
        if (writtenByAuto && flight.ifMatch === undefined) problems.push(`auto-save sent content ${flight.version} without If-Match: it could overwrite what someone else saved, unasked`);
        if (flight.ifMatch === undefined && controller.currentETag !== undefined && allowed === 0) {
          problems.push(`content ${flight.version} was sent without If-Match although the ETag ${controller.currentETag} is known and nobody allowed an overwrite since the server last took a save`);
        }
        if (open().length > 1) problems.push(`two requests on their way at once (content ${open().map((other) => other.version).join(' and ')})`);
      })) as unknown as typeof fetch,
  });

  /** The server takes the request: refused when its If-Match names another file than the one it holds. */
  const take = (flight: Flight): { status: number; etag?: string } => {
    if (flight.ifMatch !== undefined && flight.ifMatch !== server.etag) return { status: 412 };
    const newest = server.writes.length ? server.writes[server.writes.length - 1] : 0;
    if (flight.version < newest) problems.push(`older content ${flight.version} was written over ${newest}`);
    server.body = flight.version;
    server.writes.push(flight.version);
    server.etag = `"e${++server.serial}"`;
    return { status: 200, etag: server.etag };
  };
  const respond = (flight: Flight, status: number, etag?: string | null, more: { [name: string]: unknown } = {}): void => {
    flight.done = true;
    // What trying again does not mend: auto-save holds still until the user acts.
    if (status === 412 || status === 401 || status === 423 || status === 302 || status === 0 || more.url !== undefined) halted = true;
    // A save the server took and said so: what was allowed before it is used up, whether it was this request or not.
    if (status >= 200 && status < 300 && more.url === undefined) allowed = 0;
    flight.resolve({ status, headers: { get: (name: string) => (name.toLowerCase() === 'etag' && etag !== undefined ? etag : null) }, ...more });
  };
  const timeOut = (): void => {
    const timer = timers.filter((entry) => !entry.off && entry.request).pop();
    if (!timer) return;
    timer.off = true;
    timer.callback();
  };
  /** Time passes: the timers of auto-save that come due run, the earliest first. */
  const pass = async (ms: number): Promise<void> => {
    const end = clock + ms;
    for (let guard = 0; guard < 200; guard++) {
      let next: (typeof timers)[number] | undefined;
      for (const entry of timers) if (!entry.off && !entry.request && entry.at <= end && (!next || entry.at < next.at)) next = entry;
      if (!next) break;
      clock = Math.max(clock, next.at);
      next.off = true;
      next.callback();
      await settle();
    }
    clock = end;
  };

  const trace: string[] = [];
  const check = (step: string): void => {
    const state = controller.state;
    if (!controller.dirty && last2xxBody !== version) problems.push(`${step}: nothing unsaved, but the last 2xx answer was to content ${last2xxBody} and the content is ${version}`);
    if (!controller.dirty && server.body !== version) problems.push(`${step}: nothing unsaved, but the server holds ${server.body} and the content is ${version}`);
    if (controller.saving && open().length === 0) problems.push(`${step}: saving with no request on its way`);
    if (!controller.saving && open().length > 0) problems.push(`${step}: a request on its way and not saving`);
    if (controller.asking && open().length > 0) problems.push(`${step}: a request on its way while the user is asked about a conflict`);
    // The state says that auto-save waits exactly when it is on and has nothing to guard a save with, or was stopped.
    const waits = state.auto === true && (controller.currentETag === undefined || controller.autoSaveStopped);
    if ((state.autoPaused === true) !== waits) problems.push(`${step}: the state says autoPaused ${state.autoPaused}; auto-save is ${state.auto ? 'on' : 'off'}, the ETag is ${controller.currentETag}, stopped ${controller.autoSaveStopped}`);
    if (state.autoPaused === true && controller.currentETag === undefined) seen.paused++;
    const waiting = timers.filter((entry) => !entry.off && !entry.request).length;
    if (waiting > 1) problems.push(`${step}: ${waiting} timers of auto-save wait at once`);
    if (!autoOn && waiting > 0) problems.push(`${step}: a timer of auto-save waits while it is off`);
    if (autoOn && state.dirty && !state.saving && !controller.asking && state.autoPaused !== true && writable && !readOnly && !unwoken && waiting === 0) {
      problems.push(`${step}: auto-save is on, there are changes and nothing is in its way, and no timer of its waits`);
    }
    if (controller.asking && !controller.dirty) problems.push(`${step}: the user is asked about a conflict of a file with nothing unsaved`);
    if (controller.asking !== (asked().length === 1)) problems.push(`${step}: the controller says asking ${controller.asking}, and ${asked().length} question(s) were not taken away`);
    if (questions.some((entry) => entry.closed > 1)) problems.push(`${step}: a question was taken away more than once`);
    if (told === undefined ? state.dirty || state.saving || state.error !== undefined : told.dirty !== state.dirty || told.saving !== state.saving || told.error !== state.error) {
      problems.push(`${step}: the page was last told ${JSON.stringify(told)} and the state is ${JSON.stringify(state)}`);
    }
  };

  for (let step = 0; step < STEPS; step++) {
    const pick = random();
    const flight = open()[0];
    let what = 'nothing';
    let saved = false;
    const extra = auto ? random() : 1;
    if (extra < 0.3) {
      // What only happens with auto-save: time, typing, a view that is only read, a hidden page, the switch.
      const which = random();
      if (which < 0.6) {
        const ms = 700 * (1 + Math.floor(random() * 10));
        await pass(ms);
        what = `${ms} ms pass`;
      } else if (which < 0.75) {
        typing = !typing;
        what = `typing ${typing}`;
      } else if (which < 0.83) {
        readOnly = !readOnly;
        if (!readOnly) controller.wake();
        what = `read-only ${readOnly}`;
      } else if (which < 0.93) {
        controller.pageHidden();
        what = 'page hidden';
      } else {
        autoOn = !autoOn;
        controller.setAutoSave(autoOn ? { delayMs: 1000, maxWaitMs: 4000 } : false);
        what = `auto-save ${autoOn}`;
      }
    } else if (asked().length > 0 && random() < 0.4) {
      // The user answers the question, or an answer comes for one that is gone.
      const target = random() < 0.9 ? asked()[0] : questions[0];
      const overwrite = random() < 0.5;
      what = `answer ${overwrite ? 'overwrite' : 'cancel'}${target.closed > 0 ? ' (late)' : ''}`;
      if (overwrite && target.closed === 0) {
        seen.overwrites++;
        allowed++;
      }
      target.question.answer(overwrite);
    } else if (pick < 0.25) {
      version++;
      controller.changed();
      unwoken = false;
      what = 'change';
    } else if (pick < 0.4) {
      call();
      what = 'saveNow';
    } else if (pick < 0.48) {
      call({ overwrite: true });
      what = 'overwrite';
    } else if (pick < 0.85 && flight) {
      const answer = random();
      if (answer < 0.4) {
        const taken = take(flight);
        const kind = random();
        // The answer carries the new ETag, none, or a weak one.
        respond(flight, taken.status, taken.status === 200 ? (kind < 0.6 ? taken.etag : kind < 0.8 ? null : 'W/"w"') : null);
        if (taken.status === 200) {
          last2xxBody = flight.version;
          saved = true;
        }
        what = `answer ${taken.status}`;
      } else if (answer < 0.5) {
        respond(flight, 412);
        what = '412';
      } else if (answer < 0.55) {
        respond(flight, 423);
        what = '423';
      } else if (answer < 0.6) {
        respond(flight, 401);
        what = '401';
      } else if (answer < 0.68) {
        respond(flight, 500);
        what = '500';
      } else if (answer < 0.72) {
        respond(flight, 302);
        what = '302';
      } else if (answer < 0.76) {
        respond(flight, 0, null, { type: 'opaqueredirect' });
        what = 'opaque redirect';
      } else if (answer < 0.8) {
        respond(flight, 200, '"login"', { url: 'http://files.example/login' });
        what = '200 from another address';
      } else if (answer < 0.87) {
        flight.done = true;
        flight.reject(new TypeError('network'));
        what = 'connection lost';
      } else if (answer < 0.94) {
        // The server wrote the file, and its answer never came.
        take(flight);
        timeOut();
        what = 'written, then timeout';
      } else {
        timeOut();
        what = 'timeout';
      }
    } else if (pick < 0.88) {
      writable = !writable;
      if (writable) unwoken = true;
      what = `writable ${writable}`;
    } else if (pick < 0.93) {
      stuck = !stuck;
      what = `stuck ${stuck}`;
    }
    await settle();
    trace.push(`${step} ${what}: content ${version}, ${JSON.stringify(controller.state)}, server ${server.body} ${server.etag}, held ETag ${controller.currentETag}, on their way ${open().length}`);
    check(`step ${step} (${what})`);
    // A save that succeeded and holds everything leaves no "could not be saved" behind.
    if (saved && !stuck && reenter === 'none' && !controller.dirty && controller.state.error !== undefined) problems.push(`step ${step}: saved, and the state still says "${controller.state.error}"`);
    if (problems.length > 0) break;
  }

  // Everything that is still on its way is answered, and nothing new starts by itself.
  if (auto && problems.length === 0) {
    autoOn = false;
    controller.setAutoSave(false);
    await settle();
  }
  writable = true;
  stuck = false;
  for (let i = 0; i < 50 && problems.length === 0; i++) {
    const flight = open()[0];
    const waiting = asked()[0];
    if (!flight && !waiting) break;
    if (flight) {
      const taken = take(flight);
      respond(flight, taken.status, taken.etag);
      if (taken.status === 200) last2xxBody = flight.version;
    } else waiting.question.answer(false);
    await settle();
    check(`end ${i}`);
  }
  if (problems.length === 0 && asked().length > 0) problems.push('a question about a conflict is open for ever at the end');
  if (open().length > 0) problems.push('requests go on without end');
  else {
    if (controller.saving) problems.push('saving for ever at the end');
    const waiting = promises.filter((entry) => !entry.settled).length;
    if (waiting > 0) problems.push(`${waiting} saveNow() call(s) were never answered`);
  }
  return problems.length > 0 ? [`seed ${seed}, calling back from ${reenter}:`, ...problems.slice(0, 4), ...trace.slice(-8)] : [];
}

for (const reenter of MODES) {
  test(`random saves, calling back into the controller from ${reenter}: one request at a time, and never "saved" for what the server does not hold`, async () => {
    let requests = 0;
    seen.questions = 0;
    seen.overwrites = 0;
    for (let seed = 1; seed <= SEEDS; seed++) {
      const problems = await run(seed, reenter);
      assert.deepEqual(problems, [], problems.join('\n'));
      requests++;
    }
    assert.equal(requests, SEEDS, 'every sequence ran');
    assert.equal(seen.questions >= 8 && seen.overwrites >= 2, true, `the sequences ran into conflicts that were asked about and overwritten (${seen.questions} questions, ${seen.overwrites} overwrites)`);
  });

  test(`random saves with auto-save, calling back into the controller from ${reenter}: nothing is sent over what is being typed, nor by itself after a conflict`, async () => {
    seen.questions = 0;
    seen.overwrites = 0;
    seen.autoWrites = 0;
    seen.deferred = 0;
    seen.unguarded = 0;
    seen.paused = 0;
    for (let seed = 1; seed <= SEEDS; seed++) {
      const problems = await run(seed, reenter, true);
      assert.deepEqual(problems, [], problems.join('\n'));
    }
    assert.equal(seen.autoWrites >= 25, true, `auto-save wrote the file often enough to be tested (${seen.autoWrites} times)`);
    assert.equal(seen.deferred >= 20, true, `auto-save met something that was being typed often enough (${seen.deferred} times)`);
    assert.equal(seen.questions >= 4, true, `the sequences ran into conflicts that were asked about (${seen.questions} questions)`);
    assert.equal(seen.unguarded, 0, `auto-save never wrote the file with no ETag to guard it (${seen.unguarded} times)`);
    assert.equal(seen.paused >= 50, true, `the server left auto-save without an ETag often enough (${seen.paused} steps)`);
  });
}
