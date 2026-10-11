import assert from 'node:assert/strict';
import { test } from 'node:test';
import { SaveController, resolveSaveTarget, type SaveNowOptions, type SaveState } from '../src/embed/save-to';
import { STRINGS_EN } from '../src/ui/strings';

// Random sequences of changes, saves, overwrites and answers of a server that
// keeps an ETag, with a page that calls back into the controller from inside
// each of its callbacks. Whatever the order, the controller keeps to this:
//
// - never two requests on their way at once;
// - "nothing unsaved" only when the last 2xx answer was to the content as it is now;
// - content that is older never goes over content that is newer;
// - "saving" only while a request is on its way, and every saveNow() is answered.
//
// The sequences come from a seed: one that fails is named in the message and runs again as it was.

type Reenter = 'none' | 'onState' | 'onError' | 'onSaved' | 'notice' | 'strings' | 'commit' | 'bytes' | 'info';
const MODES: Reenter[] = ['none', 'onState', 'onError', 'onSaved', 'notice', 'strings', 'commit', 'bytes', 'info'];
const SEEDS = 120;
const STEPS = 60;

function generator(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

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

async function run(seed: number, reenter: Reenter): Promise<string[]> {
  const random = generator(seed * 7919 + MODES.indexOf(reenter));
  /** The content of the file now. */
  let version = 0;
  const server = { body: 0, etag: '"e0"', serial: 0, writes: [] as number[] };
  const flights: Flight[] = [];
  const timers: { callback: () => void; off: boolean }[] = [];
  const problems: string[] = [];
  const promises: { settled: boolean }[] = [];
  let last2xxBody = 0;
  let writable = true;
  let stuck = false;
  let told: SaveState | undefined;
  let depth = 0;
  let controller: SaveController;

  const call = (options?: SaveNowOptions): void => {
    const entry = { settled: false };
    promises.push(entry);
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
      back('bytes');
      return new Uint8Array([taken & 255, taken >> 8]);
    },
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
    timer: (callback) => {
      const entry = { callback, off: false };
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
    flight.resolve({ status, headers: { get: (name: string) => (name.toLowerCase() === 'etag' && etag !== undefined ? etag : null) }, ...more });
  };
  const timeOut = (): void => {
    const timer = timers.filter((entry) => !entry.off).pop();
    if (!timer) return;
    timer.off = true;
    timer.callback();
  };

  const trace: string[] = [];
  const check = (step: string): void => {
    const state = controller.state;
    if (!controller.dirty && last2xxBody !== version) problems.push(`${step}: nothing unsaved, but the last 2xx answer was to content ${last2xxBody} and the content is ${version}`);
    if (!controller.dirty && server.body !== version) problems.push(`${step}: nothing unsaved, but the server holds ${server.body} and the content is ${version}`);
    if (controller.saving && open().length === 0) problems.push(`${step}: saving with no request on its way`);
    if (!controller.saving && open().length > 0) problems.push(`${step}: a request on its way and not saving`);
    if (told === undefined ? state.dirty || state.saving || state.error !== undefined : told.dirty !== state.dirty || told.saving !== state.saving || told.error !== state.error) {
      problems.push(`${step}: the page was last told ${JSON.stringify(told)} and the state is ${JSON.stringify(state)}`);
    }
  };

  for (let step = 0; step < STEPS; step++) {
    const pick = random();
    const flight = open()[0];
    let what = 'nothing';
    let saved = false;
    if (pick < 0.25) {
      version++;
      controller.changed();
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

  // Everything that is still on its way is answered.
  writable = true;
  stuck = false;
  for (let i = 0; i < 50 && problems.length === 0; i++) {
    const flight = open()[0];
    if (!flight) break;
    const taken = take(flight);
    respond(flight, taken.status, taken.etag);
    if (taken.status === 200) last2xxBody = flight.version;
    await settle();
    check(`end ${i}`);
  }
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
    for (let seed = 1; seed <= SEEDS; seed++) {
      const problems = await run(seed, reenter);
      assert.deepEqual(problems, [], problems.join('\n'));
      requests++;
    }
    assert.equal(requests, SEEDS, 'every sequence ran');
  });
}
