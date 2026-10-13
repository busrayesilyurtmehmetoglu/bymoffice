// The transports that come with the engine: a WebSocket that connects again by itself, and a
// channel between the tabs of one browser (for tests and demos; one tab holds the room, see
// MemoryServer.serve()). An application with a channel of its own (a broadcast service, a
// message queue) writes the four methods of CollabTransport itself.

import { channelName, type ChannelLike } from './server';
import type { CollabTransport } from './session';

/** A transport that can be cut and connected again by hand, as a network that comes and goes does it. */
export interface PausableTransport extends CollabTransport {
  suspend(): void;
  resume(): void;
}

/** What the transport needs of a WebSocket. */
interface SocketLike {
  readyState: number;
  send(data: string): void;
  close(): void;
  onopen: ((event: unknown) => void) | null;
  onclose: ((event: unknown) => void) | null;
  onerror: ((event: unknown) => void) | null;
  onmessage: ((event: { data: unknown }) => void) | null;
}

export interface SocketTransportOptions {
  /** Default: the global WebSocket. */
  WebSocket?: new (url: string) => SocketLike;
  /** The first wait before connecting again, which doubles up to `maxDelayMs`. Defaults 500 and 30000. */
  minDelayMs?: number;
  maxDelayMs?: number;
  /** How often a ping is sent, and how long nothing may arrive before the connection is taken to be dead. Defaults 25000 and 60000. */
  pingMs?: number;
  idleMs?: number;
  /** Default: Math.random. */
  random?(): number;
}

const OPEN = 1;

/** Whether an address is one a WebSocket connects to. */
export function isSocketUrl(value: unknown): value is string {
  return typeof value === 'string' && /^wss?:\/\/[^\s]+$/i.test(value);
}

/**
 * A WebSocket to a relay. A connection that is lost is made again after a wait that grows
 * (with some chance in it, so that not everybody comes back at once), and at once when the
 * browser says the network is back. While there is no connection what is sent is dropped: the
 * session sends what matters again when it has joined.
 */
export function socketTransport(url: string, options: SocketTransportOptions = {}): PausableTransport {
  if (!isSocketUrl(url)) throw new TypeError('The address of a relay starts with "ws://" or "wss://"');
  const Socket = options.WebSocket || ((typeof WebSocket !== 'undefined' ? WebSocket : undefined) as unknown as new (url: string) => SocketLike);
  if (!Socket) throw new TypeError('There is no WebSocket here');
  const minDelay = options.minDelayMs === undefined ? 500 : options.minDelayMs;
  const maxDelay = options.maxDelayMs === undefined ? 30000 : options.maxDelayMs;
  const pingMs = options.pingMs === undefined ? 25000 : options.pingMs;
  const idleMs = options.idleMs === undefined ? 60000 : options.idleMs;
  const random = options.random || Math.random;
  let listener: ((text: string) => void) | undefined;
  let state: ((state: 'open' | 'closed') => void) | undefined;
  let socket: SocketLike | undefined;
  let wanted = false;
  let over = false;
  let attempts = 0;
  let retry: ReturnType<typeof setTimeout> | undefined;
  let beat: ReturnType<typeof setInterval> | undefined;
  let heard = 0;
  let isOpen = false;

  const lost = (from: SocketLike): void => {
    if (socket !== from) return;
    socket = undefined;
    from.onopen = from.onclose = from.onerror = from.onmessage = null;
    try {
      from.close();
    } catch {
      // Closed already.
    }
    if (beat !== undefined) clearInterval(beat);
    beat = undefined;
    if (isOpen) {
      isOpen = false;
      if (state) state('closed');
    }
    if (!wanted || over) return;
    const wait = Math.min(maxDelay, minDelay * Math.pow(2, Math.min(attempts, 16))) * (0.5 + random() * 0.5);
    attempts++;
    retry = setTimeout(connect, wait);
  };

  function connect(): void {
    if (retry !== undefined) clearTimeout(retry);
    retry = undefined;
    if (!wanted || over || socket) return;
    let made: SocketLike;
    try {
      made = new Socket(url);
    } catch {
      retry = setTimeout(connect, maxDelay);
      return;
    }
    socket = made;
    made.onopen = () => {
      if (socket !== made) return;
      attempts = 0;
      heard = Date.now();
      isOpen = true;
      beat = setInterval(() => {
        if (socket !== made) return;
        if (Date.now() - heard > idleMs) return lost(made);
        try {
          made.send('{"t":"ping"}');
        } catch {
          lost(made);
        }
      }, pingMs);
      if (state) state('open');
    };
    made.onmessage = (event) => {
      if (socket !== made) return;
      heard = Date.now();
      // Text only: the contract has no binary messages.
      if (typeof event.data === 'string' && listener) listener(event.data);
    };
    made.onclose = () => lost(made);
    made.onerror = () => lost(made);
  }

  const online = (): void => {
    if (!socket && wanted && !over) {
      attempts = 0;
      connect();
    }
  };
  if (typeof window !== 'undefined' && typeof window.addEventListener === 'function') window.addEventListener('online', online);

  return {
    send: (text) => {
      if (socket && socket.readyState === OPEN) socket.send(text);
    },
    onMessage: (callback) => {
      listener = callback;
    },
    onState: (callback) => {
      state = callback;
      wanted = true;
      connect();
    },
    close: () => {
      over = true;
      wanted = false;
      if (retry !== undefined) clearTimeout(retry);
      retry = undefined;
      if (typeof window !== 'undefined' && typeof window.removeEventListener === 'function') window.removeEventListener('online', online);
      if (socket) lost(socket);
    },
    suspend: () => {
      wanted = false;
      if (retry !== undefined) clearTimeout(retry);
      retry = undefined;
      if (socket) lost(socket);
    },
    resume: () => {
      if (over) return;
      wanted = true;
      attempts = 0;
      connect();
    },
  };
}

export interface BroadcastTransportOptions {
  /** Default: a BroadcastChannel of the name. */
  channel?: ChannelLike;
  /** How often the room is asked for until it answers. Default 300. */
  retryMs?: number;
}

/**
 * A transport to a room that another tab (or this one) of the same browser holds with
 * MemoryServer.serve(name). Until the room answers it is asked for again and again; when the
 * tab that holds it goes, the transport is closed until a room answers again.
 */
/** How often a tab tells the relay of a page that it is still there; three times as long without a word, and the relay lets it go. */
export const TAB_ALIVE_MS = 10000;

export function broadcastTransport(name: string, options: BroadcastTransportOptions = {}): PausableTransport {
  if (!options.channel && typeof BroadcastChannel === 'undefined') throw new TypeError('There is no BroadcastChannel here');
  const port = options.channel || (new BroadcastChannel(channelName(String(name))) as unknown as ChannelLike);
  const id = `${Math.random().toString(36).slice(2)}${Date.now().toString(36)}`;
  let listener: ((text: string) => void) | undefined;
  let state: ((state: 'open' | 'closed') => void) | undefined;
  let isOpen = false;
  let wanted = false;
  let over = false;
  let asking: ReturnType<typeof setInterval> | undefined;

  const ask = (): void => {
    if (asking !== undefined || !wanted || over) return;
    const knock = (): void => port.postMessage({ to: 'relay', from: id, kind: 'open' });
    asking = setInterval(knock, options.retryMs === undefined ? 300 : options.retryMs);
    knock();
  };
  /** While it is open the relay hears of this tab every few seconds: one that is closed, or gone to sleep for good, is let go of there. */
  let beat: ReturnType<typeof setInterval> | undefined;
  const alive = (): void => {
    if (beat !== undefined) return;
    beat = setInterval(() => port.postMessage({ to: 'relay', from: id, kind: 'alive' }), TAB_ALIVE_MS);
    const timer = beat as unknown as { unref?(): void };
    if (typeof timer.unref === 'function') timer.unref();
  };
  // A page that goes says so, where the browser lets it.
  const gone = (): void => {
    if (isOpen) port.postMessage({ to: 'relay', from: id, kind: 'close' });
  };
  const page = typeof window !== 'undefined' && typeof window.addEventListener === 'function' ? window : undefined;
  if (page) page.addEventListener('pagehide', gone);
  const cut = (tell: boolean): void => {
    if (asking !== undefined) clearInterval(asking);
    asking = undefined;
    if (beat !== undefined) clearInterval(beat);
    beat = undefined;
    if (isOpen) {
      isOpen = false;
      if (tell) port.postMessage({ to: 'relay', from: id, kind: 'close' });
      if (state) state('closed');
    }
  };
  port.onmessage = (event) => {
    const data = event.data as { to?: unknown; kind?: unknown; text?: unknown } | null;
    if (!data || data.to !== id || over) return;
    if (data.kind === 'accept') {
      if (!wanted) return;
      if (asking !== undefined) clearInterval(asking);
      asking = undefined;
      // The room answered a second knock: the connection it knows is a new one.
      if (isOpen && state) state('closed');
      isOpen = true;
      alive();
      if (state) state('open');
    } else if (data.kind === 'bye') {
      cut(false);
      ask();
    } else if (data.kind === 'msg' && typeof data.text === 'string' && isOpen && listener) listener(data.text);
  };
  return {
    send: (text) => {
      if (isOpen) port.postMessage({ to: 'relay', from: id, kind: 'msg', text });
    },
    onMessage: (callback) => {
      listener = callback;
    },
    onState: (callback) => {
      state = callback;
      wanted = true;
      ask();
    },
    close: () => {
      cut(true);
      over = true;
      wanted = false;
      if (page) page.removeEventListener('pagehide', gone);
      port.onmessage = null;
      port.close();
    },
    suspend: () => {
      wanted = false;
      cut(true);
    },
    resume: () => {
      if (over) return;
      wanted = true;
      ask();
    },
  };
}
