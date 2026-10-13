// Who is in a room, as a list to show: this participant first, then the others in the order
// the relay named them. Nothing here needs a page.

import type { CollabPeer, CollabState } from './session';
import { cleanName } from './protocol';

/** One participant of a room, as shown. */
export interface CollabPerson {
  /** The name of the connection in the room (the relay's): two tabs of one user are two of them. */
  id: string;
  /** The name the participant gave (or the relay gave it), or "Guest N" in the language of the view. What a participant says of itself: no proof of who it is. */
  name: string;
  /** "#RRGGBB": the colour its selection is drawn in. */
  color: string;
  /** This participant itself. */
  self: boolean;
  readOnly: boolean;
  /** The sheet it is on (0: the first), when it has said where it is. */
  sheet?: number;
}

/** Most participants that are listed; the rest is counted. */
export const MAX_PEOPLE = 200;

export interface People {
  list: CollabPerson[];
  /** How many more are in the room than are listed. */
  more: number;
  /** Everybody in the room, this participant included. */
  count: number;
}

function person(peer: CollabPeer, self: boolean, guest: (n: number) => string): CollabPerson {
  const given = peer.name !== undefined ? cleanName(peer.name) : '';
  const out: CollabPerson = { id: peer.id, name: given !== '' ? given : guest(peer.guest), color: /^#[0-9A-Fa-f]{6}$/.test(peer.color) ? peer.color : '#5F6368', self, readOnly: peer.readOnly === true };
  if (peer.presence && Number.isSafeInteger(peer.presence.sheet) && peer.presence.sheet >= 0) out.sheet = peer.presence.sheet;
  return out;
}

/**
 * The list. Out of the room (offline, detached, closed, not in yet) nobody else is known to be
 * there: the list is this participant alone. An entry of the others that carries this
 * participant's own name of the room, or one that is there twice, is listed once.
 */
export function listPeople(self: CollabPeer, others: readonly CollabPeer[], state: CollabState, guest: (n: number) => string, limit = MAX_PEOPLE): People {
  const list: CollabPerson[] = [person(self, true, guest)];
  const seen = new Set<string>([self.id]);
  let count = 1;
  if (state === 'online') {
    for (const peer of others) {
      if (seen.has(peer.id)) continue;
      seen.add(peer.id);
      count++;
      if (list.length < Math.max(1, limit)) list.push(person(peer, false, guest));
    }
  }
  return { list, more: count - list.length, count };
}

/** The letter (or two) a participant is shown by in a small circle. */
export function initials(name: string): string {
  const words = name.split(' ').filter((word) => word !== '');
  if (words.length === 0) return '?';
  const first = Array.from(words[0])[0] || '?';
  const last = words.length > 1 ? Array.from(words[words.length - 1])[0] || '' : '';
  return (first + last).toUpperCase();
}
