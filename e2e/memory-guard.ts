// Stops a run of the browser tests before it takes the machine down with it.
//
// Three browsers at a time have exhausted the memory of a development machine
// more than once: the system then kills thousands of processes, the window
// server among them, and the desktop session is gone. The run watches how
// much memory the system says is left and ends itself, browsers and all,
// while there still is some.

import { execFile } from 'node:child_process';

/** Percent of memory the system counts as available below which the run is stopped. */
const FLOOR = Number(process.env.E2E_MEMORY_FLOOR ?? 15);
const EVERY_MS = 3000;
/** Readings below the floor in a row before the run is stopped: one low reading is not a trend. */
const READINGS = 2;

/** What the system counts as available, in percent; undefined where it cannot be asked. */
function available(): Promise<number | undefined> {
  return new Promise((resolve) => {
    if (process.platform !== 'darwin') return resolve(undefined);
    execFile('/usr/sbin/sysctl', ['-n', 'kern.memorystatus_level'], { timeout: 2000 }, (error, stdout) => {
      const value = error ? NaN : Number(String(stdout).trim());
      resolve(value >= 0 && value <= 100 ? value : undefined);
    });
  });
}

export default async function memoryGuard(): Promise<() => void> {
  const first = await available();
  if (first === undefined || !(FLOOR > 0)) return () => undefined;
  if (first < FLOOR) throw new Error(`The browser tests were not started: ${first}% of the memory is available, the floor is ${FLOOR}% (E2E_MEMORY_FLOOR). Close something and run them again.`);
  let low = 0;
  let stopped = false;
  const timer = setInterval(() => {
    void available().then((now) => {
      if (stopped || now === undefined) return;
      low = now < FLOOR ? low + 1 : 0;
      if (low < READINGS) return;
      stopped = true;
      console.error(`\nThe browser tests were stopped: ${now}% of the memory is available, the floor is ${FLOOR}% (E2E_MEMORY_FLOOR). The run did not finish and proves nothing.\n`);
      process.exitCode = 1;
      // The runner closes its browsers on an interrupt; the group is asked, so that the server and the workers go too.
      try {
        process.kill(process.pid, 'SIGINT');
      } catch {
        process.exit(1);
      }
    });
  }, EVERY_MS);
  timer.unref();
  return () => clearInterval(timer);
}
