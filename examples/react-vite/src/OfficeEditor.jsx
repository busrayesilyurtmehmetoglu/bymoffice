import { createBlank, mount } from 'bymoffice';
import { useEffect, useRef, useState } from 'react';

/** Shows `file` (an address) in the editor, or a new file of `kind` ('docx' or 'xlsx') when there is none. */
export default function OfficeEditor({ file, kind }) {
  const host = useRef(null);
  const [error, setError] = useState('');

  useEffect(() => {
    // mount() is asynchronous: the component can be gone before it resolves. In development <StrictMode> runs
    // every effect twice (mount, clean up, mount), so the first editor arrives after its cleanup and is
    // destroyed then; exactly one editor stays.
    let view = null;
    let gone = false;
    setError('');
    (async () => {
      const mounted = await mount(host.current, {
        file: file || (await createBlank(kind)),
        // Where vite.config.js put the font folder of the package: next to the page, under the base of the app.
        fontsUrl: import.meta.env.BASE_URL + 'bymoffice-fonts/',
      });
      if (gone) return mounted.destroy();
      view = mounted;
      window.bymEditor = mounted; // to try it out in the browser console
    })().catch((e) => {
      if (!gone) setError(e instanceof Error ? e.message : String(e));
    });
    return () => {
      gone = true;
      if (view) view.destroy();
      if (window.bymEditor === view) window.bymEditor = undefined;
    };
  }, [file, kind]);

  return (
    <>
      {error && <p className="error" role="alert">{error}</p>}
      <div ref={host} className="editor" />
    </>
  );
}
