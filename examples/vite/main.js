import { createBlank, mount } from 'bymoffice';

// What to open: ?file=<address of a docx, xlsx, doc, xls or RTF file>, or a new file (?new=docx, default xlsx).
const params = new URLSearchParams(location.search);
const host = document.getElementById('editor');
const toggle = document.getElementById('toggle');
const error = document.getElementById('error');

let mounted = null;
/** Counts the openings: an opening that was overtaken by a close (or another opening) is thrown away. */
let opening = 0;

async function open() {
  const mine = ++opening;
  error.textContent = '';
  try {
    const file = params.get('file') || (await createBlank(params.get('new') === 'docx' ? 'docx' : 'xlsx'));
    const view = await mount(host, {
      file,
      // Where vite.config.js put the font folder of the package: next to the page, under the base of the app.
      fontsUrl: import.meta.env.BASE_URL + 'bymoffice-fonts/',
    });
    if (mine !== opening) return view.destroy();
    mounted = view;
    window.bymEditor = view; // to try it out in the browser console
  } catch (e) {
    if (mine === opening) error.textContent = e instanceof Error ? e.message : String(e);
  }
}

function close() {
  opening++;
  if (mounted) mounted.destroy();
  mounted = null;
  window.bymEditor = undefined;
}

toggle.addEventListener('click', () => {
  const shown = toggle.textContent === 'Close editor';
  if (shown) close();
  else void open();
  toggle.textContent = shown ? 'Open editor' : 'Close editor';
});

void open();
