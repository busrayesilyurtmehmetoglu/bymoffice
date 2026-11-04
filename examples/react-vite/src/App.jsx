import { useState } from 'react';
import OfficeEditor from './OfficeEditor.jsx';

// What to open: ?file=<address of a docx, xlsx, doc, xls or RTF file>, or a new file (?new=xlsx, default docx).
const params = new URLSearchParams(location.search);
const file = params.get('file');
const kind = params.get('new') === 'xlsx' ? 'xlsx' : 'docx';

export default function App() {
  const [shown, setShown] = useState(true);
  return (
    <>
      <header>
        <button type="button" onClick={() => setShown((value) => !value)}>
          {shown ? 'Close editor' : 'Open editor'}
        </button>
      </header>
      {shown && <OfficeEditor file={file} kind={kind} />}
    </>
  );
}
