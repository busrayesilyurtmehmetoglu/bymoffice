// The live editor of the site. A file chosen or dropped here is read in this page only: it is
// handed to BymOffice.mount() as a File and never sent anywhere. Messages are set as text.
(function () {
  'use strict';
  var site = window.bymSite;
  var api = window.BymOffice;
  var $ = function (id) {
    return document.getElementById(id);
  };
  var container = $('editor');
  var empty = $('empty');
  var statusEl = $('status');
  var readOnly = $('read-only');

  /** What is open: the mounted editor, its name, and the bytes it was opened from. */
  var current = null;
  var name = '';
  var source = null;
  /** Each opening counts; an older one that ends late does not replace a newer one. */
  var serial = 0;

  var MIME = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pdf: 'application/pdf',
  };
  var LEGACY = { doc: true, xls: true, rtf: true };

  function say(key, values, kind) {
    statusEl.textContent = site.t(key, values);
    statusEl.setAttribute('data-kind', kind || 'info');
  }
  function message(error) {
    return String((error && error.message) || error).slice(0, 300);
  }
  function setBusy(busy) {
    $('download').disabled = busy || !current;
    $('pdf').disabled = busy || !current;
  }

  function close() {
    if (current) {
      try {
        current.destroy();
      } catch (error) {
        // Already gone.
      }
    }
    current = null;
    window.bymEditor = undefined;
    empty.hidden = false;
    setBusy(false);
  }

  /** Opens bytes (or a File) in the editor, replacing what was open. */
  function show(file, fileName) {
    var mine = ++serial;
    close();
    empty.hidden = true;
    name = fileName;
    say('editor.loading');
    var bytes = file instanceof Uint8Array ? Promise.resolve(file) : file.arrayBuffer().then(function (buffer) {
      return new Uint8Array(buffer);
    });
    return bytes
      .then(function (data) {
        if (mine !== serial) return undefined;
        source = data;
        return api.mount(container, {
          // A copy: the editor may keep what it is given.
          file: data.slice(),
          fileName: fileName,
          locale: site.lang,
          readOnly: readOnly.checked,
          // doc, xls and RTF become editable as docx and xlsx (unless read-only is asked for).
          convert: true,
          onNotice: function (notice) {
            statusEl.textContent = String(notice.message);
            statusEl.setAttribute('data-kind', notice.kind === 'error' ? 'error' : 'info');
          },
        });
      })
      .then(function (mounted) {
        if (!mounted) return;
        if (mine !== serial) {
          mounted.destroy();
          return;
        }
        current = mounted;
        window.bymEditor = mounted;
        setBusy(false);
        if (mounted.conversion) say('editor.converted', { name: fileName, to: mounted.conversion.to });
        else if (LEGACY[mounted.file.format]) say('editor.legacyReadOnly', { name: fileName });
        else say('editor.opened', { name: fileName });
      })
      .catch(function (error) {
        if (mine !== serial) return;
        close();
        say('editor.failed', { message: message(error) }, 'error');
      });
  }

  /** The name the file is downloaded under: its format after a conversion. */
  function savedName(format) {
    var base = name.replace(/\.[A-Za-z0-9]{1,5}$/, '') || 'file';
    return base + '.' + format;
  }

  function offer(bytes, fileName, type) {
    var url = URL.createObjectURL(new Blob([bytes], { type: type }));
    var link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.hidden = true;
    document.body.appendChild(link);
    link.click();
    link.remove();
    // Some browsers read the address after the click returns.
    window.setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 60000);
  }

  $('download').addEventListener('click', function () {
    if (!current) return;
    var mounted = current;
    setBusy(true);
    say('editor.saving');
    mounted
      .save()
      .then(function (bytes) {
        var format = mounted.conversion ? mounted.conversion.to : mounted.file.format;
        var fileName = savedName(format);
        offer(bytes, fileName, MIME[format] || 'application/octet-stream');
        say('editor.downloaded', { name: fileName });
      })
      .catch(function (error) {
        say('editor.saveFailed', { message: message(error) }, 'error');
      })
      .then(function () {
        setBusy(false);
      });
  });

  $('pdf').addEventListener('click', function () {
    if (!current) return;
    var mounted = current;
    setBusy(true);
    say('editor.pdfMaking');
    mounted
      .exportPdf()
      .then(function (bytes) {
        var fileName = savedName('pdf');
        offer(bytes, fileName, MIME.pdf);
        var pages = (new TextDecoder('latin1').decode(bytes).match(/\/Type\s*\/Page\b/g) || []).length;
        say('editor.pdfDone', { name: fileName, pages: site.t('editor.pages', { count: pages }) });
      })
      .catch(function (error) {
        say('editor.pdfFailed', { message: message(error) }, 'error');
      })
      .then(function () {
        setBusy(false);
      });
  });

  readOnly.addEventListener('change', function () {
    if (!current) return;
    // A doc, xls or RTF shown read-only was not converted: it is opened again to be edited.
    if (!readOnly.checked && LEGACY[current.file.format] && !current.conversion && source) {
      show(source, name);
      return;
    }
    current.editor.setReadOnly(readOnly.checked);
  });

  function blank(kind) {
    var fileName = site.t(kind === 'docx' ? 'editor.blankDocName' : 'editor.blankSheetName');
    api.createBlank(kind, { locale: site.lang }).then(
      function (bytes) {
        show(bytes, fileName);
      },
      function (error) {
        say('editor.failed', { message: message(error) }, 'error');
      },
    );
  }
  $('blank-doc').addEventListener('click', function () {
    blank('docx');
  });
  $('blank-sheet').addEventListener('click', function () {
    blank('xlsx');
  });

  /** The samples are files of the site, made when it was built. */
  function sample(file, key) {
    var mine = serial;
    fetch('samples/' + file.replace('{lang}', site.lang))
      .then(function (response) {
        if (!response.ok) throw new Error('HTTP ' + response.status);
        return response.arrayBuffer();
      })
      .then(function (buffer) {
        if (mine === serial) show(new Uint8Array(buffer), site.t(key));
      })
      .catch(function (error) {
        say('editor.failed', { message: message(error) }, 'error');
      });
  }
  $('sample-sheet').addEventListener('click', function () {
    sample('sample-sheet-{lang}.xlsx', 'editor.sampleSheetName');
  });
  $('sample-doc').addEventListener('click', function () {
    sample('sample-doc-{lang}.docx', 'editor.sampleDocName');
  });

  $('pick').addEventListener('change', function (event) {
    var file = event.target.files && event.target.files[0];
    if (file) show(file, file.name);
    event.target.value = '';
  });

  var drop = $('drop');
  drop.addEventListener('dragover', function (event) {
    event.preventDefault();
    drop.classList.add('over');
  });
  drop.addEventListener('dragleave', function () {
    drop.classList.remove('over');
  });
  drop.addEventListener('drop', function (event) {
    event.preventDefault();
    drop.classList.remove('over');
    var file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
    if (file) show(file, file.name);
  });

  // A file dropped beside the box would make the browser leave the page to show it.
  window.addEventListener('dragover', function (event) {
    if (event.target !== drop && !drop.contains(event.target)) event.preventDefault();
  });
  window.addEventListener('drop', function (event) {
    if (event.target !== drop && !drop.contains(event.target)) event.preventDefault();
  });

  site.onLanguage(function (lang) {
    if (!current) return;
    try {
      current.editor.setLocale(lang);
    } catch (error) {
      // A dialog with changes is open: the editor said why, and keeps its language.
    }
  });
})();
