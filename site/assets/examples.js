// The examples page: for each example, its page running in an iframe and its code beside it. The
// code comes from examples/code.json, which the build makes from the very files the iframes show.
(function () {
  'use strict';
  var site = window.bymSite;
  var EXAMPLES = [
    { id: 'script' },
    { id: 'module' },
    { id: 'viewer' },
    { id: 'save' },
    { id: 'autosave' },
    { id: 'collab', tall: true },
    { id: 'pdf' },
    { id: 'bridge', tall: true },
    { id: 'locale' },
  ];
  var list = document.getElementById('examples');
  var toc = document.getElementById('toc');
  var frames = [];

  function el(tag, attrs, key) {
    var node = document.createElement(tag);
    for (var name in attrs) if (Object.prototype.hasOwnProperty.call(attrs, name)) node.setAttribute(name, attrs[name]);
    if (key) {
      node.setAttribute('data-i18n', key);
      node.textContent = site.t(key);
    }
    return node;
  }
  function address(id) {
    return 'examples/' + id + '.html?lang=' + site.lang;
  }

  EXAMPLES.forEach(function (example) {
    var id = example.id;
    var item = el('li');
    var link = el('a', { href: '#ex-' + id }, 'ex.' + id + '.title');
    item.appendChild(link);
    toc.appendChild(item);

    var section = el('section', { class: 'example', id: 'ex-' + id, 'aria-labelledby': 'ex-' + id + '-title', 'data-example': id });
    section.appendChild(el('h2', { id: 'ex-' + id + '-title' }, 'ex.' + id + '.title'));
    section.appendChild(el('p', {}, 'ex.' + id + '.body'));

    var liveMeta = el('div', { class: 'meta' });
    liveMeta.appendChild(el('h3', {}, 'examples.live'));
    var open = el('a', { href: address(id), target: '_blank', rel: 'noopener' }, 'examples.open');
    liveMeta.appendChild(open);
    section.appendChild(liveMeta);
    // Started when it comes near the screen (see below): a page of nine editors at once is heavy on a phone.
    var frame = el('iframe', { class: 'frame' + (example.tall ? ' tall' : ''), 'data-src': address(id), 'data-i18n-attr': 'title:ex.' + id + '.title', title: site.t('ex.' + id + '.title') });
    section.appendChild(frame);
    frames.push({ id: id, frame: frame, open: open, started: false });

    var codeMeta = el('div', { class: 'meta' });
    codeMeta.appendChild(el('h3', {}, 'examples.code'));
    section.appendChild(codeMeta);
    var box = el('div', { class: 'code' });
    var pre = el('pre', { id: 'code-' + id, tabindex: '0', 'data-i18n-attr': 'aria-label:examples.code', 'aria-label': site.t('examples.code') });
    var code = el('code');
    code.textContent = site.t('examples.loadingCode');
    pre.appendChild(code);
    box.appendChild(pre);
    var copy = el('button', { type: 'button', class: 'copy', 'data-copy-for': 'code-' + id }, 'code.copy');
    box.appendChild(copy);
    section.appendChild(box);
    list.appendChild(section);
  });
  site.wireCopy(list);

  fetch('examples/code.json')
    .then(function (response) {
      if (!response.ok) throw new Error('HTTP ' + response.status);
      return response.json();
    })
    .then(function (all) {
      EXAMPLES.forEach(function (example) {
        var code = document.querySelector('#code-' + example.id + ' code');
        code.textContent = typeof all[example.id] === 'string' ? all[example.id] : site.t('examples.codeFailed');
      });
      document.body.setAttribute('data-code', 'ready');
    })
    .catch(function () {
      EXAMPLES.forEach(function (example) {
        document.querySelector('#code-' + example.id + ' code').textContent = site.t('examples.codeFailed');
      });
    });

  function start(entry) {
    if (entry.started) return;
    entry.started = true;
    entry.frame.src = entry.frame.getAttribute('data-src');
  }
  if (typeof IntersectionObserver === 'function') {
    var observer = new IntersectionObserver(
      function (seen) {
        seen.forEach(function (item) {
          if (!item.isIntersecting) return;
          frames.forEach(function (entry) {
            if (entry.frame === item.target) start(entry);
          });
          observer.unobserve(item.target);
        });
      },
      { rootMargin: '200px 0px' },
    );
    frames.forEach(function (entry) {
      observer.observe(entry.frame);
    });
  } else {
    frames.forEach(start);
  }

  // The examples speak the language of the site: those already running are opened again in the new one.
  site.onLanguage(function () {
    frames.forEach(function (entry) {
      entry.frame.setAttribute('data-src', address(entry.id));
      if (entry.started) entry.frame.src = address(entry.id);
      entry.open.href = address(entry.id);
    });
  });
})();
