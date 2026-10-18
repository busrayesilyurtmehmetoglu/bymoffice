// What every page of the site shares: the language (chosen, remembered, applied to the page) and
// the copy buttons of code blocks. Texts are only ever set with textContent.
(function () {
  'use strict';
  var TEXT = window.BYM_SITE_TEXT;
  var KEY = 'bymoffice-site-lang';
  var listeners = [];

  function stored() {
    try {
      var value = window.localStorage.getItem(KEY);
      return value === 'tr' || value === 'en' ? value : null;
    } catch (error) {
      return null;
    }
  }
  function remember(lang) {
    try {
      window.localStorage.setItem(KEY, lang);
    } catch (error) {
      // A private window or blocked storage: the choice lasts as long as the page.
    }
  }
  /** `?lang=tr` or `?lang=en` in the address: a link that opens the site in a given language. */
  function asked() {
    var match = /[?&]lang=(tr|en)(?:&|#|$)/.exec(window.location.search);
    return match ? match[1] : null;
  }
  // English first, whatever the language of the browser; Turkish when the address asks for it
  // or the visitor chose it before.
  function initial() {
    var given = asked();
    if (given) {
      remember(given);
      return given;
    }
    return stored() || 'en';
  }

  var lang = initial();

  /** The text of a key in the language of the page, with "{name}" filled from `values`. */
  function t(key, values) {
    var text = (TEXT[lang] && TEXT[lang][key]) || TEXT.en[key] || key;
    if (values) {
      text = text.replace(/\{(\w+)\}/g, function (all, name) {
        return Object.prototype.hasOwnProperty.call(values, name) ? String(values[name]) : all;
      });
    }
    return text;
  }

  function apply() {
    document.documentElement.lang = lang;
    var nodes = document.querySelectorAll('[data-i18n]');
    for (var i = 0; i < nodes.length; i++) nodes[i].textContent = t(nodes[i].getAttribute('data-i18n'));
    var attributed = document.querySelectorAll('[data-i18n-attr]');
    for (var j = 0; j < attributed.length; j++) {
      var pairs = attributed[j].getAttribute('data-i18n-attr').split(';');
      for (var k = 0; k < pairs.length; k++) {
        var pair = pairs[k].split(':');
        if (pair.length === 2) attributed[j].setAttribute(pair[0], t(pair[1]));
      }
    }
    var title = document.body.getAttribute('data-title');
    if (title) document.title = t(title);
    var select = document.getElementById('lang');
    if (select) select.value = lang;
  }

  function setLanguage(next) {
    if (next !== 'tr' && next !== 'en') return;
    remember(next);
    if (next === lang) return;
    lang = next;
    apply();
    for (var i = 0; i < listeners.length; i++) {
      try {
        listeners[i](lang);
      } catch (error) {
        if (window.console) console.warn(error);
      }
    }
  }

  /** Copies text; tells the button what happened. */
  function copy(text, button) {
    function said(key) {
      button.textContent = t(key);
      window.setTimeout(function () {
        button.textContent = t('code.copy');
      }, 2000);
    }
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(
          function () {
            said('code.copied');
          },
          function () {
            said('code.copyFailed');
          },
        );
        return;
      }
    } catch (error) {
      // Falls through.
    }
    said('code.copyFailed');
  }

  /** Wires each button with data-copy-for="id": it copies the text of that element. */
  function wireCopy(root) {
    var blocks = (root || document).querySelectorAll('[data-copy-for]');
    for (var i = 0; i < blocks.length; i++) {
      (function (button) {
        if (button.getAttribute('data-wired')) return;
        button.setAttribute('data-wired', '1');
        button.addEventListener('click', function () {
          var target = document.getElementById(button.getAttribute('data-copy-for'));
          if (target) copy(target.textContent, button);
        });
      })(blocks[i]);
    }
  }

  window.bymSite = {
    get lang() {
      return lang;
    },
    t: t,
    setLanguage: setLanguage,
    onLanguage: function (listener) {
      listeners.push(listener);
    },
    wireCopy: wireCopy,
  };

  function start() {
    apply();
    var select = document.getElementById('lang');
    if (select) {
      select.addEventListener('change', function () {
        setLanguage(select.value);
      });
    }
    wireCopy(document);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
