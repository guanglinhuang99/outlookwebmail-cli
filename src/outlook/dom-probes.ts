export const PAGE_DOM_INVENTORY = String.raw`
(() => {
  const allowedDataAttrs = new Set(['data-testid', 'data-automationid', 'data-icon-name']);

  function visible(el) {
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
  }

  function selectedAttrs(el) {
    return Object.fromEntries(
      Array.from(el.attributes)
        .filter(a =>
          a.name.startsWith('aria-') ||
          allowedDataAttrs.has(a.name) ||
          ['role', 'title', 'datetime', 'contenteditable'].includes(a.name)
        )
        .map(a => [a.name, a.value.slice(0, 300)])
    );
  }

  return Array.from(document.querySelectorAll('*'))
    .filter(visible)
    .map((el, index) => {
      const r = el.getBoundingClientRect();
      const text = (el.innerText || '').trim();
      const attrs = selectedAttrs(el);
      return {
        index,
        tag: el.tagName,
        role: el.getAttribute('role'),
        ariaLabel: el.getAttribute('aria-label'),
        title: el.getAttribute('title'),
        datetime: el.getAttribute('datetime'),
        contentEditable: el.getAttribute('contenteditable'),
        text: text.slice(0, 240),
        textLength: text.length,
        attrs,
        childCount: el.children.length,
        scrollHeight: el.scrollHeight,
        clientHeight: el.clientHeight,
        rect: {
          x: Math.round(r.x), y: Math.round(r.y),
          width: Math.round(r.width), height: Math.round(r.height)
        }
      };
    })
    .filter(x => x.role || x.ariaLabel || x.title || x.datetime || Object.keys(x.attrs).length > 0 || x.textLength > 0)
    .slice(0, 1500);
})()
`;

export const LIST_CANDIDATES = String.raw`
(() => {
  const selector = ['[role="list"]', '[role="listbox"]', '[role="grid"]', '[role="treegrid"]', '[role="feed"]'].join(',');
  return Array.from(document.querySelectorAll(selector)).map((el, index) => {
    const r = el.getBoundingClientRect();
    const text = (el.innerText || '').trim();
    return {
      index,
      tag: el.tagName,
      role: el.getAttribute('role'),
      ariaLabel: el.getAttribute('aria-label'),
      text: text.slice(0, 3000),
      childCount: el.children.length,
      scrollHeight: el.scrollHeight,
      clientHeight: el.clientHeight,
      rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
    };
  });
})()
`;

export const MESSAGE_ROW_CANDIDATES = String.raw`
(() => {
  const allowedDataAttrs = new Set(['data-testid', 'data-automationid', 'data-icon-name']);
  const attrs = el => Object.fromEntries(
    Array.from(el.attributes)
      .filter(a => a.name.startsWith('aria-') || allowedDataAttrs.has(a.name) || ['role', 'title'].includes(a.name))
      .map(a => [a.name, a.value.slice(0, 300)])
  );

  return Array.from(document.querySelectorAll('[role="option"], [role="listitem"], [role="row"]'))
    .map((row, index) => ({
      index,
      tag: row.tagName,
      role: row.getAttribute('role'),
      ariaLabel: row.getAttribute('aria-label'),
      text: (row.innerText || '').trim().slice(0, 2000),
      attrs: attrs(row),
      children: Array.from(row.querySelectorAll('*')).slice(0, 80).map((el, childIndex) => ({
        childIndex,
        tag: el.tagName,
        role: el.getAttribute('role'),
        ariaLabel: el.getAttribute('aria-label'),
        title: el.getAttribute('title'),
        datetime: el.getAttribute('datetime'),
        text: (el.innerText || '').trim().slice(0, 300),
        attrs: attrs(el)
      }))
    }))
    .filter(x => x.text.length > 0)
    .slice(0, 100);
})()
`;

export const SCROLL_CANDIDATES = String.raw`
(() => Array.from(document.querySelectorAll('*'))
  .filter(el => el.scrollHeight > el.clientHeight + 20 && el.clientHeight > 100)
  .map((el, index) => {
    const r = el.getBoundingClientRect();
    return {
      index,
      tag: el.tagName,
      role: el.getAttribute('role'),
      ariaLabel: el.getAttribute('aria-label'),
      text: (el.innerText || '').trim().slice(0, 500),
      scrollTop: el.scrollTop,
      scrollHeight: el.scrollHeight,
      clientHeight: el.clientHeight,
      rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
    };
  })
  .filter(x => x.rect.width > 0 && x.rect.height > 0)
  .slice(0, 100))()
`;

export const IFRAME_CANDIDATES = String.raw`
(() => Array.from(document.querySelectorAll('iframe')).map((el, index) => {
  const r = el.getBoundingClientRect();
  let src = null;
  try {
    const url = new URL(el.src);
    src = url.origin + url.pathname;
  } catch {}
  return {
    index,
    title: el.getAttribute('title'),
    ariaLabel: el.getAttribute('aria-label'),
    src,
    rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
  };
}))()
`;

export const MESSAGE_BODY_CANDIDATES = String.raw`
(() => {
  const selector = ['[role="document"]', '[role="main"]', 'article', 'iframe'].join(',');
  return Array.from(document.querySelectorAll(selector)).map((el, index) => {
    const r = el.getBoundingClientRect();
    const text = (el.innerText || '').trim();
    return {
      index,
      tag: el.tagName,
      role: el.getAttribute('role'),
      ariaLabel: el.getAttribute('aria-label'),
      title: el.getAttribute('title'),
      textLength: text.length,
      text: text.slice(0, 8000),
      childCount: el.children.length,
      rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
    };
  }).filter(x => x.textLength > 50 || x.tag === 'IFRAME');
})()
`;

export const MESSAGE_HEADER_CANDIDATES = String.raw`
(() => Array.from(document.querySelectorAll('h1,h2,h3,[role="heading"],span[title],button[aria-label]'))
  .map((el, index) => {
    const r = el.getBoundingClientRect();
    return {
      index,
      tag: el.tagName,
      role: el.getAttribute('role'),
      ariaLabel: el.getAttribute('aria-label'),
      title: el.getAttribute('title'),
      text: (el.innerText || '').trim().slice(0, 1000),
      rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
    };
  })
  .filter(x => x.rect.width > 0 && x.rect.height > 0 && (x.text || x.ariaLabel || x.title))
  .slice(0, 500))()
`;

export const ATTACHMENT_CANDIDATES = String.raw`
(() => Array.from(document.querySelectorAll('button,[role="button"],[download],[title]'))
  .map((el, index) => {
    const r = el.getBoundingClientRect();
    return {
      index,
      tag: el.tagName,
      role: el.getAttribute('role'),
      ariaLabel: el.getAttribute('aria-label'),
      title: el.getAttribute('title'),
      text: (el.innerText || '').trim().slice(0, 1000),
      rect: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
    };
  })
  .filter(x => x.rect.width > 0 && x.rect.height > 0 && /附件|下载|attachment|download|\.[a-z0-9]{2,6}\b/i.test([x.ariaLabel, x.title, x.text].filter(Boolean).join(' ')))
  .slice(0, 300))()
`;
