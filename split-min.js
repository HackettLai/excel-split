let e = [],
  t = [],
  n = [],
  o = null;
const l = document.getElementById('fileInput'),
  a = document.getElementById('fileName'),
  r = document.getElementById('columnSection'),
  c = document.getElementById('columnInput'),
  i = document.getElementById('confirmBtn'),
  d = document.getElementById('loadingIndicator'),
  s = document.getElementById('errorMessage'),
  u = document.getElementById('resultSection'),
  m = document.getElementById('resultTable'),
  h = document.getElementById('downloadBtn'),
  p = document.getElementById('inputHint'),
  g = document.querySelectorAll('input[name="inputType"]'),
  y = document.getElementById('hasHeaderCheckbox'),
  f = document.getElementById('colNameList'),
  E = document.getElementById('columnListWrapper');
function S(e) {
  const t = e.target.files[0];
  t && ((o = t), (a.textContent = `Selected: ${t.name}`), b(t));
}
function b(t) {
  const o = new FileReader();
  (o.onload = function (t) {
    try {
      const o = new Uint8Array(t.target.result),
        l = XLSX.read(o, { type: 'array', raw: !1, codepage: 65001 }),
        a = l.Sheets[l.SheetNames[0]],
        c = XLSX.utils.sheet_to_json(a, { header: 1, defval: '', raw: !1 });
      if (c.length < 1) return void B('File is empty or contains no data');
      if (y.checked) {
        if (c.length < 2) return void B('File must contain at least a header row and one data row');
        (n = c[0].map((e) => String(e))), (e = c.slice(1)), X(`✓ Loaded with headers: ${n.join(', ')}`);
      } else {
        const t = c[0]?.length || 0;
        (n = Array.from({ length: t }, (e, t) => `Column ${t + 1}`)), (e = c), X(`✓ Loaded without headers (${e.length} rows, ${t} columns)`);
      }
      v(), (r.style.display = 'block'), (u.style.display = 'none');
    } catch (e) {
      B('Error reading file: ' + e.message);
    }
  }),
    o.readAsArrayBuffer(t);
}
function v() {
  if (0 === n.length) return void (E.style.display = 'none');
  let e = '';
  n.forEach((t, n) => {
    const o = y.checked ? t : `${n + 1}. ${t}`;
    e += `<span class="column-badge clickable" data-column="${t}" data-index="${n + 1}">${w(o)}</span>`;
  }),
    (f.innerHTML = e),
    (E.style.display = 'block'),
    document.querySelectorAll('.column-badge').forEach((e) => {
      e.addEventListener('click', (e) => {
        const t = e.target.getAttribute('data-column'),
          n = e.target.getAttribute('data-index');
        y.checked && document.querySelector('input[name="inputType"][value="name"]').checked ? (c.value = t) : (c.value = n), c.focus();
      });
    });
}
function L() {
  const e = document.querySelector('input[name="inputType"]:checked').value,
    o = c.value.trim();
  if (!o) return void B('Please enter a column ' + ('name' === e ? 'name' : 'number'));
  let l = -1;
  if ('name' === e) {
    if (((l = n.findIndex((e) => String(e).toLowerCase() === o.toLowerCase())), -1 === l)) return void B(`Column "${o}" not found. Available columns: ${n.join(', ')}`);
  } else if ('index' === e) {
    const e = parseInt(o);
    if (isNaN(e) || e < 1 || e > n.length) return void B(`Invalid column number. Please enter a number between 1 and ${n.length}`);
    l = e - 1;
  }
  (d.style.display = 'block'),
    C(),
    setTimeout(() => {
      try {
        (t = k(l)), I(), (d.style.display = 'none'), (u.style.display = 'block');
      } catch (e) {
        (d.style.display = 'none'), B('Error processing data: ' + e.message);
      }
    }, 100);
}
function k(t) {
  const o = [],
    l = [',', ';', '/', '&', '\n', '，', '；', '／'],
    a = ['n/a', 'N/A', 'n\\a', 'N\\A'];
  return (
    e.forEach((e) => {
      const r = [...e];
      for (; r.length < n.length; ) r.push('');
      const c = String(r[t] || '').trim();
      if (!c) return void o.push({ data: r, isSplit: !1 });
      if (a.some((e) => c.toLowerCase() === e.toLowerCase())) return void o.push({ data: r, isSplit: !1 });
      let i = [c];
      for (const e of l)
        if (c.includes(e)) {
          i = c
            .split(e)
            .map((e) => e.trim())
            .filter((e) => e);
          break;
        }
      i.length > 1
        ? i.forEach((e) => {
            const n = [...r];
            (n[t] = e), o.push({ data: n, isSplit: !0 });
          })
        : o.push({ data: r, isSplit: !1 });
    }),
    o
  );
}
function I() {
  let e = '<thead><tr>';
  n.forEach((t) => {
    e += `<th>${w(String(t))}</th>`;
  }),
    (e += '</tr></thead><tbody>'),
    t.forEach((t) => {
      const o = t.isSplit ? 'highlighted' : '';
      e += `<tr class="${o}">`;
      for (let o = 0; o < n.length; o++) {
        const n = void 0 !== t.data[o] ? t.data[o] : '';
        e += `<td>${w(String(n))}</td>`;
      }
      e += '</tr>';
    }),
    (e += '</tbody>'),
    (m.innerHTML = e);
}
function x() {
  const e = ['[SPLIT ROW]', ...n],
    o = t.map((e) => {
      const t = [...e.data];
      for (; t.length < n.length; ) t.push('');
      return [e.isSplit ? '✓ SPLIT' : '', ...t];
    }),
    l = XLSX.utils.aoa_to_sheet([e, ...o]),
    a = e.map((t, n) => {
      const l = Math.max(String(e[n]).length, ...o.map((e) => String(e[n] || '')).map((e) => e.length));
      return { wch: Math.min(Math.max(l + 2, 10), 50) };
    });
  l['!cols'] = a;
  const r = XLSX.utils.decode_range(l['!ref']);
  for (let e = 1; e <= r.e.r; e++)
    if (t[e - 1]?.isSplit)
      for (let t = r.s.c; t <= r.e.c; t++) {
        const n = XLSX.utils.encode_cell({ r: e, c: t });
        l[n] && (l[n].s || (l[n].s = {}), (l[n].s.fill = { fgColor: { rgb: 'FFF9E6' } }));
      }
  const c = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(c, l, 'Split Data'), XLSX.writeFile(c, 'split_data.xlsx');
}
function w(e) {
  const t = document.createElement('div');
  return (t.textContent = e), t.innerHTML;
}
function B(e) {
  (s.textContent = e), (s.style.display = 'block');
}
function C() {
  s.style.display = 'none';
}
function X(e) {
  C(),
    (s.textContent = e),
    (s.style.background = '#d4edda'),
    (s.style.color = '#155724'),
    (s.style.borderLeft = '4px solid #28a745'),
    (s.style.display = 'block'),
    setTimeout(() => {
      (s.style.background = ''), (s.style.color = ''), (s.style.borderLeft = ''), C();
    }, 3e3);
}
g.forEach((e) => {
  e.addEventListener('change', (e) => {
    'name' === e.target.value ? ((c.placeholder = 'Enter column name (e.g., Email)'), (p.textContent = 'Enter the exact column header name')) : ((c.placeholder = 'Enter column number (e.g., 7 for 7th column)'), (p.textContent = 'Enter the column position (1 = first column, 2 = second column, etc.)')), (c.value = '');
  });
}),
  y.addEventListener('change', (e) => {
    e.target.checked ? ((document.querySelector('input[name="inputType"][value="name"]').disabled = !1), (document.querySelector('input[name="inputType"][value="name"]').checked = !0), (c.placeholder = 'Enter column name (e.g., Email)'), (p.textContent = 'Enter the exact column header name')) : ((document.querySelector('input[name="inputType"][value="index"]').checked = !0), (document.querySelector('input[name="inputType"][value="name"]').disabled = !0), (c.placeholder = 'Enter column number (e.g., 1 for first column)'), (p.textContent = 'Enter the column position (1 = first column, 2 = second column, etc.)')), (c.value = ''), o && b(o);
  }),
  l.addEventListener('change', S),
  i.addEventListener('click', L),
  h.addEventListener('click', x);
