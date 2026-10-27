'use strict';
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const token = $('meta[name="klikatko-token"]').content;
const fields = ['name','number','location','investor','stage','author','date','output_dir','filename'];
let state, project = {}, selected = new Set(), expanded = new Set(['A','B']), nodes = new Map(), parents = new Map(), leaves = new Map();
let currentStep = 1, dirty = false, filenameEdited = false, toastTimer, folderState, latestExport, coverUrl, road, pinnedHelp = null;
const normal = (s) => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const escapeHTML = (s) => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const ICON = {
  chevron: '<svg class="i" viewBox="0 0 24 24" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>',
  folder: '<svg class="i" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 7h6l2 2h10v10H3z"/></svg>',
  close: '<svg class="i" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>',
};
const PT = 1240 / 595.3;   // preview page pixels per typographic point (A4 width)
// Car pictures: wheel centres and rim radius as fractions of the image, measured from the source drawings.
const CARS = {
  seat:  {aspect: .3458, rim: .0555, wheels: [[.2083, .8072], [.8187, .8072]], bottom: 1},
  mazda: {aspect: .3523, rim: .0605, wheels: [[.1835, .7376], [.8265, .749]], bottom: .9809},
};
const CAR_W = 120;
const reducedMotion = matchMedia('(prefers-reduced-motion: reduce)').matches;
const ease = u => u < .5 ? 4 * u * u * u : 1 - Math.pow(-2 * u + 2, 3) / 2;
// Czech typesetting: a one-letter preposition or conjunction never ends a line
const nbsp = s => String(s).replace(/(^|\s)([kKsSvVzZoOuUaAiI]) /g, '$1$2 ');
function chosenCar() { try { return localStorage.getItem('klikatko-car') === 'mazda' ? 'mazda' : 'seat'; } catch { return 'seat'; } }

async function api(path, body) {
  const response = await fetch(path, {method: body === undefined ? 'GET' : 'POST', headers: {'X-Klikatko-Token': token, 'Content-Type':'application/json'}, body: body === undefined ? undefined : JSON.stringify(body)});
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || 'Tohle se nepovedlo.');
  return result;
}
function toast(message, error=false) {
  clearTimeout(toastTimer);
  $('#toast').textContent = message;
  $('#toast').classList.remove('hidden');
  $('#toast').classList.toggle('error', error);
  toastTimer = setTimeout(() => $('#toast').classList.add('hidden'), error ? 10000 : 4200);
}
function setSaved(text, ok) {
  $('#save-state').textContent = text;
  $('#save-state').classList.toggle('ok', ok);
  $('#save-project').classList.toggle('hidden', ok);
}
function clearExport() { $('#export-result').classList.add('hidden'); $('#step-3').classList.remove('is-done'); if (road) { $('.pin[data-step="3"]').classList.remove('done'); driveCar(); } }
function markDirty() { dirty = true; setSaved('Neuložené změny', false); clearExport(); }
function acceptDiscard() { return !dirty || window.confirm('Máš neuložené změny. Opravdu je chceš zahodit?'); }
function collectProject() { fields.forEach(f => project[f] = $('#' + f).value.trim()); project.selected = [...selected]; return {...project}; }
function normalizeSelection(ids) { const result = new Set(); (ids || []).forEach(id => (leaves.get(id) || []).forEach(leaf => result.add(leaf))); return result; }
function indexNodes(items, parent=null) {
  for (const node of items) {
    nodes.set(node.id,node); parents.set(node.id,parent);
    indexNodes(node.children || [], node.id);
    leaves.set(node.id, node.children.length ? node.children.flatMap(n => leaves.get(n.id)) : [node.id]);
  }
}
function included(id) { return leaves.get(id).some(x => selected.has(x)); }
function counts() { let total=0; for (const id of nodes.keys()) if (included(id)) total++; return total; }
function makeDefault() {
  const today = new Date(state.today + 'T12:00:00');
  project = {id:'',name:'',number:'',location:'',investor:'',author:'Ing. Jiří Koudelka',date:today.toLocaleDateString('cs-CZ',{month:'long',year:'numeric'}),stage:'Dokumentace pro povolení záměru',output_dir:state.output_dir,filename:'',cover_image:''};
  selected = normalizeSelection(['A1','A2','A3']);
  filenameEdited=false; dirty=false; fillForm(); expanded = new Set(['A','B']);
  $('#search').value=''; $('#selected-only').checked=false;
  renderTree(); updateSummary(); showStep(1); setSaved('Zatím neuloženo', false);
}
function fillForm() {
  fields.forEach(f => $('#' + f).value = project[f] || '');
  $('#image-status').textContent = project.cover_image ? 'Tvůj vlastní obrázek.' : 'Ten z předlohy.';
  $('#reset-image').classList.toggle('hidden', !project.cover_image);
  clearExport();
  loadCover();
}

/* Cover image: the template's own, or the one he picked. Fetched with the token, shown via a blob URL. */
async function loadCover() {
  try {
    const response = await fetch('/api/cover?path=' + encodeURIComponent(project.cover_image || ''), {headers: {'X-Klikatko-Token': token}});
    if (!response.ok) throw new Error();
    const url = URL.createObjectURL(await response.blob());
    if (coverUrl) URL.revokeObjectURL(coverUrl);
    coverUrl = url; $('#cover-map').src = url; $('#cover-thumb').src = url;
  } catch { $('#cover-map').removeAttribute('src'); $('#cover-thumb').removeAttribute('src'); }
}
function renderCoverStatic() {
  const band = $('#cover-band');
  band.replaceChildren();
  for (const paragraph of state.cover.letterhead) {
    const p = document.createElement('p');
    for (const run of paragraph) {
      const span = document.createElement(run.link ? 'a' : 'span');
      span.textContent = run.text;
      span.style.fontSize = (run.size * PT) + 'px';
      span.style.fontWeight = run.bold ? '700' : '400';
      p.append(span);
    }
    band.append(p);
  }
  $('#cover-title').textContent = state.cover.title;
  $('#cover-att').textContent = state.cover.attachment;
}

/* Cars: body picture + two rims that spin with the distance driven. Built with DOM calls: the page's CSP forbids inline styles. */
function makeCar(name, parent, soft=false) {
  const m = CARS[name], h = CAR_W * m.aspect, el = document.createElement('div');
  el.className = 'vcar' + (soft ? ' soft' : '');
  let under = null;
  if (soft) {
    under = document.createElement('div'); under.className = 'under';
    Object.assign(under.style, {width: CAR_W + 'px', height: h + 'px', webkitMaskImage: `url(/static/cars/${name}.png)`, maskImage: `url(/static/cars/${name}.png)`});
    el.append(under);
  }
  const body = document.createElement('div'); body.className = 'body'; body.style.width = CAR_W + 'px';
  const img = document.createElement('img'); img.className = 'b'; img.src = `/static/cars/${name}.png`; img.alt = ''; body.append(img);
  const rims = m.wheels.map(([x, y]) => {
    const r = document.createElement('img'); r.className = 'rim'; r.src = `/static/cars/${name}-rim.png`; r.alt = '';
    Object.assign(r.style, {left: x * 100 + '%', top: y * 100 + '%', width: 2 * m.rim * CAR_W + 'px'}); body.append(r); return r;
  });
  el.append(body); parent.append(el);
  const ax = CAR_W / 2, ay = m.bottom * h, rimR = m.rim * CAR_W;
  el.style.transformOrigin = `${ax}px ${ay}px`;
  return {
    el,
    // put the point where the wheels touch the road on (x, y), tilt by deg; dist turns the rims
    pose(x, y, deg, dist, flip=false) {
      el.style.transform = `translate(${x - ax}px, ${y - ay}px) rotate(${deg}deg)`;
      body.style.transform = flip ? 'scaleX(-1)' : '';
      if (under) under.style.transform = body.style.transform;
      const turn = dist / rimR * 180 / Math.PI;
      rims.forEach(r => r.style.transform = `translate(-50%,-50%) rotate(${turn}deg)`);
    },
  };
}
function sampler(path) {   // arc-length lookups on an SVG path
  const L = path.getTotalLength();
  const at = s => { s = Math.max(0, Math.min(L, s)); const p = path.getPointAtLength(s), a = path.getPointAtLength(Math.max(0, s - 3)), b = path.getPointAtLength(Math.min(L, s + 3)); return {x: p.x, y: p.y, deg: Math.atan2(b.y - a.y, b.x - a.x) * 180 / Math.PI}; };
  const lenAtX = x => { let lo = 0, hi = L; for (let k = 0; k < 30; k++) { const m = (lo + hi) / 2; path.getPointAtLength(m).x < x ? lo = m : hi = m; } return lo; };
  return {L, at, lenAtX};
}

/* Road stepper: pins at equal distances; the car waits just before the current step and drives past "Word" once the document exists. */
let topCar, carS = null, carAnim;
function buildRoad() {
  const el = $('#road'), w = el.clientWidth, h = el.clientHeight, base = h - 40;
  const d = `M0 ${base} C ${w*.18} ${base} ${w*.24} ${base-34} ${w*.38} ${base-30} S ${w*.6} ${base+4} ${w*.72} ${base-8} S ${w*.9} ${base-34} ${w} ${base-30}`;
  el.querySelector('svg.track')?.remove();
  el.insertAdjacentHTML('afterbegin', `<svg class="track" viewBox="0 0 ${w} ${h}" aria-hidden="true"><path d="${d}" stroke="#5e6266" stroke-width="40" fill="none"/><path d="${d}" stroke="#fff" stroke-width="3" fill="none" stroke-dasharray="10 20"/></svg>`);
  const path = sampler(el.querySelector('svg.track path')), xs = [w / 6, w / 2, w * 5 / 6];
  $$('.pin').forEach((pin, i) => {
    const point = path.at(path.lenAtX(xs[i]));
    pin.style.left = point.x + 'px';
    pin.querySelector('.stem').style.height = Math.max(8, point.y - 16 - 16 - pin.querySelector('.lbl').offsetHeight) + 'px';
  });
  road = {path, stops: [...xs.map(x => path.lenAtX(x - 16 - CAR_W / 2 - 24)), path.lenAtX(xs[2] + 16 + CAR_W / 2 + 24)]};
  if (!topCar) topCar = makeCar(chosenCar(), el);
  carS = null; driveCar(false);
}
function carStop() { return currentStep === 3 && !$('#export-result').classList.contains('hidden') ? 3 : currentStep - 1; }
function driveCar(animate=true) {
  if (!road) return;
  const target = road.stops[carStop()], from = carS ?? target;
  cancelAnimationFrame(carAnim);
  const put = s => { carS = s; const p = road.path.at(s); topCar.pose(p.x, p.y, p.deg, s); };
  if (!animate || reducedMotion || from === target) return put(target);
  const steps = Math.abs(target - from) / Math.abs(road.stops[1] - road.stops[0]);
  const ms = 3000 * Math.min(Math.max(steps, .5), 1.6), t0 = performance.now();   // about 3 s from one step to the next
  const tick = now => { const u = Math.min((now - t0) / ms, 1); put(from + (target - from) * ease(u)); if (u < 1) carAnim = requestAnimationFrame(tick); };
  carAnim = requestAnimationFrame(tick);
}
function setCar(name) {
  try { localStorage.setItem('klikatko-car', name); } catch {}
  $$('#car-pop .opt').forEach(o => o.setAttribute('aria-checked', String(o.dataset.car === name)));
  $('#car-button-img').src = `/static/cars/${name}.png`;
  $('#finish-car').setAttribute('href', `/static/cars/${name}.png`);
  if (topCar) { const s = carS; topCar.el.remove(); topCar = makeCar(name, $('#road')); carS = s; if (s !== null) { const p = road.path.at(s); topCar.pose(p.x, p.y, p.deg, s); } }
}
// Buttons at the bottom of a step scroll back up, so the car's drive on the road is in view.
function scrollToRoad() {
  const from = scrollY, t0 = performance.now(), ms = 600;
  if (!from) return;
  if (reducedMotion) return scrollTo(0, 0);
  const tick = now => { const u = Math.min((now - t0) / ms, 1); scrollTo(0, from * (1 - (1 - Math.pow(1 - u, 3)))); if (u < 1) requestAnimationFrame(tick); };
  requestAnimationFrame(tick);
}
function showStep(step) {
  currentStep = Number(step);
  document.body.dataset.step = currentStep;
  for (let n = 1; n <= 3; n++) $('#step-' + n).classList.toggle('hidden', n !== currentStep);
  $$('.pin').forEach(pin => {
    const n = Number(pin.dataset.step);
    pin.classList.toggle('on', n === currentStep);
    pin.classList.toggle('done', n < currentStep || (n === 3 && carStop() === 3));
    pin.setAttribute('aria-current', n === currentStep ? 'step' : 'false');
  });
  driveCar();
  updateSummary();
}

/* Bottom road: full width, below the first screen; both cars drive back and forth on the dashed line. */
function buildTraffic() {
  const el = $('#bottom-road'), w = el.clientWidth, h = el.clientHeight;
  const pts = []; for (let x = -200; x <= w + 200; x += 8) pts.push(`${x} ${h / 2 + 24 + 34 * Math.sin(2 * Math.PI * x / 820)}`);
  const d = 'M' + pts.join(' L');
  el.innerHTML = `<svg viewBox="0 0 ${w} ${h}"><path d="${d}" stroke="#5e6266" stroke-width="32" fill="none"/><path d="${d}" stroke="#fff" stroke-width="3" fill="none" stroke-dasharray="10 20"/></svg>`;
  const path = sampler(el.querySelector('path')), cars = [makeCar('seat', el, true), makeCar('mazda', el, true)];
  const LOOP = 28;   // seconds: across, turn round off-screen, back
  const frame = t => cars.forEach((car, i) => {
    const u = ((t / LOOP + i * .25) % 1 + 1) % 1, back = u >= .5, v = back ? (1 - u) * 2 : u * 2;   // quarter-cycle apart: one car is always in view
    const s = path.L * (.02 + .96 * ease(v)), p = path.at(s);
    car.pose(p.x, p.y, p.deg, back ? -s : s, back);
  });
  frame(LOOP * .3);
  if (reducedMotion) return;
  let visible = false, raf;
  const t0 = performance.now() - LOOP * .3 * 1000;
  const loop = now => { frame((now - t0) / 1000); if (visible) raf = requestAnimationFrame(loop); };
  new IntersectionObserver(([entry]) => { visible = entry.isIntersecting; cancelAnimationFrame(raf); if (visible) raf = requestAnimationFrame(loop); }).observe(el);   // runs only while on screen
}

/* Step 2 help panel: the chapter under the mouse, otherwise the one last clicked. */
function showHelp(id) {
  const n = nodes.get(id);
  if (!n) { $('#help-title').textContent = ''; $('#help-text').textContent = 'Najeď myší na kapitolu nebo na ni klikni a uvidíš, co do ní patří.'; return; }
  $('#help-title').innerHTML = `<span class="tech">${escapeHTML(n.code)}</span><span>${escapeHTML(nbsp(n.title))}</span>`;
  $('#help-text').textContent = nbsp(n.hint || 'Tahle část je nadpis pro podkapitoly.');
}

function renderTree() {
  const query=normal($('#search').value.trim()), only=$('#selected-only').checked;
  const matches = new Map();
  function visible(node) {
    const childMatch=node.children.map(visible).some(Boolean);
    const own=normal(node.code+' '+node.title+' '+node.id).includes(query);
    const match=(!query || own || childMatch) && (!only || included(node.id));
    matches.set(node.id,match);return match;
  }
  state.catalog.nodes.forEach(visible);
  let html='';
  function walk(items,depth=0,ancestorMatch=false) {
    for(const n of items) {
      const ownMatch=query && normal(n.code+' '+n.title+' '+n.id).includes(query);
      const isVisible=(matches.get(n.id) || ancestorMatch) && (!only || included(n.id));
      if(!isVisible)continue;
      const branch=n.children.length>0, open=expanded.has(n.id)||Boolean(query)||only;
      const leafIds=leaves.get(n.id), amount=leafIds.filter(x=>selected.has(x)).length;
      html+=`<div class="chapter-row ${depth===0?'root':''}" data-depth="${depth}">
      ${branch?`<button class="expand" data-expand="${n.id}" aria-expanded="${open}" aria-label="${open?'Sbalit':'Rozbalit'} ${escapeHTML(n.code)}">${ICON.chevron}</button>`:'<span class="expand-spacer"></span>'}
      <input type="checkbox" data-check="${n.id}" ${amount===leafIds.length?'checked':''} aria-label="Vybrat ${escapeHTML(n.id+' '+n.title)}" ${amount>0&&amount<leafIds.length?'data-partial="true"':''}>
      <button class="chapter-name" data-help="${n.id}"><strong>${escapeHTML(n.code)}</strong><span>${escapeHTML(nbsp(n.title))}</span></button>
      ${branch?`<span class="chapter-total">${amount} / ${leafIds.length}</span>`:''}</div>`;
      if(branch && open)walk(n.children,depth+1,ancestorMatch||ownMatch);
    }
  }
  walk(state.catalog.nodes);
  $('#chapter-tree').innerHTML=html||'<div class="empty-state">Nic takového v osnově není.<br>Zkus jiné slovo nebo vypni „Jen vybrané“.</div>';
  $$('.chapter-row').forEach(row=>row.style.setProperty('--depth',row.dataset.depth));
  if(pinnedHelp)$(`[data-help="${pinnedHelp}"]`)?.closest('.chapter-row').classList.add('is-focused');
  $$('[data-partial]').forEach(el=>el.indeterminate=true);
  $('#selection-label').textContent=`Vybráno ${selected.size} položek. Nadřazené nadpisy přidám sám.`;
}
function updateSummary() {
  collectProject();
  $('#preview-name').textContent = project.name || 'Název stavby';
  $('#preview-stage').textContent = `${project.stage || 'Druh dokumentace'}   z.č. ${project.number}`;
  $('#preview-date').textContent = 'Datum: ' + project.date;
  $('#preview-author').textContent = 'Vypracoval: ' + project.author;
  $('#stamp-number').textContent = project.number || '–';
  $('#summary-count').textContent = counts();
  const includedTop=[...nodes.values()].filter(n=>parents.get(n.id) && ['A','B'].includes(parents.get(n.id)) && included(n.id)).map(n=>n.code);
  $('#summary-outline').textContent = includedTop.length ? includedTop.slice(0,10).join(' · ') + (includedTop.length>10 ? ` +${includedTop.length-10}` : '') : '–';
  const checks=[{ok:Boolean(project.name),text:project.name?'Název stavby je vyplněný':'Doplň název stavby'}, {ok:selected.size>0,text:selected.size?`Vybráno ${counts()} kapitol`:'Vyber aspoň jednu kapitolu'}, {ok:Boolean(project.output_dir),text:project.output_dir?'Složka je nastavená':'Vyber složku'}];
  $('#review-checks').innerHTML=checks.map(c=>`<li class="${c.ok?'':'warning'}">${escapeHTML(c.text)}</li>`).join('');
}
function updateLibraryCounts() { $('#projects-count').textContent=state.projects.length;$('#presets-count').textContent=state.presets.length; }
async function saveProject() {
  if(!$('#name').value.trim()){showStep(1);$('#name').focus();throw new Error('Nejdřív napiš název stavby.');}
  project=await api('/api/projects',collectProject());
  state.projects=state.projects.filter(p=>p.id!==project.id).concat(project);
  dirty=false;setSaved('Uloženo',true);updateLibraryCounts();toast('Zakázka je uložená. Najdeš ji v Uložených zakázkách.');
}
function showLibrary(kind) {
  $('#library-title').textContent=kind==='projects'?'Uložené zakázky':'Předvolby';
  const items=state[kind];
  $('#library-list').innerHTML=items.length?items.map(item=>`<div class="library-entry"><div><strong>${escapeHTML(item.name)}</strong><small>${kind==='projects'?escapeHTML(item.number||'Bez čísla zakázky')+' · ':''}${item.selected.length} vybraných položek</small></div><button class="btn ghost small" data-load="${item.id}" data-kind="${kind}">Načíst</button><button class="icon-button" data-delete="${item.id}" data-kind="${kind}" aria-label="Smazat ${escapeHTML(item.name)}">${ICON.close}</button></div>`).join(''):`<div class="empty-state">${kind==='projects'?'Zatím tu nic není.<br>Zakázku uložíš tlačítkem Uložit nahoře.':'Zatím tu nic není.<br>Vyber kapitoly a dej „Uložit tento výběr“.'}</div>`;
  $('#library-dialog').showModal();
}
async function browseFolder(path) {
  folderState=await api('/api/folders?path='+encodeURIComponent(path));
  $('#folder-path').value=folderState.path;
  $('#folder-list').innerHTML=folderState.children.length?folderState.children.map(d=>`<button class="folder-entry" data-folder="${escapeHTML(d.path)}">${ICON.folder}${escapeHTML(d.name)}</button>`).join(''):'<div class="empty-state">Tahle složka už nemá podsložky.<br>Můžeš ji použít pro uložení zprávy.</div>';
  $('#folder-drives').innerHTML=folderState.drives.map(d=>`<button class="btn ghost small" data-folder="${escapeHTML(d)}">${escapeHTML(d)}</button>`).join('');
}
async function downloadWord() {
  if(!latestExport)return;
  const response=await fetch('/api/download?id='+encodeURIComponent(latestExport.id),{headers:{'X-Klikatko-Token':token}});
  if(!response.ok)throw new Error('Soubor už není k dispozici. Vytvoř dokument znovu.');
  const url=URL.createObjectURL(await response.blob());
  const link=document.createElement('a');link.href=url;link.download=latestExport.name;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),10000);
}
async function generateWord() {
  if(!$('#name').value.trim()){showStep(1);$('#name').focus();throw new Error('Nejdřív napiš název stavby, pak vytvořím Word.');}
  if(!selected.size){showStep(2);throw new Error('Vyber aspoň jednu kapitolu.');}
  if(!$('#output_dir').value.trim()){showStep(3);throw new Error('Vyber složku, kam Word uložit.');}
  showStep(3);
  const button=$('#generate-main');button.disabled=true;button.textContent='Chystám dokument…';
  try {
    latestExport=await api('/api/generate',collectProject());
    const folder=latestExport.path.slice(0, latestExport.path.length-latestExport.name.length).replace(/[\\/]$/,'');
    $('#export-name').textContent=latestExport.name;
    $('#export-path').textContent='ve složce '+folder;
    $('#summary-dir').textContent=folder;
    $('#summary-file').textContent=latestExport.name;
    $('#export-tip').textContent=(/_\d{2}\.docx$/.test(latestExport.name)?'Starší verze zůstala vedle, nic se nepřepsalo. ':'')+'Ve Wordu ještě mrkni na titulní stranu a zalomení.';
    const result=$('#export-result');
    result.classList.remove('hidden','play'); void result.offsetWidth; result.classList.add('play');
    $('#step-3').classList.add('is-done');
    $('.pin[data-step="3"]').classList.add('done');
    driveCar();
    scrollToRoad();
  } finally {button.disabled=false;button.textContent='Vytvořit Word';}
}
function handle(fn) {return async e=>{try{await fn(e);}catch(error){toast(error.message,true);}};}

async function init() {
  document.body.classList.add('loading');
  try {state=await api('/api/state'); indexNodes(state.catalog.nodes); renderCoverStatic(); makeDefault(); updateLibraryCounts();}
  catch(e){toast('Aplikace se nenačetla: '+e.message,true);return;}
  finally{document.body.classList.remove('loading');}
  setCar(chosenCar()); buildRoad(); buildTraffic(); showHelp(null);
  let resizeTimer; window.addEventListener('resize',()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(()=>{buildRoad();buildTraffic();},150);});
  document.fonts.ready.then(buildRoad);
  const pop=$('#car-pop'), carBtn=$('#car-button');
  const closePop=()=>{pop.classList.add('hidden');carBtn.setAttribute('aria-expanded','false');};
  carBtn.addEventListener('click',()=>{const open=pop.classList.toggle('hidden')===false;carBtn.setAttribute('aria-expanded',String(open));if(open)pop.querySelector('[aria-checked=true]')?.focus();});
  pop.addEventListener('click',e=>{const o=e.target.closest('.opt');if(o){setCar(o.dataset.car);closePop();carBtn.focus();}});
  document.addEventListener('click',e=>{if(!e.target.closest('.carpick'))closePop();});
  document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!pop.classList.contains('hidden')){closePop();carBtn.focus();}});
  $('#chapter-tree').addEventListener('mouseover',e=>{const r=e.target.closest('.chapter-row');if(!r)return;showHelp(r.querySelector('[data-help]').dataset.help);$$('.chapter-row.is-hovered').forEach(x=>x.classList.remove('is-hovered'));r.classList.add('is-hovered');});
  $('#chapter-tree').addEventListener('mouseleave',()=>{$$('.chapter-row.is-hovered').forEach(x=>x.classList.remove('is-hovered'));showHelp(pinnedHelp);});
  $$('[data-step]').forEach(b=>b.addEventListener('click',()=>{if(!b.closest('.road'))scrollToRoad();showStep(b.dataset.step);}));
  fields.forEach(f=>$('#'+f).addEventListener('input',()=>{
    if(f==='filename')filenameEdited=true;
    if(f==='name'&&!filenameEdited)$('#filename').value=$('#name').value?'A-B '+$('#name').value:'';
    markDirty();updateSummary();
  }));
  $('#project-form').addEventListener('submit',e=>e.preventDefault());
  $('#chapter-tree').addEventListener('click',e=>{
    const expand=e.target.closest('[data-expand]');if(expand){const id=expand.dataset.expand;expanded.has(id)?expanded.delete(id):expanded.add(id);renderTree();$(`[data-expand="${id}"]`)?.focus();}
    const help=e.target.closest('[data-help]');if(help){pinnedHelp=help.dataset.help;showHelp(pinnedHelp);$$('.chapter-row').forEach(row=>row.classList.remove('is-focused'));help.closest('.chapter-row').classList.add('is-focused');}
  });
  $('#chapter-tree').addEventListener('change',e=>{
    if(!e.target.dataset.check)return;const id=e.target.dataset.check;
    leaves.get(id).forEach(leaf=>e.target.checked?selected.add(leaf):selected.delete(leaf));
    markDirty();renderTree();updateSummary();$(`[data-check="${id}"]`)?.focus();
  });
  $('#search').addEventListener('input',renderTree);$('#selected-only').addEventListener('change',renderTree);
  $('#select-all').addEventListener('click',()=>{selected=normalizeSelection(['A','B']);markDirty();renderTree();updateSummary();});
  $('#clear-selection').addEventListener('click',()=>{selected.clear();markDirty();renderTree();updateSummary();});
  $('#save-project').addEventListener('click',handle(saveProject));
  $('#new-project').addEventListener('click',()=>{if(acceptDiscard())makeDefault();});
  $('#projects-button').addEventListener('click',()=>showLibrary('projects'));
  $('#presets-button').addEventListener('click',()=>showLibrary('presets'));
  $('#load-preset').addEventListener('click',()=>showLibrary('presets'));
  $('#save-preset').addEventListener('click',()=>{if(!selected.size){toast('Nejdřív vyber kapitoly.',true);return;}$('#preset-name').value='';$('#preset-dialog').showModal();$('#preset-name').focus();});
  $('#confirm-preset').addEventListener('click',handle(async()=>{
    const name=$('#preset-name').value.trim();if(!name){$('#preset-name').focus();throw new Error('Pojmenuj předvolbu.');}
    const old=state.presets.find(p=>normal(p.name)===normal(name));
    if(old&&!window.confirm('Předvolba s tímhle názvem už existuje. Přepsat její výběr?'))return;
    const item=await api('/api/presets',{name,selected:[...selected],id:old?.id||''});
    state.presets=state.presets.filter(p=>p.id!==item.id).concat(item);updateLibraryCounts();$('#preset-dialog').close();toast('Předvolba „'+name+'“ je uložená.');
  }));
  $('#preset-name').addEventListener('keydown',e=>{if(e.key==='Enter')$('#confirm-preset').click();});
  $('#library-list').addEventListener('click',handle(async e=>{
    const load=e.target.closest('[data-load]'),del=e.target.closest('[data-delete]');
    if(load){const kind=load.dataset.kind,item=state[kind].find(x=>x.id===load.dataset.load);
      if(kind==='projects') {if(!acceptDiscard())return;project={...item};selected=normalizeSelection(item.selected);filenameEdited=true;dirty=false;fillForm();setSaved('Zakázka načtená',true);showStep(1);}
      else{selected=normalizeSelection(item.selected);markDirty();showStep(2);}
      renderTree();updateSummary();$('#library-dialog').close();toast(kind==='projects'?'Zakázka je načtená.':'Výběr kapitol je načtený, údaje o stavbě zůstaly.');
    }
    if(del){const kind=del.dataset.kind,item=state[kind].find(x=>x.id===del.dataset.delete);if(!window.confirm('Smazat „'+item.name+'“? Vytvořené Wordy zůstanou.'))return;await api('/api/delete',{kind,id:item.id});state[kind]=state[kind].filter(x=>x.id!==item.id);if(kind==='projects'&&project.id===item.id){project.id='';markDirty();}updateLibraryCounts();$('#library-dialog').close();showLibrary(kind);}
  }));
  $$('.close-dialog').forEach(b=>b.addEventListener('click',()=>b.closest('dialog').close()));
  $$('dialog').forEach(d=>d.addEventListener('click',e=>{if(e.target===d){const r=d.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)d.close();}}));
  $('#choose-folder').addEventListener('click',handle(async()=>{await browseFolder($('#output_dir').value||state.output_dir);$('#folder-dialog').showModal();}));
  $('#folder-go').addEventListener('click',handle(()=>browseFolder($('#folder-path').value)));
  $('#folder-path').addEventListener('keydown',e=>{if(e.key==='Enter')$('#folder-go').click();});
  $('#folder-up').addEventListener('click',handle(()=>browseFolder(folderState.parent)));
  $('#folder-dialog').addEventListener('click',handle(e=>{const b=e.target.closest('[data-folder]');if(b)return browseFolder(b.dataset.folder);}));
  $('#confirm-folder').addEventListener('click',handle(async()=>{await browseFolder($('#folder-path').value);$('#output_dir').value=folderState.path;markDirty();updateSummary();$('#folder-dialog').close();}));
  $('#choose-image').addEventListener('click',()=>$('#image-file').click());
  $('#image-file').addEventListener('change',handle(async e=>{
    const file=e.target.files[0];if(!file)return;if(file.size>10*1024*1024)throw new Error('Vyber obrázek menší než 10 MB.');
    const content=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result.split(',')[1]);reader.onerror=()=>reject(new Error('Obrázek se nepodařilo načíst.'));reader.readAsDataURL(file);});
    const result=await api('/api/image',{content});project.cover_image=result.path;$('#image-status').textContent='Vlastní obrázek: '+file.name;$('#reset-image').classList.remove('hidden');markDirty();loadCover();toast('Obrázek je na titulní straně.');
  }));
  $('#reset-image').addEventListener('click',()=>{project.cover_image='';$('#image-file').value='';$('#image-status').textContent='Ten z předlohy.';$('#reset-image').classList.add('hidden');markDirty();loadCover();});
  $('#generate-main').addEventListener('click',handle(generateWord));
  $('#download-word').addEventListener('click',handle(downloadWord));
  $('#open-file').addEventListener('click',handle(()=>api('/api/open',{id:latestExport.id,target:'file'})));
  $('#open-folder').addEventListener('click',handle(()=>api('/api/open',{id:latestExport.id,target:'folder'})));
  $$('[data-edit-export]').forEach(b=>b.addEventListener('click',()=>{clearExport();$('#output_dir').focus();}));
  document.addEventListener('keydown',e=>{if(e.key==='/'&&!['INPUT','TEXTAREA'].includes(document.activeElement.tagName)&&!$('dialog[open]')){e.preventDefault();showStep(2);$('#search').focus();}});
  window.addEventListener('beforeunload',e=>{if(dirty){e.preventDefault();e.returnValue='';}});
}
init();
