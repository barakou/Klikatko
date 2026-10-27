'use strict';
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const token = $('meta[name="klikatko-token"]').content;
const fields = ['name','number','location','investor','stage','author','date','output_dir','filename'];
let state, project = {}, selected = new Set(), expanded = new Set(['A','B']), nodes = new Map(), parents = new Map(), leaves = new Map();
let currentStep = 1, dirty = false, filenameEdited = false, toastTimer, folderState, latestExport;
const normal = (s) => String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
const escapeHTML = (s) => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

async function api(path, body) {
  const response = await fetch(path, {method: body === undefined ? 'GET' : 'POST', headers: {'X-Klikatko-Token': token, 'Content-Type':'application/json'}, body: body === undefined ? undefined : JSON.stringify(body)});
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || 'Operace se nezdařila.');
  return result;
}
function toast(message, error=false) {
  clearTimeout(toastTimer);
  $('#toast').textContent = message;
  $('#toast').classList.remove('hidden');
  $('#toast').classList.toggle('error', error);
  toastTimer = setTimeout(() => $('#toast').classList.add('hidden'), error ? 10000 : 4200);
}
function markDirty() { dirty = true; $('#save-state').textContent = 'Neuložené změny'; $('#export-result').classList.add('hidden'); }
function acceptDiscard() { return !dirty || window.confirm('Máte neuložené změny zakázky. Opravdu je chcete opustit?'); }
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
  project = {id:'',name:'',number:'',location:'',investor:'',author:'Ing. Koudelka',date:today.toLocaleDateString('cs-CZ',{month:'long',year:'numeric'}),stage:'Dokumentace pro povolení záměru',output_dir:state.output_dir,filename:'',cover_image:''};
  selected = normalizeSelection(['A1','A2','A3']);
  filenameEdited=false; dirty=false; fillForm(); expanded = new Set(['A','B']);
  $('#search').value=''; $('#selected-only').checked=false;
  renderTree(); updateSummary(); showStep(1); $('#save-state').textContent='Dosud neuloženo';
}
function fillForm() {
  fields.forEach(f => $('#' + f).value = project[f] || '');
  $('#image-status').textContent=project.cover_image ? 'Použijeme váš vlastní obrázek.' : 'Zachováme rozložení a obrázek z dodané předlohy.';
  $('#reset-image').classList.toggle('hidden',!project.cover_image);
  $('#export-result').classList.add('hidden');
}
function showStep(step) {
  currentStep=Number(step);
  for (let n=1;n<=3;n++) $('#step-'+n).classList.toggle('hidden',n!==currentStep);
  $$('.step').forEach(b=> { b.classList.toggle('active',Number(b.dataset.step)===currentStep); b.classList.toggle('completed',Number(b.dataset.step)<currentStep); b.setAttribute('aria-current',Number(b.dataset.step)===currentStep?'step':'false'); });
  updateSummary();
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
      html+=`<div class="chapter-row ${depth===0?'root':''}" style="--depth:${depth}">
      ${branch?`<button class="expand" data-expand="${n.id}" aria-expanded="${open}" aria-label="${open?'Sbalit':'Rozbalit'} ${escapeHTML(n.code)}">▸</button>`:'<span class="expand-spacer"></span>'}
      <input type="checkbox" data-check="${n.id}" ${amount===leafIds.length?'checked':''} aria-label="Vybrat ${escapeHTML(n.id+' '+n.title)}" ${amount>0&&amount<leafIds.length?'data-partial="true"':''}>
      <button class="chapter-name" data-help="${n.id}"><strong>${escapeHTML(n.code)}</strong><span>${escapeHTML(n.title)}</span></button>
      ${branch?`<span class="chapter-total">${amount} / ${leafIds.length}</span>`:''}</div>`;
      if(branch && open)walk(n.children,depth+1,ancestorMatch||ownMatch);
    }
  }
  walk(state.catalog.nodes);
  $('#chapter-tree').innerHTML=html||'<div class="empty-state">Žádné odpovídající kapitoly.<br>Zkuste jiné hledání nebo zrušte filtr.</div>';
  $$('[data-partial]').forEach(el=>el.indeterminate=true);
  $('#selection-label').textContent=`${selected.size} vybraných položek · nadpisy přidáme automaticky`;
}
function updateSummary() {
  collectProject();
  $('#breadcrumb-name').textContent=project.name||'Nová zakázka';
  $('#preview-name').textContent=project.name||'Název vaší stavby';
  $('#preview-stage').textContent=project.stage||'Druh dokumentace';
  $('#preview-date').textContent=project.date;
  $('#preview-author').textContent=project.author;
  $('#summary-count').textContent=counts();
  const includedTop=[...nodes.values()].filter(n=>parents.get(n.id) && ['A','B'].includes(parents.get(n.id)) && included(n.id));
  $('#summary-outline').innerHTML=includedTop.slice(0,12).map(n=>`<span title="${escapeHTML(n.title)}">${escapeHTML(n.code)}</span>`).join('')+(includedTop.length>12?`<span>+${includedTop.length-12}</span>`:'');
  const checks=[{ok:Boolean(project.name),text:project.name?'Název stavby je vyplněný.':'Doplňte název stavby.'}, {ok:selected.size>0,text:selected.size?`Vybráno ${selected.size} položek, s nadpisy celkem ${counts()} částí.`:'Vyberte alespoň jednu kapitolu.'}, {ok:Boolean(project.output_dir),text:project.output_dir?'Výstupní složka je nastavená.':'Vyberte složku pro uložení.'}];
  $('#review-checks').innerHTML=checks.map(c=>`<li class="${c.ok?'':'warning'}">${c.ok?'✓':'!'} &nbsp;${escapeHTML(c.text)}</li>`).join('');
}
function updateLibraryCounts() { $('#projects-count').textContent=state.projects.length;$('#presets-count').textContent=state.presets.length; }
async function saveProject() {
  if(!$('#name').value.trim()){showStep(1);$('#name').focus();throw new Error('Nejdřív zadejte název stavby.');}
  project=await api('/api/projects',collectProject());
  state.projects=state.projects.filter(p=>p.id!==project.id).concat(project);
  dirty=false;$('#save-state').textContent='✓ Zakázka uložená';updateLibraryCounts();toast('Zakázka je uložená. Můžete se k ní kdykoli vrátit.');
}
function showLibrary(kind) {
  $('#library-title').textContent=kind==='projects'?'Uložené zakázky':'Moje předvolby';
  const items=state[kind];
  $('#library-list').innerHTML=items.length?items.map(item=>`<div class="library-entry"><div><strong>${escapeHTML(item.name)}</strong><small>${kind==='projects'?escapeHTML(item.number||'Bez čísla zakázky')+' · ':''}${item.selected.length} vybraných položek</small></div><button class="button outline small" data-load="${item.id}" data-kind="${kind}">Načíst</button><button class="icon-button" data-delete="${item.id}" data-kind="${kind}" aria-label="Smazat ${escapeHTML(item.name)}">×</button></div>`).join(''):`<div class="empty-state">${kind==='projects'?'Zatím tu nejsou žádné zakázky.<br>První uložíte tlačítkem „Uložit zakázku“.':'Zatím tu nejsou žádné předvolby.<br>Vyberte kapitoly a zvolte „Uložit výběr“.'}</div>`;
  $('#library-dialog').showModal();
}
async function browseFolder(path) {
  folderState=await api('/api/folders?path='+encodeURIComponent(path));
  $('#folder-path').value=folderState.path;
  $('#folder-list').innerHTML=folderState.children.length?folderState.children.map(d=>`<button class="folder-entry" data-folder="${escapeHTML(d.path)}">▱ &nbsp; ${escapeHTML(d.name)}</button>`).join(''):'<div class="empty-state">Tato složka nemá další podsložky.<br>Můžete ji vybrat pro uložení zprávy.</div>';
  $('#folder-drives').innerHTML=folderState.drives.map(d=>`<button class="text-button" data-folder="${escapeHTML(d)}">${escapeHTML(d)}</button>`).join('');
}
async function downloadWord() {
  if(!latestExport)return;
  const response=await fetch('/api/download?id='+encodeURIComponent(latestExport.id),{headers:{'X-Klikatko-Token':token}});
  if(!response.ok)throw new Error('Soubor není dostupný. Vytvořte dokument znovu.');
  const url=URL.createObjectURL(await response.blob());
  const link=document.createElement('a');link.href=url;link.download=latestExport.name;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),10000);
}
async function generateWord() {
  if(!$('#name').value.trim()){showStep(1);$('#name').focus();throw new Error('Vyplňte název stavby a pak vytvořte Word.');}
  if(!selected.size){showStep(2);throw new Error('Vyberte alespoň jednu kapitolu.');}
  if(!$('#output_dir').value.trim()){showStep(3);throw new Error('Vyberte výstupní složku.');}
  showStep(3);
  const buttons=[$('#generate-main'),$('#generate-sidebar')];buttons.forEach(b=>{b.disabled=true;b.dataset.original=b.innerHTML;b.textContent='Připravuji dokument…';});
  try {
    latestExport=await api('/api/generate',collectProject());
    $('#export-path').textContent=latestExport.path;
    $('#export-result').classList.remove('hidden');
    toast('Word je připravený. Můžete jej stáhnout a doplnit.');
    $('#export-result').scrollIntoView({behavior:'smooth',block:'nearest'});
  } finally {buttons.forEach(b=>{b.disabled=false;b.innerHTML=b.dataset.original;});}
}
function handle(fn) {return async e=>{try{await fn(e);}catch(error){toast(error.message,true);}};}

async function init() {
  document.body.classList.add('loading');
  try {state=await api('/api/state'); indexNodes(state.catalog.nodes); makeDefault();updateLibraryCounts();}
  catch(e){toast('Aplikaci se nepodařilo načíst: '+e.message,true);return;}
  finally{document.body.classList.remove('loading');}
  $$('[data-step]').forEach(b=>b.addEventListener('click',()=>showStep(b.dataset.step)));
  fields.forEach(f=>$('#'+f).addEventListener('input',()=>{
    if(f==='filename')filenameEdited=true;
    if(f==='name'&&!filenameEdited)$('#filename').value=$('#name').value?'A-B '+$('#name').value:'';
    markDirty();updateSummary();
  }));
  $('#project-form').addEventListener('submit',e=>e.preventDefault());
  $('#chapter-tree').addEventListener('click',e=>{
    const expand=e.target.closest('[data-expand]');if(expand){const id=expand.dataset.expand;expanded.has(id)?expanded.delete(id):expanded.add(id);renderTree();$(`[data-expand="${id}"]`)?.focus();}
    const help=e.target.closest('[data-help]');if(help){const n=nodes.get(help.dataset.help);$('#help-title').textContent=n.id+' — '+n.title;$('#help-text').textContent=n.hint;$$('.chapter-row').forEach(row=>row.classList.remove('is-focused'));help.closest('.chapter-row').classList.add('is-focused');}
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
  $('#save-preset').addEventListener('click',()=>{if(!selected.size){toast('Nejdřív vyberte kapitoly.',true);return;}$('#preset-name').value='';$('#preset-dialog').showModal();$('#preset-name').focus();});
  $('#confirm-preset').addEventListener('click',handle(async()=>{
    const name=$('#preset-name').value.trim();if(!name){$('#preset-name').focus();throw new Error('Pojmenujte svou předvolbu.');}
    const old=state.presets.find(p=>normal(p.name)===normal(name));
    if(old&&!window.confirm('Předvolba tohoto názvu už existuje. Chcete aktualizovat její výběr?'))return;
    const item=await api('/api/presets',{name,selected:[...selected],id:old?.id||''});
    state.presets=state.presets.filter(p=>p.id!==item.id).concat(item);updateLibraryCounts();$('#preset-dialog').close();toast('Předvolba „'+name+'“ je uložená.');
  }));
  $('#preset-name').addEventListener('keydown',e=>{if(e.key==='Enter')$('#confirm-preset').click();});
  $('#library-list').addEventListener('click',handle(async e=>{
    const load=e.target.closest('[data-load]'),del=e.target.closest('[data-delete]');
    if(load){const kind=load.dataset.kind,item=state[kind].find(x=>x.id===load.dataset.load);
      if(kind==='projects') {if(!acceptDiscard())return;project={...item};selected=normalizeSelection(item.selected);filenameEdited=true;dirty=false;fillForm();$('#save-state').textContent='✓ Zakázka načtená';showStep(1);}
      else{selected=normalizeSelection(item.selected);markDirty();showStep(2);}
      renderTree();updateSummary();$('#library-dialog').close();toast(kind==='projects'?'Zakázka je načtená.':'Výběr kapitol je načtený. Údaje zakázky zůstaly zachované.');
    }
    if(del){const kind=del.dataset.kind,item=state[kind].find(x=>x.id===del.dataset.delete);if(!window.confirm('Smazat „'+item.name+'“? Vytvořené Wordy zůstanou zachované.'))return;await api('/api/delete',{kind,id:item.id});state[kind]=state[kind].filter(x=>x.id!==item.id);if(kind==='projects'&&project.id===item.id){project.id='';markDirty();}updateLibraryCounts();$('#library-dialog').close();showLibrary(kind);}
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
    const file=e.target.files[0];if(!file)return;if(file.size>10*1024*1024)throw new Error('Vyberte obrázek menší než 10 MB.');
    const content=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result.split(',')[1]);reader.onerror=()=>reject(new Error('Obrázek se nepodařilo přečíst.'));reader.readAsDataURL(file);});
    const result=await api('/api/image',{content});project.cover_image=result.path;$('#image-status').textContent='Vlastní obrázek: '+file.name;$('#reset-image').classList.remove('hidden');markDirty();toast('Obrázek je připravený pro titulní stranu.');
  }));
  $('#reset-image').addEventListener('click',()=>{project.cover_image='';$('#image-file').value='';$('#image-status').textContent='Zachováme rozložení a obrázek z dodané předlohy.';$('#reset-image').classList.add('hidden');markDirty();});
  $('#generate-main').addEventListener('click',handle(generateWord));$('#generate-sidebar').addEventListener('click',handle(generateWord));$('#download-word').addEventListener('click',handle(downloadWord));
  document.addEventListener('keydown',e=>{if(e.key==='/'&&!['INPUT','TEXTAREA'].includes(document.activeElement.tagName)&&!$('dialog[open]')){e.preventDefault();showStep(2);$('#search').focus();}});
  window.addEventListener('beforeunload',e=>{if(dirty){e.preventDefault();e.returnValue='';}});
}
init();
