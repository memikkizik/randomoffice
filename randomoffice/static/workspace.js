/* Shared desktop workspace: document tools, persistent formats and Windows integration. */
let pageLayout = { orientation: 'portrait', paper: 'A4', margin: 'normal' };
let startupDocumentRequested = false;
let restoringNativeWorkspace = false;
let nativeWorkspaceReady = false;
let workspaceUserEdited = false;
let nativeStateWriteChain = Promise.resolve();
let workspaceFindIndex = -1;
let workspaceFindMatches = [];
const documentSourceTokens = {};
const cleanDocumentFingerprints = {};
const nativeDirtyReported = {};
let recoveredTrackingState = null;
let documentTrackingReady = false;

function documentFingerprint(app) {
  const value=app==='word'?[document.getElementById('word-editor').innerHTML,pageLayout]:app==='excel'?excelModel.sheets:pptDeck;
  const text=JSON.stringify(value);
  let a=2166136261,b=5381;
  for(let i=0;i<text.length;i++){const c=text.charCodeAt(i);a=Math.imul(a^c,16777619);b=Math.imul(b,33)^c;}
  return `${text.length}:${a>>>0}:${b>>>0}`;
}
function initializeDocumentTracking() {
  for(const app of ['word','excel','powerpoint'])cleanDocumentFingerprints[app]=documentFingerprint(app);
  documentTrackingReady=true;
  if(recoveredTrackingState)restoreDocumentTracking(recoveredTrackingState);
  refreshUnsavedState(true);
}
function restoreDocumentTracking(state) {
  for(const app of ['word','excel','powerpoint']){
    const baseline=state.cleanFingerprints?.[app];
    if(typeof baseline==='string')cleanDocumentFingerprints[app]=baseline;
    else if(state.wordHtml||state.excelData||state.pptData)cleanDocumentFingerprints[app]='recovered-draft';
    if(typeof state.sourceTokens?.[app]==='string')documentSourceTokens[app]=state.sourceTokens[app];
  }
  refreshUnsavedState(true);
}
function isDocumentDirty(app) {
  return documentTrackingReady&&cleanDocumentFingerprints[app]!==documentFingerprint(app);
}
function markDocumentSaved(app,snapshot=documentFingerprint(app)) {
  cleanDocumentFingerprints[app]=snapshot;
  refreshUnsavedState(true);
}
function refreshUnsavedState(force=false) {
  if(!documentTrackingReady)return;
  for(const app of ['word','excel','powerpoint']){
    const dirty=isDocumentDirty(app);
    if(force||nativeDirtyReported[app]!==dirty){
      nativeDirtyReported[app]=dirty;
      window.pywebview?.api?.set_unsaved_state?.(app,dirty).catch(()=>{delete nativeDirtyReported[app];});
    }
  }
  const dirty=isDocumentDirty(currentOfficeApp);
  document.getElementById('doc-name-editor').setAttribute('aria-description',dirty?'Kaydedilmemiş değişiklikler var':'Dosya kaydedildi');
  document.title=`${dirty?'● ':''}${documentNames[currentOfficeApp]} — RandomOffice`;
  const badge=document.getElementById('unsaved-indicator');
  if(badge){badge.hidden=!dirty;badge.title='Dosyada kaydedilmemiş değişiklikler var';}
}
function confirmDiscardDocument(app) {
  const dirty=app?isDocumentDirty(app):['word','excel','powerpoint'].some(isDocumentDirty);
  return !dirty||confirm('Kaydedilmemiş değişiklikler var. Aynı türdeki belgeyi değiştirirseniz bu değişiklikler silinir. Devam etmek istiyor musunuz?');
}
window.addEventListener('beforeunload',event=>{
  if(window.pywebview?.api)return;
  if(['word','excel','powerpoint'].some(isDocumentDirty)){event.preventDefault();event.returnValue='';}
});

function setupProfessionalWorkspace() {
  document.getElementById('word-editor').innerHTML='<p><br></p>';
  document.addEventListener('input',event=>{if(event.target.closest('#word-editor,#excel-workspace,#ppt-workspace,#doc-name-editor'))workspaceUserEdited=true;});
  document.querySelectorAll('select[onchange*="fontSize"] option').forEach(option=>{option.textContent=({1:8,2:10,3:12,4:14,5:18,6:24,7:36})[option.value];});
  const toolbar = document.querySelector('.ribbon-toolbar');
  const tabs = document.querySelector('.ribbon-tabs-container');
  const addTab = (id, text, handler) => {
    const button = document.createElement('button');
    button.id = id; button.className = 'ribbon-tab-btn'; button.textContent = text;
    button.onclick = handler;
    tabs.insertBefore(button, document.querySelector('.app-switcher-dock'));
  };
  addTab('tab-layout', 'Düzen', () => selectRibbonTab('layout'));
  addTab('tab-review', 'Bul ve Değiştir', () => openWorkspacePanel('find'));
  const layout = document.createElement('div');
  layout.id = 'ribbon-sec-layout'; layout.className = 'ribbon-section'; layout.style.display = 'none';
  layout.innerHTML = `<div class="ribbon-group-box"><div class="ribbon-row">
    <label>Yön <select id="layout-orientation" class="r-select" onchange="setPageLayout('orientation',this.value)"><option value="portrait">Dikey</option><option value="landscape">Yatay</option></select></label>
    <label>Kağıt <select id="layout-paper" class="r-select" onchange="setPageLayout('paper',this.value)"><option>A4</option><option>Letter</option></select></label>
    <label>Kenar <select id="layout-margin" class="r-select" onchange="setPageLayout('margin',this.value)"><option value="normal">Normal</option><option value="narrow">Dar</option><option value="wide">Geniş</option></select></label>
    </div><div class="group-caption">Sayfa düzeni • Word ve baskı</div></div>
    <div class="ribbon-group-box"><div class="ribbon-row">
      <button class="r-tool" onclick="editWordTable('row')">Satır ekle</button><button class="r-tool" onclick="editWordTable('column')">Sütun ekle</button><button class="r-tool" onclick="editWordTable('delete-row')">Satırı sil</button>
    </div><div class="group-caption">Seçili tablo</div></div>`;
  toolbar.appendChild(layout);
  const data = document.createElement('div'); data.className = 'ribbon-group-box';
  data.innerHTML = `<div class="ribbon-row"><button class="r-tool" onclick="sortExcelRows(false)">A → Z sırala</button><button class="r-tool" onclick="sortExcelRows(true)">Z → A sırala</button><button class="r-tool" onclick="toggleSheetFilter()">Filtre</button><button class="r-tool" onclick="toggleFreezeRow()">İlk satırı sabitle</button></div><div class="group-caption">Veri araçları • İlk satır başlık olarak korunur</div>`;
  document.getElementById('ribbon-sec-formulas').appendChild(data);
  const right = document.querySelector('.top-right-cluster');
  const saveAs = document.createElement('button'); saveAs.id = 'save-as-button'; saveAs.className = 'header-action-btn';
  saveAs.textContent = 'Farklı Kaydet'; saveAs.title = 'Farklı Kaydet (Ctrl+Shift+S)'; saveAs.onclick = () => saveDocument(true); right.prepend(saveAs);
  const settings = document.createElement('button'); settings.id = 'workspace-settings-button'; settings.className = 'header-action-btn';
  settings.textContent = 'Ayarlar'; settings.onclick = () => openWorkspacePanel('settings'); right.appendChild(settings);
  document.querySelector('.autosave-pill').title = 'Yerel taslak otomatik korunur. Office dosyasına yazmak için Kaydet kullanın.';
  document.querySelector('.autosave-pill span').textContent = 'Yerel taslak';
  const unsaved=document.createElement('span');unsaved.id='unsaved-indicator';unsaved.textContent='●';unsaved.hidden=true;unsaved.style.color='#e4a51c';
  document.getElementById('doc-name-editor').after(unsaved);
  const panel = document.createElement('aside'); panel.id = 'workspace-panel'; panel.hidden = true; panel.setAttribute('aria-label','Belge araçları');
  panel.innerHTML = `<div class="workspace-panel-heading"><strong id="workspace-panel-title">Belge araçları</strong><button class="r-tool" onclick="closeWorkspacePanel()" aria-label="Paneli kapat">✕</button></div>
    <section id="workspace-find-section"><label>Bul<input id="workspace-find-input" type="search" placeholder="Belgede ara" oninput="findWorkspaceText(true)" onkeydown="if(event.key==='Enter')findWorkspaceText(false)"></label>
      <div id="workspace-find-status" role="status">Aranacak metni yazın.</div><button class="panel-action" onclick="findWorkspaceText(false)">Sonraki eşleşme</button>
      <label>Yeni metin<input id="workspace-replace-input" placeholder="Yerine yazılacak metin"></label><button class="panel-action" onclick="replaceWorkspaceText(false)">Seçileni değiştir</button><button class="panel-action" onclick="replaceWorkspaceText(true)">Tümünü değiştir</button></section>
    <section id="workspace-settings-section" hidden><p class="panel-description">RandomOffice Masaüstü</p><p id="desktop-file-status">Tarayıcı modu</p><p id="windows-install-status">Windows entegrasyonu denetleniyor.</p><button id="open-default-apps-button" class="panel-action" onclick="openWindowsDefaults()">Windows varsayılan uygulamaları</button><p class="panel-description">DOCX, XLSX ve PPTX için RandomOffice'i seçebilirsiniz. Windows seçimi size bırakır.</p>
      <hr><label><input id="workspace-compact-check" type="checkbox" onchange="setCompactRibbon(this.checked)"> Kompakt araç çubuğu</label><button class="panel-action" onclick="fitWorkspace()">Sayfayı pencereye sığdır</button><p class="panel-description">Ctrl+S Kaydet · Ctrl+Shift+S Farklı Kaydet · Ctrl+F Bul · Ctrl+H Değiştir · Ctrl+O Aç</p></section>`;
  document.getElementById('workspace-stage').appendChild(panel);
  const filter = document.createElement('input'); filter.id = 'sheet-filter-input'; filter.type = 'search'; filter.placeholder = 'Satırlarda filtrele…'; filter.hidden = true; filter.setAttribute('aria-label','Tablo satırlarını filtrele'); filter.oninput = filterSheetRows;
  document.querySelector('.excel-fx-bar').appendChild(filter);
  document.getElementById('excel-sheet-grid').addEventListener('paste', handleSheetPaste);
  const fit = document.createElement('button'); fit.className='r-tool'; fit.id='fit-window-button'; fit.textContent='Pencereye sığdır'; fit.onclick=fitWorkspace;
  document.querySelector('.zoom-ctrl').prepend(fit);
  document.addEventListener('keydown', event => {
    if (!(event.ctrlKey || event.metaKey)) return;
    const key=event.key.toLowerCase();
    if (key==='f'||key==='h') {event.preventDefault();openWorkspacePanel('find'); if(key==='h')document.getElementById('workspace-replace-input').focus();}
    if (key==='o') {event.preventDefault();triggerOpenDialog();}
  });
  const printSelects = [...document.querySelectorAll('#hub-section-print select')];
  if(printSelects.length===4) {
    printSelects[0].parentElement.replaceChildren(Object.assign(document.createElement('p'),{className:'panel-description',textContent:'Sayfa aralığı ve kopya sayısı sistem yazdırma penceresinde seçilir.'}));
    printSelects[1].id='print-paper-select'; printSelects[1].options[0].value='A4';printSelects[1].options[1].value='Letter';
    printSelects[1].onchange=()=>setPageLayout('paper',printSelects[1].value);
    printSelects[2].onchange=()=>setPageLayout('orientation',printSelects[2].value);
    printSelects[3].id='print-color-select';printSelects[3].options[0].value='color';printSelects[3].options[1].value='grayscale';
    printSelects[3].onchange=()=>document.body.classList.toggle('print-grayscale',printSelects[3].value==='grayscale');
  }
  applyPageLayout();
  new ResizeObserver(() => {if(document.body.dataset.zoomMode==='fit')fitWorkspace();}).observe(document.getElementById('workspace-stage'));
}

async function ingestDesktopStartup() {
  const api=window.pywebview?.api;
  if(!api?.get_startup_document || startupDocumentRequested) return;
  startupDocumentRequested=true;
  restoringNativeWorkspace=true;
  try {
    if(api.read_workspace_state){
      const state=await api.read_workspace_state();
      if(state&&!state.error&&!workspaceUserEdited)restoreEditorState(state.state||state);
      if(state?.error)showToast(state.error);
    }
    const data=await api.get_startup_document();
    if(data?.error) showToast(data.error);
    else if(data?.app) ingestDocumentData(data);
    await updateDesktopState();
  } catch(error) {showToast('Başlangıç dosyası açılamadı: '+error.message);}
  finally {restoringNativeWorkspace=false;nativeWorkspaceReady=true;persistState();}
}

function restoreEditorState(state){
  if(!state || typeof state!=='object' || !['word','excel','powerpoint'].includes(state.app))return;
  restoringState=true;
  try {
    if(typeof state.wordHtml==='string')document.getElementById('word-editor').replaceChildren(safeWordFragment(state.wordHtml));
    if(state.excelData?.sheets?.length)excelModel=state.excelData;
    if(state.pptData?.length)pptDeck=state.pptData;
    if(state.names)Object.assign(documentNames,state.names);
    if(state.title)documentNames[state.app]=state.title;
    if(state.pageLayout){Object.assign(pageLayout,state.pageLayout);applyPageLayout();}
    currentSheetIndex=Math.max(0,Math.min(state.sheetIndex||0,excelModel.sheets.length-1));
    renderExcelTabs();initExcelGrid();activeSlideIndex=0;renderPptDeckSidebar();activateOfficeApp(state.app);
    document.body.classList.toggle('compact-ribbon',Boolean(state.compactRibbon));
    document.getElementById('workspace-compact-check').checked=Boolean(state.compactRibbon);
  } finally {restoringState=false;}
  restoreDocumentTracking(state);
}
function persistNativeWorkspace(state){
  const api=window.pywebview?.api;
  if(!nativeWorkspaceReady||!api?.write_workspace_state)return false;
  const snapshot=JSON.parse(JSON.stringify(state));
  nativeStateWriteChain=nativeStateWriteChain.catch(()=>{}).then(()=>api.write_workspace_state(snapshot)).then(result=>{
    if(!result?.success)throw new Error(result?.error||'Yerel taslak kaydedilemedi');
    document.getElementById('cloud-sync-status').innerText=`✓ Cihazda taslak (${state.updated})`;
  }).catch(error=>{document.getElementById('cloud-sync-status').innerText='Taslak kaydedilemedi — dosyanızı kaydedin';});
  return true;
}

async function updateDesktopState() {
  const label=document.getElementById('desktop-file-status');
  try {
    const api=window.pywebview?.api;
    if(!api?.get_desktop_state) {if(label)label.textContent='Tarayıcı modu — dosyalar indirilir.';return;}
    const state=await api.get_desktop_state();
    const file=state.current_files?.[currentOfficeApp];
    if(label)label.textContent=file?.saved_path || 'Yeni dosya — ilk kayıtta konum seçilecek.';
    document.querySelector('.btn-save-direct').title=file?.saved_path ? 'Dosyadaki değişiklikleri kaydet (Ctrl+S)' : 'Kaydetme konumu seç (Ctrl+S)';
  } catch(error) {if(label)label.textContent='Masaüstü dosya durumu okunamadı.';}
}

function updateWorkspaceContext() {
  const word=currentOfficeApp==='word';
  const tab=document.getElementById('tab-layout');if(tab)tab.hidden=!word;
  document.getElementById('ribbon-styles-box').hidden=!word;
  const mode=document.getElementById('sb-mode-info');mode.textContent=currentOfficeApp==='excel'?'Çalışma sayfası':currentOfficeApp==='powerpoint'?'Slayt düzeni':'Sayfa düzeni';
  updateDesktopState();
}

async function openWorkspacePanel(section) {
  const panel=document.getElementById('workspace-panel');panel.hidden=false;
  document.getElementById('workspace-find-section').hidden=section!=='find';
  document.getElementById('workspace-settings-section').hidden=section!=='settings';
  document.getElementById('workspace-panel-title').textContent=section==='find'?'Bul ve Değiştir':'Uygulama Ayarları';
  if(section==='find') {document.getElementById('workspace-find-input').focus();return;}
  await updateDesktopState();
  const button=document.getElementById('open-default-apps-button');
  if(!window.pywebview?.api?.get_windows_integration_state) {button.disabled=true;document.getElementById('windows-install-status').textContent='Bu ayar masaüstü sürümünde kullanılabilir.';return;}
  try {
    const state=await window.pywebview.api.get_windows_integration_state();
    button.disabled=!state.supported;
    document.getElementById('windows-install-status').textContent=state.registered?'Windows dosya türleri için kayıtlı.':"Windows'a kaydolmak için indirme paketindeki kurulum dosyasını çalıştırın.";
  } catch(error) {document.getElementById('windows-install-status').textContent='Windows kurulum bilgisi alınamadı.';}
}
function closeWorkspacePanel(){document.getElementById('workspace-panel').hidden=true;}
async function openWindowsDefaults(){try{const result=await window.pywebview.api.open_default_apps_settings();if(!result?.success)throw new Error(result?.error||'Ayarlar açılamadı');}catch(error){showToast(error.message);}}
function setCompactRibbon(compact){document.body.classList.toggle('compact-ribbon',compact);persistState();}

function setPageLayout(key,value){
  const choices={orientation:['portrait','landscape'],paper:['A4','Letter'],margin:['normal','narrow','wide']};
  if(!choices[key]?.includes(value))return;
  pageLayout[key]=value;applyPageLayout();persistState();
  if(document.getElementById('backstage-hub').style.display==='flex')populatePrintPreview();
}
function applyPageLayout(){
  document.body.dataset.pageOrientation=pageLayout.orientation;
  document.body.dataset.pageMargin=pageLayout.margin;
  const dimensions=pageLayout.paper==='Letter'?[216,279]:[210,297];
  const width=pageLayout.orientation==='landscape'?dimensions[1]:dimensions[0];
  document.documentElement.style.setProperty('--paper-width',width+'mm');
  for(const [id,value] of [['layout-orientation',pageLayout.orientation],['layout-paper',pageLayout.paper],['layout-margin',pageLayout.margin],['print-orientation-select',pageLayout.orientation],['print-paper-select',pageLayout.paper]]) {const select=document.getElementById(id);if(select)select.value=value;}
  let style=document.getElementById('workspace-print-settings');
  if(!style){style=document.createElement('style');style.id='workspace-print-settings';document.head.appendChild(style);}
  const margin={normal:20,narrow:10,wide:30}[pageLayout.margin];
  style.textContent=`@page {size: ${pageLayout.paper} ${pageLayout.orientation};margin:${margin}mm;}`;
}
function fitWorkspace(){
  const stage=document.getElementById(currentOfficeApp==='powerpoint'?'ppt-workspace':currentOfficeApp==='word'?'word-workspace':'excel-workspace');
  const target=document.getElementById(currentOfficeApp==='powerpoint'?'slide-canvas-card':currentOfficeApp==='word'?'word-editor':'excel-sheet-grid');
  if(currentOfficeApp==='excel'){handleZoomScale(100);document.body.dataset.zoomMode='fit';return;}
  const raw=target.offsetWidth;
  const available=currentOfficeApp==='powerpoint'?document.querySelector('.ppt-canvas-stage').clientWidth:stage.clientWidth;
  const scale=Math.min(1.25,Math.max(.35,(available-48)/raw));
  handleZoomScale(Math.round(scale*100));document.body.dataset.zoomMode='fit';
}

function findWorkspaceText(reset=false){
  const query=document.getElementById('workspace-find-input').value.toLocaleLowerCase('tr-TR');
  workspaceFindMatches=[];
  if(query){
    if(currentOfficeApp==='excel'){
      document.querySelectorAll('#excel-sheet-grid td').forEach(cell=>{if(cell.innerText.toLocaleLowerCase('tr-TR').includes(query))workspaceFindMatches.push({cell});});
    }else{
      const roots=currentOfficeApp==='word'?[document.getElementById('word-editor')]:[document.getElementById('ppt-slide-header'),document.getElementById('ppt-slide-body')];
      roots.forEach(root=>{const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);let node;
        while(node=walker.nextNode()){const lower=node.textContent.toLocaleLowerCase('tr-TR');let index=0;
          while((index=lower.indexOf(query,index))!==-1){workspaceFindMatches.push({node,start:index,end:index+query.length});index+=query.length;}}
      });
    }
  }
  if(reset)workspaceFindIndex=-1;
  if(workspaceFindMatches.length){workspaceFindIndex=(workspaceFindIndex+1)%workspaceFindMatches.length;const match=workspaceFindMatches[workspaceFindIndex];
    if(match.cell){match.cell.focus();match.cell.scrollIntoView({block:'nearest'});}else{const range=document.createRange();range.setStart(match.node,match.start);range.setEnd(match.node,match.end);const selection=getSelection();selection.removeAllRanges();selection.addRange(range);editorSelection=range.cloneRange();match.node.parentElement.scrollIntoView({block:'nearest'});}
  }
  document.getElementById('workspace-find-status').textContent=query?`${workspaceFindMatches.length?workspaceFindIndex+1:0} / ${workspaceFindMatches.length} eşleşme`:'Aranacak metni yazın.';
}
function replaceWorkspaceText(all){
  const query=document.getElementById('workspace-find-input').value;
  if(!query)return;
  const replacement=document.getElementById('workspace-replace-input').value;
  const escaped=query.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  const expression=new RegExp(escaped,all?'giu':'iu');
  if(all){
    if(currentOfficeApp==='excel'){excelModel.sheets[currentSheetIndex].rows.forEach(row=>row.forEach((value,index)=>{if(typeof value==='string'&&!value.startsWith('='))row[index]=value.replace(expression,()=>replacement);}));renderExcelTableDOM();}
    else {const roots=currentOfficeApp==='word'?[document.getElementById('word-editor')]:[document.getElementById('ppt-slide-header'),document.getElementById('ppt-slide-body')];roots.forEach(root=>{const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);let node;while(node=walker.nextNode())node.textContent=node.textContent.replace(expression,()=>replacement);});}
  }else{
    const match=workspaceFindMatches[workspaceFindIndex];if(!match)return;
    if(match.cell){const r=Number(match.cell.dataset.row),c=Number(match.cell.dataset.col);const raw=excelModel.sheets[currentSheetIndex].rows[r][c];if(typeof raw==='string'&&!raw.startsWith('=')){excelModel.sheets[currentSheetIndex].rows[r][c]=raw.replace(expression,()=>replacement);renderExcelTableDOM();}}
    else {match.node.textContent=match.node.textContent.slice(0,match.start)+replacement+match.node.textContent.slice(match.end);}
  }
  if(currentOfficeApp==='powerpoint')onPptSlideChange();else if(currentOfficeApp==='word')onWordDocumentChange();
  persistState();findWorkspaceText(true);
}

function editWordTable(action){
  const container=editorSelection?.startContainer;const element=container?.nodeType===1?container:container?.parentElement;
  const cell=element?.closest('td,th');const row=cell?.parentElement;const table=cell?.closest('table');
  if(!table){showToast('Önce bir tablo hücresine tıklayın.');return;}
  if(action==='row'){const next=table.insertRow(row.rowIndex+1);for(let i=0;i<row.cells.length;i++)next.insertCell().textContent='';}
  if(action==='column'){for(const item of table.rows)item.insertCell(cell.cellIndex+1).textContent='';}
  if(action==='delete-row'){if(table.rows.length===1)table.remove();else table.deleteRow(row.rowIndex);}
  onWordDocumentChange();
}

function handleExcelFormatting(command,value){
  const toggles={bold:'bold',italic:'italic',underline:'underline',strikeThrough:'strike'};
  const keys={foreColor:'color',hiliteColor:'background',fontName:'font_name',fontSize:'font_size',justifyLeft:'alignment',justifyCenter:'alignment',justifyRight:'alignment',justifyFull:'alignment'};
  const key=toggles[command]||keys[command];if(!key)return false;
  const sheet=excelModel.sheets[currentSheetIndex];sheet.cell_styles||={};
  const id=`${activeCellPos.r}:${activeCellPos.c}`;const style=sheet.cell_styles[id]||={};
  if(toggles[command])style[key]=!style[key];
  else if(command.startsWith('justify'))style[key]={justifyLeft:'left',justifyCenter:'center',justifyRight:'right',justifyFull:'justify'}[command];
  else style[key]=command==='fontSize'?({1:8,2:10,3:12,4:14,5:18,6:24,7:36})[value]:value;
  renderExcelTableDOM();persistState();return true;
}
function applyExcelCellStyle(cell,row,column){
  const style=excelModel.sheets[currentSheetIndex].cell_styles?.[`${row}:${column}`];if(!style)return;
  cell.style.fontWeight=style.bold?'700':'';cell.style.fontStyle=style.italic?'italic':'';
  cell.style.textDecoration=[style.underline?'underline':'',style.strike?'line-through':''].join(' ');
  if(style.color)cell.style.color=style.color;if(style.background)cell.style.backgroundColor=style.background;
  if(style.font_name)cell.style.fontFamily=style.font_name;if(style.font_size)cell.style.fontSize=style.font_size+'pt';
  if(style.alignment)cell.style.textAlign=style.alignment;
}
function sortExcelRows(descending){
  const sheet=excelModel.sheets[currentSheetIndex],column=activeCellPos.c;
  const indexed=sheet.rows.slice(1).map((row,index)=>({row,old:index+1}));
  indexed.sort((a,b)=>{const av=a.row[column]??'',bv=b.row[column]??'';const numberA=Number(av),numberB=Number(bv);const result=av!==''&&bv!==''&&Number.isFinite(numberA)&&Number.isFinite(numberB)?numberA-numberB:String(av).localeCompare(String(bv),'tr',{numeric:true});return descending?-result:result;});
  for(const key of ['cell_styles','number_formats']){if(!sheet[key])continue;const next={};for(const [position,value]of Object.entries(sheet[key])){const [r,c]=position.split(':').map(Number);const nextRow=r===0?0:indexed.findIndex(item=>item.old===r)+1;if(nextRow>=0)next[`${nextRow}:${c}`]=value;}sheet[key]=next;}
  sheet.rows=[sheet.rows[0],...indexed.map(item=>item.row)];renderExcelTableDOM();persistState();
}
function toggleSheetFilter(){const input=document.getElementById('sheet-filter-input');input.hidden=!input.hidden;if(input.hidden){input.value='';filterSheetRows();}else input.focus();}
function filterSheetRows(){const query=document.getElementById('sheet-filter-input').value.toLocaleLowerCase('tr-TR');document.querySelectorAll('#excel-sheet-grid tr').forEach((row,index)=>{row.hidden=index>1&&!row.innerText.toLocaleLowerCase('tr-TR').includes(query);});}
function toggleFreezeRow(){const sheet=excelModel.sheets[currentSheetIndex];sheet.freeze_first_row=!sheet.freeze_first_row;document.getElementById('excel-sheet-grid').classList.toggle('freeze-first-row',sheet.freeze_first_row);persistState();}
function handleSheetPaste(event){
  const text=event.clipboardData?.getData('text/plain');if(!text||!/[\t\n]/.test(text))return;event.preventDefault();
  const rows=text.replace(/\r/g,'').replace(/\n$/,'').split('\n').map(row=>row.split('\t'));
  const sheet=excelModel.sheets[currentSheetIndex],start={...activeCellPos};
  if((start.r+rows.length)*(start.c+Math.max(...rows.map(row=>row.length)))>200000){showToast('Yapıştırılan tablo 200.000 hücre sınırını aşıyor.');return;}
  rows.forEach((row,r)=>{sheet.rows[start.r+r]||=[];row.forEach((value,c)=>{sheet.rows[start.r+r][start.c+c]=value;});});
  renderExcelTableDOM();persistState();
}
