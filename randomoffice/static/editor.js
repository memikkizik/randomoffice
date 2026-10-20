let currentOfficeApp = 'word';
    let currentDocTitle = 'Belge1.docx';
    let currentActiveTab = 'home';
    let currentSheetIndex = 0;
    let excelModel = { sheets: [{ name: 'Sayfa1', rows: [] }] };
    let pptDeck = [
      { id: 1, title: 'Başlıksız sunum', bullets: [] }
    ];
    let activeSlideIndex = 0;
    let activeCellPos = { r: 0, c: 0 };
    let presenterSlideIndex = 0;
    let isDesktopModeActive = false;
    let restoringState = false;
    const documentNames = { word: 'Belge1.docx', excel: 'Tablo1.xlsx', powerpoint: 'Sunum1.pptx' };
    let editorSelection = null;
    const modelHistory = {excel: [], powerpoint: []};
    const historyIndex = {excel: -1, powerpoint: -1};
    let replayingHistory = false;
    let savingDocument = false;
    let toastTimer = null;

    // Boot
    window.addEventListener('DOMContentLoaded', () => {
      setupEditorElements();
      setupProfessionalWorkspace();
      document.body.dataset.activeApp = 'word';
      checkDesktopMode();
      restoreLocalCache();
      renderExcelTabs();
      initExcelGrid();
      renderPptDeckSidebar();
      onWordDocumentChange();
      bindKeyShortcuts();
      bindContextMenu();
      bindDragAndDrop();
      bindGlobalEvents();
      activateOfficeApp(currentOfficeApp);
      initializeDocumentTracking();
      ingestDesktopStartup();
    });

    function setupEditorElements() {
      const stage = document.querySelector('.ppt-canvas-stage');
      if (!document.getElementById('ppt-speaker-notes')) {
        const label = document.createElement('label');
        label.className = 'speaker-notes';
        label.textContent = 'Konuşmacı notları';
        const notes = document.createElement('textarea');
        notes.id = 'ppt-speaker-notes';
        notes.placeholder = 'Sunum sırasında kullanacağınız özel notlar';
        notes.oninput = onPptSlideChange;
        label.appendChild(notes);
        stage.appendChild(label);
      }
      document.getElementById('ppt-speaker-notes').oninput = onPptSlideChange;
      for (const id of ['ppt-slide-images', 'ppt-slide-tables']) {
        if (!document.getElementById(id)) {
          const element = document.createElement('div');
          element.id = id;
          document.getElementById('slide-canvas-card').appendChild(element);
        }
      }
      document.getElementById('system-toast').setAttribute('role', 'status');
      document.getElementById('system-toast').setAttribute('aria-live', 'polite');
    }

    // Check desktop mode via pywebview
    function checkDesktopMode() {
      if (window.pywebview && window.pywebview.api) {
        setDesktopModeUI();
      } else {
        window.addEventListener('pywebviewready', () => {
          setDesktopModeUI();
        });
      }
    }

    function setDesktopModeUI() {
      isDesktopModeActive = true;
      const btn = document.getElementById('btn-desktop-download');
      if (btn) {
        btn.innerText = 'Masaüstü';
        btn.classList.add('active-desktop');
        btn.removeAttribute('href');
        btn.onclick = () => openWorkspacePanel('settings');
      }
      ingestDesktopStartup();
    }

    // 1. Ribbon Tab Switching
    function selectRibbonTab(tab) {
      currentActiveTab = tab;
      document.querySelectorAll('.ribbon-tab-btn').forEach(btn => {
        if (!btn.classList.contains('file-tab-btn')) btn.classList.remove('active');
      });

      const activeBtn = document.getElementById(`tab-${tab}`);
      if (activeBtn) activeBtn.classList.add('active');

      // Hide all ribbon sections
      document.querySelectorAll('.ribbon-section').forEach(section => section.style.display = 'none');

      // Show selected section
      const sec = document.getElementById(`ribbon-sec-${tab}`);
      if (sec) sec.style.display = 'flex';
    }

    // 2. App Switcher
    function activateOfficeApp(app) {
      if (!['word','excel','powerpoint'].includes(app)) return;
      if(app!==currentOfficeApp)editorSelection=null;
      if (!restoringState) documentNames[currentOfficeApp] = document.getElementById('doc-name-editor').value || currentDocTitle;
      currentOfficeApp = app;
      currentDocTitle = documentNames[app];
      document.body.dataset.activeApp = app;
      const root = document.documentElement;

      document.querySelectorAll('.app-pill').forEach(p => p.className = 'app-pill');
      
      const tabExcel = document.getElementById('tab-excel-tools');
      const tabPpt = document.getElementById('tab-ppt-tools');
      tabExcel.style.display = 'none';
      tabPpt.style.display = 'none';

      if (app === 'word') {
        root.style.setProperty('--theme-color', 'var(--word-base)');
        root.style.setProperty('--theme-dark', 'var(--word-dark)');
        root.style.setProperty('--theme-light', 'var(--word-light)');
        root.style.setProperty('--theme-accent', 'var(--word-accent)');
        document.getElementById('pill-word').className = 'app-pill active-word';
        document.getElementById('app-logo-badge').innerText = '📝';
        document.getElementById('doc-name-editor').value = currentDocTitle.endsWith('.docx') ? currentDocTitle : 'Belge1.docx';

        document.getElementById('word-workspace').style.display = 'flex';
        document.getElementById('excel-workspace').style.display = 'none';
        document.getElementById('ppt-workspace').style.display = 'none';

        selectRibbonTab('home');
        onWordDocumentChange();
      } else if (app === 'excel') {
        root.style.setProperty('--theme-color', 'var(--excel-base)');
        root.style.setProperty('--theme-dark', 'var(--excel-dark)');
        root.style.setProperty('--theme-light', 'var(--excel-light)');
        root.style.setProperty('--theme-accent', 'var(--excel-accent)');
        document.getElementById('pill-excel').className = 'app-pill active-excel';
        document.getElementById('app-logo-badge').innerText = '📊';
        document.getElementById('doc-name-editor').value = currentDocTitle.endsWith('.xlsx') ? currentDocTitle : 'Tablo1.xlsx';

        document.getElementById('word-workspace').style.display = 'none';
        document.getElementById('excel-workspace').style.display = 'flex';
        document.getElementById('ppt-workspace').style.display = 'none';

        tabExcel.style.display = 'inline-block';
        selectRibbonTab('home');

        document.getElementById('sb-doc-page').innerText = 'Hazır';
        document.getElementById('sb-word-count').innerText = 'Hücre: A1';
        document.getElementById('sb-reading-time').innerText = '100%';
      } else if (app === 'powerpoint') {
        root.style.setProperty('--theme-color', 'var(--ppt-base)');
        root.style.setProperty('--theme-dark', 'var(--ppt-dark)');
        root.style.setProperty('--theme-light', 'var(--ppt-light)');
        root.style.setProperty('--theme-accent', 'var(--ppt-accent)');
        document.getElementById('pill-ppt').className = 'app-pill active-ppt';
        document.getElementById('app-logo-badge').innerText = '📽️';
        document.getElementById('doc-name-editor').value = currentDocTitle.endsWith('.pptx') ? currentDocTitle : 'Sunum1.pptx';

        document.getElementById('word-workspace').style.display = 'none';
        document.getElementById('excel-workspace').style.display = 'none';
        document.getElementById('ppt-workspace').style.display = 'flex';

        tabPpt.style.display = 'inline-block';
        selectRibbonTab('home');

        document.getElementById('sb-doc-page').innerText = `Slayt ${activeSlideIndex + 1} / ${pptDeck.length}`;
        document.getElementById('sb-word-count').innerText = 'Türkçe';
        document.getElementById('sb-reading-time').innerText = '100%';
      }
      persistState();
      updateWorkspaceContext();
    }

    // 3. Shortcuts & Global Events
    function bindKeyShortcuts() {
      window.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
          e.preventDefault();
          saveDocument(e.shiftKey);
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
          e.preventDefault();
          openBackstageHub('print');
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
          e.preventDefault();
          execOfficeAction('undo');
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
          e.preventDefault();
          execOfficeAction('redo');
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'b') {
          e.preventDefault();
          execOfficeAction('bold');
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'i') {
          e.preventDefault();
          execOfficeAction('italic');
        } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'u') {
          e.preventDefault();
          execOfficeAction('underline');
        } else if (e.key === 'F5' && currentOfficeApp === 'powerpoint') {
          e.preventDefault();
          launchPresenterShow();
        } else if (e.key === 'Escape') {
          closeBackstageHub();
          exitPresenterShow();
          hideColorPickers();
          closeTellMeDropdown();
          closeWorkspacePanel();
        } else if (currentOfficeApp === 'powerpoint' && document.getElementById('presenter-modal').style.display === 'flex') {
          if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
            e.preventDefault();
            advanceSlide();
          } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
            e.preventDefault();
            prevSlide();
          }
        }
      });
    }

    function bindContextMenu() {
      const menu = document.getElementById('desktop-ctx-menu');
      window.addEventListener('contextmenu', (e) => {
        if (e.target.closest('#workspace-stage')) {
          e.preventDefault();
          menu.style.display = 'block';
          menu.style.left = `${Math.min(e.clientX, window.innerWidth - 180)}px`;
          menu.style.top = `${Math.min(e.clientY, window.innerHeight - 200)}px`;
        }
      });
      window.addEventListener('click', () => {
        menu.style.display = 'none';
      });
    }

    function bindGlobalEvents() {
      document.addEventListener('selectionchange', () => {
        const selection = getSelection();
        const element = selection.anchorNode?.nodeType === 1 ? selection.anchorNode : selection.anchorNode?.parentElement;
        if (selection.rangeCount && element?.closest('[contenteditable="true"]')) editorSelection = selection.getRangeAt(0).cloneRange();
      });
      document.querySelector('.ribbon-toolbar').addEventListener('mousedown', event => {
        if (event.target.closest('button, .color-swatch')) event.preventDefault();
      });
      window.addEventListener('click', (e) => {
        if (!e.target.closest('.color-btn-wrap')) hideColorPickers();
      });
    }

    function bindDragAndDrop() {
      const overlay = document.getElementById('drag-overlay');
      window.addEventListener('dragenter', (e) => {
        e.preventDefault();
        overlay.style.display = 'flex';
      });
      overlay.addEventListener('dragover', (e) => {
        e.preventDefault();
      });
      overlay.addEventListener('dragleave', (e) => {
        e.preventDefault();
        overlay.style.display = 'none';
      });
      overlay.addEventListener('drop', (e) => {
        e.preventDefault();
        overlay.style.display = 'none';
        const file = e.dataTransfer.files[0];
        if (file) handleDirectFile(file);
      });
    }

    function hideColorPickers() {
      document.querySelectorAll('.color-dropdown').forEach(d => d.style.display = 'none');
    }

    function toggleColorPicker(type) {
      hideColorPickers();
      const drop = document.getElementById(`dropdown-${type}`);
      if (drop) {
        const bounds = drop.parentElement.getBoundingClientRect();
        drop.style.position='fixed'; drop.style.top=(bounds.bottom+8)+'px';
        drop.style.left=Math.min(bounds.left,innerWidth-160)+'px'; drop.style.display='grid';
      }
    }

    function applyColor(cmd, hex) {
      execOfficeAction(cmd, hex);
      hideColorPickers();
      persistState();
    }

    // 4. Tell Me / Search Commands
    const tellMeActions = [
      { title: '🖨️ Belgeyi Yazdır (Ctrl+P)', exec: () => openBackstageHub('print') },
      { title: '💾 Kaydet & İndir (Ctrl+S)', exec: () => saveDocument() },
      { title: '▦ Tablo Ekle', exec: () => insertGridTable() },
      { title: '🖼️ Resim Ekle', exec: () => triggerImageUpload() },
      { title: '🔗 Bağlantı (Link) Ekle', exec: () => insertHyperlink() },
      { title: '📅 Tarih & Saat Ekle', exec: () => insertCurrentDateTime() },
      { title: '<b>B</b> Metni Kalın Yap', exec: () => execOfficeAction('bold') },
      { title: '<i>I</i> Metni İtalik Yap', exec: () => execOfficeAction('italic') },
      { title: '📏 Cetveli Aç / Kapat', exec: () => toggleRuler() },
      { title: '🌓 Koyu / Açık Tuval', exec: () => toggleDarkCanvas() },
      { title: '🔲 Tam Ekran Görünüm', exec: () => toggleFullscreen() },
      { title: '📝 Word Uygulamasına Geç', exec: () => activateOfficeApp('word') },
      { title: '📊 Excel Uygulamasına Geç', exec: () => activateOfficeApp('excel') },
      { title: '📽️ PowerPoint Uygulamasına Geç', exec: () => activateOfficeApp('powerpoint') },
      { title: '▶️ Slayt Gösterisi Başlat (F5)', exec: () => launchPresenterShow() },
      { title: '∑ Excel Sütun Toplamı Al', exec: () => computeSumExcel() }
    ];

    function handleTellMeSearch(q) {
      const list = document.getElementById('tell-me-results');
      if (!q.trim()) {
        list.style.display = 'none';
        return;
      }
      const matches = tellMeActions.filter(a => a.title.toLowerCase().includes(q.toLowerCase()));
      if (matches.length === 0) {
        list.innerHTML = '<div style="padding:10px 14px;font-size:12px;color:#8a8886;">Eşleşen komut bulunamadı</div>';
        list.style.display = 'block';
        return;
      }
      list.innerHTML = '';
      matches.forEach(m => {
        const row = document.createElement('div');
        row.className = 'tell-me-item';
        row.innerHTML = m.title;
        row.onclick = () => {
          m.exec();
          closeTellMeDropdown();
          document.getElementById('tell-me-input').value = '';
        };
        list.appendChild(row);
      });
      list.style.display = 'block';
    }

    function closeTellMeDropdown() {
      document.getElementById('tell-me-results').style.display = 'none';
    }

    function closeTellMeDropdownDelay() {
      setTimeout(closeTellMeDropdown, 250);
    }

    // 5. Native Print & Print Preview
    function triggerNativePrint() {
      document.body.dataset.activeApp = currentOfficeApp;
      closeBackstageHub();
      showToast('Yazdırma motoru başlatılıyor...');
      setTimeout(() => {
        window.print();
      }, 150);
    }

    function executeActualPrint() {
      closeBackstageHub();
      triggerNativePrint();
    }

    function populatePrintPreview() {
      const box = document.getElementById('print-live-preview-box');
      box.innerHTML = '';
      if (currentOfficeApp === 'word') {
        const editor = document.getElementById('word-editor');
        box.innerHTML = editor.innerHTML;
      } else if (currentOfficeApp === 'excel') {
        const grid = document.getElementById('excel-sheet-grid');
        const clone = grid.cloneNode(true);
        clone.style.width = '100%';
        clone.style.fontSize = '9px';
        box.appendChild(clone);
      } else if (currentOfficeApp === 'powerpoint') {
        const card = document.getElementById('slide-canvas-card');
        const clone = card.cloneNode(true);
        clone.style.width = '100%';
        clone.style.boxShadow = 'none';
        clone.style.padding = '16px';
        box.appendChild(clone);
      }
      box.querySelectorAll('[id]').forEach(node=>node.removeAttribute('id'));
      box.querySelectorAll('[contenteditable]').forEach(node=>node.removeAttribute('contenteditable'));
    }

    // 6. View Controls
    function toggleRuler() {
      const ruler = document.getElementById('word-ruler');
      ruler.style.display = (ruler.style.display === 'none') ? 'flex' : 'none';
      showToast('Cetvel görünümü değiştirildi');
    }

    function toggleDarkCanvas() {
      const sheet = document.getElementById('word-editor');
      sheet.classList.toggle('dark-sheet');
      showToast(sheet.classList.contains('dark-sheet') ? 'Koyu Tuval Modu Aktif' : 'Açık Tuval Modu Aktif');
    }

    function toggleFullscreen() {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen().catch(() => {});
      } else {
        document.exitFullscreen().catch(() => {});
      }
    }

    // 7. Persistence & State
    function persistState() {
      if (restoringState || restoringNativeWorkspace) return;
      if (typeof refreshUnsavedState === 'function') refreshUnsavedState();
      currentDocTitle = document.getElementById('doc-name-editor').value || currentDocTitle;
      documentNames[currentOfficeApp] = currentDocTitle;
      if (!replayingHistory && modelHistory[currentOfficeApp]) {
        const history = modelHistory[currentOfficeApp];
        const snapshot = JSON.stringify(currentOfficeApp === 'excel' ? excelModel : pptDeck);
        if (history[historyIndex[currentOfficeApp]] !== snapshot) {
          history.splice(historyIndex[currentOfficeApp] + 1);
          history.push(snapshot);
          while (history.length > 30 || (history.length > 1 && history.reduce((n,s) => n+s.length,0) > 8000000)) history.shift();
          historyIndex[currentOfficeApp] = history.length - 1;
        }
      }
      const state = {
        app: currentOfficeApp,
        title: currentDocTitle,
        wordHtml: document.getElementById('word-editor').innerHTML,
        excelData: excelModel,
        pptData: pptDeck,
        names: documentNames,
        pageLayout,
        compactRibbon: document.body.classList.contains('compact-ribbon'),
        sheetIndex: currentSheetIndex,
        cleanFingerprints: {...cleanDocumentFingerprints},
        sourceTokens: {...documentSourceTokens},
        updated: new Date().toLocaleTimeString()
      };
      if(persistNativeWorkspace(state))return;
      try {
        localStorage.setItem('randomoffice_cache_v3', JSON.stringify(state));
        document.getElementById('cloud-sync-status').innerText = `✓ Yerel taslak (${state.updated})`;
      } catch (error) {
        document.getElementById('cloud-sync-status').innerText = 'Taslak kaydedilemedi — dosyanızı indirin';
      }
    }

    function restoreLocalCache() {
      restoringState = true;
      try {
        const raw = localStorage.getItem('randomoffice_cache_v3') || localStorage.getItem('randomoffice_cache_v2');
        if (raw) {
          const s = JSON.parse(raw);
          recoveredTrackingState = s;
          if (s.names) Object.assign(documentNames, s.names);
          if (s.pageLayout) {Object.assign(pageLayout,s.pageLayout);applyPageLayout();}
          if (s.compactRibbon) {document.body.classList.add('compact-ribbon');document.getElementById('workspace-compact-check').checked=true;}
          if (s.app && s.title) documentNames[s.app] = s.title;
          if (typeof s.wordHtml==='string') document.getElementById('word-editor').replaceChildren(safeWordFragment(s.wordHtml));
          if (s.excelData?.sheets?.length) excelModel = s.excelData;
          currentSheetIndex = Math.max(0, Math.min(s.sheetIndex || 0, excelModel.sheets.length - 1));
          if (s.pptData && s.pptData.length) pptDeck = s.pptData;
          if (s.title) {
            currentDocTitle = s.title;
            document.getElementById('doc-name-editor').value = s.title;
          }
          if (s.app) activateOfficeApp(s.app);
        }
      } catch (e) {
        document.getElementById('cloud-sync-status').innerText = 'Önceki taslak okunamadı';
      } finally { restoringState = false; }
    }

    // 8. Backstage Hub & Templates
    function openBackstageHub(sec = 'new') {
      document.getElementById('backstage-hub').style.display = 'flex';
      switchHubSection(sec);
    }
    function closeBackstageHub() {
      document.getElementById('backstage-hub').style.display = 'none';
    }
    function switchHubSection(sec) {
      document.querySelectorAll('.hub-nav-item').forEach(i => i.classList.remove('active'));
      const activeNav = document.getElementById(`hub-nav-${sec}`);
      if (activeNav) activeNav.classList.add('active');

      document.getElementById('hub-section-new').style.display = sec === 'new' ? 'block' : 'none';
      document.getElementById('hub-section-print').style.display = sec === 'print' ? 'block' : 'none';
      document.getElementById('hub-section-about').style.display = sec === 'about' ? 'block' : 'none';

      if (sec === 'print') {
        populatePrintPreview();
      }
    }

    function createBlankFromHub(app) {
      if (!confirmDiscardDocument(app)) return;
      closeBackstageHub();
      activateOfficeApp(app);
      if (app === 'word') {
        document.getElementById('word-editor').innerHTML = '<p><br></p>';
        pageLayout={orientation:'portrait',paper:'A4',margin:'normal'};applyPageLayout();
        currentDocTitle = 'Belge1.docx';
      } else if (app === 'excel') {
        excelModel = { sheets: [{ name: 'Sayfa1', rows: [] }] };
        currentSheetIndex = 0;
        initExcelGrid();
        renderExcelTabs();
        currentDocTitle = 'Tablo1.xlsx';
      } else if (app === 'powerpoint') {
        pptDeck = [{ id: 1, title: 'Yeni Sunum', bullets: ['Sunum içeriği buraya gelir'] }];
        activeSlideIndex = 0;
        currentDocTitle = 'Sunum1.pptx';
        renderPptDeckSidebar();
      }
      document.getElementById('doc-name-editor').value = currentDocTitle;
      if(modelHistory[app]) {modelHistory[app]=[];historyIndex[app]=-1;}
      window.pywebview?.api?.reset_current_file?.(app).then(()=>updateDesktopState()).catch(error=>showToast(error.message));
      if(app==='word')onWordDocumentChange();
      delete documentSourceTokens[app];
      markDocumentSaved(app);
      persistState();
    }

    function createTemplateFromHub(tpl) {
      const targetApp=tpl==='word-report'?'word':tpl==='excel-budget'?'excel':null;
      if(!targetApp || !confirmDiscardDocument(targetApp))return;
      closeBackstageHub();
      if (tpl === 'word-report') {
        activateOfficeApp('word');
        currentDocTitle = 'Proje_Faaliyet_Raporu.docx';
        document.getElementById('word-editor').innerHTML = `
          <h1>Proje Faaliyet ve İlerleme Raporu</h1>
          <p><b>Tarih:</b> ${new Date().toLocaleDateString('tr-TR')} | <b>Hazırlayan:</b> RandomOffice Proje Ekibi</p>
          <hr style="margin: 16px 0;">
          <h2>1. Yönetici Özeti</h2>
          <p>Bu rapor projenin genel ilerlemesini, hedeflere ulaşma oranını ve önümüzdeki çeyrek planlarını özetlemektedir.</p>
          <h2>2. Performans Göstergeleri</h2>
          <table>
            <tr><th>Metrik</th><th>Hedef</th><th>Gerçekleşen</th><th>Durum</th></tr>
            <tr><td>Modül Teslimleri</td><td>%100</td><td>%100</td><td>Tamamlandı</td></tr>
            <tr><td>Kullanıcı Memnuniyeti</td><td>%90</td><td>%98</td><td>Başarılı</td></tr>
            <tr><td>Lisans Maliyeti</td><td>0 ₺</td><td>0 ₺</td><td>Ücretsiz</td></tr>
          </table>
          <h2>3. Sonuç ve Öneriler</h2>
          <p>Tüm hedefler zamanında ve bütçe aşımı olmaksızın başarıyla tamamlanmıştır.</p>
        `;
      } else if (tpl === 'excel-budget') {
        activateOfficeApp('excel');
        currentDocTitle = 'Aylik_Butce.xlsx';
        const rows = [
          ['Kategori', 'Açıklama', 'Gelir', 'Gider', 'Net Bakiye'],
          ['Maaş', 'Aylık net gelir', '55000', '0', '55000'],
          ['Kira', 'Ev kirası', '0', '18000', '-18000'],
          ['Faturalar', 'Elektrik, su, internet', '0', '3500', '-3500'],
          ['Market', 'Gıda ve ihtiyaçlar', '0', '9500', '-9500'],
          ['Tasarruf', 'Banka mevduat', '10000', '0', '10000']
        ];
        excelModel = { sheets: [{ name: 'Bütçe', rows }] };
        currentSheetIndex = 0;
        renderExcelTabs();
        renderExcelTableDOM();
      }
      document.getElementById('doc-name-editor').value = currentDocTitle;
      if(modelHistory[currentOfficeApp]) {modelHistory[currentOfficeApp]=[];historyIndex[currentOfficeApp]=-1;}
      if(currentOfficeApp==='word')onWordDocumentChange();
      delete documentSourceTokens[targetApp];
      window.pywebview?.api?.reset_current_file?.(targetApp).then(()=>{if(typeof refreshUnsavedState==='function')refreshUnsavedState(true);updateDesktopState();}).catch(error=>showToast(error.message));
      persistState();
      showToast('Şablon yüklendi!');
    }

    // 9. Word Functions
    function onWordDocumentChange() {
      const text = document.getElementById('word-editor').innerText.trim();
      const words = text ? text.split(/\s+/).length : 0;
      const chars = text ? text.length : 0;
      const minutes = Math.max(1, Math.round(words / 150));
      document.getElementById('sb-doc-page').innerText = 'Sayfa 1 / 1';
      document.getElementById('sb-word-count').innerText = `${words} Sözcük`;
      document.getElementById('sb-reading-time').innerText = `~${minutes} dk okuma (${chars} harf)`;
      persistState();
    }

    function execOfficeAction(cmd, val = null) {
      if(currentOfficeApp==='excel'&&handleExcelFormatting(cmd,val))return;
      if (['undo','redo'].includes(cmd) && modelHistory[currentOfficeApp]) {
        const next = historyIndex[currentOfficeApp] + (cmd === 'undo' ? -1 : 1);
        const snapshot = modelHistory[currentOfficeApp][next];
        if (!snapshot) { showToast('Geri alınacak başka değişiklik yok.'); return; }
        historyIndex[currentOfficeApp] = next;
        replayingHistory = true;
        if (currentOfficeApp === 'excel') {
          excelModel = JSON.parse(snapshot);
          currentSheetIndex = Math.min(currentSheetIndex, excelModel.sheets.length - 1);
          renderExcelTabs(); renderExcelTableDOM();
        } else {
          pptDeck = JSON.parse(snapshot);
          activeSlideIndex = Math.min(activeSlideIndex, pptDeck.length - 1);
          renderPptDeckSidebar();
        }
        persistState();
        replayingHistory = false;
        return;
      }
      const selection = getSelection();
      const element = selection.anchorNode?.nodeType === 1 ? selection.anchorNode : selection.anchorNode?.parentElement;
      const workspace=document.getElementById(currentOfficeApp==='word'?'word-workspace':currentOfficeApp==='powerpoint'?'ppt-workspace':'excel-workspace');
      if (selection.rangeCount && element?.closest('[contenteditable="true"]')&&workspace.contains(element)) {
        editorSelection = selection.getRangeAt(0).cloneRange();
      } else if (editorSelection && workspace.contains(editorSelection.commonAncestorContainer)) {
        selection.removeAllRanges(); selection.addRange(editorSelection);
      } else if (currentOfficeApp!=='excel') {
        document.getElementById(currentOfficeApp==='word'?'word-editor':'ppt-slide-body').focus();
      }
      const success = document.execCommand(cmd, false, val);
      if (!success && ['paste','cut','copy'].includes(cmd)) showToast('Pano için Ctrl+V / Ctrl+X / Ctrl+C kullanın.');
      if(currentOfficeApp==='powerpoint')onPptSlideChange();else persistState();
    }

    function formatHeadTag(tag) {
      execOfficeAction('formatBlock', tag);
      persistState();
    }

    function insertGridTable() {
      const html = `<table style="width:100%;border-collapse:collapse;margin:16px 0;">
        <tr><th style="border:1px solid #8a8886;padding:8px;background:#f3f2f1;">Sütun 1</th><th style="border:1px solid #8a8886;padding:8px;background:#f3f2f1;">Sütun 2</th><th style="border:1px solid #8a8886;padding:8px;background:#f3f2f1;">Sütun 3</th></tr>
        <tr><td style="border:1px solid #8a8886;padding:8px;">Veri 1</td><td style="border:1px solid #8a8886;padding:8px;">Veri 2</td><td style="border:1px solid #8a8886;padding:8px;">Veri 3</td></tr>
        <tr><td style="border:1px solid #8a8886;padding:8px;">Veri 4</td><td style="border:1px solid #8a8886;padding:8px;">Veri 5</td><td style="border:1px solid #8a8886;padding:8px;">Veri 6</td></tr>
      </table><p><br></p>`;
      execOfficeAction('insertHTML', html);
      persistState();
      showToast('Tablo eklendi');
    }

    function insertHyperlink() {
      const url = prompt('Bağlantı adresini girin (URL):', 'https://');
      if (url) {
        if (!/^(https?:|mailto:)/i.test(url)) { showToast('Geçerli bir https:// veya mailto: adresi girin.'); return; }
        execOfficeAction('createLink', url);
        persistState();
      }
    }

    function insertHorizontalRule() {
      execOfficeAction('insertHorizontalRule');
      persistState();
    }

    function insertCurrentDateTime() {
      const now = new Date().toLocaleString('tr-TR');
      execOfficeAction('insertText', now);
      persistState();
    }

    function triggerImageUpload() {
      document.getElementById('image-upload-input').click();
    }

    function handleImagePicked(e) {
      const file = e.target.files[0];
      if (!file) return;
      if (!['image/png','image/jpeg','image/gif'].includes(file.type) || file.size > 10 * 1024 * 1024) {
        showToast('En fazla 10 MB boyutunda PNG, JPEG veya GIF seçin.'); return;
      }

      const reader = new FileReader();
      reader.onload = (ev) => {
        const base64 = ev.target.result;
        if (currentOfficeApp === 'word') {
          execOfficeAction('insertImage', base64);
        } else if (currentOfficeApp === 'powerpoint') {
          const slide = pptDeck[activeSlideIndex];
          slide.images = [...(slide.images || []), base64];
          loadSlideToStage(activeSlideIndex);
        }
        showToast('Resim eklendi.');
        persistState();
      };
      reader.readAsDataURL(file);
      e.target.value = '';
    }

    // 10. Excel Functions & Formula Engine
    function initExcelGrid() {
      if (!excelModel.sheets || excelModel.sheets.length === 0) {
        excelModel.sheets = [{ name: 'Sayfa1', rows: [] }];
      }
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      if (!sheet.rows || sheet.rows.length === 0) {
        const rows = [];
        for (let r = 0; r < 35; r++) {
          const row = [];
          for (let c = 0; c < 15; c++) row.push('');
          rows.push(row);
        }
        sheet.rows = rows;
      }
      renderExcelTableDOM();
    }

    function renderExcelTabs() {
      const dock = document.getElementById('excel-sheet-tabs-list');
      dock.innerHTML = '';
      excelModel.sheets.forEach((sheet, idx) => {
        const card = document.createElement('div');
        card.className = `sheet-tab-card ${idx === currentSheetIndex ? 'active' : ''}`;
        const name = document.createElement('span');
        name.textContent = sheet.name || 'Sayfa ' + (idx + 1);
        card.appendChild(name);
        if (excelModel.sheets.length > 1) {
          const closeBtn = document.createElement('span');
          closeBtn.className = 'sheet-tab-close';
          closeBtn.innerText = '✕';
          closeBtn.onclick = (e) => {
            e.stopPropagation();
            removeExcelSheet(idx);
          };
          card.appendChild(closeBtn);
        }
        card.onclick = () => switchExcelSheet(idx);
        card.ondblclick = () => renameExcelSheet(idx);
        dock.appendChild(card);
      });
    }

    function switchExcelSheet(idx) {
      if (!excelModel.sheets[idx]) return;
      currentSheetIndex = idx;
      activeCellPos = { r: 0, c: 0 };
      document.getElementById('excel-fx-input').value = '';
      document.getElementById('excel-cell-id').textContent = 'A1';
      renderExcelTabs();
      renderExcelTableDOM();
      persistState();
    }

    function addNewSheetTab() {
      let count = excelModel.sheets.length + 1;
      while (excelModel.sheets.some(sheet => sheet.name === `Sayfa${count}`)) count++;
      const rows = [];
      for (let r = 0; r < 35; r++) {
        const row = [];
        for (let c = 0; c < 15; c++) row.push('');
        rows.push(row);
      }
      excelModel.sheets.push({ name: `Sayfa${count}`, rows });
      currentSheetIndex = excelModel.sheets.length - 1;
      renderExcelTabs();
      renderExcelTableDOM();
      persistState();
      showToast(`Sayfa${count} oluşturuldu.`);
    }

    function removeExcelSheet(idx) {
      if (excelModel.sheets.length <= 1 || !excelModel.sheets[idx]) return;
      excelModel.sheets.splice(idx, 1);
      if (idx < currentSheetIndex) currentSheetIndex--;
      currentSheetIndex = Math.min(currentSheetIndex, excelModel.sheets.length - 1);
      activeCellPos = {r: 0, c: 0};
      document.getElementById('excel-fx-input').value = '';
      document.getElementById('excel-cell-id').textContent = 'A1';
      renderExcelTabs();
      renderExcelTableDOM();
      persistState();
    }

    function renameExcelSheet(idx) {
      const current = excelModel.sheets[idx].name || `Sayfa${idx+1}`;
      const name = prompt('Sayfa adını girin:', current);
      if (name && name.trim()) {
        if (name.trim().length > 31 || /[\\/?*\[\]:]/.test(name) || excelModel.sheets.some((s,i) => i !== idx && s.name.toLowerCase() === name.trim().toLowerCase())) {
          showToast('Sayfa adı benzersiz, en fazla 31 karakter olmalı; \\ / ? * [ ] : içeremez.'); return;
        }
        excelModel.sheets[idx].name = name.trim();
        renderExcelTabs();
        persistState();
      }
    }

    function renderExcelTableDOM() {
      const table = document.getElementById('excel-sheet-grid');
      table.innerHTML = '';
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      const rows = sheet.rows;
      if (!rows || rows.length === 0) { initExcelGrid(); return; }

      const colCount = rows.reduce((count,row) => Math.max(count,row.length),15);
      const headerTr = document.createElement('tr');
      const cornerTh = document.createElement('th');
      cornerTh.className = 'row-idx';
      cornerTh.innerText = '';
      headerTr.appendChild(cornerTh);

      for (let c = 0; c < colCount; c++) {
        const th = document.createElement('th');
        th.innerText = columnLabel(c);
        headerTr.appendChild(th);
      }
      table.appendChild(headerTr);

      rows.forEach((row, rIdx) => {
        const tr = document.createElement('tr');
        const rTh = document.createElement('th');
        rTh.className = 'row-idx';
        rTh.innerText = rIdx + 1;
        tr.appendChild(rTh);

        for (let c = 0; c < colCount; c++) {
          const td = document.createElement('td');
          td.contentEditable = "true";
          const rawVal = row[c] ?? '';
          td.innerText = displayExcelCell(rawVal,rIdx,c);
          applyExcelCellStyle(td,rIdx,c);
          td.dataset.row = rIdx;
          td.dataset.col = c;
          td.onfocus = () => selectExcelCell(td, rIdx, c);
          td.onblur = () => finalizeExcelCell(td, rIdx, c);
          td.oninput = () => {
            sheet.rows[rIdx][c] = td.innerText;
            document.getElementById('excel-fx-input').value = td.innerText;
            persistState();
          };
          td.onkeydown = (e) => handleExcelKeyNav(e, rIdx, c);
          tr.appendChild(td);
        }
        table.appendChild(tr);
      });
      table.classList.toggle('freeze-first-row',Boolean(sheet.freeze_first_row));
      filterSheetRows();
    }

    function selectExcelCell(td, r, c) {
      activeCellPos = { r, c };
      document.querySelectorAll('#excel-sheet-grid td').forEach(cell => cell.classList.remove('cell-focused'));
      td.classList.add('cell-focused');
      const addr = `${columnLabel(c)}${r + 1}`;
      document.getElementById('excel-cell-id').innerText = addr;
      
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      const rawVal = (sheet.rows[r] && sheet.rows[r][c] !== undefined) ? sheet.rows[r][c] : '';
      document.getElementById('excel-fx-input').value = rawVal;
      td.innerText = rawVal ?? '';
      document.getElementById('sb-word-count').innerText = `Hücre: ${addr}`;
    }

    function finalizeExcelCell(td, r, c) {
      const val = td.innerText.trim();
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      if (!sheet.rows[r]) sheet.rows[r] = [];
      sheet.rows[r][c] = val;
      td.innerText = displayExcelCell(val,r,c);
      document.querySelectorAll('#excel-sheet-grid td').forEach(cell => {
        if (cell !== document.activeElement && cell !== td) {
          const row = Number(cell.dataset.row), col = Number(cell.dataset.col);
          cell.innerText = displayExcelCell(sheet.rows[row]?.[col] ?? '',row,col);
        }
      });
      persistState();
    }

    function handleExcelKeyNav(e, r, c) {
      if (['Enter','Tab','ArrowUp','ArrowDown'].includes(e.key)) {
        e.preventDefault();
        let nextRow=r, nextCol=c;
        if(e.key==='Tab') nextCol+=e.shiftKey?-1:1;
        else nextRow+=(e.key==='ArrowUp'||e.shiftKey)?-1:1;
        const next = document.querySelector(`#excel-sheet-grid td[data-row="${nextRow}"][data-col="${nextCol}"]`);
        if (next) next.focus();
      }
    }

    function onFormulaTyped(val) {
      const { r, c } = activeCellPos;
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      if (!sheet.rows[r]) sheet.rows[r] = [];
      sheet.rows[r][c] = val;
      const cell = document.querySelector(`#excel-sheet-grid td[data-row="${r}"][data-col="${c}"]`);
      if (cell) cell.innerText = val;
      persistState();
    }

    function onFormulaKeyDown(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        const { r, c } = activeCellPos;
        const cell = document.querySelector(`#excel-sheet-grid td[data-row="${r}"][data-col="${c}"]`);
        if (cell) {
          finalizeExcelCell(cell, r, c);
          const next = document.querySelector(`#excel-sheet-grid td[data-row="${r+1}"][data-col="${c}"]`);
          if (next) next.focus();
        }
      }
    }

    function addTableRowExcel() {
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      const rows = sheet.rows;
      const cols = rows.length > 0 ? rows[0].length : 15;
      rows.push(new Array(cols).fill(''));
      renderExcelTableDOM();
      persistState();
      showToast('Satır eklendi');
    }

    function removeTableRowExcel() {
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      if (sheet.rows.length <= 1) return;
      sheet.rows.pop();
      activeCellPos.r=Math.min(activeCellPos.r,sheet.rows.length-1);
      renderExcelTableDOM();
      persistState();
      showToast('Satır silindi');
    }

    function addTableColExcel() {
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      const width=sheet.rows.reduce((count,row)=>Math.max(count,row.length),15);
      sheet.rows.forEach(row => { while(row.length < width+1) row.push(''); });
      renderExcelTableDOM();
      persistState();
      showToast('Sütun eklendi');
    }

    function removeTableColExcel() {
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      const width=sheet.rows.reduce((count,row)=>Math.max(count,row.length),0);
      if(width<=1) return;
      sheet.rows.forEach(row => { if(row.length>=width) row.pop(); });
      activeCellPos.c=Math.min(activeCellPos.c,width-2);
      renderExcelTableDOM();
      persistState();
      showToast('Sütun silindi');
    }

    function clearSelectedCell() {
      const { r, c } = activeCellPos;
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      if (sheet.rows[r]) sheet.rows[r][c] = '';
      document.getElementById('excel-fx-input').value = '';
      const cell = document.querySelector(`#excel-sheet-grid td[data-row="${r}"][data-col="${c}"]`);
      if (cell) cell.innerText = '';
      persistState();
      showToast('Hücre temizlendi');
    }

    function computeSumExcel() {
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      let sum = 0;
      sheet.rows.forEach(r => {
        try { const val = formulaValue(r[activeCellPos.c]); if(typeof val==='number'&&Number.isFinite(val)) sum+=val; } catch(error) {}
      });
      alert(`Seçili sütun toplamı (SUM): ${sum.toLocaleString('tr-TR')}`);
    }

    function computeAvgExcel() {
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      let sum = 0, count = 0;
      sheet.rows.forEach(r => {
        try { const val=formulaValue(r[activeCellPos.c]); if(typeof val==='number'&&Number.isFinite(val)) {sum+=val;count++;} } catch(error) {}
      });
      const avg = count > 0 ? (sum / count) : 0;
      alert(`Seçili sütun ortalaması (AVERAGE): ${avg.toLocaleString('tr-TR', { maximumFractionDigits: 2 })}`);
    }

    function computeMinMaxExcel(type) {
      const sheet = excelModel.sheets[currentSheetIndex] || excelModel.sheets[0];
      const numbers = [];
      sheet.rows.forEach(r => {
        try {const val=formulaValue(r[activeCellPos.c]); if(typeof val==='number'&&Number.isFinite(val)) numbers.push(val);} catch(error) {}
      });
      if (numbers.length === 0) {
        alert('Sütunda sayısal veri bulunamadı.');
        return;
      }
      const res = type === 'min' ? Math.min(...numbers) : Math.max(...numbers);
      alert(`Seçili sütun ${type.toUpperCase()}: ${res.toLocaleString('tr-TR')}`);
    }

    function formatCurrencyExcel() { setExcelNumberFormat('"\u20ba" #,##0.00'); }
    function formatPercentExcel() { setExcelNumberFormat('0.00%'); }

    // 11. PowerPoint Functions
    function renderPptDeckSidebar() {
      const pane = document.getElementById('ppt-slides-sidebar');
      pane.innerHTML = '';
      pptDeck.forEach((s, idx) => {
        const card = document.createElement('div');
        card.className = `slide-thumb-node ${idx === activeSlideIndex ? 'active' : ''}`;
        card.innerHTML = `
          <div class="slide-thumb-num">${idx + 1}</div>
          <div class="slide-thumb-text"></div>
          <div class="slide-thumb-actions">
            <button class="thumb-mini-btn" title="Çoğalt" onclick="duplicateSlideFromThumb(event, ${idx})">⧉</button>
            <button class="thumb-mini-btn" title="Sil" onclick="removeSlideFromThumb(event, ${idx})">🗑</button>
          </div>
        `;
        card.querySelector('.slide-thumb-text').textContent=s.title || 'Başlıksız Slayt';
        card.onclick = () => {
          activeSlideIndex = idx;
          renderPptDeckSidebar();
          loadSlideToStage(idx);
        };
        pane.appendChild(card);
      });
      loadSlideToStage(activeSlideIndex);
    }

    function loadSlideToStage(idx) {
      const s = pptDeck[idx];
      if (!s) return;
      const header=document.getElementById('ppt-slide-header'),body=document.getElementById('ppt-slide-body');
      if(s.title_html!==undefined)header.replaceChildren(safeWordFragment(s.title_html));else header.innerText=s.title||'';
      if(s.body_html!==undefined)body.replaceChildren(safeWordFragment(s.body_html));else body.innerText=(s.bullets||[]).map(b=>`• ${b}`).join('\n');
      document.getElementById('ppt-speaker-notes').value = s.notes || '';
      document.getElementById('slide-canvas-card').className = `slide-canvas-card theme-${s.theme || 'default'}`;
      const gallery = document.getElementById('ppt-slide-images');
      gallery.replaceChildren();
      (s.images || []).forEach(source => {
        if (!/^data:image\/(png|jpeg|gif);base64,/i.test(source)) return;
        const image = document.createElement('img');
        image.src = source; image.alt = 'Slayt resmi'; gallery.appendChild(image);
      });
      const tables = document.getElementById('ppt-slide-tables');
      tables.replaceChildren();
      (s.tables || []).forEach(rows => {
        const table = document.createElement('table');
        rows.forEach(values => {
          const row = table.insertRow();
          values.forEach(value => { row.insertCell().textContent = value; });
        });
        tables.appendChild(table);
      });
      document.getElementById('sb-doc-page').innerText = `Slayt ${idx + 1} / ${pptDeck.length}`;
    }

    function onPptSlideChange() {
      if (pptDeck[activeSlideIndex]) {
        pptDeck[activeSlideIndex].notes = document.getElementById('ppt-speaker-notes').value;
        pptDeck[activeSlideIndex].title = document.getElementById('ppt-slide-header').innerText.trim();
        pptDeck[activeSlideIndex].title_html=document.getElementById('ppt-slide-header').innerHTML;
        pptDeck[activeSlideIndex].body_html=document.getElementById('ppt-slide-body').innerHTML;
        const raw = document.getElementById('ppt-slide-body').innerText.trim();
        pptDeck[activeSlideIndex].bullets = raw.split('\n').map(l => l.replace(/^[•\s]+/, '')).filter(Boolean);
        const cardTitle = document.querySelector(`.slide-thumb-node:nth-child(${activeSlideIndex + 1}) .slide-thumb-text`);
        if (cardTitle) cardTitle.innerText = pptDeck[activeSlideIndex].title || 'Başlıksız';
        persistState();
      }
    }

    function onPptNotesChange() { onPptSlideChange(); }

    function createNewSlide() {
      pptDeck.push({ id: pptDeck.length + 1, title: 'Yeni Slayt', bullets: ['Yeni konu başlığı'] });
      activeSlideIndex = pptDeck.length - 1;
      renderPptDeckSidebar();
      persistState();
      showToast('Yeni slayt eklendi');
    }

    function duplicateCurrentSlide() {
      const current = pptDeck[activeSlideIndex];
      if (!current) return;
      const clone = {
        ...structuredClone(current),
        id: pptDeck.length + 1,
        title: (current.title || 'Slayt') + ' (Kopya)',
        bullets: [...(current.bullets || [])]
      };
      pptDeck.splice(activeSlideIndex + 1, 0, clone);
      activeSlideIndex++;
      renderPptDeckSidebar();
      persistState();
      showToast('Slayt çoğaltıldı');
    }

    function duplicateSlideFromThumb(e, idx) {
      e.stopPropagation();
      const current = pptDeck[idx];
      if (!current) return;
      const clone = {
        ...structuredClone(current),
        id: pptDeck.length + 1,
        title: (current.title || 'Slayt') + ' (Kopya)',
        bullets: [...(current.bullets || [])]
      };
      pptDeck.splice(idx + 1, 0, clone);
      if(idx<activeSlideIndex)activeSlideIndex++;
      renderPptDeckSidebar();
      persistState();
    }

    function removeCurrentSlide() {
      if (pptDeck.length <= 1) {
        alert('Sunumda en az bir slayt bulunmalıdır.');
        return;
      }
      pptDeck.splice(activeSlideIndex, 1);
      activeSlideIndex = Math.max(0, activeSlideIndex - 1);
      renderPptDeckSidebar();
      persistState();
      showToast('Slayt silindi');
    }

    function removeSlideFromThumb(e, idx) {
      e.stopPropagation();
      if (pptDeck.length <= 1) {
        alert('Sunumda en az bir slayt bulunmalıdır.');
        return;
      }
      pptDeck.splice(idx, 1);
      if(idx<activeSlideIndex)activeSlideIndex--;
      if (activeSlideIndex >= pptDeck.length) activeSlideIndex = pptDeck.length - 1;
      renderPptDeckSidebar();
      persistState();
    }

    function setSlideTheme(theme) {
      const card = document.getElementById('slide-canvas-card');
      card.className = `slide-canvas-card theme-${theme}`;
      if (pptDeck[activeSlideIndex]) pptDeck[activeSlideIndex].theme = theme;
      persistState();
      showToast(`Sunum teması: ${theme}`);
    }

    function launchPresenterShow(startIdx = 0) {
      onPptSlideChange();
      presenterSlideIndex = startIdx;
      document.getElementById('presenter-modal').style.display = 'flex';
      renderPresenterSlide();
    }

    function renderPresenterSlide() {
      const s = pptDeck[presenterSlideIndex];
      if (!s) { exitPresenterShow(); return; }
      if(s.title_html!==undefined)document.getElementById('presenter-title').replaceChildren(safeWordFragment(s.title_html));
      else document.getElementById('presenter-title').innerText = s.title || '';
      document.getElementById('presenter-modal').dataset.theme=s.theme||'white';
      const list = document.getElementById('presenter-list');
      list.innerHTML = '';
      (s.bullets || []).forEach(b => {
        const li = document.createElement('li');
        li.innerText = b;
        list.appendChild(li);
      });
      if(s.body_html!==undefined){list.replaceChildren(safeWordFragment(s.body_html));list.style.listStyle='none';}
      else list.style.listStyle='';
      let media=document.getElementById('presenter-media');
      if(!media){media=document.createElement('div');media.id='presenter-media';list.after(media);}
      media.replaceChildren();
      (s.images||[]).forEach(source=>{if(/^data:image\/(png|jpeg|gif);base64,/i.test(source)){const image=document.createElement('img');image.src=source;image.alt='Sunum resmi';media.appendChild(image);}});
      (s.tables||[]).forEach(rows=>{const table=document.createElement('table');rows.forEach(values=>{const row=table.insertRow();values.forEach(value=>row.insertCell().textContent=value);});media.appendChild(table);});
      document.getElementById('presenter-hud').innerText = `Slayt ${presenterSlideIndex + 1} / ${pptDeck.length} • [Boşluk]/[Sağ Ok]: İleri • [Sol Ok]: Geri • [ESC]: Çıkış`;
    }

    function advanceSlide() {
      if (presenterSlideIndex < pptDeck.length - 1) {
        presenterSlideIndex++;
        renderPresenterSlide();
      } else {
        exitPresenterShow();
      }
    }

    function prevSlide() {
      if (presenterSlideIndex > 0) {
        presenterSlideIndex--;
        renderPresenterSlide();
      }
    }

    function exitPresenterShow() {
      document.getElementById('presenter-modal').style.display = 'none';
    }

    // 12. Zoom Scale
    function handleZoomScale(val) {
      document.body.dataset.zoomMode='manual';
      document.getElementById('sb-zoom-pct').innerText = `${val}%`;
      const slider = document.getElementById('zoom-slider');
      if (slider) slider.value = val;
      const scale = val / 100;
      if (currentOfficeApp === 'word') {
        document.getElementById('word-viewport-scaler').style.zoom=String(scale);
        document.getElementById('word-viewport-scaler').style.transform='none';
      } else if (currentOfficeApp === 'powerpoint') {
        document.getElementById('slide-canvas-card').style.zoom=String(scale);
        document.getElementById('slide-canvas-card').style.transform='none';
      } else {
        document.getElementById('excel-sheet-grid').style.zoom=String(scale);
      }
    }

    // 13. Open & Save Document (Desktop Bridge + Web API)
    function triggerOpenDialog() {
      closeBackstageHub();
      if (window.pywebview && window.pywebview.api && window.pywebview.api.open_native_file) {
        if(!confirmDiscardDocument(null))return;
        showToast('Masaüstü dosya seçici açılıyor...');
        window.pywebview.api.open_native_file().then(data => {
          if (data) {
            if (data.error) showToast(data.error);
            else ingestDocumentData(data);
          }
        }).catch(error => showToast('Dosya açılamadı: ' + error.message));
        return;
      }
      document.getElementById('file-upload-input').click();
    }

    function handleDirectFile(file) {
      if (!/\.(docx|xlsx|pptx)$/i.test(file.name)) {
        showToast('DOCX, XLSX veya PPTX dosyası seçin.'); return Promise.resolve();
      }
      const targetApp=({docx:'word',xlsx:'excel',pptx:'powerpoint'})[file.name.split('.').pop().toLowerCase()];
      if(!confirmDiscardDocument(targetApp))return Promise.resolve();
      if(window.pywebview?.api?.open_file_data) {
        return new Promise((resolve,reject)=>{
          const reader=new FileReader();
          reader.onload=()=>resolve(reader.result.split(',')[1]);
          reader.onerror=()=>reject(new Error('Dosya okunamadı'));
          reader.readAsDataURL(file);
        }).then(encoded=>window.pywebview.api.open_file_data(file.name,encoded))
          .then(data=>{if(data.error) throw new Error(data.error);ingestDocumentData(data);})
          .catch(error=>showToast('Dosya açılamadı: '+error.message));
      }
      showToast(`${file.name} yükleniyor...`);
      const formData = new FormData();
      formData.append('file', file);

      return fetch('/api/v1/office/open', { method: 'POST', body: formData })
        .then(async res => {
          const data = await res.json();
          if (!res.ok) throw new Error(data.detail || 'Dosya açılamadı');
          return data;
        })
        .then(data => {
          data.filename = file.name;
          ingestDocumentData(data);
        })
        .catch(err => showToast('Dosya açılamadı: ' + err.message));
    }

    function handleFilePicked(e) {
      const file = e.target.files[0];
      if (file) handleDirectFile(file).finally(() => { e.target.value = ''; });
    }

    function safeWordFragment(source) {
      const template = document.createElement('template');
      template.innerHTML = source;
      const allowed = new Set(['P','DIV','SPAN','B','STRONG','I','EM','U','S','STRIKE','BR','H1','H2','H3','UL','OL','LI','TABLE','THEAD','TBODY','TR','TD','TH','IMG','FONT','A']);
      template.content.querySelectorAll('*').forEach(node => {
        if (!allowed.has(node.tagName)) { node.replaceWith(document.createTextNode(node.textContent)); return; }
        [...node.attributes].forEach(attr => {
          if (!['style','src','alt','color','face','size','href'].includes(attr.name)) node.removeAttribute(attr.name);
        });
        if (node.tagName === 'IMG' && !/^data:image\/(png|jpeg|gif);base64,/i.test(node.getAttribute('src') || '')) node.remove();
        if (node.hasAttribute('href') && !/^(https?:|mailto:)/i.test(node.getAttribute('href'))) node.removeAttribute('href');
        const styles = ['color','background-color','font-size','font-family','font-weight','font-style','text-decoration','text-align'].map(key => [key,node.style.getPropertyValue(key)]);
        node.removeAttribute('style');
        styles.forEach(([key,value]) => { if (value && !/url\s*\(|expression/i.test(value)) node.style.setProperty(key,value); });
      });
      return template.content;
    }

    function ingestDocumentData(data) {
      if (!['word','excel','powerpoint'].includes(data.app)) throw new Error('Geçersiz belge verisi');
      const incomingName = data.filename || `Belge${data.format || '.docx'}`;
      if (data.app === currentOfficeApp) document.getElementById('doc-name-editor').value = incomingName;
      documentNames[data.app] = incomingName;

      if (data.app === 'word') {
        pageLayout=data.page_layout||{orientation:'portrait',paper:'A4',margin:'normal'};applyPageLayout();
        const editor = document.getElementById('word-editor');
        editor.innerHTML = '';
        (data.items || []).forEach(item => {
          if (item.type === 'paragraph') {
            const tag = ['h1','h2','h3'].includes(item.level) ? item.level : ['bullet','number'].includes(item.level) ? 'li' : 'p';
            const el = document.createElement(tag);
            if (item.html !== undefined) el.appendChild(safeWordFragment(item.html));
            else el.innerText = item.text;
            if (['left','center','right','justify'].includes(item.alignment)) el.style.textAlign = item.alignment;
            if (tag === 'li') {
              const list = document.createElement(item.level === 'number' ? 'ol' : 'ul');
              list.appendChild(el); editor.appendChild(list);
            } else editor.appendChild(el);
          } else if (item.type === 'table') {
            const tbl = document.createElement('table');
            (item.rows || []).forEach(row => {
              const tr = document.createElement('tr');
              row.forEach(cell => {
                const td = document.createElement('td');
                td.innerText = cell;
                tr.appendChild(td);
              });
              tbl.appendChild(tr);
            });
            editor.appendChild(tbl);
          }
        });
        activateOfficeApp('word');
      } else if (data.app === 'excel') {
        excelModel = data;
        currentSheetIndex = 0;
        renderExcelTabs();
        renderExcelTableDOM();
        activateOfficeApp('excel');
      } else if (data.app === 'powerpoint') {
        pptDeck = (data.slides || []).map((s, i) => ({
          id: i + 1,
          title: s.title || `Slayt ${i+1}`,
          bullets: s.paragraphs || [],
          notes: s.notes || '', tables: s.tables || [], images: s.images || [], theme: s.theme || 'default', title_html:s.title_html,body_html:s.body_html
        }));
        if (!pptDeck.length) pptDeck = [{id:1,title:'',bullets:[],notes:'',tables:[],images:[]}];
        activeSlideIndex = 0;
        renderPptDeckSidebar();
        activateOfficeApp('powerpoint');
      }
      showToast(`${currentDocTitle} başarıyla açıldı!`);
      if(data.source_token)documentSourceTokens[data.app]=data.source_token;
      else delete documentSourceTokens[data.app];
      markDocumentSaved(data.app);
      if(modelHistory[currentOfficeApp]) {modelHistory[currentOfficeApp]=[];historyIndex[currentOfficeApp]=-1;}
      persistState();
    }

    async function saveDocument(saveAs=false) {
      if (savingDocument) return;
      savingDocument = true;
      const buttons = [...document.querySelectorAll('button[onclick="saveDocument()"]')];
      buttons.forEach(button => { button.disabled = true; button.setAttribute('aria-busy', 'true'); });
      try {
        return await performDocumentSave(saveAs);
      } catch (error) {
        showToast('Kaydetme başarısız: ' + error.message);
        return false;
      } finally {
        savingDocument = false;
        buttons.forEach(button => { button.disabled = false; button.removeAttribute('aria-busy'); });
      }
    }

    async function performDocumentSave(saveAs=false) {
      closeBackstageHub();
      showToast('Belge derleniyor ve kaydediliyor...');
      let filename = document.getElementById('doc-name-editor').value.trim() || 'Belge.docx';
      filename=filename.replace(/\.(docx|xlsx|pptx)$/i,'')+({word:'.docx',excel:'.xlsx',powerpoint:'.pptx'})[currentOfficeApp];
      let payload = {};
      let docType = currentOfficeApp;

      if (currentOfficeApp === 'word') {
        if (!filename.endsWith('.docx')) filename += '.docx';
        const editor = document.getElementById('word-editor');
        const items = [];
        editor.childNodes.forEach(node => {
          if (node.nodeType === 1) {
            const tag = node.tagName.toLowerCase();
            if (tag === 'table') {
              const rows = [];
              node.querySelectorAll('tr').forEach(tr => {
                const r = [];
                tr.querySelectorAll('td, th').forEach(c => r.push(c.innerText.trim()));
                if (r.length > 0) rows.push(r);
              });
              if (rows.length > 0) items.push({ type: 'table', rows });
            } else if (tag === 'img') {
              // skip raw img tag
            } else {
              items.push({
                type: 'paragraph',
                text: node.innerText.trim(),
                level: tag === 'h1' ? 'h1' : tag === 'h2' ? 'h2' : tag === 'h3' ? 'h3' : 'p'
              });
            }
          }
        });
        payload = { app: 'word', title: filename, items, html: editor.innerHTML, page_layout:pageLayout };
      } else if (currentOfficeApp === 'excel') {
        if (!filename.endsWith('.xlsx')) filename += '.xlsx';
        payload = { app: 'excel', sheets: excelModel.sheets };
      } else if (currentOfficeApp === 'powerpoint') {
        if (!filename.endsWith('.pptx')) filename += '.pptx';
        onPptSlideChange();
        payload = {
          app: 'powerpoint',
          slides: pptDeck.map(s => ({
            slide_number: s.id,
            title: s.title,
            paragraphs: s.bullets,
            notes: s.notes || '', tables: s.tables || [], images: s.images || [], theme: s.theme || 'default',title_html:s.title_html,body_html:s.body_html
          }))
        };
      }

      const savedSnapshot=documentFingerprint(docType);
      if(documentSourceTokens[docType])payload.source_token=documentSourceTokens[docType];
      // Native Desktop Mode Save
      if (window.pywebview && window.pywebview.api && window.pywebview.api.save_native_file) {
        const api=window.pywebview.api;
        let res = await (!saveAs && api.save_current_file ? api.save_current_file(docType,filename,payload) : api.save_native_file(docType,filename,payload));
        if(res?.conflict && confirm('Bu dosya başka bir uygulamada değiştirildi veya taşındı. Değişikliklerinizi farklı bir dosyaya kaydetmek ister misiniz?'))res=await api.save_native_file(docType,filename,payload);
        if(res?.requires_save_as && confirm('Bu düzenleme, belgedeki bazı gelişmiş nesnelerin korunmasına izin vermiyor. Yeni bir kopya oluşturulsun mu? Yeni kopyada desteklenmeyen nesneler ve bazı biçimler kaybolabilir. Asıl dosyanız korunacak.'))res=await api.save_native_file(docType,filename,payload,true);
        if (res && res.success) {
          documentNames[docType]=res.filename||filename;
          if(currentOfficeApp===docType)onNameEdited(res.filename || filename);
          markDocumentSaved(docType,savedSnapshot);persistState();
          showToast(`${res.filename || filename} kaydedildi!`);
          updateDesktopState();
          return true;
        } else if (res && res.cancelled) showToast('Kaydetme iptal edildi.');
        else throw new Error(res?.error || 'Masaüstü kaydetme işlemi tamamlanamadı');
        return false;
      }

      // Web API Mode Save
      try {
        const params = new URLSearchParams({ doc_type: docType, filename });
        let res = await fetch(`/api/v1/office/export?${params.toString()}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
        if (!res.ok) {
          const details = await res.json().catch(() => ({}));
          if(details.detail?.requires_save_as && confirm('Bu düzenleme gelişmiş nesnelerin korunmasına izin vermiyor. Desteklenmeyen nesnelerin kaybolabileceği yeni bir kopya indirilsin mi? Asıl dosyanız korunacak.')){
            params.set('allow_reconstruction','true');
            res=await fetch(`/api/v1/office/export?${params}`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
            if(!res.ok)throw new Error('Yeni kopya oluşturulamadı');
            delete documentSourceTokens[docType];
          }else throw new Error(typeof details.detail==='string'?details.detail:details.detail?.error||'Dosya oluşturulamadı');
        }

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url),1000);
        documentNames[docType]=filename;
        if(currentOfficeApp===docType)onNameEdited(filename);
        markDocumentSaved(docType,savedSnapshot);persistState();
        showToast(`${filename} indirildi!`);
        return true;
      } catch (err) {
        throw err;
      }
    }

    function onNameEdited(name) {
      currentDocTitle = name;
      document.getElementById('doc-name-editor').value = name;
      persistState();
    }

    function showToast(msg) {
      const toast = document.getElementById('system-toast');
      toast.innerText = msg;
      toast.style.display = 'block';
      clearTimeout(toastTimer);
      toastTimer = setTimeout(() => { toast.style.display = 'none'; }, 5000);
    }
