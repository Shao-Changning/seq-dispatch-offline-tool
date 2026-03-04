(() => {
  const TEMPLATE_PATHS = {
    gzb: '模板文件/格致博雅.xlsx',
    hplosBgi: '模板文件/BGI-海普洛斯.xlsx',
    hplosNova: '模板文件/NOVA海普洛斯.xlsx',
    mingma: '模板文件/明码.xlsx',
    mingmaAtac: '模板文件/ATAC-明码.xlsx',
    indexKit: '模板文件/Single_Index_Kit_N_Set_A(ATAC).xlsx'
  };

  const RULE_FIELDS = [
    { key: 'vendorLibType', label: '文库类型(工厂)' },
    { key: 'vendorLibStructure', label: '文库结构' },
    { key: 'vendorLibProcess', label: '文库处理情况' },
    { key: 'defaultVolumeUl', label: '体积(ul)' },
    { key: 'defaultSpecies', label: '物种' },
    { key: 'defaultSpecialSeq', label: '特殊序列' },
    { key: 'defaultBaseBalance', label: '碱基是否均衡' },
    { key: 'defaultSeqStrategy', label: '测序策略' },
    { key: 'defaultProductType', label: '产品类型' },
    { key: 'defaultRemark', label: '备注' },
    { key: 'defaultAdapterType', label: '接头类型' },
    { key: 'defaultPhosphorylation', label: '5’磷酸化' },
    { key: 'defaultCyclization', label: '是否环化' },
    { key: 'defaultKitVersion', label: '试剂盒版本' },
    { key: 'defaultSeqNumber', label: '序列号' },
    { key: 'defaultPlatformName', label: '测序平台名(模板)' }
  ];

  const RULE_FIELDS_BY_TEMPLATE = {
    gzb: [
      { key: 'vendorLibType', label: '文库类型(工厂)' },
      { key: 'vendorLibStructure', label: '文库结构' },
      { key: 'vendorLibProcess', label: '文库处理情况' },
      { key: 'defaultVolumeUl', label: '体积(ul)' },
      { key: 'defaultSpecies', label: '物种' },
      { key: 'defaultSpecialSeq', label: '特殊序列' },
      { key: 'defaultBaseBalance', label: '碱基是否均衡' },
      { key: 'defaultRemark', label: '备注' }
    ],
    hplosBgi: [
      { key: 'vendorLibType', label: '文库类型(工厂)' },
      { key: 'vendorLibStructure', label: '文库结构' },
      { key: 'defaultKitVersion', label: '试剂盒版本' },
      { key: 'defaultPhosphorylation', label: '磷酸化' },
      { key: 'defaultCyclization', label: '环化' },
      { key: 'defaultSeqNumber', label: '序列号' },
      { key: 'defaultProductType', label: '产品类型' },
      { key: 'defaultSeqStrategy', label: '测序策略' },
      { key: 'defaultRemark', label: '备注' }
    ],
    hplosNova: [
      { key: 'vendorLibType', label: '文库类型(工厂)' },
      { key: 'defaultProductType', label: '产品类型' },
      { key: 'defaultSeqStrategy', label: '测序策略' },
      { key: 'defaultRemark', label: '备注' }
    ],
    mingma: [
      { key: 'vendorLibType', label: '文库类型(工厂)' },
      { key: 'defaultVolumeUl', label: '体积(ul)' },
      { key: 'defaultAdapterType', label: '接头类型' },
      { key: 'defaultPhosphorylation', label: '5’磷酸化' },
      { key: 'defaultCyclization', label: '是否环化' },
      { key: 'defaultPlatformName', label: '测序平台名(模板)' },
      { key: 'defaultRemark', label: '备注' }
    ],
    mingmaAtac: [
      { key: 'vendorLibType', label: '文库类型(工厂)' },
      { key: 'defaultVolumeUl', label: '体积(ul)' },
      { key: 'defaultAdapterType', label: '接头类型' },
      { key: 'defaultPhosphorylation', label: '5’磷酸化' },
      { key: 'defaultCyclization', label: '是否环化' },
      { key: 'defaultPlatformName', label: '测序平台名(模板)' },
      { key: 'defaultRemark', label: '备注' }
    ]
  };

  const DISPLAY_COLUMNS = [
    { key: 'status', label: '状态', type: 'status' },
    { key: 'customer', label: '客户', type: 'row', field: 'customer' },
    { key: 'sendLibId', label: '送测文库ID', type: 'row', field: 'sendLibId' },
    { key: 'libType', label: '内部文库类型', type: 'row', field: 'libType' },
    {
      key: 'indexI7',
      label: 'I7',
      type: 'row',
      field: 'indexI7',
      display: (row, group) => row.indexI7 || (group.templateKey === 'mingmaAtac' ? row.indexCode : '') || ''
    },
    { key: 'indexI5', label: 'I5', type: 'row', field: 'indexI5' },
    { key: 'targetGb', label: '数据量(GB)', type: 'row', field: 'targetGb' },
    { key: 'vendorLibType', label: '工厂文库类型', type: 'mapped', editable: true },
    { key: 'defaultRemark', label: '备注', type: 'mapped', editable: true, special: 'remark' }
  ];

  const state = {
    dispatches: [],
    rows: [],
    groups: [],
    rules: loadRules(),
    templateBuffers: {},
    indexKit: {},
    ui: {
      date: '',
      labSiteMode: 'auto',
      labSiteCustom: ''
    }
  };

  const elements = {
    dropZone: document.getElementById('dropZone'),
    dispatchInput: document.getElementById('dispatchInput'),
    dispatchList: document.getElementById('dispatchList'),
    sendDateInput: document.getElementById('sendDateInput'),
    labSiteSelect: document.getElementById('labSiteSelect'),
    labSiteCustom: document.getElementById('labSiteCustom'),
    templateStatus: document.getElementById('templateStatus'),
    templateDropZone: document.getElementById('templateDropZone'),
    templateAutoInput: document.getElementById('templateAutoInput'),
    templateAutoSelect: document.getElementById('templateAutoSelect'),
    templateFolderSelect: document.getElementById('templateFolderSelect'),
    groupTabs: document.getElementById('groupTabs'),
    groupPanels: document.getElementById('groupPanels'),
    downloadAllBtn: document.getElementById('downloadAllBtn'),
    copyAllMailBtn: document.getElementById('copyAllMailBtn'),
    exportRulesBtn: document.getElementById('exportRulesBtn'),
    importRulesInput: document.getElementById('importRulesInput'),
    tplGzb: document.getElementById('tpl-gzb'),
    tplHplosBgi: document.getElementById('tpl-hplos-bgi'),
    tplHplosNova: document.getElementById('tpl-hplos-nova'),
    tplMingma: document.getElementById('tpl-mingma'),
    tplMingmaAtac: document.getElementById('tpl-mingma-atac'),
    tplIndexKit: document.getElementById('tpl-index-kit'),
    toast: document.getElementById('toast')
  };

  init();

  async function init() {
    const today = new Date();
    state.ui.date = formatDateInput(today);
    elements.sendDateInput.value = state.ui.date;

    elements.sendDateInput.addEventListener('change', () => {
      state.ui.date = elements.sendDateInput.value;
      renderGroups();
    });

    elements.labSiteSelect.addEventListener('change', () => {
      state.ui.labSiteMode = elements.labSiteSelect.value;
      elements.labSiteCustom.disabled = state.ui.labSiteMode !== 'custom';
      if (state.ui.labSiteMode !== 'custom') {
        state.ui.labSiteCustom = '';
        elements.labSiteCustom.value = '';
      }
      applyLabSiteOverride();
    });

    elements.labSiteCustom.addEventListener('input', () => {
      state.ui.labSiteCustom = elements.labSiteCustom.value.trim();
      applyLabSiteOverride();
    });

    elements.dispatchInput.addEventListener('change', (event) => {
      handleDispatchFiles(Array.from(event.target.files || []));
    });

    setupDropZone();
    setupTemplateUploads();
    setupTemplateAutoMatch();

    elements.downloadAllBtn.addEventListener('click', downloadAllGroups);
    elements.copyAllMailBtn.addEventListener('click', copyAllEmails);
    elements.exportRulesBtn.addEventListener('click', exportRulesJson);
    elements.importRulesInput.addEventListener('change', handleImportRules);

    await loadTemplatesFromCache();
    await loadTemplates();
    renderTemplateStatus();
    renderGroups();
  }

  function setupDropZone() {
    ['dragenter', 'dragover'].forEach((evt) => {
      elements.dropZone.addEventListener(evt, (event) => {
        event.preventDefault();
        event.stopPropagation();
        elements.dropZone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach((evt) => {
      elements.dropZone.addEventListener(evt, (event) => {
        event.preventDefault();
        event.stopPropagation();
        elements.dropZone.classList.remove('dragover');
      });
    });

    elements.dropZone.addEventListener('drop', (event) => {
      const files = Array.from(event.dataTransfer.files || []);
      handleDispatchFiles(files);
    });
  }

  function setupTemplateUploads() {
    elements.tplGzb.addEventListener('change', () => handleTemplateUpload('gzb', elements.tplGzb.files));
    elements.tplHplosBgi.addEventListener('change', () => handleTemplateUpload('hplosBgi', elements.tplHplosBgi.files));
    elements.tplHplosNova.addEventListener('change', () => handleTemplateUpload('hplosNova', elements.tplHplosNova.files));
    elements.tplMingma.addEventListener('change', () => handleTemplateUpload('mingma', elements.tplMingma.files));
    elements.tplMingmaAtac.addEventListener('change', () => handleTemplateUpload('mingmaAtac', elements.tplMingmaAtac.files));
    elements.tplIndexKit.addEventListener('change', () => handleTemplateUpload('indexKit', elements.tplIndexKit.files));
  }

  function setupTemplateAutoMatch() {
    elements.templateAutoInput.addEventListener('change', () => {
      handleTemplateFiles(Array.from(elements.templateAutoInput.files || []));
      elements.templateAutoInput.value = '';
    });

    elements.templateAutoSelect.addEventListener('change', () => {
      handleTemplateFiles(Array.from(elements.templateAutoSelect.files || []));
      elements.templateAutoSelect.value = '';
    });

    elements.templateFolderSelect.addEventListener('change', () => {
      handleTemplateFiles(Array.from(elements.templateFolderSelect.files || []));
      elements.templateFolderSelect.value = '';
    });
    ['dragenter', 'dragover'].forEach((evt) => {
      elements.templateDropZone.addEventListener(evt, (event) => {
        event.preventDefault();
        event.stopPropagation();
        elements.templateDropZone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach((evt) => {
      elements.templateDropZone.addEventListener(evt, (event) => {
        event.preventDefault();
        event.stopPropagation();
        elements.templateDropZone.classList.remove('dragover');
      });
    });

    elements.templateDropZone.addEventListener('drop', (event) => {
      const files = Array.from(event.dataTransfer.files || []);
      handleTemplateFiles(files);
    });
  }

  async function handleTemplateUpload(key, files) {
    if (!files || !files.length) return;
    const file = files[0];
    const buffer = await file.arrayBuffer();
    state.templateBuffers[key] = buffer;
    if (key === 'indexKit') {
      state.indexKit = parseIndexKit(buffer);
    }
    await cacheTemplate(key, buffer, file.name);
    renderTemplateStatus();
    toast(`已载入模板：${file.name}`);
  }

  async function handleTemplateFiles(files) {
    if (!files || !files.length) return;
    const excelFiles = files.filter((file) => file.name.endsWith('.xlsx') || file.name.endsWith('.xls'));
    if (!excelFiles.length) {
      toast('未发现可解析的模板文件');
      return;
    }

    const unmatched = [];
    for (const file of excelFiles) {
      const key = matchTemplateKey(file.name);
      if (!key) {
        unmatched.push(file.name);
        continue;
      }
      const buffer = await file.arrayBuffer();
      state.templateBuffers[key] = buffer;
      if (key === 'indexKit') {
        state.indexKit = parseIndexKit(buffer);
      }
      await cacheTemplate(key, buffer, file.name);
    }

    renderTemplateStatus();
    if (unmatched.length) {
      toast(`未匹配模板：${unmatched.slice(0, 2).join('、')}${unmatched.length > 2 ? '…' : ''}`);
    } else {
      toast('模板已自动匹配');
    }
  }

  async function loadTemplates() {
    const entries = Object.entries(TEMPLATE_PATHS);
    await Promise.all(entries.map(async ([key, path]) => {
      try {
        const response = await fetch(encodeURI(path));
        if (!response.ok) return;
        const buffer = await response.arrayBuffer();
        state.templateBuffers[key] = buffer;
        if (key === 'indexKit') {
          state.indexKit = parseIndexKit(buffer);
        }
        await cacheTemplate(key, buffer, getTemplateNameFromPath(path));
      } catch (err) {
        // ignore
      }
    }));
    renderTemplateStatus();
  }

  function renderTemplateStatus() {
    const statuses = [
      { key: 'gzb', label: '格致博雅' },
      { key: 'hplosBgi', label: '海普洛斯(BGI)' },
      { key: 'hplosNova', label: '海普洛斯(Nova)' },
      { key: 'mingma', label: '明码(通用)' },
      { key: 'mingmaAtac', label: '明码(ATAC)' },
      { key: 'indexKit', label: 'ATAC Index Kit' }
    ];

    elements.templateStatus.innerHTML = '';
    statuses.forEach(({ key, label }) => {
      const pill = document.createElement('div');
      const loaded = Boolean(state.templateBuffers[key]);
      pill.className = `status-pill ${loaded ? '' : 'missing'}`;
      pill.textContent = `${label}：${loaded ? '已加载' : '未加载'}`;
      elements.templateStatus.appendChild(pill);
    });
  }

  function getTemplateNameFromPath(path) {
    const parts = String(path || '').split('/');
    return parts[parts.length - 1] || path;
  }

  async function loadTemplatesFromCache() {
    try {
      const db = await openTemplateDB();
      const records = await getAllTemplates(db);
      records.forEach((record) => {
        if (!record || !record.key || !record.buffer) return;
        state.templateBuffers[record.key] = record.buffer;
        if (record.key === 'indexKit') {
          state.indexKit = parseIndexKit(record.buffer);
        }
      });
    } catch (err) {
      // ignore cache failures
    }
  }

  async function cacheTemplate(key, buffer, name) {
    try {
      const db = await openTemplateDB();
      await putTemplate(db, { key, buffer, name, savedAt: Date.now() });
    } catch (err) {
      // ignore cache failures
    }
  }

  function openTemplateDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open('dispatch-template-cache', 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains('templates')) {
          db.createObjectStore('templates', { keyPath: 'key' });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  function putTemplate(db, record) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction('templates', 'readwrite');
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.objectStore('templates').put(record);
    });
  }

  function getAllTemplates(db) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction('templates', 'readonly');
      const store = tx.objectStore('templates');
      const request = store.getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error);
    });
  }

  function matchTemplateKey(fileName) {
    const name = String(fileName || '').toLowerCase();
    if (!name) return null;
    if (name.includes('single_index') || name.includes('index_kit') || (name.includes('index') && name.includes('atac'))) {
      return 'indexKit';
    }
    if (name.includes('格致博雅') || name.includes('gzb')) return 'gzb';
    if ((name.includes('海普') || name.includes('hplos') || name.includes('haipu')) && (name.includes('bgi') || name.includes('t7'))) {
      return 'hplosBgi';
    }
    if ((name.includes('海普') || name.includes('hplos') || name.includes('haipu')) && name.includes('nova')) {
      return 'hplosNova';
    }
    if (name.includes('明码') || name.includes('mingma')) {
      if (name.includes('atac') || name.includes('element')) return 'mingmaAtac';
      return 'mingma';
    }
    if (name.includes('atac') && name.includes('明码')) return 'mingmaAtac';
    return null;
  }

  async function handleDispatchFiles(files) {
    if (!files.length) return;
    const excelFiles = files.filter((file) => file.name.endsWith('.xlsx') || file.name.endsWith('.xls'));
    if (!excelFiles.length) {
      toast('未发现可解析的 Excel 文件');
      return;
    }

    const parsed = [];
    for (const file of excelFiles) {
      try {
        const dispatch = await parseDispatchFile(file);
        parsed.push(dispatch);
      } catch (err) {
        console.error(err);
        toast(`解析失败：${file.name}`);
      }
    }

    state.dispatches = parsed;
    state.rows = parsed.flatMap((item) => item.rows);
    applyLabSiteOverride();
    renderDispatchList();
    renderGroups();
  }

  function renderDispatchList() {
    elements.dispatchList.innerHTML = '';
    state.dispatches.forEach((dispatch) => {
      const item = document.createElement('div');
      item.className = 'file-item';
      const name = dispatch.sourceName;
      const info = `${dispatch.header.cxpsNo || '未知单号'} / ${dispatch.rows.length}行`;
      item.innerHTML = `<span>${name}</span><span>${info}</span>`;
      elements.dispatchList.appendChild(item);
    });
  }

  function applyLabSiteOverride() {
    if (!state.rows.length) return;
    const mode = state.ui.labSiteMode;
    const value = mode === 'custom' ? state.ui.labSiteCustom : mode;

    state.rows.forEach((row) => {
      if (mode === 'auto') return;
      if (value) row.labSite = value;
    });
    renderGroups();
  }

  async function parseDispatchFile(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, blankrows: false });

    const headerRowIndex = findHeaderRowIndex(rows);
    const headerRow = forwardFillRow(rows[headerRowIndex] || []);
    const dataStartIndex = headerRowIndex + 1;
    const dataEndIndex = findDataEndIndex(rows, dataStartIndex);

    const headerInfo = extractHeaderInfo(rows, 0, headerRowIndex);
    const footerInfo = extractFooterInfo(rows, dataEndIndex, rows.length - 1);

    const cxpsNo = headerInfo.cxpsNo || extractCxpsFromName(file.name) || '未知';
    const defaultVendor = normalizeVendor(footerInfo.vendor || headerInfo.vendor || '未知');
    const defaultPlatform = normalizePlatform(headerInfo.platform || '未知');
    const defaultLabSite = headerInfo.labSite || headerInfo.labSiteRaw || '未知';

    const parsedRows = [];
    for (let r = dataStartIndex; r < dataEndIndex; r += 1) {
      const rowData = rows[r];
      if (!rowData || rowData.every((cell) => cell === null || cell === undefined || String(cell).trim() === '')) {
        continue;
      }

      const row = {
        _id: `${file.name}-${r}`,
        cxpsNo,
        vendor: defaultVendor,
        platform: defaultPlatform,
        labSite: defaultLabSite,
        overrides: {}
      };

      headerRow.forEach((headerCell, idx) => {
        const value = rowData[idx];
        if (value === null || value === undefined || value === '') return;
        const headerKey = normalizeHeader(headerCell);
        const mappedKey = mapColumnKey(headerKey);
        if (!mappedKey) return;
        row[mappedKey] = castValue(mappedKey, value);
      });

      if (row.platform) row.platform = normalizePlatform(row.platform);
      if (row.vendor) row.vendor = normalizeVendor(row.vendor);
      if (row.labSite) row.labSite = String(row.labSite).trim();

      if (!row.sendLibId && row.sampleName) row.sendLibId = row.sampleName;
      if (!row.sendLibId && row.libId) row.sendLibId = row.libId;
      if (!row.libType && row.rowRemark) row.libType = row.rowRemark;

      if (!row.sendLibId && !row.libType && !row.indexI7 && !row.indexI5) {
        continue;
      }

      parsedRows.push(row);
    }

    return {
      sourceName: file.name,
      header: {
        cxpsNo,
        dispatcherBiz: headerInfo.dispatcherBiz || '',
        i5Direction: headerInfo.i5Direction || '',
        sendMethod: headerInfo.sendMethod || '',
        platform: defaultPlatform,
        labSite: defaultLabSite,
        remark: headerInfo.remark || '',
        vendor: defaultVendor,
        vendorAddr: footerInfo.vendorAddr || '',
        sendDate: footerInfo.sendDate || '',
        totalGb: footerInfo.totalGb || null
      },
      rows: parsedRows
    };
  }

  function findHeaderRowIndex(rows) {
    for (let i = 0; i < Math.min(rows.length, 20); i += 1) {
      const row = rows[i];
      if (!row) continue;
      const hasSeq = row.some((cell) => normalizeHeader(cell) === '序号');
      if (hasSeq) return i;
    }
    return 4;
  }

  function findDataEndIndex(rows, startIndex) {
    for (let i = startIndex; i < rows.length; i += 1) {
      const row = rows[i];
      if (!row) continue;
      const rowText = row.map((cell) => String(cell || '')).join(' ');
      if (/测序机构|测序工厂邮寄地址|总数据量|送测时间/.test(rowText)) {
        return i;
      }
    }
    return rows.length;
  }

  function extractHeaderInfo(rows, start, end) {
    const info = {};
    const labelMap = {
      cxpsNo: /送测单号/,
      dispatcherBiz: /派单商务/,
      i5Direction: /i5index方向/i,
      sendMethod: /送测方式/,
      platform: /测序平台/,
      labSite: /实验地点/,
      remark: /备注信息/,
      vendor: /测序机构/
    };

    Object.entries(labelMap).forEach(([key, regex]) => {
      const value = findValueByLabel(rows, start, end, regex);
      if (value) info[key] = value;
    });
    return info;
  }

  function extractFooterInfo(rows, start, end) {
    const info = {};
    const labelMap = {
      vendor: /测序机构/,
      vendorAddr: /测序工厂邮寄地址|邮寄地址|地址及联系方式/,
      sendDate: /送测时间/,
      totalGb: /总数据量/
    };

    Object.entries(labelMap).forEach(([key, regex]) => {
      const value = findValueByLabel(rows, start, end, regex);
      if (value) info[key] = value;
    });
    return info;
  }

  function findValueByLabel(rows, start, end, regex) {
    for (let r = start; r <= end; r += 1) {
      const row = rows[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c += 1) {
        const cell = row[c];
        if (cell && regex.test(String(cell))) {
          const candidate = findNearbyValue(rows, r, c);
          if (candidate) return candidate;
        }
      }
    }
    return '';
  }

  function findNearbyValue(rows, rowIndex, colIndex) {
    const row = rows[rowIndex] || [];
    for (let offset = 1; offset <= 4; offset += 1) {
      const candidate = row[colIndex + offset];
      if (isMeaningfulValue(candidate)) return String(candidate).trim();
    }

    for (let down = 1; down <= 3; down += 1) {
      const nextRow = rows[rowIndex + down];
      if (!nextRow) continue;
      for (let offset = 0; offset <= 4; offset += 1) {
        const candidate = nextRow[colIndex + offset];
        if (isMeaningfulValue(candidate)) return String(candidate).trim();
      }
    }

    return '';
  }

  function isMeaningfulValue(value) {
    if (value === null || value === undefined) return false;
    const text = String(value).trim();
    if (!text) return false;
    return !isLabelish(text);
  }

  function isLabelish(text) {
    const keywords = [
      '测序工厂邮寄地址',
      '邮寄地址',
      '联系方式',
      '测序机构',
      '测序工厂',
      '送测时间',
      '总数据量',
      '备注',
      '实验地点',
      '测序平台',
      '派单商务',
      '送测方式',
      'i5index方向'
    ];
    return keywords.some((keyword) => text.includes(keyword));
  }

  function forwardFillRow(row) {
    const result = [];
    let lastValue = '';
    row.forEach((cell) => {
      if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
        lastValue = String(cell).trim();
      }
      result.push(lastValue);
    });
    return result;
  }

  function normalizeHeader(value) {
    return String(value || '')
      .replace(/\s+/g, '')
      .replace(/[（）()]/g, '')
      .replace(/[\-_/]/g, '')
      .replace(/[’']/g, '')
      .replace(/[＊*]/g, '')
      .replace(/μ/g, 'u')
      .toLowerCase();
  }

  function mapColumnKey(headerKey) {
    const map = {
      序号: 'seq',
      客户: 'customer',
      备注: 'rowRemark',
      文库id: 'libId',
      样本登记名: 'sampleName',
      送测文库id: 'sendLibId',
      文库类型: 'libType',
      index编码: 'indexCode',
      indexi5: 'indexI5',
      indexi7: 'indexI7',
      待测数据量gb: 'targetGb',
      文库样本存储位置: 'storagePos',
      文库所在实验室: 'lab',
      bp值: 'bp',
      文库浓度ngul: 'concNgUl',
      质检文库浓度ngul: 'qcConcNgUl',
      使用体积量ul: 'useVolUl',
      剩余样本存储位置位置未变更打对号: 'storagePos2',
      测序平台: 'platform',
      平台: 'platform',
      测序机构: 'vendor',
      实验地点: 'labSite'
    };
    return map[headerKey] || null;
  }

  function castValue(key, value) {
    if (value === null || value === undefined) return value;
    if (['targetGb', 'bp', 'concNgUl', 'qcConcNgUl', 'useVolUl'].includes(key)) {
      const num = Number(value);
      return Number.isFinite(num) ? num : value;
    }
    return String(value).trim();
  }

  function normalizeVendor(value) {
    const text = String(value || '').trim();
    if (!text) return '未知';
    if (text.includes('格致')) return '格致博雅';
    if (text.includes('海普洛斯') || text.includes('海普')) return '海普洛斯';
    if (text.includes('明码')) return '明码';
    return text;
  }

  function normalizePlatform(value) {
    const text = String(value || '').trim();
    if (!text) return '未知';
    const lower = text.toLowerCase();
    if (lower.includes('atac')) return 'Element(ATAC)';
    if (lower.includes('element')) return 'Element';
    if (lower.includes('nova')) return 'Nova';
    if (lower.includes('bgi') || lower.includes('t7') || lower.includes('dnbseq')) return 'BGI(T7)';
    return text;
  }

  function extractCxpsFromName(name) {
    const match = String(name).match(/CXPS-[A-Za-z0-9-]+/i);
    return match ? match[0] : '';
  }

  function groupRows(rows) {
    const map = new Map();
    rows.forEach((row) => {
      const vendor = row.vendor || '未知';
      const platform = row.platform || '未知';
      const labSite = row.labSite || '未知';
      const key = `${vendor}|${platform}|${labSite}`;
      if (!map.has(key)) {
        map.set(key, {
          key,
          vendor,
          platform,
          labSite,
          rows: [],
          cxpsNos: new Set()
        });
      }
      const group = map.get(key);
      group.rows.push(row);
      group.cxpsNos.add(row.cxpsNo || '未知');
    });

    return Array.from(map.values()).map((group) => {
      const totalGb = group.rows.reduce((sum, row) => {
        const val = Number(row.targetGb);
        return Number.isFinite(val) ? sum + val : sum;
      }, 0);
      const libTypes = Array.from(new Set(group.rows.map((row) => row.libType).filter(Boolean)));
      return {
        ...group,
        totalGb,
        libTypes,
        templateKey: detectTemplateKey(group.vendor, group.platform)
      };
    });
  }

  function detectTemplateKey(vendor, platform) {
    if (vendor.includes('格致')) return 'gzb';
    if (vendor.includes('海普')) return platform.includes('Nova') ? 'hplosNova' : 'hplosBgi';
    if (vendor.includes('明码')) return platform.includes('ATAC') || platform.includes('Element') ? 'mingmaAtac' : 'mingma';
    return 'unknown';
  }

  function renderGroups() {
    const scrollY = window.scrollY;
    const activeIndex = Array.from(elements.groupTabs.querySelectorAll('.tab-btn')).findIndex((btn) => btn.classList.contains('active'));
    state.groups = groupRows(state.rows);
    elements.groupTabs.innerHTML = '';
    elements.groupPanels.innerHTML = '';

    if (!state.groups.length) {
      const empty = document.createElement('div');
      empty.className = 'tab-panel active';
      empty.innerHTML = '<p>请先导入烈冰送测单 Excel。</p>';
      elements.groupPanels.appendChild(empty);
      return;
    }

    state.groups.forEach((group, index) => {
      const tabBtn = document.createElement('button');
      const shouldActive = activeIndex === -1 ? index === 0 : index === activeIndex;
      tabBtn.className = `tab-btn ${shouldActive ? 'active' : ''}`;
      tabBtn.textContent = buildGroupTitle(group);
      tabBtn.addEventListener('click', () => activateTab(index));
      elements.groupTabs.appendChild(tabBtn);

      const panel = document.createElement('div');
      panel.className = `tab-panel ${shouldActive ? 'active' : ''}`;
      panel.appendChild(buildOverview(group));
      panel.appendChild(buildRulePanel(group));
      panel.appendChild(buildDataTable(group));
      panel.appendChild(buildGroupActions(group));
      elements.groupPanels.appendChild(panel);
    });

    if (scrollY) {
      window.requestAnimationFrame(() => {
        window.scrollTo(0, scrollY);
      });
    }
  }

  function activateTab(index) {
    const tabs = elements.groupTabs.querySelectorAll('.tab-btn');
    const panels = elements.groupPanels.querySelectorAll('.tab-panel');
    tabs.forEach((tab, idx) => tab.classList.toggle('active', idx === index));
    panels.forEach((panel, idx) => panel.classList.toggle('active', idx === index));
  }

  function buildGroupTitle(group) {
    const date = state.ui.date ? formatDateDisplay(state.ui.date) : formatDateDisplay(formatDateInput(new Date()));
    const count = group.rows.length;
    return `${group.vendor}-${group.platform}-${date}送测-${count}个样本`;
  }

  function buildOverview(group) {
    const missingRuleCount = group.rows.filter((row) => !getRule(group.vendor, group.platform, row.libType)).length;
    const wrapper = document.createElement('div');
    wrapper.className = 'overview';

    const cards = [
      { label: '测序工厂', value: group.vendor },
      { label: '测序平台', value: group.platform },
      { label: '送样地点', value: group.labSite },
      { label: '送测单数', value: `${group.cxpsNos.size}` },
      { label: '样本数', value: `${group.rows.length}` },
      { label: '总数据量(GB)', value: group.totalGb ? group.totalGb.toFixed(2) : '0' },
      { label: '未匹配规则', value: `${missingRuleCount}` }
    ];

    cards.forEach((card) => {
      const el = document.createElement('div');
      el.className = 'overview-card';
      el.innerHTML = `<h4>${card.label}</h4><p>${card.value}</p>`;
      wrapper.appendChild(el);
    });

    return wrapper;
  }

  function buildRulePanel(group) {
    const panel = document.createElement('div');
    panel.className = 'rule-panel';

    const missingTypes = group.libTypes.filter((type) => !getRule(group.vendor, group.platform, type));
    const fieldSet = getRuleFields(group);

    panel.innerHTML = `
      <h3>规则面板</h3>
      <p>规则作用域：测序工厂 × 测序平台 × 文库类型</p>
      <p>当前工厂：${group.vendor}</p>
      <div><strong>未匹配文库类型：</strong>${missingTypes.length ? missingTypes.join('、') : '无'}</div>
    `;

    const select = document.createElement('select');
    select.innerHTML = group.libTypes.map((type) => `<option value="${type}">${type}</option>`).join('');
    if (!group.libTypes.length) {
      const option = document.createElement('option');
      option.value = '';
      option.textContent = '暂无文库类型';
      select.appendChild(option);
    }

    const formGrid = document.createElement('div');
    formGrid.className = 'rule-grid';

    const inputs = {};
    fieldSet.forEach((field) => {
      const wrapper = document.createElement('div');
      const label = document.createElement('label');
      label.textContent = field.label;
      const input = document.createElement('input');
      input.type = 'text';
      input.dataset.key = field.key;
      wrapper.appendChild(label);
      wrapper.appendChild(input);
      formGrid.appendChild(wrapper);
      inputs[field.key] = input;
    });

    const fillForm = () => {
      const type = select.value;
      const rule = getRule(group.vendor, group.platform, type) || {};
      fieldSet.forEach((field) => {
        if (inputs[field.key]) inputs[field.key].value = rule[field.key] || '';
      });
    };

    select.addEventListener('change', fillForm);

    const actions = document.createElement('div');
    actions.className = 'rule-actions';

    const saveBtn = document.createElement('button');
    saveBtn.textContent = '保存规则';
    saveBtn.className = 'primary';
    saveBtn.addEventListener('click', () => {
      const type = select.value;
      if (!type) return;
      const existing = getRule(group.vendor, group.platform, type) || {};
      const rule = {
        ...existing,
        vendor: group.vendor,
        platform: group.platform,
        internalLibType: type
      };
      fieldSet.forEach((field) => {
        if (inputs[field.key]) rule[field.key] = inputs[field.key].value.trim();
      });
      setRule(rule);
      renderGroups();
      toast('规则已保存');
    });

    const cloneBtn = document.createElement('button');
    cloneBtn.textContent = '复制规则';
    cloneBtn.addEventListener('click', () => {
      const type = select.value;
      if (!type) return;
      const newType = prompt('请输入新的内部文库类型');
      if (!newType) return;
      const current = getRule(group.vendor, group.platform, type) || {};
      setRule({ ...current, vendor: group.vendor, platform: group.platform, internalLibType: newType });
      renderGroups();
      toast('规则已复制');
    });

    actions.appendChild(saveBtn);
    actions.appendChild(cloneBtn);

    panel.appendChild(select);
    panel.appendChild(formGrid);
    panel.appendChild(actions);

    fillForm();

    return panel;
  }

  function buildDataTable(group) {
    const wrapper = document.createElement('div');
    wrapper.className = 'table-wrap';

    const table = document.createElement('table');
    table.className = 'data-table';
    const columns = getTableColumns(group);

    table.innerHTML = `
      <thead><tr>${columns.map((col) => `<th>${col.label}</th>`).join('')}</tr></thead>
      <tbody></tbody>
    `;

    const tbody = table.querySelector('tbody');

    group.rows.forEach((row) => {
      const status = validateRow(row, group);
      const rule = getRule(group.vendor, group.platform, row.libType) || {};
      const mapped = buildMappedValues(row, rule, group);

      const tr = document.createElement('tr');

      columns.forEach((col) => {
        const td = document.createElement('td');
        if (col.type === 'status') {
          td.innerHTML = renderStatusPill(status);
          tr.appendChild(td);
          return;
        }

        if (col.type === 'mapped') {
          const value = mapped[col.key] || '';
          if (col.special === 'remark') {
            const wrapper = document.createElement('div');
            wrapper.className = 'remark-cell';
            const urgentBtn = document.createElement('button');
            urgentBtn.type = 'button';
            urgentBtn.className = 'urgent-btn';
            urgentBtn.textContent = '加急';
            const input = document.createElement('input');
            input.className = 'remark-input';
            input.value = value;
            const syncRemark = () => {
              row.overrides[col.key] = input.value.trim();
              input.classList.toggle('urgent', input.value.includes('加急'));
            };
            urgentBtn.addEventListener('click', () => {
              const current = input.value.trim();
              if (!current.includes('加急')) {
                input.value = current ? `${current}；加急` : '加急';
              }
              syncRemark();
            });
            input.addEventListener('input', syncRemark);
            syncRemark();
            wrapper.appendChild(urgentBtn);
            wrapper.appendChild(input);
            td.appendChild(wrapper);
          } else if (col.editable) {
            const input = document.createElement('input');
            input.value = value;
            input.addEventListener('input', () => {
              row.overrides[col.key] = input.value.trim();
            });
            td.appendChild(input);
          } else {
            td.textContent = safeText(value);
          }
          tr.appendChild(td);
          return;
        }

        const rawValue = col.display ? col.display(row, group) : (row[col.field] ?? '');
        td.textContent = safeText(rawValue);
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    wrapper.appendChild(table);
    return wrapper;
  }

  function buildGroupActions(group) {
    const wrapper = document.createElement('div');
    wrapper.className = 'group-actions';

    const downloadBtn = document.createElement('button');
    downloadBtn.textContent = '生成并下载该组 Excel';
    downloadBtn.className = 'primary';
    downloadBtn.addEventListener('click', () => downloadGroup(group));

    const copyBtn = document.createElement('button');
    copyBtn.textContent = '复制邮件正文';
    copyBtn.addEventListener('click', () => copyEmail(group));

    wrapper.appendChild(downloadBtn);
    wrapper.appendChild(copyBtn);
    return wrapper;
  }

  function getRuleFields(group) {
    return RULE_FIELDS_BY_TEMPLATE[group.templateKey] || RULE_FIELDS;
  }

  function getTableColumns(group) {
    return DISPLAY_COLUMNS;
  }

  function buildMappedValues(row, rule, group) {
    const defaults = {
      vendorLibType: rule.vendorLibType || '',
      vendorLibStructure: rule.vendorLibStructure || '',
      vendorLibProcess: rule.vendorLibProcess || '',
      defaultVolumeUl: resolveVolume(rule, group),
      defaultSpecies: rule.defaultSpecies || '',
      defaultSpecialSeq: rule.defaultSpecialSeq || '',
      defaultBaseBalance: rule.defaultBaseBalance || '',
      defaultSeqStrategy: rule.defaultSeqStrategy || '150+10+10+150',
      defaultProductType: rule.defaultProductType || 'NovaSeq X plus PE150包数据量',
      defaultRemark: rule.defaultRemark || '',
      defaultAdapterType: rule.defaultAdapterType || '',
      defaultPhosphorylation: rule.defaultPhosphorylation || '未磷酸化',
      defaultCyclization: rule.defaultCyclization || '未环化',
      defaultKitVersion: rule.defaultKitVersion || '',
      defaultSeqNumber: rule.defaultSeqNumber || '',
      defaultPlatformName: rule.defaultPlatformName || ''
    };

    const mapped = { ...defaults };
    Object.keys(row.overrides || {}).forEach((key) => {
      if (row.overrides[key] !== undefined && row.overrides[key] !== '') {
        mapped[key] = row.overrides[key];
      }
    });

    return mapped;
  }

  function resolveVolume(rule, group) {
    if (rule.defaultVolumeUl) return rule.defaultVolumeUl;
    if (group.templateKey === 'gzb') return '15';
    if (group.templateKey === 'mingma' || group.templateKey === 'mingmaAtac') return '10';
    return '';
  }

  function validateRow(row, group) {
    const errors = [];
    const warns = [];
    if (!row.sendLibId) errors.push('送测文库ID缺失');
    if (!isFiniteNumber(row.targetGb)) errors.push('待测数据量缺失/非数字');

    const templateKey = group.templateKey;
    const isAtac = templateKey === 'mingmaAtac';
    if (isAtac) {
      if (!row.indexCode) errors.push('ATAC index编码缺失');
      const code = normalizeIndexCode(row.indexCode);
      if (code && !state.indexKit[code]) warns.push('ATAC index未匹配');
    } else {
      if (!row.indexI7) errors.push('indexI7缺失');
      if (requiresI5(templateKey) && !row.indexI5) errors.push('indexI5缺失');
    }

    const rule = getRule(group.vendor, group.platform, row.libType);
    if (!rule) warns.push('未匹配规则');

    if (errors.length) return { level: 'error', message: errors.join('，') };
    if (warns.length) return { level: 'warn', message: warns.join('，') };
    return { level: 'ok', message: '已映射' };
  }

  function requiresI5(templateKey) {
    return templateKey !== 'mingmaAtac';
  }

  function renderStatusPill(status) {
    const levelMap = {
      error: 'error',
      warn: 'warn',
      ok: 'ok'
    };
    const labelMap = {
      error: '❌',
      warn: '⚠️',
      ok: '✅'
    };
    const level = levelMap[status.level] || 'ok';
    return `<span class="status-pill small ${level}">${labelMap[status.level] || '✅'} ${status.message}</span>`;
  }

  function safeText(value) {
    if (value === null || value === undefined) return '';
    return String(value);
  }

  function formatDateInput(date) {
    const year = date.getFullYear();
    const month = `${date.getMonth() + 1}`.padStart(2, '0');
    const day = `${date.getDate()}`.padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function formatDateDisplay(dateString) {
    const date = new Date(dateString);
    const week = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'][date.getDay()];
    const month = date.getMonth() + 1;
    const day = date.getDate();
    return `${date.getFullYear()}年${month}月${day}日${week}`;
  }

  function isFiniteNumber(value) {
    const num = Number(value);
    return Number.isFinite(num);
  }

  function getRule(vendor, platform, internalLibType) {
    if (!internalLibType) return null;
    const key = buildRuleKey(vendor, platform, internalLibType);
    return state.rules[key] || null;
  }

  function setRule(rule) {
    const key = buildRuleKey(rule.vendor, rule.platform, rule.internalLibType);
    state.rules[key] = rule;
    saveRules();
  }

  function buildRuleKey(vendor, platform, internalLibType) {
    return `${vendor}|${platform}|${internalLibType}`;
  }

  function loadRules() {
    try {
      const raw = localStorage.getItem('dispatchRules');
      if (!raw) return {};
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        const map = {};
        parsed.forEach((rule) => {
          if (rule.vendor && rule.platform && rule.internalLibType) {
            map[buildRuleKey(rule.vendor, rule.platform, rule.internalLibType)] = rule;
          }
        });
        return map;
      }
      return parsed || {};
    } catch (err) {
      return {};
    }
  }

  function saveRules() {
    const list = Object.values(state.rules);
    localStorage.setItem('dispatchRules', JSON.stringify(list, null, 2));
  }

  function exportRulesJson() {
    const list = Object.values(state.rules);
    const blob = new Blob([JSON.stringify(list, null, 2)], { type: 'application/json' });
    downloadBlob(blob, `rules-${Date.now()}.json`);
  }

  async function handleImportRules(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    const text = await file.text();
    try {
      const parsed = JSON.parse(text);
      const list = Array.isArray(parsed) ? parsed : Object.values(parsed);
      list.forEach((rule) => {
        if (rule.vendor && rule.platform && rule.internalLibType) {
          state.rules[buildRuleKey(rule.vendor, rule.platform, rule.internalLibType)] = rule;
        }
      });
      saveRules();
      renderGroups();
      toast('规则已导入');
    } catch (err) {
      toast('规则 JSON 解析失败');
    }
  }

  async function downloadAllGroups() {
    if (!state.groups.length) return;
    for (const group of state.groups) {
      await downloadGroup(group, true);
    }
  }

  function copyAllEmails() {
    const content = state.groups.map((group) => buildEmailContent(group)).join('\n\n-----\n\n');
    copyText(content);
  }

  async function downloadGroup(group, silent) {
    if (group.templateKey === 'unknown') {
      toast('未识别的测序机构/平台，无法导出');
      return;
    }
    const buffer = state.templateBuffers[group.templateKey];
    if (!buffer) {
      toast('模板未加载，请先上传');
      return;
    }

    try {
      if (window.ExcelJS && window.ExcelJS.Workbook) {
        const workbook = new window.ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        await writeGroupToExcelWorkbook(workbook, group);
        let out = await workbook.xlsx.writeBuffer({ useStyles: true, useSharedStrings: true });
        if (group.templateKey === 'gzb') {
          out = await patchGzbTemplateArtifacts(out, buffer);
        }
        const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const fileName = buildFileName(group);
        downloadBlob(blob, fileName);
        if (!silent) toast('已生成并下载');
        return;
      }

      const workbook = XLSX.read(buffer, { type: 'array' });
      await writeGroupToWorkbook(workbook, group);
      let out = XLSX.write(workbook, { type: 'array', bookType: 'xlsx', compression: true, cellStyles: true });
      if (group.templateKey === 'gzb') {
        out = await patchGzbTemplateArtifacts(out, buffer);
      }
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const fileName = buildFileName(group);
      downloadBlob(blob, fileName);
      if (!silent) toast('已生成并下载');
    } catch (err) {
      console.error(err);
      toast('导出失败，请检查模板');
    }
  }

  function buildFileName(group) {
    const date = state.ui.date ? state.ui.date.replace(/-/g, '') : formatDateInput(new Date()).replace(/-/g, '');
    const cxpsList = Array.from(group.cxpsNos)
      .map((item) => String(item || '').replace(/cxps[-]?/i, ''))
      .filter((item) => item);
    const tail = cxpsList.length ? cxpsList.join('、') : '';
    return `${group.vendor}-${group.platform}-${date}送测-${group.rows.length}个样本${tail}.xlsx`;
  }

  function buildEmailContent(group) {
    const dateLabel = state.ui.date || formatDateInput(new Date());
    const subject = `${group.labSite}-${group.vendor} ${dateLabel} 外包送测信息单`;
    const platform = formatPlatformForEmail(group.platform);
    const cxpsList = Array.from(group.cxpsNos).join('、');
    const internalLine = cxpsList ? `（内部单号记录：${cxpsList}）` : '';
    const body = `老师，您好：\n\n  附件是今天邮寄的 ${platform} 测序文库信息单，请查收，谢谢。\n  顺丰单号：\n${internalLine}\n祝好！`;
    return `Subject：${subject}\n\n${body}`;
  }

  function copyEmail(group) {
    copyText(buildEmailContent(group));
  }

  function formatPlatformForEmail(platform) {
    const text = String(platform || '').trim();
    if (!text) return '';
    if (text.toLowerCase().includes('element')) return 'Element';
    return text;
  }

  function copyText(text) {
    navigator.clipboard.writeText(text).then(() => {
      toast('已复制到剪贴板');
    }).catch(() => {
      toast('复制失败，请手动复制');
    });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  async function patchGzbTemplateArtifacts(outputBuffer, templateBuffer) {
    try {
      const JSZipLib = window.JSZip;
      if (!JSZipLib || !JSZipLib.loadAsync) return outputBuffer;
      const [zipOut, zipTpl] = await Promise.all([
        JSZipLib.loadAsync(outputBuffer),
        JSZipLib.loadAsync(templateBuffer)
      ]);

      const [workbookXml, workbookRelsXml] = await Promise.all([
        readZipText(zipOut, 'xl/workbook.xml'),
        readZipText(zipOut, 'xl/_rels/workbook.xml.rels')
      ]);
      const [tplWorkbookXml, tplWorkbookRelsXml] = await Promise.all([
        readZipText(zipTpl, 'xl/workbook.xml'),
        readZipText(zipTpl, 'xl/_rels/workbook.xml.rels')
      ]);

      const sheetNames = ['文库信息单', '文库信息'];
      const sheetPath = normalizeZipPath(findSheetPathByName(workbookXml, workbookRelsXml, sheetNames) || '');
      const tplSheetPath = normalizeZipPath(findSheetPathByName(tplWorkbookXml, tplWorkbookRelsXml, sheetNames) || sheetPath);
      if (!sheetPath || !tplSheetPath) return outputBuffer;

      const fullSheetPath = sheetPath.startsWith('xl/') ? sheetPath : `xl/${sheetPath}`;
      const tplFullSheetPath = tplSheetPath.startsWith('xl/') ? tplSheetPath : `xl/${tplSheetPath}`;
      const relPath = buildSheetRelPath(fullSheetPath);
      const tplRelPath = buildSheetRelPath(tplFullSheetPath);

      const [outSheetXml, tplSheetXml, tplRelXml, outTypesXml, tplTypesXml] = await Promise.all([
        readZipText(zipOut, fullSheetPath),
        readZipText(zipTpl, tplFullSheetPath),
        readZipText(zipTpl, tplRelPath),
        readZipText(zipOut, '[Content_Types].xml'),
        readZipText(zipTpl, '[Content_Types].xml')
      ]);

      let mergedSheetXml = outSheetXml;
      mergedSheetXml = mergeWorksheetNamespaces(mergedSheetXml, tplSheetXml);
      mergedSheetXml = injectTemplateControls(mergedSheetXml, tplSheetXml);
      if (mergedSheetXml) zipOut.file(fullSheetPath, mergedSheetXml);
      if (tplRelXml) zipOut.file(relPath, tplRelXml);
      const mergedTypesXml = mergeContentTypes(outTypesXml, tplTypesXml);
      if (mergedTypesXml) zipOut.file('[Content_Types].xml', mergedTypesXml);

      await copyZipEntries(zipOut, zipTpl, [
        /^xl\/drawings\/.+/,
        /^xl\/drawings\/_rels\/.+/,
        /^xl\/ctrlProps\/.+/,
        /^xl\/media\/.+/,
        /^customXml\/.+/,
        /^customXml\/_rels\/.+/
      ]);

      return await zipOut.generateAsync({ type: 'arraybuffer' });
    } catch (err) {
      console.warn('GZB template patch failed', err);
      return outputBuffer;
    }
  }

  function readZipText(zip, path) {
    const normalized = normalizeZipPath(path);
    const file = zip.file(normalized) || zip.file(path) || zip.file(`/${normalized}`);
    if (!file) return Promise.resolve('');
    return file.async('string');
  }

  function buildSheetRelPath(sheetPath) {
    const normalized = normalizeZipPath(sheetPath);
    const withXl = normalized.startsWith('xl/') ? normalized : `xl/${normalized}`;
    return withXl.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
  }

  function findSheetPathByName(workbookXml, relsXml, names) {
    if (!workbookXml || !relsXml || !names || !names.length) return null;
    const relMap = {};
    const relRegex = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/g;
    let match = null;
    while ((match = relRegex.exec(relsXml))) {
      relMap[match[1]] = match[2];
    }
    for (const name of names) {
      const sheetRegex = new RegExp(`<sheet[^>]*name="${escapeRegex(name)}"[^>]*r:id="([^"]+)"[^>]*/?>`, 'i');
      const sheetMatch = workbookXml.match(sheetRegex);
      if (!sheetMatch) continue;
      const target = relMap[sheetMatch[1]];
      if (target) return target.replace(/^\//, '');
    }
    return null;
  }

  function escapeRegex(value) {
    return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function normalizeZipPath(path) {
    return String(path || '').replace(/^\/+/, '');
  }

  function mergeWorksheetNamespaces(outXml, tplXml) {
    if (!outXml || !tplXml) return outXml;
    const outTagMatch = outXml.match(/<worksheet[^>]*>/);
    const tplTagMatch = tplXml.match(/<worksheet[^>]*>/);
    if (!outTagMatch || !tplTagMatch) return outXml;
    const outTag = outTagMatch[0];
    const tplTag = tplTagMatch[0];
    const attrRegex = /\s+xmlns:[^=]+="[^"]*"/g;
    const outAttrs = outTag.match(attrRegex) || [];
    const tplAttrs = tplTag.match(attrRegex) || [];
    const missing = tplAttrs.filter((attr) => !outAttrs.includes(attr));
    if (!missing.length) return outXml;
    const newTag = outTag.replace(/>$/, `${missing.join('')}>`);
    return outXml.replace(outTag, newTag);
  }

  function mergeContentTypes(outXml, tplXml) {
    if (!outXml || !tplXml) return outXml || tplXml;
    const defaultRegex = /<Default[^>]*Extension="([^"]+)"[^>]*\/>/g;
    const overrideRegex = /<Override[^>]*PartName="([^"]+)"[^>]*\/>/g;
    const extract = (xml, regex) => {
      const map = new Map();
      let match = null;
      while ((match = regex.exec(xml))) {
        map.set(match[1], match[0]);
      }
      return map;
    };
    const outDefaults = extract(outXml, defaultRegex);
    const outOverrides = extract(outXml, overrideRegex);
    const tplDefaults = extract(tplXml, defaultRegex);
    const tplOverrides = extract(tplXml, overrideRegex);
    const missingDefaults = Array.from(tplDefaults.entries()).filter(([key]) => !outDefaults.has(key)).map(([, tag]) => tag);
    const missingOverrides = Array.from(tplOverrides.entries()).filter(([key]) => !outOverrides.has(key)).map(([, tag]) => tag);
    if (!missingDefaults.length && !missingOverrides.length) return outXml;
    return outXml.replace('</Types>', `${missingDefaults.join('')}${missingOverrides.join('')}</Types>`);
  }

  function injectTemplateControls(outXml, tplXml) {
    if (!outXml || !tplXml) return outXml;
    const tailMatch = tplXml.match(/<drawing[\s\S]*<\/worksheet>/);
    if (!tailMatch) return outXml;
    const tail = tailMatch[0].replace(/<\/worksheet>\s*$/, '');
    let cleaned = outXml;
    cleaned = cleaned.replace(/<drawing[^>]*\/>/g, '');
    cleaned = cleaned.replace(/<legacyDrawing[^>]*\/>/g, '');
    cleaned = cleaned.replace(/<mc:AlternateContent[\s\S]*?<controls>[\s\S]*?<\/controls>[\s\S]*?<\/mc:AlternateContent>/g, '');
    if (!/<\/worksheet>\s*$/.test(cleaned)) return outXml;
    return cleaned.replace(/<\/worksheet>\s*$/, `${tail}</worksheet>`);
  }

  async function copyZipEntries(zipOut, zipTpl, patterns) {
    const entries = [];
    zipTpl.forEach((path, file) => {
      if (file.dir) return;
      if (patterns.some((regex) => regex.test(path))) {
        entries.push({ path, file });
      }
    });
    for (const entry of entries) {
      const data = await entry.file.async('uint8array');
      zipOut.file(entry.path, data);
    }
  }

  function toast(message) {
    elements.toast.textContent = message;
    elements.toast.classList.add('show');
    setTimeout(() => elements.toast.classList.remove('show'), 2200);
  }

  function parseIndexKit(buffer) {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const mapping = {};
    rows.forEach((row) => {
      const flatRow = normalizeIndexKitRow(row);
      if (!flatRow.length) return;
      const codeRaw = flatRow[0];
      const code = normalizeIndexCode(codeRaw);
      if (!code || /index/i.test(codeRaw) || String(codeRaw).includes('序号')) return;
      const seqs = flatRow.slice(1, 5).map((cell) => String(cell || '').trim()).filter(Boolean);
      if (seqs.length >= 4) {
        mapping[code] = seqs.slice(0, 4);
      }
    });
    return mapping;
  }

  async function writeGroupToWorkbook(workbook, group) {
    const templateKey = group.templateKey;
    if (templateKey === 'gzb') return writeGzb(workbook, group);
    if (templateKey === 'hplosBgi') return writeHplosBgi(workbook, group);
    if (templateKey === 'hplosNova') return writeHplosNova(workbook, group);
    if (templateKey === 'mingma') return writeMingma(workbook, group);
    if (templateKey === 'mingmaAtac') return writeMingmaAtac(workbook, group);
  }

  async function writeGroupToExcelWorkbook(workbook, group) {
    const templateKey = group.templateKey;
    if (templateKey === 'gzb') return writeGzbExcel(workbook, group);
    if (templateKey === 'hplosBgi') return writeHplosBgiExcel(workbook, group);
    if (templateKey === 'hplosNova') return writeHplosNovaExcel(workbook, group);
    if (templateKey === 'mingma') return writeMingmaExcel(workbook, group);
    if (templateKey === 'mingmaAtac') return writeMingmaAtacExcel(workbook, group);
  }

  function writeGzb(workbook, group) {
    const sheetName = findSheetName(workbook, ['文库信息', '文库信息单']);
    const sheet = workbook.Sheets[sheetName];
    const rows = sheetToRows(sheet);

    const detailHeaderRowIndex = findHeaderRow(rows, ['样本名称', '文库类型', '文库结构']);
    const splitHeaderRowIndex = findHeaderRow(rows, ['子文库名称', 'index', 'indexf', 'indexi7']);

    if (detailHeaderRowIndex === -1 || splitHeaderRowIndex === -1) throw new Error('模板头未匹配');

    const detailHeaderRow = trimHeaderRow(rows[detailHeaderRowIndex]);
    const splitHeaderRow = trimHeaderRow(rows[splitHeaderRowIndex]);

    const splitColMap = buildColMap(splitHeaderRow, {
      seq: ['序号'],
      sampleName: ['样本名称'],
      subLibName: ['子文库名称'],
      species: ['物种'],
      specialSeq: ['有无特殊序列'],
      baseBalance: ['碱基', '碱基是否均衡'],
      dataGb: ['数据量', '数据量(raw/g)'],
      indexCode: ['index编号', 'index编码'],
      indexI7: ['indexf', 'indexi7', 'indexf(i7)'],
      indexI5: ['indexr', 'indexi5', 'indexr(i5)'],
      remark: ['备注']
    });

    const detailColMap = buildColMap(detailHeaderRow, {
      seq: ['序号'],
      sampleName: ['样本名称'],
      libType: ['文库类型'],
      libStructure: ['文库结构'],
      libProcess: ['处理情况', '文库处理情况'],
      bp: ['文库片段大小', 'bp', 'bp值'],
      conc: ['浓度', '浓度(ng/ul)', '浓度(ng/μl)'],
      volume: ['体积(ul)', '体积'],
      dataGb: ['数据量', '数据量(raw/g)'],
      remark: ['备注']
    });

    writeRowsToSheet(sheet, detailHeaderRowIndex + 1, detailColMap, group.rows, (field, row, idx) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      switch (field) {
        case 'seq': return idx + 1;
        case 'sampleName': return row.sendLibId || '';
        case 'libType': return mapped.vendorLibType;
        case 'libStructure': return mapped.vendorLibStructure;
        case 'libProcess': return mapped.vendorLibProcess;
        case 'bp': return row.bp || '';
        case 'conc': return row.concNgUl || '';
        case 'volume': return mapped.defaultVolumeUl;
        case 'dataGb': return row.targetGb || '';
        case 'remark': return mapped.defaultRemark;
        default: return '';
      }
    });

    writeRowsToSheet(sheet, splitHeaderRowIndex + 1, splitColMap, group.rows, (field, row, idx) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      switch (field) {
        case 'seq': return idx + 1;
        case 'sampleName': return row.sendLibId || '';
        case 'subLibName': return row.sendLibId || '';
        case 'species': return mapped.defaultSpecies || null;
        case 'specialSeq': return mapped.defaultSpecialSeq || null;
        case 'baseBalance': return mapped.defaultBaseBalance || null;
        case 'dataGb': return row.targetGb || '';
        case 'indexCode': return row.indexCode || '';
        case 'indexI7': return row.indexI7 || '';
        case 'indexI5': return row.indexI5 || '';
        case 'remark': return '';
        default: return '';
      }
    });
  }

  function writeHplosBgi(workbook, group) {
    const sheetName = findSheetName(workbook, ['文库信息']);
    const sheet = workbook.Sheets[sheetName];
    const rows = sheetToRows(sheet);

    const headerRowIndex = findHeaderRow(rows, ['子文库名称', '测序策略', 'index']);
    if (headerRowIndex === -1) throw new Error('模板头未匹配');

    const headerRow = trimHeaderRow(rows[headerRowIndex]);
    const colMap = buildColMap(headerRow, {
      originName: ['原样本名称', '试管上的样本名称'],
      mixName: ['混文库名称'],
      subName: ['子文库名称'],
      kitVersion: ['试剂盒版本'],
      libType: ['文库类型'],
      libStructure: ['文库结构'],
      phosphorylation: ['磷酸化'],
      cyclization: ['环化'],
      seqNumber: ['序列号'],
      indexI7: ['index-i7', 'indexi7'],
      indexI5: ['index-i5', 'indexi5'],
      subGb: ['子文库数据量'],
      mixGb: ['混库总数据量'],
      productType: ['产品类型'],
      seqStrategy: ['测序策略'],
      remark: ['备注']
    });
    refineIndexColumnsInRangeSheet(headerRow, colMap, [
      'originName',
      'subName',
      'libType',
      'subGb',
      'mixGb',
      'seqStrategy',
      'remark'
    ]);

    const dataStartRow = headerRowIndex + 2;
    writeRowsToSheet(sheet, dataStartRow, colMap, group.rows, (field, row) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      const exon = isExon(row.libType);
      switch (field) {
        case 'originName': return exon ? '' : row.sendLibId || '';
        case 'mixName': return exon ? '' : row.sendLibId || '';
        case 'subName': return row.sendLibId || '';
        case 'kitVersion': return mapped.defaultKitVersion || '';
        case 'libType': return mapped.vendorLibType;
        case 'libStructure': return mapped.vendorLibStructure || '';
        case 'phosphorylation': return mapped.defaultPhosphorylation || '未磷酸化';
        case 'cyclization': return mapped.defaultCyclization || '未环化';
        case 'seqNumber': return mapped.defaultSeqNumber || '';
        case 'indexI7': return row.indexI7 || '';
        case 'indexI5': return row.indexI5 || '';
        case 'subGb': return row.targetGb || '';
        case 'mixGb': return row.targetGb || '';
        case 'productType': return mapped.defaultProductType || '';
        case 'seqStrategy': return mapped.defaultSeqStrategy || '150+10+10+150';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });
  }

  function writeHplosNova(workbook, group) {
    const sheetName = findSheetName(workbook, ['客户自建库样本信息', '样本信息']);
    const sheet = workbook.Sheets[sheetName];
    const rows = sheetToRows(sheet);

    const headerRowIndex = findHeaderRow(rows, ['子文库名称', '测序策略', '产品类型']);
    if (headerRowIndex === -1) throw new Error('模板头未匹配');

    const headerRow = trimHeaderRow(rows[headerRowIndex]);
    const colMap = buildColMap(headerRow, {
      originName: ['原样本名称', '试管上的名称'],
      mixName: ['混文库名称'],
      subName: ['子文库名称'],
      libType: ['文库类型'],
      indexI7: ['index-i7', 'indexi7'],
      indexI5: ['index-i5', 'indexi5'],
      subGb: ['子文库数据量'],
      mixGb: ['混库总数据量'],
      productType: ['产品类型'],
      seqStrategy: ['测序策略'],
      remark: ['备注']
    });
    refineIndexColumnsInRangeSheet(headerRow, colMap, [
      'originName',
      'subName',
      'libType',
      'subGb',
      'mixGb',
      'seqStrategy',
      'remark'
    ]);

    const dataStartRow = headerRowIndex + 2;
    writeRowsToSheet(sheet, dataStartRow, colMap, group.rows, (field, row) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      const exon = isExon(row.libType);
      switch (field) {
        case 'originName': return exon ? '' : row.sendLibId || '';
        case 'mixName': return exon ? '' : row.sendLibId || '';
        case 'subName': return row.sendLibId || '';
        case 'libType': return mapped.vendorLibType;
        case 'indexI7': return row.indexI7 || '';
        case 'indexI5': return row.indexI5 || '';
        case 'subGb': return row.targetGb || '';
        case 'mixGb': return row.targetGb || '';
        case 'productType': return mapped.defaultProductType || 'NovaSeq X plus PE150包数据量';
        case 'seqStrategy': return mapped.defaultSeqStrategy || '150+10+10+150';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });
  }

  function writeMingma(workbook, group) {
    const sheetName1 = findSheetName(workbook, ['文库样本信息']);
    const sheetName2 = findSheetName(workbook, ['文库及其子文库信息', '子文库信息']);
    const sheet1 = workbook.Sheets[sheetName1];
    const sheet2 = workbook.Sheets[sheetName2];

    const rows1 = sheetToRows(sheet1);
    const rows2 = sheetToRows(sheet2);

    const headerRowIndex1 = findHeaderRow(rows1, ['文库名称', '测序平台', '测序数据量']);
    const headerRowIndex2 = findHeaderRow(rows2, ['子文库名称', 'indexi7', '文库类型']);
    if (headerRowIndex1 === -1 || headerRowIndex2 === -1) throw new Error('模板头未匹配');

    const headerRow1 = trimHeaderRow(rows1[headerRowIndex1]);
    const headerRow2 = trimHeaderRow(rows2[headerRowIndex2]);

    const colMap1 = buildColMap(headerRow1, {
      seq: ['序号'],
      libName: ['文库名称'],
      species: ['物种'],
      tubeCount: ['管数'],
      conc: ['浓度'],
      volume: ['体积'],
      platform: ['测序平台'],
      adapter: ['接头类型'],
      phosphorylation: ['磷酸化', '5’磷酸化', '文库末端5’磷酸化'],
      cyclization: ['是否环化'],
      dataGb: ['测序数据量']
    });

    const colMap2 = buildColMap(headerRow2, {
      libName: ['文库名称'],
      subName: ['子文库名称'],
      indexI7: ['indexi7', 'index i7'],
      indexI5: ['indexi5', 'index i5'],
      libType: ['文库类型'],
      remark: ['备注']
    });
    refineMingmaIndexColumnsSheet(headerRow2, colMap2);

    writeRowsToSheet(sheet1, headerRowIndex1 + 1, colMap1, group.rows, (field, row, idx) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      const platformMap = resolveMingmaPlatform(group.platform, mapped);
      switch (field) {
        case 'seq': return idx + 1;
        case 'libName': return row.sendLibId || '';
        case 'species': return mapped.defaultSpecies || '';
        case 'tubeCount': return 1;
        case 'conc': return row.concNgUl || '';
        case 'volume': return mapped.defaultVolumeUl || '10';
        case 'platform': return platformMap.platformName;
        case 'adapter': return platformMap.adapterType;
        case 'phosphorylation': return platformMap.phosphorylation;
        case 'cyclization': return mapped.defaultCyclization || '未环化';
        case 'dataGb': return row.targetGb || '';
        default: return '';
      }
    });

    writeRowsToSheet(sheet2, headerRowIndex2 + 1, colMap2, group.rows, (field, row) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      switch (field) {
        case 'libName': return row.sendLibId || '';
        case 'subName': return row.sendLibId || '';
        case 'indexI7': return row.indexI7 || '';
        case 'indexI5': return row.indexI5 || '';
        case 'libType': return mapped.vendorLibType || '';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });
  }

  function writeMingmaAtac(workbook, group) {
    const sheetName1 = findSheetName(workbook, ['文库样本信息']);
    const sheetName2 = findSheetName(workbook, ['文库及其子文库信息', '子文库信息']);
    const sheet1 = workbook.Sheets[sheetName1];
    const sheet2 = workbook.Sheets[sheetName2];

    const rows1 = sheetToRows(sheet1);
    const rows2 = sheetToRows(sheet2);

    const headerRowIndex1 = findHeaderRow(rows1, ['文库名称', '测序平台']);
    const headerRowIndex2 = findHeaderRow(rows2, ['子文库名称', 'indexi7']);
    if (headerRowIndex1 === -1 || headerRowIndex2 === -1) throw new Error('模板头未匹配');

    const headerRow1 = trimHeaderRow(rows1[headerRowIndex1]);
    const headerRow2 = trimHeaderRow(rows2[headerRowIndex2]);

    const colMap1 = buildColMap(headerRow1, {
      seq: ['序号'],
      libName: ['文库名称'],
      species: ['物种'],
      tubeCount: ['管数'],
      conc: ['浓度'],
      volume: ['体积'],
      platform: ['测序平台'],
      adapter: ['接头类型'],
      phosphorylation: ['磷酸化', '5’磷酸化'],
      cyclization: ['是否环化'],
      dataGb: ['测序数据量'],
      remark: ['备注']
    });

    const colMap2 = buildColMap(headerRow2, {
      libName: ['文库名称'],
      subName: ['子文库名称'],
      indexI7: ['indexi7', 'index i7'],
      indexI5: ['indexi5', 'index i5'],
      libType: ['文库类型'],
      remark: ['备注']
    });
    refineMingmaIndexColumnsSheet(headerRow2, colMap2);

    writeRowsToSheet(sheet1, headerRowIndex1 + 1, colMap1, group.rows, (field, row, idx) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      const platformMap = resolveMingmaPlatform('Element', mapped);
      switch (field) {
        case 'seq': return idx + 1;
        case 'libName': return row.sendLibId || '';
        case 'species': return mapped.defaultSpecies || '';
        case 'tubeCount': return 1;
        case 'conc': return row.concNgUl || '';
        case 'volume': return mapped.defaultVolumeUl || '10';
        case 'platform': return mapped.defaultPlatformName || 'Element';
        case 'adapter': return platformMap.adapterType;
        case 'phosphorylation': return platformMap.phosphorylation;
        case 'cyclization': return mapped.defaultCyclization || '未环化';
        case 'dataGb': return row.targetGb || '';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });

    const atacRows = [];
    group.rows.forEach((row) => {
      const code = normalizeIndexCode(row.indexCode);
      const seqs = state.indexKit[code] || ['', '', '', ''];
      seqs.forEach((seq, idx) => {
        atacRows.push({
          base: row,
          subName: `${row.sendLibId || ''}-${idx + 1}`,
          indexI7: seq
        });
      });
    });

    writeRowsToSheet(sheet2, headerRowIndex2 + 1, colMap2, atacRows, (field, item) => {
      const row = item.base;
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      switch (field) {
        case 'libName': return row.sendLibId || '';
        case 'subName': return item.subName;
        case 'indexI7': return item.indexI7 || '';
        case 'indexI5': return '';
        case 'libType': return mapped.vendorLibType || '';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });
  }

  function writeGzbExcel(workbook, group) {
    const sheet = findWorksheet(workbook, ['文库信息', '文库信息单']);
    if (!sheet) throw new Error('模板头未匹配');

    const detailHeaderRowIndex = findHeaderRowInWorksheet(sheet, ['样本名称', '文库类型', '文库结构']);
    let splitHeaderRowIndex = findHeaderRowInWorksheet(sheet, ['子文库名称', 'index', 'indexf', 'indexi7']);
    if (detailHeaderRowIndex === -1 || splitHeaderRowIndex === -1) throw new Error('模板头未匹配');
    let splitTitleRowIndex = findSectionTitleRow(sheet, detailHeaderRowIndex + 1, splitHeaderRowIndex - 1, '拆分样本信息');
    if (splitTitleRowIndex) {
      const candidate = splitTitleRowIndex + 1;
      if (rowHasSplitHeader(sheet, candidate)) {
        splitHeaderRowIndex = candidate;
      }
    }

    const detailColMap = buildColMapFromWorksheet(sheet, detailHeaderRowIndex, {
      seq: ['序号'],
      sampleName: ['样本名称'],
      libType: ['文库类型'],
      libStructure: ['文库结构'],
      libProcess: ['处理情况', '文库处理情况'],
      bp: ['文库片段大小', 'bp', 'bp值'],
      conc: ['浓度', '浓度(ng/ul)', '浓度(ng/μl)'],
      volume: ['体积(ul)', '体积'],
      dataGb: ['数据量', '数据量(raw/g)'],
      remark: ['备注']
    });

    const splitColMap = buildColMapFromWorksheet(sheet, splitHeaderRowIndex, {
      seq: ['序号'],
      sampleName: ['样本名称'],
      subLibName: ['子文库名称'],
      species: ['物种'],
      specialSeq: ['有无特殊序列', '特殊序列'],
      baseBalance: ['插入片段碱基是否均衡', '碱基是否均衡', '碱基'],
      dataGb: ['数据量', '数据量(raw/g)'],
      indexCode: ['index编号', 'index编码'],
      indexI7: ['indexf', 'indexi7', 'indexf(i7)'],
      indexI5: ['indexr', 'indexi5', 'indexr(i5)'],
      remark: ['备注']
    });

    const detailStart = detailHeaderRowIndex + 1;
    let detailEndRow = (splitTitleRowIndex || splitHeaderRowIndex) - 1;
    const detailAvailable = Math.max(0, detailEndRow - detailStart + 1);
    const desiredCount = group.rows.length;
    if (desiredCount > detailAvailable) {
      const extra = desiredCount - detailAvailable;
      const templateRow = Math.max(detailEndRow, detailStart);
      sheet.duplicateRow(templateRow, extra, true);
      detailEndRow += extra;
      if (splitTitleRowIndex) splitTitleRowIndex += extra;
      splitHeaderRowIndex += extra;
    }

    const splitStart = splitHeaderRowIndex + 1;
    let splitExistingLast = getLastDataRowInWorksheet(sheet, splitStart, splitColMap);
    const splitAvailable = splitExistingLast >= splitStart ? (splitExistingLast - splitStart + 1) : 0;
    if (desiredCount > splitAvailable) {
      const extra = desiredCount - splitAvailable;
      const templateRow = splitAvailable > 0 ? splitExistingLast : splitStart;
      sheet.duplicateRow(templateRow, extra, true);
    }

    clearRowsWithTitle(sheet, detailEndRow + 1, splitHeaderRowIndex - 1, '拆分样本信息');
    const detailExistingLast = getLastDataRowInWorksheet(sheet, detailStart, detailColMap, detailEndRow);
    splitExistingLast = getLastDataRowInWorksheet(sheet, splitStart, splitColMap);
    clearWorksheetData(sheet, detailStart, detailExistingLast, detailColMap);
    const splitClearMap = { ...splitColMap };
    delete splitClearMap.species;
    delete splitClearMap.specialSeq;
    delete splitClearMap.baseBalance;
    clearWorksheetData(sheet, splitStart, splitExistingLast, splitClearMap);

    writeRowsToWorksheet(sheet, detailHeaderRowIndex + 1, detailColMap, group.rows, (field, row, idx) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      switch (field) {
        case 'seq': return idx + 1;
        case 'sampleName': return row.sendLibId || '';
        case 'libType': return mapped.vendorLibType;
        case 'libStructure': return mapped.vendorLibStructure;
        case 'libProcess': return mapped.vendorLibProcess;
        case 'bp': return row.bp || '';
        case 'conc': return row.concNgUl || '';
        case 'volume': return mapped.defaultVolumeUl;
        case 'dataGb': return row.targetGb || '';
        case 'remark': return mapped.defaultRemark;
        default: return '';
      }
    });

    writeRowsToWorksheet(sheet, splitHeaderRowIndex + 1, splitColMap, group.rows, (field, row, idx) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      switch (field) {
        case 'seq': return idx + 1;
        case 'sampleName': return row.sendLibId || '';
        case 'subLibName': return row.sendLibId || '';
        case 'species': return mapped.defaultSpecies || null;
        case 'specialSeq': return mapped.defaultSpecialSeq || null;
        case 'baseBalance': return mapped.defaultBaseBalance || null;
        case 'dataGb': return row.targetGb || '';
        case 'indexCode': return row.indexCode || '';
        case 'indexI7': return row.indexI7 || '';
        case 'indexI5': return row.indexI5 || '';
        case 'remark': return '';
        default: return '';
      }
    });

    const detailNewLast = detailStart + group.rows.length - 1;
    const splitNewLast = splitStart + group.rows.length - 1;
    const seqInfoLast = findLastRowWithKeywords(sheet, ['测序信息', '测序平台', '测序要求', '是否包lane', '是否拆分index']);
    const keepLast = Math.max(
      detailHeaderRowIndex,
      splitHeaderRowIndex,
      detailExistingLast,
      splitExistingLast,
      detailNewLast,
      splitNewLast,
      seqInfoLast || 0
    );
    trimWorksheetRows(sheet, keepLast);
    setWorksheetPrintArea(sheet, 1, keepLast, [detailColMap, splitColMap]);
  }

  function writeHplosBgiExcel(workbook, group) {
    const sheet = findWorksheet(workbook, ['文库信息']);
    if (!sheet) throw new Error('模板头未匹配');

    const headerRowIndex = findHeaderRowInWorksheet(sheet, ['子文库名称', '测序策略', 'index']);
    if (headerRowIndex === -1) throw new Error('模板头未匹配');

    const colMap = buildColMapFromWorksheet(sheet, headerRowIndex, {
      originName: ['原样本名称', '试管上的样本名称'],
      mixName: ['混文库名称'],
      subName: ['子文库名称'],
      kitVersion: ['试剂盒版本'],
      libType: ['文库类型'],
      libStructure: ['文库结构'],
      phosphorylation: ['磷酸化'],
      cyclization: ['环化'],
      seqNumber: ['序列号'],
      indexI7: ['index-i7', 'indexi7'],
      indexI5: ['index-i5', 'indexi5'],
      subGb: ['子文库数据量'],
      mixGb: ['混库总数据量'],
      productType: ['产品类型'],
      seqStrategy: ['测序策略'],
      remark: ['备注']
    });
    refineIndexColumnsInRange(sheet, headerRowIndex, colMap, [
      'originName',
      'subName',
      'libType',
      'subGb',
      'mixGb',
      'seqStrategy',
      'remark'
    ]);

    const dataStartRow = headerRowIndex + 2;
    const existingLast = getLastDataRowInWorksheet(sheet, dataStartRow, colMap);
    clearWorksheetData(sheet, dataStartRow, existingLast, colMap);

    writeRowsToWorksheet(sheet, dataStartRow, colMap, group.rows, (field, row) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      const exon = isExon(row.libType);
      switch (field) {
        case 'originName': return exon ? '' : row.sendLibId || '';
        case 'mixName': return exon ? '' : row.sendLibId || '';
        case 'subName': return row.sendLibId || '';
        case 'kitVersion': return mapped.defaultKitVersion || '';
        case 'libType': return mapped.vendorLibType;
        case 'libStructure': return mapped.vendorLibStructure || '';
        case 'phosphorylation': return mapped.defaultPhosphorylation || '未磷酸化';
        case 'cyclization': return mapped.defaultCyclization || '未环化';
        case 'seqNumber': return mapped.defaultSeqNumber || '';
        case 'indexI7': return row.indexI7 || '';
        case 'indexI5': return row.indexI5 || '';
        case 'subGb': return row.targetGb || '';
        case 'mixGb': return row.targetGb || '';
        case 'productType': return mapped.defaultProductType || '';
        case 'seqStrategy': return mapped.defaultSeqStrategy || '150+10+10+150';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });

    const newLast = dataStartRow + group.rows.length - 1;
    trimWorksheetRows(sheet, Math.max(headerRowIndex + 1, existingLast, newLast));
    const keepLast = Math.max(headerRowIndex + 1, existingLast, newLast);
    setWorksheetPrintArea(sheet, 1, keepLast, [colMap]);
  }

  function writeHplosNovaExcel(workbook, group) {
    const sheet = findWorksheet(workbook, ['客户自建库样本信息', '样本信息']);
    if (!sheet) throw new Error('模板头未匹配');

    const headerRowIndex = findHeaderRowInWorksheet(sheet, ['子文库名称', '测序策略', '产品类型']);
    if (headerRowIndex === -1) throw new Error('模板头未匹配');

    const colMap = buildColMapFromWorksheet(sheet, headerRowIndex, {
      originName: ['原样本名称', '试管上的名称'],
      mixName: ['混文库名称'],
      subName: ['子文库名称'],
      libType: ['文库类型'],
      indexI7: ['index-i7', 'indexi7'],
      indexI5: ['index-i5', 'indexi5'],
      subGb: ['子文库数据量'],
      mixGb: ['混库总数据量'],
      productType: ['产品类型'],
      seqStrategy: ['测序策略'],
      remark: ['备注']
    });
    refineIndexColumnsInRange(sheet, headerRowIndex, colMap, [
      'originName',
      'subName',
      'libType',
      'subGb',
      'mixGb',
      'seqStrategy',
      'remark'
    ]);

    const dataStartRow = headerRowIndex + 2;
    const existingLast = getLastDataRowInWorksheet(sheet, dataStartRow, colMap);
    clearWorksheetData(sheet, dataStartRow, existingLast, colMap);

    writeRowsToWorksheet(sheet, dataStartRow, colMap, group.rows, (field, row) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      const exon = isExon(row.libType);
      switch (field) {
        case 'originName': return exon ? '' : row.sendLibId || '';
        case 'mixName': return exon ? '' : row.sendLibId || '';
        case 'subName': return row.sendLibId || '';
        case 'libType': return mapped.vendorLibType;
        case 'indexI7': return row.indexI7 || '';
        case 'indexI5': return row.indexI5 || '';
        case 'subGb': return row.targetGb || '';
        case 'mixGb': return row.targetGb || '';
        case 'productType': return mapped.defaultProductType || 'NovaSeq X plus PE150包数据量';
        case 'seqStrategy': return mapped.defaultSeqStrategy || '150+10+10+150';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });

    const newLast = dataStartRow + group.rows.length - 1;
    trimWorksheetRows(sheet, Math.max(headerRowIndex + 1, existingLast, newLast));
    const keepLast = Math.max(headerRowIndex + 1, existingLast, newLast);
    setWorksheetPrintArea(sheet, 1, keepLast, [colMap]);
  }

  function writeMingmaExcel(workbook, group) {
    const sheet1 = findWorksheet(workbook, ['文库样本信息']);
    const sheet2 = findWorksheet(workbook, ['文库及其子文库信息', '子文库信息']);
    if (!sheet1 || !sheet2) throw new Error('模板头未匹配');

    const headerRowIndex1 = findHeaderRowInWorksheet(sheet1, ['文库名称', '测序平台', '测序数据量']);
    const mingmaSubHeader = findMingmaSubLibHeaderRow(sheet2);
    const headerRowIndex2 = mingmaSubHeader !== -1
      ? mingmaSubHeader
      : findHeaderRowInWorksheet(sheet2, ['子文库名称', 'indexi7', '文库类型']);
    if (headerRowIndex1 === -1 || headerRowIndex2 === -1) throw new Error('模板头未匹配');

    const colMap1 = buildColMapFromWorksheet(sheet1, headerRowIndex1, {
      seq: ['序号'],
      libName: ['文库名称'],
      species: ['物种'],
      tubeCount: ['管数'],
      conc: ['浓度'],
      volume: ['体积'],
      platform: ['测序平台'],
      adapter: ['接头类型'],
      phosphorylation: ['磷酸化', '5’磷酸化', '文库末端5’磷酸化'],
      cyclization: ['是否环化'],
      dataGb: ['测序数据量']
    });

    const colMap2 = buildColMapFromWorksheet(sheet2, headerRowIndex2, {
      libName: ['文库名称'],
      subName: ['子文库名称'],
      indexI7: ['indexi7', 'index i7'],
      indexI5: ['indexi5', 'index i5'],
      libType: ['文库类型'],
      remark: ['备注']
    });
    if (!colMap2.libName || colMap2.libName === colMap2.subName) {
      const fixed = findColumnByAliasExcluding(sheet2, headerRowIndex2, ['文库名称'], ['子文库名称']);
      if (fixed) colMap2.libName = fixed;
    }
    refineMingmaIndexColumns(sheet2, headerRowIndex2, colMap2);

    const startRow1 = headerRowIndex1 + 1;
    const startRow2 = headerRowIndex2 + 1;
    const existingLast1 = getLastDataRowInWorksheet(sheet1, startRow1, colMap1);
    const existingLast2 = getLastDataRowInWorksheet(sheet2, startRow2, colMap2);
    clearWorksheetData(sheet1, startRow1, existingLast1, colMap1);
    clearWorksheetData(sheet2, startRow2, existingLast2, colMap2);

    writeRowsToWorksheet(sheet1, headerRowIndex1 + 1, colMap1, group.rows, (field, row, idx) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      const platformMap = resolveMingmaPlatform(group.platform, mapped);
      switch (field) {
        case 'seq': return idx + 1;
        case 'libName': return row.sendLibId || '';
        case 'species': return mapped.defaultSpecies || '';
        case 'tubeCount': return 1;
        case 'conc': return row.concNgUl || '';
        case 'volume': return mapped.defaultVolumeUl || '10';
        case 'platform': return platformMap.platformName;
        case 'adapter': return platformMap.adapterType;
        case 'phosphorylation': return platformMap.phosphorylation;
        case 'cyclization': return mapped.defaultCyclization || '未环化';
        case 'dataGb': return row.targetGb || '';
        default: return '';
      }
    });

    writeRowsToWorksheet(sheet2, headerRowIndex2 + 1, colMap2, group.rows, (field, row) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      switch (field) {
        case 'libName': return row.sendLibId || '';
        case 'subName': return row.sendLibId || '';
        case 'indexI7': return row.indexI7 || '';
        case 'indexI5': return row.indexI5 || '';
        case 'libType': return mapped.vendorLibType || '';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });

    const newLast1 = startRow1 + group.rows.length - 1;
    const newLast2 = startRow2 + group.rows.length - 1;
    trimWorksheetRows(sheet1, Math.max(headerRowIndex1, existingLast1, newLast1));
    trimWorksheetRows(sheet2, Math.max(headerRowIndex2, existingLast2, newLast2));
    const keepLast1 = Math.max(headerRowIndex1, existingLast1, newLast1);
    const keepLast2 = Math.max(headerRowIndex2, existingLast2, newLast2);
    setWorksheetPrintArea(sheet1, 1, keepLast1, [colMap1]);
    setWorksheetPrintArea(sheet2, 1, keepLast2, [colMap2]);
  }

  function writeMingmaAtacExcel(workbook, group) {
    const sheet1 = findWorksheet(workbook, ['文库样本信息']);
    const sheet2 = findWorksheet(workbook, ['文库及其子文库信息', '子文库信息']);
    if (!sheet1 || !sheet2) throw new Error('模板头未匹配');

    const headerRowIndex1 = findHeaderRowInWorksheet(sheet1, ['文库名称', '测序平台']);
    const mingmaSubHeader = findMingmaSubLibHeaderRow(sheet2);
    const headerRowIndex2 = mingmaSubHeader !== -1
      ? mingmaSubHeader
      : findHeaderRowInWorksheet(sheet2, ['子文库名称', 'indexi7']);
    if (headerRowIndex1 === -1 || headerRowIndex2 === -1) throw new Error('模板头未匹配');

    const colMap1 = buildColMapFromWorksheet(sheet1, headerRowIndex1, {
      seq: ['序号'],
      libName: ['文库名称'],
      species: ['物种'],
      tubeCount: ['管数'],
      conc: ['浓度'],
      volume: ['体积'],
      platform: ['测序平台'],
      adapter: ['接头类型'],
      phosphorylation: ['磷酸化', '5’磷酸化'],
      cyclization: ['是否环化'],
      dataGb: ['测序数据量'],
      remark: ['备注']
    });

    const colMap2 = buildColMapFromWorksheet(sheet2, headerRowIndex2, {
      libName: ['文库名称'],
      subName: ['子文库名称'],
      indexI7: ['indexi7', 'index i7'],
      indexI5: ['indexi5', 'index i5'],
      libType: ['文库类型'],
      remark: ['备注']
    });
    if (!colMap2.libName || colMap2.libName === colMap2.subName) {
      const fixed = findColumnByAliasExcluding(sheet2, headerRowIndex2, ['文库名称'], ['子文库名称']);
      if (fixed) colMap2.libName = fixed;
    }
    refineMingmaIndexColumns(sheet2, headerRowIndex2, colMap2);

    const startRow1 = headerRowIndex1 + 1;
    const startRow2 = headerRowIndex2 + 1;
    const existingLast1 = getLastDataRowInWorksheet(sheet1, startRow1, colMap1);
    const existingLast2 = getLastDataRowInWorksheet(sheet2, startRow2, colMap2);
    clearWorksheetData(sheet1, startRow1, existingLast1, colMap1);
    clearWorksheetData(sheet2, startRow2, existingLast2, colMap2);

    writeRowsToWorksheet(sheet1, headerRowIndex1 + 1, colMap1, group.rows, (field, row, idx) => {
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      const platformMap = resolveMingmaPlatform('Element', mapped);
      switch (field) {
        case 'seq': return idx + 1;
        case 'libName': return row.sendLibId || '';
        case 'species': return mapped.defaultSpecies || '';
        case 'tubeCount': return 1;
        case 'conc': return row.concNgUl || '';
        case 'volume': return mapped.defaultVolumeUl || '10';
        case 'platform': return mapped.defaultPlatformName || 'Element';
        case 'adapter': return platformMap.adapterType;
        case 'phosphorylation': return platformMap.phosphorylation;
        case 'cyclization': return mapped.defaultCyclization || '未环化';
        case 'dataGb': return row.targetGb || '';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });

    const atacRows = [];
    group.rows.forEach((row) => {
      const code = normalizeIndexCode(row.indexCode);
      const seqs = state.indexKit[code] || ['', '', '', ''];
      seqs.forEach((seq, idx) => {
        atacRows.push({
          base: row,
          subName: `${row.sendLibId || ''}-${idx + 1}`,
          indexI7: seq
        });
      });
    });

    writeRowsToWorksheet(sheet2, headerRowIndex2 + 1, colMap2, atacRows, (field, item) => {
      const row = item.base;
      const mapped = buildMappedValues(row, getRule(group.vendor, group.platform, row.libType) || {}, group);
      switch (field) {
        case 'libName': return row.sendLibId || '';
        case 'subName': return item.subName;
        case 'indexI7': return item.indexI7 || '';
        case 'indexI5': return '';
        case 'libType': return mapped.vendorLibType || '';
        case 'remark': return mapped.defaultRemark || '';
        default: return '';
      }
    });

    const newLast1 = startRow1 + group.rows.length - 1;
    const newLast2 = startRow2 + atacRows.length - 1;
    trimWorksheetRows(sheet1, Math.max(headerRowIndex1, existingLast1, newLast1));
    trimWorksheetRows(sheet2, Math.max(headerRowIndex2, existingLast2, newLast2));
    const keepLast1 = Math.max(headerRowIndex1, existingLast1, newLast1);
    const keepLast2 = Math.max(headerRowIndex2, existingLast2, newLast2);
    setWorksheetPrintArea(sheet1, 1, keepLast1, [colMap1]);
    setWorksheetPrintArea(sheet2, 1, keepLast2, [colMap2]);
  }

  function findWorksheet(workbook, hints) {
    if (!workbook || !workbook.worksheets || !workbook.worksheets.length) return null;
    const sheets = workbook.worksheets;
    for (const hint of hints) {
      const normalized = normalizeHeader(hint);
      const found = sheets.find((sheet) => normalizeHeader(sheet.name).includes(normalized));
      if (found) return found;
    }
    return sheets[0] || null;
  }

  function findMingmaSubLibHeaderRow(worksheet) {
    if (!worksheet) return -1;
    const maxRow = Math.min(worksheet.rowCount || 200, 200);
    let bestRow = -1;
    let bestScore = -1;

    for (let r = 1; r <= maxRow; r += 1) {
      const row = worksheet.getRow(r);
      let hasLibName = false;
      let hasSubName = false;
      let hasIndexI7 = false;
      let hasIndexI5 = false;
      let hasLibType = false;

      row.eachCell({ includeEmpty: false }, (cell) => {
        const text = normalizeHeader(getCellText(cell));
        if (!text) return;
        if (text.startsWith('子文库名称')) {
          hasSubName = true;
          return;
        }
        if (text.startsWith('文库名称')) {
          hasLibName = true;
          return;
        }
        if (text.includes('indexi7')) hasIndexI7 = true;
        if (text.includes('indexi5')) hasIndexI5 = true;
        if (text.includes('文库类型')) hasLibType = true;
      });

      if (!hasLibName || !hasSubName) continue;
      const score = (hasLibName ? 2 : 0)
        + (hasSubName ? 2 : 0)
        + (hasIndexI7 ? 1 : 0)
        + (hasIndexI5 ? 1 : 0)
        + (hasLibType ? 1 : 0);
      if (score > bestScore) {
        bestScore = score;
        bestRow = r;
      }
    }

    return bestRow;
  }

  function findHeaderRowInWorksheet(worksheet, requiredAliases) {
    if (!worksheet) return -1;
    const normalized = requiredAliases.map((alias) => normalizeHeader(alias));
    const maxRow = Math.min(worksheet.rowCount || 200, 200);
    for (let r = 1; r <= maxRow; r += 1) {
      const texts = collectWorksheetRowTexts(worksheet, r);
      if (!texts.length) continue;
      const hits = normalized.filter((alias) => texts.some((cellText) => cellText.includes(alias)));
      if (hits.length >= Math.min(2, normalized.length)) {
        return r;
      }
    }
    return -1;
  }

  function buildColMapFromWorksheet(worksheet, rowIndex, fieldAliases) {
    const cells = collectWorksheetRowCells(worksheet, rowIndex);
    const map = {};
    Object.entries(fieldAliases).forEach(([field, aliases]) => {
      const normalizedAliases = aliases.map((alias) => normalizeHeader(alias));
      let found = null;
      for (const alias of normalizedAliases) {
        found = cells.find((cell) => cell.text === alias);
        if (found) break;
      }
      if (!found) {
        for (const alias of normalizedAliases) {
          found = cells.find((cell) => cell.text.startsWith(alias));
          if (found) break;
        }
      }
      if (!found) {
        for (const alias of normalizedAliases) {
          found = cells.find((cell) => cell.text.includes(alias));
          if (found) break;
        }
      }
      if (found) map[field] = found.col;
    });
    return map;
  }

  function findColumnByAliasExcluding(worksheet, rowIndex, aliases, excludeAliases) {
    if (!worksheet) return null;
    const cells = collectWorksheetRowCells(worksheet, rowIndex);

    const normalizedAliases = aliases.map((alias) => normalizeHeader(alias));
    const normalizedExclude = (excludeAliases || []).map((alias) => normalizeHeader(alias));
    const isExcluded = (text) => normalizedExclude.some((alias) => text.includes(alias));

    const search = (matcher) => {
      for (const alias of normalizedAliases) {
        const found = cells.find((cell) => !isExcluded(cell.text) && matcher(cell.text, alias));
        if (found) return found.col;
      }
      return null;
    };

    return (
      search((text, alias) => text === alias) ||
      search((text, alias) => text.startsWith(alias)) ||
      search((text, alias) => text.includes(alias))
    );
  }

  function collectWorksheetRowTexts(worksheet, rowIndex) {
    return collectWorksheetRowCells(worksheet, rowIndex).map((cell) => cell.text);
  }

  function collectWorksheetRowCells(worksheet, rowIndex) {
    if (!worksheet || !rowIndex) return [];
    const row = worksheet.getRow(rowIndex);
    const maxCol = Math.max(row.cellCount || 0, worksheet.columnCount || 0);
    const limit = Math.min(maxCol, 200);
    const cells = [];
    for (let col = 1; col <= limit; col += 1) {
      const cell = row.getCell(col);
      let text = getCellText(cell);
      if ((!text || !String(text).trim()) && cell && cell.isMerged && cell.master) {
        text = getCellText(cell.master);
      }
      const normalized = normalizeHeader(text);
      if (normalized) cells.push({ col, text: normalized });
    }
    return cells;
  }

  function findLastRowWithKeywords(worksheet, keywords) {
    if (!worksheet || !keywords || !keywords.length) return null;
    const normalizedKeys = keywords.map((key) => normalizeHeader(key)).filter(Boolean);
    if (!normalizedKeys.length) return null;
    const maxRow = worksheet.rowCount || 0;
    let last = null;
    for (let r = 1; r <= maxRow; r += 1) {
      const texts = collectWorksheetRowTexts(worksheet, r);
      if (!texts.length) continue;
      const hit = normalizedKeys.some((key) => texts.some((text) => text.includes(key)));
      if (hit) last = r;
    }
    return last;
  }

  function refineMingmaIndexColumns(worksheet, rowIndex, colMap) {
    if (!worksheet || !colMap) return;
    const row = worksheet.getRow(rowIndex);
    const indexI7Cols = [];
    const indexI5Cols = [];
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      const text = normalizeHeader(getCellText(cell));
      if (!text) return;
      if (text.includes('indexi7')) indexI7Cols.push(col);
      if (text.includes('indexi5')) indexI5Cols.push(col);
    });

    const rangeCols = [colMap.libName, colMap.subName, colMap.libType, colMap.remark]
      .filter((col) => typeof col === 'number');
    const minCol = rangeCols.length ? Math.min(...rangeCols) : null;
    const maxCol = rangeCols.length ? Math.max(...rangeCols) : null;

    const chooseWithinRange = (cols) => {
      if (minCol === null || maxCol === null) return null;
      return cols.find((col) => col >= minCol && col <= maxCol) || null;
    };

    const pickI7 = chooseWithinRange(indexI7Cols) || indexI7Cols[0];
    const pickI5 = chooseWithinRange(indexI5Cols) || indexI5Cols[0];

    if (pickI7) colMap.indexI7 = pickI7;
    if (pickI5) colMap.indexI5 = pickI5;

    if (colMap.indexI7 && colMap.indexI5 && colMap.indexI7 === colMap.indexI5) {
      const alt = indexI5Cols.find((col) => col !== colMap.indexI7);
      if (alt) colMap.indexI5 = alt;
    }
  }

  function refineMingmaIndexColumnsSheet(headerRow, colMap) {
    if (!headerRow || !colMap) return;
    const indexI7Cols = [];
    const indexI5Cols = [];
    headerRow.forEach((cell, idx) => {
      const text = normalizeHeader(cell);
      if (!text) return;
      if (text.includes('indexi7')) indexI7Cols.push(idx);
      if (text.includes('indexi5')) indexI5Cols.push(idx);
    });

    const rangeCols = [colMap.libName, colMap.subName, colMap.libType, colMap.remark]
      .filter((col) => typeof col === 'number');
    const minCol = rangeCols.length ? Math.min(...rangeCols) : null;
    const maxCol = rangeCols.length ? Math.max(...rangeCols) : null;

    const chooseWithinRange = (cols) => {
      if (minCol === null || maxCol === null) return null;
      return cols.find((col) => col >= minCol && col <= maxCol) || null;
    };

    const pickI7 = chooseWithinRange(indexI7Cols) || indexI7Cols[0];
    const pickI5 = chooseWithinRange(indexI5Cols) || indexI5Cols[0];

    if (typeof pickI7 === 'number') colMap.indexI7 = pickI7;
    if (typeof pickI5 === 'number') colMap.indexI5 = pickI5;

    if (colMap.indexI7 !== undefined && colMap.indexI5 !== undefined && colMap.indexI7 === colMap.indexI5) {
      const alt = indexI5Cols.find((col) => col !== colMap.indexI7);
      if (typeof alt === 'number') colMap.indexI5 = alt;
    }
  }

  function refineIndexColumnsInRange(worksheet, rowIndex, colMap, rangeFields) {
    if (!worksheet || !colMap) return;
    const row = worksheet.getRow(rowIndex);
    const indexI7Cols = [];
    const indexI5Cols = [];
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      const text = normalizeHeader(getCellText(cell));
      if (!text) return;
      if (text.includes('indexi7')) indexI7Cols.push(col);
      if (text.includes('indexi5')) indexI5Cols.push(col);
    });

    const rangeCols = (rangeFields || [])
      .map((field) => colMap[field])
      .filter((col) => typeof col === 'number' && Number.isFinite(col));
    const minCol = rangeCols.length ? Math.min(...rangeCols) : null;
    const maxCol = rangeCols.length ? Math.max(...rangeCols) : null;

    const chooseWithinRange = (cols) => {
      if (minCol === null || maxCol === null) return null;
      return cols.find((col) => col >= minCol && col <= maxCol) || null;
    };

    const pickI7 = chooseWithinRange(indexI7Cols) || indexI7Cols[0];
    const pickI5 = chooseWithinRange(indexI5Cols) || indexI5Cols[0];

    if (pickI7) colMap.indexI7 = pickI7;
    if (pickI5) colMap.indexI5 = pickI5;

    if (colMap.indexI7 && colMap.indexI5 && colMap.indexI7 === colMap.indexI5) {
      const alt = indexI5Cols.find((col) => col !== colMap.indexI7);
      if (alt) colMap.indexI5 = alt;
    }
  }

  function refineIndexColumnsInRangeSheet(headerRow, colMap, rangeFields) {
    if (!headerRow || !colMap) return;
    const indexI7Cols = [];
    const indexI5Cols = [];
    headerRow.forEach((cell, idx) => {
      const text = normalizeHeader(cell);
      if (!text) return;
      if (text.includes('indexi7')) indexI7Cols.push(idx);
      if (text.includes('indexi5')) indexI5Cols.push(idx);
    });

    const rangeCols = (rangeFields || [])
      .map((field) => colMap[field])
      .filter((col) => typeof col === 'number' && Number.isFinite(col));
    const minCol = rangeCols.length ? Math.min(...rangeCols) : null;
    const maxCol = rangeCols.length ? Math.max(...rangeCols) : null;

    const chooseWithinRange = (cols) => {
      if (minCol === null || maxCol === null) return null;
      return cols.find((col) => col >= minCol && col <= maxCol) || null;
    };

    const pickI7 = chooseWithinRange(indexI7Cols) || indexI7Cols[0];
    const pickI5 = chooseWithinRange(indexI5Cols) || indexI5Cols[0];

    if (typeof pickI7 === 'number') colMap.indexI7 = pickI7;
    if (typeof pickI5 === 'number') colMap.indexI5 = pickI5;

    if (colMap.indexI7 !== undefined && colMap.indexI5 !== undefined && colMap.indexI7 === colMap.indexI5) {
      const alt = indexI5Cols.find((col) => col !== colMap.indexI7);
      if (typeof alt === 'number') colMap.indexI5 = alt;
    }
  }

  function writeRowsToWorksheet(worksheet, startRow, colMap, rows, getter) {
    if (!worksheet || !rows || !rows.length) return;
    rows.forEach((row, idx) => {
      const excelRow = worksheet.getRow(startRow + idx);
      Object.entries(colMap).forEach(([field, col]) => {
        const value = getter(field, row, idx);
        if (value === null || value === undefined) return;
        excelRow.getCell(col).value = value;
      });
      applyRowWrapAndHeight(worksheet, excelRow, colMap);
      if (excelRow.commit) excelRow.commit();
    });
  }

  function getCellText(cell) {
    const value = cell ? cell.value : null;
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') {
      if (value.text) return value.text;
      if (value.richText) return value.richText.map((item) => item.text).join('');
      if (value.result !== undefined && value.result !== null) return String(value.result);
      if (value.hyperlink) return value.text || value.hyperlink;
      if (value.formula) return String(value.result || '');
    }
    return String(value);
  }

  function normalizeIndexCode(value) {
    return String(value || '')
      .replace(/\s+/g, '')
      .replace(/\uFEFF/g, '')
      .toUpperCase();
  }

  function normalizeIndexKitRow(row) {
    if (!row || !row.length) return [];
    const first = String(row[0] || '').trim();
    if (row.length === 1 && first.includes(',')) {
      return first.split(',').map((part) => part.trim()).filter(Boolean);
    }
    const merged = row.filter((cell) => cell !== null && cell !== undefined && String(cell).trim() !== '');
    if (merged.length === 1 && String(merged[0]).includes(',')) {
      return String(merged[0]).split(',').map((part) => part.trim()).filter(Boolean);
    }
    return row.map((cell) => String(cell || '').trim());
  }

  function clearWorksheetData(worksheet, startRow, endRow, colMap) {
    if (!worksheet || !colMap || !Object.keys(colMap).length) return;
    if (endRow < startRow) return;
    for (let r = startRow; r <= endRow; r += 1) {
      const row = worksheet.getRow(r);
      Object.values(colMap).forEach((col) => {
        row.getCell(col).value = null;
      });
      if (row.commit) row.commit();
    }
  }

  function getLastDataRowInWorksheet(worksheet, startRow, colMap, endRow) {
    if (!worksheet || !colMap || !Object.keys(colMap).length) return startRow - 1;
    const rowCount = endRow && endRow >= startRow ? endRow : (worksheet.rowCount || startRow);
    let lastRow = startRow - 1;
    for (let r = startRow; r <= rowCount; r += 1) {
      const row = worksheet.getRow(r);
      const hasValue = Object.values(colMap).some((col) => {
        const cell = row.getCell(col);
        return isMeaningfulCellValue(cell);
      });
      if (hasValue) lastRow = r;
    }
    return lastRow;
  }

  function findSectionTitleRow(worksheet, startRow, endRow, title) {
    if (!worksheet || !title) return null;
    const start = Math.max(1, startRow || 1);
    const end = Math.max(start, endRow || (worksheet.rowCount || start));
    const target = normalizeHeader(title);
    for (let r = start; r <= end; r += 1) {
      const row = worksheet.getRow(r);
      let found = false;
      row.eachCell({ includeEmpty: false }, (cell) => {
        const text = normalizeHeader(getCellText(cell));
        if (text && text.includes(target)) found = true;
      });
      if (found) return r;
    }
    return null;
  }

  function clearRowsWithTitle(worksheet, startRow, endRow, title) {
    if (!worksheet || !title) return;
    const start = Math.max(1, startRow || 1);
    const end = Math.max(start, endRow || start);
    const target = normalizeHeader(title);
    for (let r = start; r <= end; r += 1) {
      const row = worksheet.getRow(r);
      let found = false;
      let combined = '';
      row.eachCell({ includeEmpty: false }, (cell) => {
        const text = normalizeHeader(getCellText(cell));
        if (text && text.includes(target)) found = true;
        if (text) combined += text;
      });
      if (!found && combined.includes(target)) found = true;
      if (found) {
        clearRowValues(row);
      }
    }
  }

  function clearRowValues(row) {
    if (!row) return;
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.value = null;
    });
    if (row.commit) row.commit();
  }

  function rowHasSplitHeader(worksheet, rowIndex) {
    if (!worksheet || !rowIndex) return false;
    const row = worksheet.getRow(rowIndex);
    let hasSample = false;
    let hasSub = false;
    let hasIndex = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const text = normalizeHeader(getCellText(cell));
      if (!text) return;
      if (text.includes('样本名称')) hasSample = true;
      if (text.includes('子文库名称')) hasSub = true;
      if (text.includes('indexi7') || text.includes('indexf')) hasIndex = true;
    });
    return hasSample && hasSub && hasIndex;
  }

  function isMeaningfulCellValue(cell) {
    if (!cell) return false;
    const value = cell.value;
    if (value === null || value === undefined || value === '') return false;
    if (typeof value === 'object') {
      const text = getCellText(cell);
      return text.trim() !== '';
    }
    return String(value).trim() !== '';
  }

  function trimWorksheetRows(worksheet, keepLastRow) {
    if (!worksheet || !keepLastRow || keepLastRow < 1) return;
    const total = worksheet.rowCount || keepLastRow;
    if (total > keepLastRow) {
      worksheet.spliceRows(keepLastRow + 1, total - keepLastRow);
    }
  }

  function setWorksheetPrintArea(worksheet, startRow, endRow, colMaps) {
    if (!worksheet || !startRow || !endRow || startRow > endRow) return;
    const cols = [];
    (colMaps || []).forEach((map) => {
      if (!map) return;
      Object.values(map).forEach((col) => {
        if (typeof col === 'number' && Number.isFinite(col)) cols.push(col);
      });
    });
    if (!cols.length) return;
    const minCol = Math.min(...cols);
    const maxCol = Math.max(...cols);
    const startCol = columnToLetter(minCol);
    const endCol = columnToLetter(maxCol);
    worksheet.pageSetup = worksheet.pageSetup || {};
    worksheet.pageSetup.printArea = `${startCol}${startRow}:${endCol}${endRow}`;
  }

  function applyRowWrapAndHeight(worksheet, row, colMap) {
    if (!worksheet || !row || !colMap) return;
    let maxLines = 1;
    const cols = Array.from(new Set(Object.values(colMap).filter((col) => typeof col === 'number')));
    cols.forEach((col) => {
      const cell = row.getCell(col);
      const value = cell.value;
      if (typeof value !== 'string' || !value.trim()) return;
      cell.alignment = { ...(cell.alignment || {}), wrapText: true };
      const width = worksheet.getColumn(col).width || 18;
      const lines = estimateLineCount(value, width);
      if (lines > maxLines) maxLines = lines;
    });
    if (maxLines > 1) {
      const baseHeight = 18;
      const target = baseHeight * maxLines;
      row.height = Math.max(row.height || 0, target);
    }
  }

  function estimateLineCount(text, width) {
    const safeWidth = Math.max(4, Math.floor(width || 18));
    const parts = String(text).split(/\r?\n/);
    let maxLines = 1;
    parts.forEach((part) => {
      const len = getDisplayLength(part);
      const lines = Math.max(1, Math.ceil(len / safeWidth));
      if (lines > maxLines) maxLines = lines;
    });
    return maxLines;
  }

  function getDisplayLength(text) {
    let len = 0;
    for (const ch of String(text)) {
      len += /[^\x00-\xff]/.test(ch) ? 2 : 1;
    }
    return len;
  }

  function columnToLetter(col) {
    let num = col;
    let letters = '';
    while (num > 0) {
      const mod = (num - 1) % 26;
      letters = String.fromCharCode(65 + mod) + letters;
      num = Math.floor((num - mod) / 26);
    }
    return letters;
  }

  function resolveMingmaPlatform(platform, mapped) {
    if (mapped.defaultPlatformName) {
      return {
        platformName: mapped.defaultPlatformName,
        adapterType: mapped.defaultAdapterType || 'illumina文库',
        phosphorylation: mapped.defaultPhosphorylation || '5’未磷酸化'
      };
    }
    const key = String(platform || '').toLowerCase();
    if (key.includes('bgi') || key.includes('t7')) {
      return { platformName: 'DNBSEQ-T7', adapterType: 'illumina文库', phosphorylation: '5’未磷酸化' };
    }
    if (key.includes('nova')) {
      return { platformName: 'NovaSeq', adapterType: 'illumina文库', phosphorylation: '5’未磷酸化' };
    }
    if (key.includes('element') || key.includes('atac')) {
      return { platformName: 'Element', adapterType: 'Element', phosphorylation: '5’未磷酸化' };
    }
    return { platformName: platform, adapterType: mapped.defaultAdapterType || '', phosphorylation: mapped.defaultPhosphorylation || '5’未磷酸化' };
  }

  function trimHeaderRow(row) {
    const lastIndex = findLastNonEmptyIndex(row || []);
    if (lastIndex === -1) return [];
    return row.slice(0, lastIndex + 1);
  }

  function findLastNonEmptyIndex(row) {
    for (let i = row.length - 1; i >= 0; i -= 1) {
      const cell = row[i];
      if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
        return i;
      }
    }
    return -1;
  }

  function writeRowsToSheet(sheet, startRow, colMap, rows, getter) {
    if (!sheet || !rows || !rows.length) return;
    const range = sheet['!ref']
      ? XLSX.utils.decode_range(sheet['!ref'])
      : { s: { r: startRow, c: 0 }, e: { r: startRow, c: 0 } };

    rows.forEach((row, idx) => {
      const r = startRow + idx;
      Object.entries(colMap).forEach(([field, colIndex]) => {
        const value = getter(field, row, idx);
        if (isEmptyValue(value)) return;
        setCellValue(sheet, r, colIndex, value);
        updateRange(range, r, colIndex);
      });
    });

    sheet['!ref'] = XLSX.utils.encode_range(range);
  }

  function setCellValue(sheet, r, c, value) {
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = sheet[addr] || {};
    cell.v = value;
    if (typeof value === 'number') {
      cell.t = 'n';
    } else if (typeof value === 'boolean') {
      cell.t = 'b';
    } else {
      cell.t = 's';
    }
    sheet[addr] = cell;
  }

  function updateRange(range, r, c) {
    if (r < range.s.r) range.s.r = r;
    if (c < range.s.c) range.s.c = c;
    if (r > range.e.r) range.e.r = r;
    if (c > range.e.c) range.e.c = c;
  }

  function isEmptyValue(value) {
    return value === null || value === undefined;
  }

  function buildColMap(headerRow, fieldAliases) {
    const map = {};
    Object.entries(fieldAliases).forEach(([field, aliases]) => {
      const idx = findColIndex(headerRow, aliases);
      if (idx !== -1) map[field] = idx;
    });
    return map;
  }

  function findColIndex(headerRow, aliases) {
    const normalizedAliases = aliases.map((alias) => normalizeHeader(alias));
    const cells = headerRow.map((cell, idx) => ({ idx, text: normalizeHeader(cell) }));
    for (const alias of normalizedAliases) {
      const exact = cells.find((cell) => cell.text === alias);
      if (exact) return exact.idx;
      const starts = cells.find((cell) => cell.text.startsWith(alias));
      if (starts) return starts.idx;
    }
    for (const alias of normalizedAliases) {
      const contains = cells.find((cell) => cell.text.includes(alias));
      if (contains) return contains.idx;
    }
    return -1;
  }

  function sheetToRows(sheet) {
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  }

  function findSheetName(workbook, hints) {
    const names = workbook.SheetNames;
    for (const hint of hints) {
      const normalized = normalizeHeader(hint);
      const found = names.find((name) => normalizeHeader(name).includes(normalized));
      if (found) return found;
    }
    return names[0];
  }

  function findHeaderRow(rows, requiredAliases) {
    const normalized = requiredAliases.map((alias) => normalizeHeader(alias));
    for (let r = 0; r < rows.length; r += 1) {
      const row = rows[r];
      if (!row) continue;
      const rowNormalized = row.map((cell) => normalizeHeader(cell));
      const hits = normalized.filter((alias) => rowNormalized.some((cell) => cell.includes(alias)));
      if (hits.length >= Math.min(2, normalized.length)) {
        return r;
      }
    }
    return -1;
  }

  function isExon(libType) {
    return String(libType || '').toLowerCase().includes('exon');
  }
})();
