// リネーム後: app.js (旧: app_jtccm.js)
// 仕様はJTCCM準拠だがファイル名から jtccm 文字列を除去

(function(){
  'use strict';

  // === State ===
  let rawData = [];
  let header = [];
  // envelopeData: 現在表示・編集中の包絡線（表示用に間引き済みの場合あり）
  let envelopeData = [];
  // インポートや再計算前のフル包絡線保持（再間引きで点数を増やす際に利用）
  let originalEnvelopeBySide = {}; // side -> full envelope array
  let analysisResults = {};
  let relayoutHandlerAttached = false; // Plotly autoscale対策のイベント重複防止
  // イベントハンドラ参照（重複登録防止用）
  let _plotClickHandler = null;
  let _keydownHandler = null;
  // 包絡線フィット範囲キャッシュ
  let cachedEnvelopeRange = null;

  // 編集済み包絡線の保持（設定変更しても点情報を維持）
  const editedEnvelopeBySide = { positive: null, negative: null };
  const editedDirtyBySide = { positive: false, negative: false };
  function getCurrentSide(){
    return (envelope_side && envelope_side.value === 'negative') ? 'negative' : 'positive';
  }
  function setEditedEnvelopeForCurrentSide(env){
    const side = getCurrentSide();
    editedEnvelopeBySide[side] = Array.isArray(env) ? env.map(p=>({gamma:p.gamma, Load:p.Load, gamma0:p.gamma0})) : null;
    editedDirtyBySide[side] = !!(editedEnvelopeBySide[side] && editedEnvelopeBySide[side].length >= 2);
  }

  // 目標点数に極力合わせるシンプルな均等弧長サンプリング
  function sampleEnvelopeExact(envelope, targetPoints, mandatoryGammas){
    try{
      if(!Array.isArray(envelope) || envelope.length <= targetPoints) return envelope.map(pt=>({...pt}));
      const pts = envelope.map(pt => ({x: pt.gamma, y: pt.Load}));
      // 弧長配列
      const arc = [0];
      for(let i=1;i<pts.length;i++){
        const dx = pts[i].x - pts[i-1].x;
        const dy = pts[i].y - pts[i-1].y;
        arc.push(arc[i-1] + Math.hypot(dx, dy));
      }
      const total = arc[arc.length-1];
      // 必須点
      const mandatory = new Set();
      mandatory.add(0);
      mandatory.add(pts.length-1);
      // 最大荷重点
      let pmaxIdx = 0, pmaxAbs=-Infinity;
      for(let i=0;i<pts.length;i++){ const a = Math.abs(pts[i].y); if(a>pmaxAbs){ pmaxAbs=a; pmaxIdx=i; } }
      mandatory.add(pmaxIdx);
      if(Array.isArray(mandatoryGammas)){
        mandatoryGammas.forEach(gTarget => {
          if(!Number.isFinite(gTarget)) return;
          let best=-1, bestDiff=Infinity;
          for(let i=0;i<pts.length;i++){
            const d = Math.abs(pts[i].x - gTarget);
            if(d < bestDiff){ bestDiff=d; best=i; }
          }
          if(best>=0) mandatory.add(best);
        });
      }
      // ループピーク: 局所最大（絶対値）も追加（ただし過剰なら後で調整）
      for(let i=1;i<pts.length-1;i++){
        const y0=Math.abs(pts[i-1].y), y1=Math.abs(pts[i].y), y2=Math.abs(pts[i+1].y);
        if(y1>=y0 && y1>=y2) mandatory.add(i);
      }
      if(mandatory.size >= targetPoints){
        return Array.from(mandatory).sort((a,b)=>a-b).slice(0,targetPoints).map(i=>({...envelope[i]}));
      }
      const need = targetPoints - mandatory.size;
      const selected = new Set(mandatory);
      // 均等弧長ステップ
      const step = total / (need + 1);
      for(let j=1;j<=need;j++){
        const tArc = j * step;
        let best=-1, bestDiff=Infinity;
        for(let i=0;i<pts.length;i++){
          if(selected.has(i)) continue;
          const d = Math.abs(arc[i] - tArc);
          if(d < bestDiff){ bestDiff=d; best=i; }
        }
        if(best>=0) selected.add(best);
      }
      return Array.from(selected).sort((a,b)=>a-b).map(i=>({...envelope[i]}));
    }catch(err){ console.warn('sampleEnvelopeExact エラー', err); return envelope; }
  }

  // === Elements ===
  const gammaInput = document.getElementById('gammaInput');
  const loadInput = document.getElementById('loadInput');
  const pasteGammaButton = document.getElementById('pasteGammaButton');
  const pasteLoadButton = document.getElementById('pasteLoadButton');
  const alpha_factor = document.getElementById('alpha_factor');
  const max_ultimate_deformation = document.getElementById('max_ultimate_deformation');
  const envelope_side = document.getElementById('envelope_side');
  const specimen_name = document.getElementById('specimen_name');
  const show_annotations = document.getElementById('show_annotations');
  const envelope_thinning_rate = document.getElementById('envelope_thinning_rate');
  const thinning_rate_value = document.getElementById('thinning_rate_value');
  // 手動解析ボタンは廃止
  const processButton = null;
  const downloadExcelButton = document.getElementById('downloadExcelButton');
  const generatePdfButton = document.getElementById('generatePdfButton');
  const clearDataButton = document.getElementById('clearDataButton');
  
  const plotDiv = document.getElementById('plot');
  const pointTooltip = document.getElementById('pointTooltip');
  const undoButton = document.getElementById('undoButton');
  const redoButton = document.getElementById('redoButton');
  const openPointEditButton = null; // ボタンは廃止
  const pointEditDialog = document.getElementById('pointEditDialog');
  const editGammaInput = document.getElementById('edit_gamma');
  const editLoadInput = document.getElementById('edit_load');
  const applyPointEditButton = document.getElementById('applyPointEdit');
  const cancelPointEditButton = document.getElementById('cancelPointEdit');
  const toggleDragMoveButton = document.getElementById('toggleDragMove');
  const toggleRangeSelectButton = document.getElementById('toggleRangeSelect');
  const selectPrevPointButton = document.getElementById('selectPrevPoint');
  const selectNextPointButton = document.getElementById('selectNextPoint');
  const importExcelButton = document.getElementById('importExcelButton');
  const importExcelInput = document.getElementById('importExcelInput');
  // 表示間引き（包絡線点数）
  const thin_target_points = document.getElementById('thin_target_points');
  const applyThinningButton = document.getElementById('applyThinningButton');

  // ドラッグ移動モード（ボタンONの間のみ点ドラッグ可能。パン/ズームを抑止）
  let dragMoveEnabled = false;
  // 範囲選択モード（ONの間はBox/Lassoを有効化し、削除ボタンで複数削除）
  let rangeSelectEnabled = false;
  // 複数選択インデックスのグローバル共有（削除ボタンから参照するため）
  if(typeof window !== 'undefined' && typeof window._selectedEnvelopePoints === 'undefined'){
    window._selectedEnvelopePoints = [];
  }

  // 履歴管理 (Undo/Redo)
  let historyStack = [];
  let redoStack = [];
  const MAX_HISTORY = 100;

  function cloneEnvelope(env){
    return env.map(pt => ({gamma: pt.gamma, Load: pt.Load}));
  }

  // === Utilities ===
  // 角度[rad]を 1/N 表記の文字列へ変換（Nは四捨五入した整数）
  function formatReciprocal(rad){
    const v = Number(rad);
    if(!isFinite(v) || v <= 0) return '-';
    const denom = Math.round(1 / v);
    if(!isFinite(denom) || denom <= 0) return '-';
    return '1/' + denom.toLocaleString('ja-JP');
  }

  // 対象接合部プリセット機能は廃止

  // 自動解析スケジューラ（タイプ中の過剰実行を防止）
  let _autoRunTimer = null;
  function scheduleAutoRun(delay=150){
    if(_autoRunTimer) clearTimeout(_autoRunTimer);
    _autoRunTimer = setTimeout(() => {
      try{ processDataDirect(); }catch(e){ console.warn('auto-run error', e); }
    }, delay);
  }

  function pushHistory(current){
    if(!current) return;
    historyStack.push(cloneEnvelope(current));
    if(historyStack.length > MAX_HISTORY){ historyStack.shift(); }
    redoStack = [];
    updateHistoryButtons();
  }

  function updateHistoryButtons(){
    if(undoButton) undoButton.disabled = historyStack.length <= 1;
    if(redoButton) redoButton.disabled = redoStack.length === 0;
  }

  function performUndo(){
    if(historyStack.length <= 1) return;
    const current = historyStack.pop();
    redoStack.push(current);
    const prev = cloneEnvelope(historyStack[historyStack.length - 1]);
    envelopeData = prev;
  appendLog('Undo: 包絡線を前状態へ戻しました');
  recalculateFromEnvelope(envelopeData);
    window._selectedEnvelopePoint = -1;
    updateHistoryButtons();
  }

  function performRedo(){
    if(redoStack.length === 0) return;
    const next = redoStack.pop();
    historyStack.push(cloneEnvelope(next));
    envelopeData = cloneEnvelope(next);
  appendLog('Redo: 包絡線編集をやり直しました');
  recalculateFromEnvelope(envelopeData);
    window._selectedEnvelopePoint = -1;
    updateHistoryButtons();
  }

  function openPointEditDialog(){
    if(window._selectedEnvelopePoint < 0 || !envelopeData) return;
    const idx = window._selectedEnvelopePoint;
    const pt = envelopeData[idx];
    console.debug('[openPointEditDialog] 開始 idx='+idx+' γ='+pt.gamma+' P='+pt.Load);

    const originalGamma = pt.gamma;
    const originalLoad = pt.Load;
    pointEditDialog.dataset.originalGamma = originalGamma.toString();
    pointEditDialog.dataset.originalLoad = originalLoad.toString();

    editGammaInput.value = pt.gamma.toFixed(4);
    editLoadInput.value = pt.Load.toFixed(1);

    // 右端固定表示（CSS custom-positionが制御）
    pointEditDialog.classList.add('custom-position');
    pointEditDialog.style.display = 'flex';

    // ドラッグ移動モードボタン初期化
    if(toggleDragMoveButton){
      const updateBtnUI = () => {
        toggleDragMoveButton.textContent = dragMoveEnabled ? 'マウスドラッグ移動OFF' : 'マウスドラッグ移動ON';
        if(dragMoveEnabled){
          toggleDragMoveButton.classList.add('active-bright');
        } else {
          toggleDragMoveButton.classList.remove('active-bright');
        }
      };
      updateBtnUI();
      toggleDragMoveButton.onclick = function(){
        dragMoveEnabled = !dragMoveEnabled;
        if(dragMoveEnabled){
          // Drag ON にしたら範囲選択はOFFへ
          rangeSelectEnabled = false; updateRangeSelectUI(false);
          // ドラッグ移動はパン/ズーム抑止
        }
        updateBtnUI();
        // カーソルはホバー時の条件でのみ変更（常時 hand 表示しない）
        if(plotDiv){ plotDiv.style.cursor = 'default'; }
      };
    }

    // 範囲選択ONボタン初期化
    function updateRangeSelectUI(force){
      if(!toggleRangeSelectButton) return;
      const on = (typeof force === 'boolean') ? force : rangeSelectEnabled;
      toggleRangeSelectButton.textContent = on ? '範囲選択OFF' : '範囲選択ON';
      toggleRangeSelectButton.style.background = on ? '#e0f0ff' : '';
    }
    if(toggleRangeSelectButton){
      updateRangeSelectUI();
      toggleRangeSelectButton.onclick = function(){
        rangeSelectEnabled = !rangeSelectEnabled;
        if(rangeSelectEnabled){
          // 範囲選択を有効化したらドラッグ移動はOFFへ
          dragMoveEnabled = false; if(toggleDragMoveButton){ toggleDragMoveButton.textContent='マウスドラッグ移動ON'; toggleDragMoveButton.classList.remove('active-bright'); }
          // Box選択モードへ
          try{ if(window.Plotly && plotDiv){ safeRelayout(plotDiv, {'dragmode':'select'}); } }catch(_){/* noop */}
        } else {
          // 通常のパンへ戻す
          try{ if(window.Plotly && plotDiv){ safeRelayout(plotDiv, {'dragmode':'pan'}); } }catch(_){/* noop */}
        }
        updateRangeSelectUI();
      };
    }

    // 編集中リアルタイム反映
    editGammaInput.oninput = function(){
      const v = parseFloat(editGammaInput.value);
      if(!isNaN(v)){
        envelopeData[idx].gamma = v;
        renderPlot(envelopeData, analysisResults);
      }
    };
    editLoadInput.oninput = function(){
      const v = parseFloat(editLoadInput.value);
      if(!isNaN(v)){
        envelopeData[idx].Load = v;
        renderPlot(envelopeData, analysisResults);
      }
    };

    cancelPointEditButton.onclick = function(){
      if(idx >= 0 && envelopeData && envelopeData[idx]){
        envelopeData[idx].gamma = originalGamma;
        envelopeData[idx].Load = originalLoad;
        renderPlot(envelopeData, analysisResults);
      }
      closePointEditDialog();
    };
  }

  function closePointEditDialog(){ 
    pointEditDialog.style.display = 'none'; 
    pointEditDialog.classList.remove('custom-position');
    const content = document.getElementById('pointEditContent');
    if(content){
      content.style.position = '';
      content.style.left = '';
      content.style.top = '';
      content.style.margin = '';
      content.style.transform = '';
    }
    console.debug('[ダイアログ] 閉じました');
    // ダイアログを閉じたらドラッグ移動モードを自動OFF
    dragMoveEnabled = false;
  if(toggleDragMoveButton){ toggleDragMoveButton.textContent = 'マウスドラッグ移動ON'; toggleDragMoveButton.style.filter=''; toggleDragMoveButton.classList.remove('active-bright'); }
  if(plotDiv){
    plotDiv.style.cursor = 'default';
    plotDiv.classList.remove('drag-point-active');
    plotDiv.classList.remove('drag-point-hover');
    plotDiv.classList.remove('pointer-mode');
  }
    // 範囲選択も自動OFFしてパンへ戻す
    rangeSelectEnabled = false;
    if(toggleRangeSelectButton){ toggleRangeSelectButton.textContent='範囲選択ON'; toggleRangeSelectButton.style.background=''; }
    try{ if(window.Plotly && plotDiv){ safeRelayout(plotDiv, {'dragmode':'pan'}); } }catch(_){/* noop */}
  }

  function applyPointEdit(){
    console.debug('[適用ボタン] クリックされました');
    if(window._selectedEnvelopePoint < 0 || !envelopeData) {
      console.warn('[適用ボタン] 選択点またはデータが無効です');
      return;
    }
    const g = parseFloat(editGammaInput.value);
    const l = parseFloat(editLoadInput.value);
    if(isNaN(g) || isNaN(l)){ 
      alert('数値が不正です'); 
      console.warn('[適用ボタン] 数値が不正です: γ='+g+', P='+l);
      return; 
    }
    envelopeData[window._selectedEnvelopePoint].gamma = g;
    envelopeData[window._selectedEnvelopePoint].Load = l;
    appendLog('点を数値編集しました (γ='+g+', P='+l+')');
    pushHistory(envelopeData);
    console.debug('[適用ボタン] closePointEditDialog()を呼び出します');
    closePointEditDialog();
    // 先に再計算して範囲を再評価
    recalculateFromEnvelope(envelopeData);
    // 描画後、requestAnimationFrameで包絡線範囲を適用して全体フィット化を阻止
    requestAnimationFrame(()=>{
      if(cachedEnvelopeRange){
        safeRelayout(plotDiv, {
          'xaxis.autorange': false,
          'yaxis.autorange': false,
          'xaxis.range': cachedEnvelopeRange.xRange,
          'yaxis.range': cachedEnvelopeRange.yRange
        });
      }else{
        fitEnvelopeRanges('点編集後キャッシュ無し');
      }
    });
  }
  
  // キャンセルボタンのハンドラは openPointEditDialog 内で動的に設定されるため、ここでは不要
  
  // 削除ボタン
  const deletePointEditButton = document.getElementById('deletePointEdit');
  if(deletePointEditButton){
    deletePointEditButton.onclick = function(){
      // 優先: 範囲選択で複数点が選ばれている場合は一括削除
      const sel = Array.isArray(window._selectedEnvelopePoints) ? window._selectedEnvelopePoints.slice() : [];
      if(sel.length > 0 && Array.isArray(envelopeData)){
        // 残点が2点未満にならないように保護
        const remaining = envelopeData.length - sel.length;
        if(remaining < 2){
          alert('包絡線には最低2点が必要です。削除できません。');
          return;
        }
        // 変更前履歴
        pushHistory(envelopeData);
        // 降順で安全に削除
        sel.sort((a,b)=>b-a).forEach(idx => {
          if(idx >= 0 && idx < envelopeData.length){ envelopeData.splice(idx,1); }
        });
        // 状態クリア
        window._selectedEnvelopePoints = [];
        window._selectedEnvelopePoint = -1;
        // 再計算・再描画
        recalculateFromEnvelope(envelopeData);
        appendLog(`包絡線点 ${sel.length}個 をボタンから一括削除しました`);
        closePointEditDialog();
        return;
      }
      // 単一点選択の通常削除
      if(window._selectedEnvelopePoint >= 0 && Array.isArray(envelopeData)){
        deleteEnvelopePoint(window._selectedEnvelopePoint, envelopeData);
        window._selectedEnvelopePoint = -1;
        closePointEditDialog();
      }
    };
  }
  // 追加ボタン
  const addPointEditButton = document.getElementById('addPointEdit');
  if(addPointEditButton){
    addPointEditButton.onclick = function(){
      if(window._selectedEnvelopePoint >= 0 && envelopeData){
        const idx = window._selectedEnvelopePoint;
        if(idx >= envelopeData.length - 1){
          alert('最後の点の次には追加できません。');
          return;
        }
        // 選択点と次の点の中間値を計算
        const pt1 = envelopeData[idx];
        const pt2 = envelopeData[idx + 1];
        let midGamma = (pt1.gamma + pt2.gamma) / 2;
        let midLoad = (pt1.Load + pt2.Load) / 2;
        
        // 実験データ折れ線への吸着処理（側面判定: 荷重符号で判断）
        if(rawData && rawData.length > 1){
          const side = midLoad >= 0 ? 'positive' : 'negative';
          const snapped = snapToNearestRawDataSegment(midGamma, midLoad, rawData, side);
          if(snapped){
            midGamma = snapped.gamma;
            midLoad = snapped.Load;
          }
        }
        
        // 履歴に保存
        pushHistory(envelopeData);
        
        // 新しい点を挿入
        envelopeData.splice(idx + 1, 0, {
          gamma: midGamma,
          Load: midLoad,
          gamma0: midGamma
        });
        
        appendLog('包絡線点を追加しました（γ=' + midGamma.toFixed(6) + ', P=' + midLoad.toFixed(3) + '）');
        renderPlot(envelopeData, analysisResults);
        recalculateFromEnvelope(envelopeData);
        
        // 新しく追加した点を選択
        window._selectedEnvelopePoint = idx + 1;
        
        // ダイアログを閉じて新しい点のダイアログを開く
        closePointEditDialog();
        setTimeout(function(){
          openPointEditDialog();
        }, 100);
      }
    };
  }
  // ダイアログドラッグ移動
  (function enableDialogDrag(){
    const content = document.getElementById('pointEditContent');
    const handle = content ? content.querySelector('.drag-handle') : null;
    if(!content || !handle) return;
    let dragging = false; let startX=0, startY=0, origLeft=0, origTop=0;
    handle.addEventListener('mousedown', function(e){
      dragging = true; startX = e.clientX; startY = e.clientY;
      const rect = content.getBoundingClientRect();
      origLeft = rect.left; origTop = rect.top;
      // ドラッグ開始時にabsolute配置に切り替え
      content.style.position = 'absolute';
      content.style.left = origLeft + 'px';
      content.style.top = origTop + 'px';
      content.style.margin = '0';
      content.style.transform = '';
      document.body.style.userSelect='none';
      e.preventDefault(); // デフォルト動作を抑制
    });
    window.addEventListener('mousemove', function(e){
      if(!dragging) return;
      const dx = e.clientX - startX; const dy = e.clientY - startY;
      content.style.left = (origLeft + dx) + 'px';
      content.style.top = (origTop + dy) + 'px';
    });
    window.addEventListener('mouseup', function(){
      dragging = false; document.body.style.userSelect='';
    });
  })();

  // ローカル(file://)でのCORS制約回避用: 組込サンプルCSV（fetch失敗時のフォールバック）
  const BUILTIN_SAMPLE_CSV = `gamma,Load\n
2.28743E-05,0
8.52363E-05,0.42
0.000109903,0.79
0.000129205,1.23
0.000204985,1.46
0.000220261,1.91
0.000272465,2.34
0.000346011,2.77
0.00039815,3.26
0.000458278,3.76
0.000584535,4.43
0.000564155,4.61
0.000711558,5.19
0.000761398,5.77
0.000871873,6.55
0.000961902,7.23
0.00105653,7.69
0.001138636,8.32
0.001247383,8.91
0.001379134,9.76
0.001195178,7.47
0.000964071,5.79
0.000792001,4.24
0.00059316,2.62
0.000385629,1.48
0.000206401,0.45
0.000163458,0
9.06139E-05,-0.14
0.000110605,-0.26
9.68618E-05,-0.39
6.81682E-05,-0.65
5.90237E-05,-0.94
3.57597E-05,-1.27
3.52402E-05,-1.69
-4.86582E-05,-2.17
-0.00011441,-2.73
-0.000192749,-3.36
-0.000284246,-3.79
-0.000332424,-4.16
-0.000347505,-4.27
-0.00041466,-4.73
-0.000511015,-5.35
-0.000574208,-5.81
-0.000660341,-6.39
-0.000715611,-6.79
-0.000777271,-7.31
-0.000868833,-7.67
-0.000896631,-8.13
-0.000980399,-8.69
-0.000882121,-6.85
-0.000740459,-5.25
-0.000600525,-3.77
-0.000398293,-2.34
-0.000252098,-1.47
-0.000246525,-0.73
-0.000139493,-0.17
-0.000166706,0
-0.000100187,0.12
-7.86507E-05,0.22
-7.66633E-05,0.37
-7.2182E-05,0.54
2.81999E-05,0.77
3.56169E-05,1.1
0.000135102,1.64
0.000199893,2.09
0.000252864,2.51
0.000335671,2.94
0.000452537,3.81
0.000494324,3.98
0.000555919,4.56
0.000683644,5.28
0.000828683,6.01
0.00089111,6.8
0.001007014,7.47
0.001136336,8.29
0.00126962,9.2
0.001365014,9.83
0.001170966,7.73
0.00100126,6.24`;

  // === Events ===
  gammaInput.addEventListener('input', handleDirectInput);
  loadInput.addEventListener('input', handleDirectInput);
  async function pasteFromClipboard(target){
    try{
      const text = await navigator.clipboard.readText();
      if(!text){ alert('クリップボードにテキストがありません'); return; }
      const raw = text.trim();
      const lines = raw.split(/\r?\n/).filter(l => l.trim().length>0);
      // 区切り検出: タブ優先、なければカンマ
      const hasTab = lines.some(l => l.includes('\t'));
      const hasComma = !hasTab && lines.some(l => l.includes(','));
      const sep = hasTab ? '\t' : (hasComma ? ',' : null);
      // 数値判定
      const isNumericString = (s) => /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(s.trim());
      let filledBoth = false;
      if(sep){
        const gCol = [];
        const lCol = [];
        lines.forEach(l => {
          const parts = l.split(new RegExp(sep)).map(s=>s.trim());
          if(parts.length >= 2 && isNumericString(parts[0]) && isNumericString(parts[1])){
            gCol.push(parts[0]);
            lCol.push(parts[1]);
          }
        });
        if(gCol.length > 0 && lCol.length === gCol.length){
          gammaInput.value = gCol.join('\n');
          loadInput.value = lCol.join('\n');
          filledBoth = true;
        }
      }
      if(!filledBoth){
        // 単一列: 対象指定 or both
        if(target === 'gamma'){
          gammaInput.value = lines.join('\n');
        } else if(target === 'load') {
          loadInput.value = lines.join('\n');
        } else if(target === 'both') {
          gammaInput.value = lines.join('\n');
          loadInput.value  = lines.join('\n');
        }
      }
      // 解析試行（どちらか空なら内部で早期return）
      handleDirectInput();
    }catch(err){
      console.warn('クリップボード読み取り失敗', err);
      alert('クリップボードの読み取りに失敗しました（ブラウザ許可を確認）');
    }
  }
  if(pasteGammaButton){ pasteGammaButton.addEventListener('click', () => pasteFromClipboard('gamma')); }
  if(pasteLoadButton){ pasteLoadButton.addEventListener('click', () => pasteFromClipboard('load')); }
  const pasteBothButton = document.getElementById('pasteBothButton');
  if(pasteBothButton){ pasteBothButton.addEventListener('click', () => pasteFromClipboard('both')); }
  // 手動ボタン削除済み: processButton クリックイベント不要

  // パラメータ変更時の自動解析
  // 設定変更時も編集済み包絡線を維持する
  function recalcOrProcess(reason){
    try{
      if(!rawData || rawData.length < 3){ return; }
      const side = getCurrentSide();
      const edited = editedEnvelopeBySide[side];
      if(Array.isArray(edited) && edited.length >= 2){
        // 既存の編集済み包絡線から再計算のみ
        envelopeData = edited.map(p=>({ ...p }));
        recalculateFromEnvelope(envelopeData);
      }else{
        processDataDirect();
      }
    }catch(err){ console.warn('recalcOrProcess error', err); }
  }
  const autoInputs = [alpha_factor, max_ultimate_deformation];
  autoInputs.forEach(el => {
    if(!el) return;
    el.addEventListener('input', () => recalcOrProcess('input-change'));
    el.addEventListener('change', () => recalcOrProcess('input-change'));
  });
  if(envelope_side){
    envelope_side.addEventListener('change', () => {
      if(!rawData || rawData.length<3){ return; }
      const side = getCurrentSide();
      const edited = editedEnvelopeBySide[side];
      if(Array.isArray(edited) && edited.length>=2){
        envelopeData = edited.map(p=>({...p}));
        recalculateFromEnvelope(envelopeData);
      } else {
        processDataDirect();
      }
    });
  }
  if(show_annotations){
    show_annotations.addEventListener('change', () => {
      if(envelopeData && analysisResults && Object.keys(analysisResults).length){
        renderPlot(envelopeData, analysisResults);
      }
    });
  }
  // 間引き率スライダーのイベントリスナー
  if(envelope_thinning_rate && thinning_rate_value){
    const updateThinningLabel = () => {
      const rate = parseInt(envelope_thinning_rate.value);
      let label = '';
      if(rate <= 10) label = '低（間引き無し）';
      else if(rate <= 40) label = '低〜中';
      else if(rate <= 60) label = '中（デフォルト）';
      else if(rate <= 80) label = '中〜高';
      else label = '高（最大間引き）';
      thinning_rate_value.textContent = label;
    };
    updateThinningLabel();
    envelope_thinning_rate.addEventListener('input', () => {
      updateThinningLabel();
      // 通常は rawData があるなら再解析をスケジュール
      if(rawData && rawData.length>=3){
        scheduleAutoRun(300); // 少し長めの待ち時間で再解析
        return;
      }
      // rawData が無い場合（Excelインポートで包絡線のみ取り込んだ等）は
      // originalEnvelopeBySide キャッシュがあればそれを使って表示用間引きを即時再適用する
      try{
        const side = getCurrentSide();
        const full = originalEnvelopeBySide[side];
        if(Array.isArray(full) && full.length > 2){
          const metrics = calculateJTCCMMetrics(full, parseFloat(max_ultimate_deformation.value), parseFloat(alpha_factor.value));
          envelopeData = reapplyDisplayThinning(full, side, metrics);
          renderPlot(envelopeData, metrics);
          renderResults(metrics);
          appendLog('表示間引きを再適用（キャッシュ包絡線）: ' + envelopeData.length + ' 点');
        }
      }catch(err){ console.warn('thinning slider handler error', err); }
    });
  }
  if(downloadExcelButton) downloadExcelButton.addEventListener('click', downloadExcel);
  if(generatePdfButton) generatePdfButton.addEventListener('click', generatePdfReport);
  clearDataButton.addEventListener('click', clearInputData);
  
  if(undoButton) undoButton.addEventListener('click', performUndo);
  if(redoButton) redoButton.addEventListener('click', performRedo);
  if(openPointEditButton) openPointEditButton.addEventListener('click', openPointEditDialog);
  if(applyPointEditButton) applyPointEditButton.addEventListener('click', applyPointEdit);
  if(cancelPointEditButton) cancelPointEditButton.addEventListener('click', closePointEditDialog);
  if(selectPrevPointButton) selectPrevPointButton.addEventListener('click', () => moveSelectedPoint(-1));
  if(selectNextPointButton) selectNextPointButton.addEventListener('click', () => moveSelectedPoint(1));
  if(importExcelButton && importExcelInput){
    importExcelButton.addEventListener('click', () => importExcelInput.click());
    importExcelInput.addEventListener('change', handleImportExcelFile);
  }

  // 間引き（再サンプリング）適用
  if(applyThinningButton){
    applyThinningButton.addEventListener('click', () => {
      try{ applyEnvelopeThinning(); }catch(e){ console.warn('apply thinning error', e); }
    });
  }
  if(thin_target_points){
    thin_target_points.addEventListener('change', () => {
      // 値変更で即適用（ユーザーの意図が明確なため）
      try{
        // 可能であればフル包絡線から再適用して点数の増加にも対応
        const side = getCurrentSide();
        const full = originalEnvelopeBySide[side];
        if(Array.isArray(full) && full.length > 2){
          const metrics = calculateJTCCMMetrics(full, parseFloat(max_ultimate_deformation.value), parseFloat(alpha_factor.value));
          envelopeData = reapplyDisplayThinning(full, side, metrics);
          renderPlot(envelopeData, metrics);
          renderResults(metrics);
          appendLog('表示間引き(目標)を再適用（キャッシュ包絡線）: ' + envelopeData.length + ' 点');
        } else {
          applyEnvelopeThinning();
        }
      }catch(e){ console.warn('apply thinning error', e); }
    });
  }


  function clearInputData(){
    gammaInput.value = '';
    loadInput.value = '';
    rawData = [];
    envelopeData = null;
    analysisResults = {};
    
  // 自動解析化に伴い、旧ボタンの状態管理は不要
  if(downloadExcelButton) downloadExcelButton.disabled = true;
  if(generatePdfButton) generatePdfButton.disabled = true;
    if(undoButton) undoButton.disabled = true;
  if(redoButton) redoButton.disabled = true;
  if(openPointEditButton) openPointEditButton.disabled = true;
  historyStack = [];
  redoStack = [];
    plotDiv.innerHTML = '';
    // 結果表示リセット
  ['val_pmax','val_py','val_dy','val_K','val_pu','val_dv','val_du','val_mu','val_ds','val_p0_a','val_p0_b','val_p0','val_pa'].forEach(id=>{
      const el = document.getElementById(id); if(el) el.textContent='-';
    });
  }

  // === PDF Generation ===
  async function generatePdfReport(){
    try{
      if(!analysisResults || !envelopeData || !envelopeData.length){
        alert('解析結果がありません');
        return;
      }
      const specimen = (specimen_name && specimen_name.value ? specimen_name.value.trim() : 'testname');
      // Ensure jsPDF and html2canvas
      const { jsPDF } = window.jspdf || {};
      if(!jsPDF){
        alert('jsPDFライブラリが読み込まれていません');
        return;
      }
      if(typeof html2canvas === 'undefined'){
        alert('html2canvasライブラリが読み込まれていません');
        return;
      }

      // Create temporary container for PDF content
      const container = document.createElement('div');
      container.style.cssText = 'position:absolute; left:-9999px; top:0; width:800px; background:white; padding:20px; font-family:sans-serif;';
      document.body.appendChild(container);

      // Build HTML content
      const r = analysisResults;
      const fmt2 = (v) => (Number.isFinite(v) && v > 0) ? v.toFixed(2) + ' mm' : '-';
      container.innerHTML = `
        <div style="text-align:center; margin-bottom:15px;">
          <h1 style="font-size:24px; margin:10px 0;">接合部性能評価レポート</h1>
          <p style="font-size:12px; color:#333; margin:5px 0;">試験体名称: ${specimen.replace(/</g,'&lt;')}</p>
        </div>
        <div id="pdf-plot" style="width:100%; height:400px; margin-bottom:20px;"></div>
        <div style="display:flex; gap:20px;">
          <div style="flex:1;">
            <h3 style="font-size:16px; margin:10px 0; border-bottom:2px solid #333; padding-bottom:5px;">入力パラメータ</h3>
            <table style="width:100%; font-size:12px; border-collapse:collapse; table-layout:fixed;">
              <colgroup>
                <col style="width:60%">
                <col style="width:40%">
              </colgroup>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">最大変位δmax</td><td style="text-align:right; padding:6px 8px;">${Number(max_ultimate_deformation.value).toLocaleString('ja-JP')} mm</td></tr>
              <tr style="border-bottom:1px solid #ccc;"><td style="padding:6px 8px;">α</td><td style="text-align:right; padding:6px 8px;">${alpha_factor.value}</td></tr>
            </table>
          </div>
          <div style="flex:1;">
            <h3 style="font-size:16px; margin:10px 0; border-bottom:2px solid #333; padding-bottom:5px;">計算結果</h3>
            <table style="width:100%; font-size:12px; border-collapse:collapse; table-layout:fixed;">
              <colgroup>
                <col style="width:60%">
                <col style="width:40%">
              </colgroup>
              <tr style="border-bottom:1px solid #ccc;"><td style="padding:6px 8px;">Pmax (kN)</td><td style="text-align:right; padding:6px 8px;">${r.Pmax?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">Py (kN)</td><td style="text-align:right; padding:6px 8px;">${r.Py?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">Pu (kN)</td><td style="text-align:right; padding:6px 8px;">${r.Pu?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">δv (mm)</td><td style="text-align:right; padding:6px 8px;">${fmt2(r.delta_v)}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">δu (mm)</td><td style="text-align:right; padding:6px 8px;">${fmt2(r.delta_u)}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">μ</td><td style="text-align:right; padding:6px 8px;">${r.mu?.toFixed(2) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">Ds</td><td style="text-align:right; padding:6px 8px;">${r.mu && r.mu>0 ? (1/Math.sqrt(2*r.mu-1)).toFixed(3) : '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">P0(a)</td><td style="text-align:right; padding:6px 8px;">${r.p0_a?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">P0(b)</td><td style="text-align:right; padding:6px 8px;">${r.p0_b?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #eee;"><td style="padding:6px 8px;">P0</td><td style="text-align:right; padding:6px 8px;">${r.P0?.toFixed(3) ?? '-'}</td></tr>
              <tr style="border-bottom:1px solid #ccc;"><td style="padding:6px 8px;">Pa (kN)</td><td style="text-align:right; padding:6px 8px;">${r.Pa?.toFixed(3) ?? '-'}</td></tr>
            </table>
          </div>
        </div>
  <div style="text-align:center; font-size:10px; color:#666; margin-top:20px;">© Arkhitek / Generated by 接合部性能評価プログラム</div>
      `;

      // Render Plotly graph to temporary container
      const pdfPlotDiv = container.querySelector('#pdf-plot');
      
      // Create simplified annotations for PDF (label only, no values)
      const simplifiedAnnotations = plotDiv.layout.annotations ? plotDiv.layout.annotations.map(ann => {
        // Extract label from text (remove values and units)
        let simplifiedText = '';
        // IMPORTANT: Detect P0 first so embedded 'Py=' or 'Pmax=' doesn't override
    if(ann.text.includes('P0(a)')) simplifiedText = 'P0(a)';
    else if(ann.text.includes('P0(b)')) simplifiedText = 'P0(b)';
        else if(ann.text.includes('δu=')) simplifiedText = 'δu';
        else if(ann.text.includes('δy=')) simplifiedText = 'δy';
        else if(ann.text.includes('Py=')) simplifiedText = 'Py';
        else if(ann.text.includes('Pu=')) simplifiedText = 'Pu';
        else if(ann.text.includes('δv=')) simplifiedText = 'δv';
        else if(ann.text.includes('Pmax=')) simplifiedText = 'Pmax';
        else return ann;
        
        return {
          ...ann,
          text: simplifiedText,
          font: {...ann.font, size: 10},
          ax: ann.ax ? ann.ax * 0.7 : 0,
          ay: ann.ay ? ann.ay * 0.7 : 0
        };
      }) : [];
      
      await Plotly.newPlot(pdfPlotDiv, plotDiv.data, {
        ...plotDiv.layout,
        width: 760,
        height: 400,
        margin: {l:60, r:20, t:40, b:60},
        showlegend: false,
        annotations: simplifiedAnnotations
      }, {displayModeBar: false});

      // Convert to image using html2canvas
      const canvas = await html2canvas(container, {
        scale: 2,
        backgroundColor: '#ffffff',
        logging: false
      });

      // Remove temporary container
      document.body.removeChild(container);

      // Create PDF
      const doc = new jsPDF({orientation:'portrait', unit:'mm', format:'a4'});
      const imgData = canvas.toDataURL('image/png');
      const pageW = 210;
      const pageH = 297;
      const imgW = pageW;
      const imgH = (canvas.height * pageW) / canvas.width;
      
      // Scale to fit page if necessary
      if(imgH > pageH - 20){
        const scale = (pageH - 20) / imgH;
        doc.addImage(imgData, 'PNG', 0, 10, imgW * scale, imgH * scale);
      } else {
        doc.addImage(imgData, 'PNG', 0, 10, imgW, imgH);
      }

      const pdfFileName = `Report_${specimen.replace(/[^a-zA-Z0-9_\-一-龥ぁ-んァ-ヶ]/g,'_')}.pdf`;
      doc.save(pdfFileName);
      appendLog('PDFレポートを生成しました');
    }catch(err){
      console.error('PDF生成エラー:', err);
      alert('PDF生成に失敗しました: ' + (err && err.message ? err.message : err));
      appendLog('PDF生成エラー: ' + (err && err.stack ? err.stack : err));
    }
  }

  // 編集モードは廃止

  // === 起動時 sample.csv 自動読込のみ ===
  function loadCsvText(text){
    const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
    const pairs = [];
    const numericRegex = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;
    for(const line of lines){
      const cols = line.split(/,|\t|;/).map(c=>c.trim());
      if(cols.length < 2) continue;
      if(!numericRegex.test(cols[0]) || !numericRegex.test(cols[1])) continue; // header/非数値は除外
      pairs.push([parseFloat(cols[0]), parseFloat(cols[1])]);
    }
    if(pairs.length === 0){
      console.warn('CSVに有効なデータがありません。');
      appendLog('警告: CSVに有効なデータがありません');
      return;
    }
    gammaInput.value = pairs.map(p=>p[0]).join('\n');
    loadInput.value = pairs.map(p=>p[1]).join('\n');
    handleDirectInput();
  }

  function autoLoadSample(){
    fetch('sample.csv', {cache:'no-cache'})
      .then(r => r.ok ? r.text() : Promise.reject(new Error('sample.csvが取得できません')))
      .then(text => loadCsvText(text))
      .catch(err => {
        // file:// でのCORS制約時は組込サンプルへフォールバック
        if(location && location.protocol === 'file:'){
          console.warn('file:// での自動読込を組込サンプルにフォールバックします。詳細:', err.message);
          appendLog('情報: sample.csv 取得失敗 → 組込サンプルを使用 ('+ err.message +')');
          loadCsvText(BUILTIN_SAMPLE_CSV);
        } else {
          console.warn('sample.csv 自動読込失敗:', err.message);
          appendLog('警告: sample.csv 自動読込失敗 ('+ err.message +')');
        }
      });
  }

  autoLoadSample();

  // 包絡線ベースの表示範囲を計算（10%マージン、ゼロ幅回避込み）
  function computeEnvelopeRanges(env){
    if(!env || env.length === 0){
      return { xRange: [-1, 1], yRange: [-1, 1] };
    }
    const xs = env.map(pt => pt.gamma);
    const ys = env.map(pt => pt.Load);
    let minX = Math.min(...xs), maxX = Math.max(...xs);
    let minY = Math.min(...ys), maxY = Math.max(...ys);
    // ゼロ幅のときは小さな幅を与える
    if(minX === maxX){ const pad = Math.max(1e-6, Math.abs(minX)*0.1 || 1e-6); minX -= pad; maxX += pad; }
    if(minY === maxY){ const pad = Math.max(1e-3, Math.abs(minY)*0.1 || 1e-3); minY -= pad; maxY += pad; }
    const mx = (maxX - minX) * 0.1;
    const my = (maxY - minY) * 0.1;
    return { xRange: [minX - mx, maxX + mx], yRange: [minY - my, maxY + my] };
  }

  // Plotly.relayout の安全ラッパ：引数検証と失敗時の抑止・診断
  function safeRelayout(gd, updates){
    try{
      if(typeof updates === 'string'){
        // 第3引数が必要なケースは呼び出し側の責務。
        // 誤用を避けるため、この分岐ではそのまま通さず警告して無視する。
        console.warn('[safeRelayout] 文字列キーのみの呼び出しはサポートしていません:', updates);
        return Promise.resolve();
      }
      if(!updates || typeof updates !== 'object' || Array.isArray(updates)){
        console.warn('[safeRelayout] invalid updates. expect plain object. got =', updates);
        return Promise.resolve();
      }
      const p = Plotly.relayout(gd, updates);
      // Promise 対応: 失敗を握りつぶして未処理拒否を防ぐ
      if(p && typeof p.then === 'function'){
        return p.catch(err => {
          console.warn('[safeRelayout] relayout rejected:', err);
        });
      }
      return Promise.resolve();
    }catch(err){
      console.warn('[safeRelayout] wrapper error', err);
      return Promise.resolve();
    }
  }

  // 軽量なグローバル抑止: 理由が undefined の未処理Promise拒否のみ握りつぶす（本番でも常時有効）
  // Plotly 内部の一部経路で reject(undefined) が発生するため、実用上のノイズを抑える。
  if(typeof window !== 'undefined' && window.addEventListener){
    window.addEventListener('unhandledrejection', function(ev){
      try{
        if(!ev) return;
        if(typeof ev.reason === 'undefined'){
          // デバッグが必要な場合は ?debug=layout を使用（詳細パッチが有効化される）
          ev.preventDefault();
        }
      }catch(_){/* noop */}
    });
  }

  // デバッグフラグ (?debug=layout または APP_CONFIG.debugLayout) が true のときのみ詳細パッチを適用
  (function patchGlobalRelayout(){
    try{
      if(!window.Plotly) return;
      if(window.__PLOTLY_RELAYOUT_PATCHED__) return;
      function isLayoutDebug(){
        try{ const u=new URL(window.location.href); if(u.searchParams.get('debug')==='layout') return true; }catch(_){/*noop*/}
        try{ if(window.APP_CONFIG && window.APP_CONFIG.debugLayout===true) return true; }catch(_){/*noop*/}
        return false;
      }
      const DEBUG = isLayoutDebug();
      if(!DEBUG){ return; } // 本番は何もしない

      const origRelayout = window.Plotly.relayout;
      if(typeof origRelayout !== 'function') return;
      window.Plotly.relayout = function patchedRelayout(gd, a, b){
        try{
          // 文字列キー指定のときは値が未指定なら無視
          if(typeof a === 'string'){
            if(typeof b === 'undefined'){
              console.warn('[patch.relayout] string key without value. ignore:', a);
              return Promise.resolve();
            }
            const pr = origRelayout.call(window.Plotly, gd, a, b);
            return (pr && typeof pr.then === 'function') ? pr.catch(err => {
              console.warn('[patch.relayout] rejected (string key):', a, b, err);
            }) : Promise.resolve();
          }
          // オブジェクト以外は無視（undefined reject を回避）
          if(!a || typeof a !== 'object' || Array.isArray(a)){
            console.warn('[patch.relayout] invalid updates (expect plain object). got =', a);
            return Promise.resolve();
          }
          const pr = origRelayout.call(window.Plotly, gd, a);
          return (pr && typeof pr.then === 'function') ? pr.catch(err => {
            console.warn('[patch.relayout] rejected (object updates):', a, err);
          }) : Promise.resolve();
        }catch(err){
          console.warn('[patch.relayout] unexpected error', err);
          return Promise.resolve();
        }
      };
      window.__PLOTLY_RELAYOUT_PATCHED__ = true;
      // Plotly.Lib.warn のうち "Relayout fail" を抑制（デバッグ時のみ）
      try{
        if(window.Plotly.Lib && typeof window.Plotly.Lib.warn === 'function' && !window.__PLOTLY_LIBWARN_PATCHED__){
          const origWarn = window.Plotly.Lib.warn;
          window.Plotly.Lib.warn = function patchedWarn(){
            try{
              const msg = arguments && arguments[0] ? String(arguments[0]) : '';
              if(msg.indexOf('Relayout fail') !== -1){
                // 2 つ目以降の引数（問題の updates など）を一応表示
                console.info('[plotly.warn suppressed] Relayout fail:', arguments[1], arguments[2]);
                return; // 抑制
              }
            }catch(_){/* noop */}
            return origWarn.apply(this, arguments);
          };
          window.__PLOTLY_LIBWARN_PATCHED__ = true;
          console.info('[patch.lib.warn] installed');
        }
      }catch(err){ console.warn('patch Plotly.Lib.warn failed', err); }

          // さらに保険として console.warn を薄くラップし、"Relayout fail" を info に降格（デバッグのみ）
          try{
            if(!window.__CONSOLE_WARN_PATCHED__ && typeof console !== 'undefined' && typeof console.warn === 'function'){
              const _origConsoleWarn = console.warn.bind(console);
              console.warn = function(){
                try{
                  const msg = arguments && arguments[0] ? String(arguments[0]) : '';
                  if(msg.indexOf('Relayout fail') !== -1){
                    console.info('[console.warn suppressed] Relayout fail:', arguments[1], arguments[2]);
                    return;
                  }
                }catch(_){/* noop */}
                return _origConsoleWarn.apply(console, arguments);
              };
              window.__CONSOLE_WARN_PATCHED__ = true;
              console.info('[patch.console.warn] installed');
            }
            // 同様に console.log/console.error でも出る可能性を抑止（デバッグのみ）
            if(!window.__CONSOLE_LOG_PATCHED__ && typeof console !== 'undefined' && typeof console.log === 'function'){
              const _origConsoleLog = console.log.bind(console);
              console.log = function(){
                try{
                  const msg = arguments && arguments[0] ? String(arguments[0]) : '';
                  if(msg.indexOf('Relayout fail') !== -1 || msg.indexOf('WARN: Relayout fail') !== -1){
                    console.info('[console.log suppressed] Relayout fail:', arguments[1], arguments[2]);
                    return;
                  }
                }catch(_){/* noop */}
                return _origConsoleLog.apply(console, arguments);
              };
              window.__CONSOLE_LOG_PATCHED__ = true;
              console.info('[patch.console.log] installed');
            }
            if(!window.__CONSOLE_ERROR_PATCHED__ && typeof console !== 'undefined' && typeof console.error === 'function'){
              const _origConsoleError = console.error.bind(console);
              console.error = function(){
                try{
                  const msg = arguments && arguments[0] ? String(arguments[0]) : '';
                  if(msg.indexOf('Relayout fail') !== -1){
                    console.info('[console.error suppressed] Relayout fail:', arguments[1], arguments[2]);
                    return;
                  }
                }catch(_){/* noop */}
                return _origConsoleError.apply(console, arguments);
              };
              window.__CONSOLE_ERROR_PATCHED__ = true;
              console.info('[patch.console.error] installed');
            }
          }catch(err){ /* ignore */ }
      // 追加: 内部 API Plots.relayout もパッチ（デバッグのみ）
      try{
        if(window.Plotly.Plots && typeof window.Plotly.Plots.relayout === 'function' && !window.__PLOTS_RELAYOUT_PATCHED__){
          const origPlotsRelayout = window.Plotly.Plots.relayout;
          let __relayoutFailSeq = 0;
          window.Plotly.Plots.relayout = function patchedPlotsRelayout(gd, a, b){
            try{
              if(typeof a === 'string'){
                if(typeof b === 'undefined'){
                  const id = ++__relayoutFailSeq;
                  console.groupCollapsed('[patch.plots.relayout]#'+id+' string key without value (ignore) key=', a);
                  console.trace('stack');
                  console.groupEnd();
                  return Promise.resolve();
                }
                const pr = origPlotsRelayout.call(window.Plotly.Plots, gd, a, b);
                return (pr && typeof pr.then === 'function') ? pr.catch(err => {
                  const id = ++__relayoutFailSeq;
                  console.groupCollapsed('[patch.plots.relayout]#'+id+' rejected (string key) key='+a);
                  console.log('value=', b);
                  console.log('error=', err);
                  console.trace('stack');
                  console.groupEnd();
                }) : Promise.resolve();
              }
              if(!a || typeof a !== 'object' || Array.isArray(a)){
                const id = ++__relayoutFailSeq;
                console.groupCollapsed('[patch.plots.relayout]#'+id+' invalid updates (expect plain object)');
                console.log('updates=', a);
                console.trace('stack');
                console.groupEnd();
                return Promise.resolve();
              }
              const pr = origPlotsRelayout.call(window.Plotly.Plots, gd, a);
              return (pr && typeof pr.then === 'function') ? pr.catch(err => {
                const id = ++__relayoutFailSeq;
                console.groupCollapsed('[patch.plots.relayout]#'+id+' rejected (object updates)');
                console.log('updates=', a);
                console.log('error=', err);
                console.trace('stack');
                console.groupEnd();
              }) : Promise.resolve();
            }catch(err){
              const id = ++__relayoutFailSeq;
              console.groupCollapsed('[patch.plots.relayout]#'+id+' unexpected error');
              console.log('error=', err);
              console.trace('stack');
              console.groupEnd();
              return Promise.resolve();
            }
          };
          window.__PLOTS_RELAYOUT_PATCHED__ = true;
          console.info('[patch.plots.relayout] installed');
        }
      }catch(err){ console.warn('patch Plots.relayout failed', err); }

      // 未処理拒否の既定ログを抑制（理由 undefined のみ／デバッグ時のみ）
      window.addEventListener('unhandledrejection', function(ev){
        try{
          if(!ev) return;
          // Plotly 由来かつ undefined 理由の拒否を抑制
          const r = ev.reason;
          const isUndefined = (typeof r === 'undefined');
          const srcMatch = (ev && ev.promise && typeof ev.promise === 'object');
          if(isUndefined){
            console.warn('[unhandledrejection] suppressed undefined reason from a promise (likely Plotly relayout)');
            ev.preventDefault();
          }
        }catch(_){/* noop */}
      });
      console.info('[patch.relayout] installed');
    }catch(err){ console.warn('patchGlobalRelayout failed', err); }
  })();

  // 包絡線範囲へフィット（初期描画・Autoscaleボタン・ダブルクリックで共通使用）
  function fitEnvelopeRanges(reason){
    try{
      if(!envelopeData || !envelopeData.length) return;
      const r = computeEnvelopeRanges(envelopeData);
      console.info('[Fit] 包絡線範囲へフィット:', reason || '');
      cachedEnvelopeRange = r; // キャッシュ更新
      safeRelayout(plotDiv, {
        'xaxis.autorange': false,
        'yaxis.autorange': false,
        'xaxis.range': r.xRange,
        'yaxis.range': r.yRange
      });
    }catch(err){ console.warn('fitEnvelopeRanges エラー', err); }
  }

  // === Direct Input Handling ===
  function handleDirectInput(){
    const gammaText = gammaInput.value.trim();
    const loadText = loadInput.value.trim();

    if(!gammaText || !loadText) return; // どちらか空なら何もしない

    try {
      // 行単位に分割（カンマ区切り等があっても先に改行を優先）
      const gammaLines = gammaText.split(/\r?\n/);
      const loadLines  = loadText.split(/\r?\n/);

      const pairCount = Math.min(gammaLines.length, loadLines.length);
      const parsed = [];
      let skipped = 0;

      // 厳密な数値判定（単位や文字が付いた行は除外）
      const isNumericString = (s) => /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(s);

      for(let i=0; i<pairCount; i++){
        const gStrRaw = gammaLines[i];
        const lStrRaw = loadLines[i];
        if(gStrRaw == null || lStrRaw == null){
          skipped++; continue;
        }
        const gStr = gStrRaw.trim();
        const lStr = lStrRaw.trim();
        if(!gStr || !lStr){ // 空白行はスキップ
          skipped++; continue;
        }

        // 数値文字列判定：行全体が数値であることを要求
        if(!isNumericString(gStr) || !isNumericString(lStr)){
          skipped++; continue; // 項目名など非数値行は無視
        }

        const gNum = parseFloat(gStr);
        const lNum = parseFloat(lStr);

        parsed.push({
          Load: lNum,
          gamma: gNum,
          gamma0: gNum // 直接入力の場合は補正なし
        });
      }

      if(parsed.length === 0){
        console.warn('有効な数値データがありません。');
        return;
      }

  rawData = parsed;
  // 生データを入れ替えたら、過去の手編集は無効化（新データと整合しないため）
  editedEnvelopeBySide.positive = null;
  editedEnvelopeBySide.negative = null;
  editedDirtyBySide.positive = false;
  editedDirtyBySide.negative = false;
      header = ['Load', 'gamma', 'gamma0'];
  // 自動解析: 旧ボタン無効化不要

      if(skipped > 0){
        console.info(`非数値または空白行を ${skipped} 行スキップしました。有効データ: ${parsed.length} 行`);
      }

      // 最低3点以上で自動解析（2点以下だと線形近似など不安定）
      if(rawData.length >= 3){
        setTimeout(() => processDataDirect(), 50);
      }
    } catch(err) {
      console.error('データ解析エラー:', err);
      appendLog('データ解析エラー: ' + (err && err.stack ? err.stack : err.message));
    }
  }

  // === Main Processing ===
  // 手動 processData 廃止（自動解析は processDataDirect 使用）

  // === Direct Input Processing ===
  function processDataDirect(){
    try {
      const alpha = parseFloat(alpha_factor.value);
      const side = envelope_side.value;
      const delta_u_max = parseFloat(max_ultimate_deformation.value);
      if(!isFinite(alpha) || !isFinite(delta_u_max) || delta_u_max <= 0){
        console.warn('入力値が不正です');
        return;
      }

      // 編集済み包絡線（手動調整 or インポート後ユーザ編集）がある場合はその点列を維持
      const edited = editedEnvelopeBySide[side];
      if(Array.isArray(edited) && edited.length >= 2){
        envelopeData = edited.map(p => ({...p}));
        // 再計算前に間引き UI の希望を適用（インポート後でも反映できるように）
        const targetInputEl = document.getElementById('thin_target_points');
        const sliderEl = typeof envelope_thinning_rate !== 'undefined' ? envelope_thinning_rate : null;
        let working = envelopeData.map(p=>({...p}));
        // 目標点数優先（元包絡線を再利用できる場合はそれを使う）
        if(targetInputEl){
          const target = parseInt(targetInputEl.value,10);
          if(Number.isFinite(target) && target > 5 && working.length > target){
            const mandatoryGammas = [];
            if(Number.isFinite(analysisResults?.delta_y)) mandatoryGammas.push(analysisResults.delta_y);
            if(Number.isFinite(analysisResults?.delta_u)) mandatoryGammas.push(analysisResults.delta_u);
            try {
              const loopGammas = detectLoopPeakGammas(rawData || [], side);
              if(Array.isArray(loopGammas) && loopGammas.length){ loopGammas.forEach(g => { if(Number.isFinite(g)) mandatoryGammas.push(g); }); }
            } catch(err){ console.warn('再間引き(編集キャッシュ): ループピーク検出失敗', err); }
            working = sampleEnvelopeExact(working, target, mandatoryGammas);
            console.info(`[re-thin edited] target=${target} -> ${working.length} 点`);
          }
        }
        // スライダーも適用（追加でさらに減らす）
        if(sliderEl){
          const thinningRate = parseInt(sliderEl.value,10);
          if(Number.isFinite(thinningRate) && thinningRate > 10 && working.length > 50){
            const mandatoryGammas = [];
            if(Number.isFinite(analysisResults?.delta_y)) mandatoryGammas.push(analysisResults.delta_y);
            if(Number.isFinite(analysisResults?.delta_u)) mandatoryGammas.push(analysisResults.delta_u);
            try {
              const loopGammas = detectLoopPeakGammas(rawData || [], side);
              if(Array.isArray(loopGammas) && loopGammas.length){ loopGammas.forEach(g => { if(Number.isFinite(g)) mandatoryGammas.push(g); }); }
            } catch(err){ console.warn('再間引き(編集キャッシュ/slider): ループピーク検出失敗', err); }
            let targetPoints;
            if(thinningRate <= 50){
              targetPoints = Math.round(100 - (thinningRate - 10) * (20 / 40));
            } else {
              targetPoints = Math.round(80 - (thinningRate - 50) * (70 / 50));
            }
            working = thinEnvelopeUniform(working, targetPoints, mandatoryGammas);
            console.info(`[re-thin edited slider] rate=${thinningRate} target=${targetPoints} -> ${working.length} 点`);
          }
        }
        envelopeData = working.map(p=>({...p}));
        analysisResults = calculateJTCCMMetrics(envelopeData, delta_u_max, alpha);
        renderPlot(envelopeData, analysisResults);
        renderResults(analysisResults);
        if(downloadExcelButton) downloadExcelButton.disabled = false;
        if(generatePdfButton) generatePdfButton.disabled = false;
        historyStack = [cloneEnvelope(envelopeData)];
        redoStack = [];
        updateHistoryButtons();
        return;
      }

      // 生データからフル包絡線生成
      const fullEnvelope = generateEnvelope(rawData, side);
      if(fullEnvelope.length === 0){
        console.warn('包絡線生成失敗');
        return;
      }
      // 指標計算は常にフル包絡線を使用（精度優先）
      analysisResults = calculateJTCCMMetrics(fullEnvelope, delta_u_max, alpha);

      let displayEnvelope = fullEnvelope;
      const targetInputEl = document.getElementById('thin_target_points');
      const sliderEl = typeof envelope_thinning_rate !== 'undefined' ? envelope_thinning_rate : null;

      if(targetInputEl){
        // 新UI: 目標表示点数 (Exact サンプリング)
        const target = parseInt(targetInputEl.value,10);
        if(Number.isFinite(target) && target > 5 && fullEnvelope.length > target){
          const mandatoryGammas = [];
          if(Number.isFinite(analysisResults.delta_y)) mandatoryGammas.push(analysisResults.delta_y);
          if(Number.isFinite(analysisResults.delta_u)) mandatoryGammas.push(analysisResults.delta_u);
          try {
            const loopGammas = detectLoopPeakGammas(rawData, side);
            if(Array.isArray(loopGammas) && loopGammas.length){
              loopGammas.forEach(g => { if(Number.isFinite(g)) mandatoryGammas.push(g); });
            }
          } catch(err){ console.warn('ループピーク検出エラー', err); }
          displayEnvelope = sampleEnvelopeExact(fullEnvelope, target, mandatoryGammas);
          console.info(`[sampleEnvelopeExact] target=${target} -> ${displayEnvelope.length} 点`);
        }
      }
      if(sliderEl){
        // 旧スライダー互換: 間引き率 (0-100)
        const thinningRate = parseInt(sliderEl.value,10);
        // 条件緩和: fullEnvelope.length が一定閾値(50)超 & thinningRate > 10 なら適用
        if(Number.isFinite(thinningRate) && fullEnvelope.length > 50 && thinningRate > 10){
          const mandatoryGammas = [];
          if(Number.isFinite(analysisResults.delta_y)) mandatoryGammas.push(analysisResults.delta_y);
          if(Number.isFinite(analysisResults.delta_u)) mandatoryGammas.push(analysisResults.delta_u);
          try {
            const loopGammas = detectLoopPeakGammas(rawData, side);
            if(Array.isArray(loopGammas) && loopGammas.length){
              loopGammas.forEach(g => { if(Number.isFinite(g)) mandatoryGammas.push(g); });
            }
          } catch(err){ console.warn('ループピーク検出エラー', err); }
          let targetPoints;
          if(thinningRate <= 50){
            targetPoints = Math.round(100 - (thinningRate - 10) * (20 / 40));
          } else {
            targetPoints = Math.round(80 - (thinningRate - 50) * (70 / 50));
          }
          displayEnvelope = thinEnvelopeUniform(fullEnvelope, targetPoints, mandatoryGammas);
          console.info(`[thinEnvelopeUniform] rate=${thinningRate} target=${targetPoints} -> ${displayEnvelope.length} 点`);
        }
      }

      envelopeData = displayEnvelope.map(p => ({...p}));
      renderPlot(envelopeData, analysisResults);
      renderResults(analysisResults);
      if(downloadExcelButton) downloadExcelButton.disabled = false;
      if(generatePdfButton) generatePdfButton.disabled = false;
      historyStack = [cloneEnvelope(envelopeData)];
      redoStack = [];
      updateHistoryButtons();
    } catch(err){
      console.error('processDataDirect エラー:', err);
      appendLog('計算エラー(自動解析): ' + (err && err.stack ? err.stack : err.message));
    }
  }



  // === Envelope Generation (Section II.3) ===
  // 均等間引き: 重要点を最小限保持し、包絡線全体で点の間隔が均等になるように間引く
  function thinEnvelopeUniform(envelope, targetPoints, mandatoryGammas){
    try{
      if(!Array.isArray(envelope) || envelope.length <= targetPoints) return envelope.map(pt=>({...pt}));
      
      const pts = envelope.map(pt => ({x: pt.gamma, y: pt.Load}));
      
      // 弧長（累積距離）を先に計算
      const arcLengths = [0];
      for(let i = 1; i < pts.length; i++){
        const dx = pts[i].x - pts[i - 1].x;
        const dy = pts[i].y - pts[i - 1].y;
        arcLengths.push(arcLengths[i - 1] + Math.hypot(dx, dy));
      }
      const totalLength = arcLengths[arcLengths.length - 1];
      
      // 最小限の重要点のみ特定
      const mandatory = new Set();
      
      // 先頭と最終点は必須
      mandatory.add(0);
      mandatory.add(pts.length - 1);
      
      // 最大荷重点（Pmax）は必ず保持
      let idxPmax = 0;
      let maxAbs = -Infinity;
      for(let i = 0; i < pts.length; i++){
        const absLoad = Math.abs(pts[i].y);
        if(absLoad > maxAbs){
          maxAbs = absLoad;
          idxPmax = i;
        }
      }
      mandatory.add(idxPmax);
      
      // 指定されたγ値に最も近い点（δy, δu等）も必須
      if(Array.isArray(mandatoryGammas)){
        mandatoryGammas.forEach(gTarget => {
          if(!Number.isFinite(gTarget)) return;
          let bestIdx = -1;
          let bestDiff = Infinity;
          for(let i = 0; i < pts.length; i++){
            const diff = Math.abs(pts[i].x - gTarget);
            if(diff < bestDiff){
              bestDiff = diff;
              bestIdx = i;
            }
          }
          if(bestIdx >= 0) mandatory.add(bestIdx);
        });
      }
      
      // 局所ピークは不要（均等配置で十分カバーされる）
      
      console.log(`[thinEnvelope] 必須点の内訳: 先頭/終端=2, Pmax=1, mandatoryGammas=${mandatoryGammas ? mandatoryGammas.length : 0}, 合計=${mandatory.size}`);
      
      // 必須点が多すぎる場合は、最小限の必須点のみに絞る
      if(mandatory.size > targetPoints * 0.5){
        console.warn(`[thinEnvelope] 必須点数(${mandatory.size})が多すぎるため、最小限の必須点のみに制限します`);
        
        // 最小限の必須点のみを残す
        mandatory.clear();
        mandatory.add(0); // 先頭
        mandatory.add(pts.length - 1); // 終端
        mandatory.add(idxPmax); // Pmax
        
        // δy, δu のみ追加（loopGammasは除外）
        if(Array.isArray(mandatoryGammas) && mandatoryGammas.length >= 2){
          for(let i = 0; i < 2; i++){
            const gTarget = mandatoryGammas[i];
            if(!Number.isFinite(gTarget)) continue;
            let bestIdx = -1;
            let bestDiff = Infinity;
            for(let j = 0; j < pts.length; j++){
              const diff = Math.abs(pts[j].x - gTarget);
              if(diff < bestDiff){
                bestDiff = diff;
                bestIdx = j;
              }
            }
            if(bestIdx >= 0) mandatory.add(bestIdx);
          }
        }
        
        console.log(`[thinEnvelope] 必須点を ${mandatory.size} 点に削減しました`);
      }
      
      // それでも必須点数が目標を超える場合
      if(mandatory.size >= targetPoints){
        console.warn(`[thinEnvelope] 必須点数(${mandatory.size})が目標点数(${targetPoints})以上のため、必須点のみを返します`);
        const indices = Array.from(mandatory).sort((a, b) => a - b);
        return indices.map(i => ({...envelope[i]}));
      }
      
      // 弧長ベースで完全に均等分割して点を配置
      const step = totalLength / (targetPoints - 1);
      const selectedSet = new Set();
      
      // デバッグ用
      console.log(`[thinEnvelope] 元の点数: ${pts.length}, 目標点数: ${targetPoints}, 必須点数: ${mandatory.size}`);
      
      // Step 1: 均等配置による点の選択（必須点は考慮せず純粋に均等）
      for(let j = 0; j < targetPoints; j++){
        const targetArc = j * step;
        
        // targetArcに最も近い点を選択
        let bestIdx = -1;
        let bestDiff = Infinity;
        
        for(let i = 0; i < pts.length; i++){
          const diff = Math.abs(arcLengths[i] - targetArc);
          if(diff < bestDiff){
            bestDiff = diff;
            bestIdx = i;
          }
        }
        
        if(bestIdx >= 0){
          selectedSet.add(bestIdx);
        }
      }
      
      console.log(`[thinEnvelope] Step1後の選択点数: ${selectedSet.size}`);
      
      // Step 2: 選択された点の中で、必須点に最も近い点を必須点に置き換え
      for(const mandatoryIdx of mandatory){
        if(selectedSet.has(mandatoryIdx)) continue; // 既に選択されていればOK
        
        // 必須点に最も近い選択済み点を探す
        let closestSelected = -1;
        let closestDist = Infinity;
        
        for(const selIdx of selectedSet){
          const dist = Math.abs(arcLengths[mandatoryIdx] - arcLengths[selIdx]);
          if(dist < closestDist){
            closestDist = dist;
            closestSelected = selIdx;
          }
        }
        
        // 最も近い選択点を必須点に置き換え（距離に関わらず）
        if(closestSelected >= 0){
          selectedSet.delete(closestSelected);
          selectedSet.add(mandatoryIdx);
        }
      }
      
      console.log(`[thinEnvelope] Step2後の選択点数: ${selectedSet.size}`);
      
      // インデックスでソートして返す
      const selected = Array.from(selectedSet).sort((a, b) => a - b);
      console.log(`[thinEnvelope] 最終結果点数: ${selected.length}, 先頭10点のインデックス: [${selected.slice(0, 10).join(', ')}]`);
      return selected.map(i => ({...envelope[i]}));
      
    }catch(err){
      console.warn('thinEnvelopeUniform エラー', err);
      return envelope;
    }
  }
  
  // 包絡線間引き（Ramer-Douglas-Peucker 風）: 重要点保持しつつ 40～50 点程度へ縮約
  function thinEnvelope(envelope, minPoints, maxPoints, mandatoryGammas){
    try{
      if(!Array.isArray(envelope) || envelope.length <= maxPoints) return envelope.map(pt=>({...pt}));
      const pts = envelope.map(pt => ({x: pt.gamma, y: pt.Load}));
      // 重要点候補（必ず保持）: 先頭, 最終, 最大荷重点
      let idxPmax = 0; let maxAbs = -Infinity;
      for(let i=0;i<pts.length;i++){ const a = Math.abs(pts[i].y); if(a>maxAbs){ maxAbs=a; idxPmax=i; } }
      const mandatory = new Set([0, pts.length-1, idxPmax]);
      // 追加必須点: 指定された γ 値（δy, δu）に最も近い点
      if(Array.isArray(mandatoryGammas)){
        mandatoryGammas.forEach(gTarget => {
          if(!Number.isFinite(gTarget)) return;
          let bestIdx = -1; let bestDiff = Infinity;
          for(let i=0;i<pts.length;i++){
            const d = Math.abs(pts[i].x - gTarget);
            if(d < bestDiff){ bestDiff = d; bestIdx = i; }
          }
          if(bestIdx >= 0) mandatory.add(bestIdx);
        });
      }
      // 各ループ（繰返し）の最大荷重点（局所ピーク）を必須保持点に追加
      // 単純な局所最大検出: y[i] >= y[i-1] かつ y[i] >= y[i+1] （絶対値ベース）
      for(let i=1;i<pts.length-1;i++){
        const y0 = Math.abs(pts[i-1].y);
        const y1 = Math.abs(pts[i].y);
        const y2 = Math.abs(pts[i+1].y);
        if(y1 >= y0 && y1 >= y2){
          mandatory.add(i);
        }
      }
      // もし必須点だけで maxPoints を超える場合は必須点から優先順位で削減
      if(mandatory.size > maxPoints){
        // 優先順位: 0(先頭), last(末尾), Pmax, δy/δu近傍, その他ループピーク
        const mustArr = Array.from(mandatory);
        const priorityCore = [0, pts.length-1, idxPmax];
        const core = mustArr.filter(i=>priorityCore.includes(i));
        const others = mustArr.filter(i=>!priorityCore.includes(i));
        // コア + 均等サンプリングで maxPoints に収める
        const need = maxPoints - core.length;
        if(need <= 0){
          return core.sort((a,b)=>a-b).map(i=>({...envelope[i]}));
        }
        const step = others.length / (need + 1);
        const chosen = [];
        for(let j=1; j<=need; j++){
          const pick = others[Math.min(others.length-1, Math.round(j*step)-1)];
          if(pick !== undefined) chosen.push(pick);
        }
        return core.concat(chosen).sort((a,b)=>a-b).map(i=>({...envelope[i]}));
      }

      // 弧長(累積距離)を計算
      const arcLengths = [0];
      for(let i=1; i<pts.length; i++){
        const dx = pts[i].x - pts[i-1].x;
        const dy = pts[i].y - pts[i-1].y;
        arcLengths.push(arcLengths[i-1] + Math.hypot(dx, dy));
      }
      const totalLength = arcLengths[arcLengths.length-1];

      // 均等分割ベースの間引き（必須点を保持しつつ、弧長に沿って均等にサンプリング）
      const targetPoints = Math.min(maxPoints, Math.max(minPoints, Math.floor((minPoints + maxPoints) / 2)));
      const selected = new Set(mandatory); // 必須点は最初に選択
      
      // 必須点をインデックス順にソート
      const mandatoryArray = Array.from(mandatory).sort((a,b) => a-b);
      
      // 必須点の間の区間を均等に補完
      for(let m=0; m<mandatoryArray.length-1; m++){
        const startIdx = mandatoryArray[m];
        const endIdx = mandatoryArray[m+1];
        const segmentStartArc = arcLengths[startIdx];
        const segmentEndArc = arcLengths[endIdx];
        const segmentLength = segmentEndArc - segmentStartArc;
        
        if(segmentLength <= 0) continue;
        
        // この区間に割り当てる追加点数（区間長に比例、最低0点）
        const additionalPointsNeeded = Math.max(0, targetPoints - selected.size);
        if(additionalPointsNeeded <= 0) break;
        
        const segmentRatio = segmentLength / totalLength;
        const pointsForSegment = Math.max(0, Math.round(additionalPointsNeeded * segmentRatio));
        
        if(pointsForSegment > 0){
          // 区間内を弧長で均等分割
          const arcStep = segmentLength / (pointsForSegment + 1);
          for(let j=1; j<=pointsForSegment; j++){
            const targetArc = segmentStartArc + j * arcStep;
            // targetArcに最も近いインデックスを探す
            let bestIdx = startIdx;
            let bestDiff = Infinity;
            for(let k=startIdx+1; k<endIdx; k++){
              const diff = Math.abs(arcLengths[k] - targetArc);
              if(diff < bestDiff){
                bestDiff = diff;
                bestIdx = k;
              }
            }
            if(bestIdx > startIdx && bestIdx < endIdx){
              selected.add(bestIdx);
            }
          }
        }
      }
      
      // まだ目標点数に達していない場合、全体から均等にサンプリング
      if(selected.size < minPoints){
        const remaining = [];
        for(let i=0; i<pts.length; i++){
          if(!selected.has(i)) remaining.push(i);
        }
        const need = minPoints - selected.size;
        if(need > 0 && remaining.length > 0){
          const step = remaining.length / (need + 1);
          for(let j=1; j<=need; j++){
            const idx = remaining[Math.min(remaining.length-1, Math.round(j*step))];
            if(idx !== undefined) selected.add(idx);
          }
        }
      }
      
      // 目標点数を超えた場合、必須点以外から削減（弧長間隔が最小のものから）
      if(selected.size > maxPoints){
        const selectedArray = Array.from(selected).sort((a,b)=>a-b);
        const canRemove = [];
        
        // 削除候補（必須でない点）の弧長間隔を計算
        for(let i=1; i<selectedArray.length; i++){
          const idx = selectedArray[i];
          if(!mandatory.has(idx)){
            const prevIdx = selectedArray[i-1];
            const interval = arcLengths[idx] - arcLengths[prevIdx];
            canRemove.push({idx, interval});
          }
        }
        
        // 間隔が小さい順にソート
        canRemove.sort((a,b) => a.interval - b.interval);
        
        // 削除
        const toRemove = selected.size - maxPoints;
        for(let i=0; i<toRemove && i<canRemove.length; i++){
          selected.delete(canRemove[i].idx);
        }
      }

      const indices = Array.from(selected).sort((a,b)=>a-b);
      return indices.map(i=> ({...envelope[i]}));
    }catch(err){ console.warn('thinEnvelope エラー', err); return envelope; }
  }

  // 現在の包絡線を「表示点数目標」に合わせて再間引き（δy / δu / ループ最大点は保持）
  function applyEnvelopeThinning(){
    try{
      if(!envelopeData || !Array.isArray(envelopeData) || envelopeData.length < 3) return;
      const target = thin_target_points ? parseInt(thin_target_points.value, 10) : 50;
      if(!Number.isFinite(target) || target <= 0) return;
      // 目標より既に十分少ない場合は何もしない
      if(envelopeData.length <= target){
        appendLog('間引き対象外: 既に点数が目標以下です ('+envelopeData.length+' <= '+target+')');
        return;
      }
      // 重要点（保持）: δy, δu, ループ最大荷重点
      const mandatoryGammas = [];
      try{
        if(Number.isFinite(analysisResults?.delta_y)) mandatoryGammas.push(analysisResults.delta_y);
        if(Number.isFinite(analysisResults?.delta_u)) mandatoryGammas.push(analysisResults.delta_u);
      }catch(_){/* noop */}
      try{
        const side = getCurrentSide();
        const loops = detectLoopPeakGammas(rawData || [], side);
        if(Array.isArray(loops) && loops.length){ loops.forEach(g => { if(Number.isFinite(g)) mandatoryGammas.push(g); }); }
      }catch(_){/* noop */}

      const thinned = sampleEnvelopeExact(envelopeData, target, mandatoryGammas);
      if(!Array.isArray(thinned) || thinned.length < 2){ return; }
      pushHistory(envelopeData);
      envelopeData = thinned.map(p=>({...p}));
      recalculateFromEnvelope(envelopeData);
      appendLog('包絡線の表示間引き(Exact)を適用: '+envelopeData.length+' 点 / 目標 '+target);
    }catch(err){ console.warn('applyEnvelopeThinning failed', err); }
  }

  // フル包絡線に対し UI (目標点数 / スライダー) を用いて表示用間引きを再適用
  // Excelインポート後や側変更時に利用し、編集前の完全包絡線から再構築できるようにする
  function reapplyDisplayThinning(fullEnvelope, side, metrics){
    try{
      if(!Array.isArray(fullEnvelope) || fullEnvelope.length === 0) return [];
      let working = fullEnvelope.map(p=>({...p}));
      const targetInputEl = document.getElementById('thin_target_points');
      const sliderEl = document.getElementById('envelope_thinning_rate');

      // 必須点候補の γ 集合 (δy, δu, ループ最大荷重点)
      const mandatoryGammas = [];
      try{
        if(metrics && Number.isFinite(metrics.delta_y)) mandatoryGammas.push(metrics.delta_y);
        if(metrics && Number.isFinite(metrics.delta_u)) mandatoryGammas.push(metrics.delta_u);
      }catch(_){/* noop */}
      try{
        const loopGammas = detectLoopPeakGammas(rawData || [], side);
        if(Array.isArray(loopGammas) && loopGammas.length){
          loopGammas.forEach(g => { if(Number.isFinite(g)) mandatoryGammas.push(g); });
        }
      }catch(err){ console.warn('ループピーク検出失敗', err); }

      // 1. 目標点数による Exact サンプリング
      if(targetInputEl){
        const target = parseInt(targetInputEl.value, 10);
        if(Number.isFinite(target) && target > 5 && working.length > target){
          working = sampleEnvelopeExact(working, target, mandatoryGammas);
          console.info(`[reapplyDisplayThinning] target=${target} -> ${working.length} 点`);
        }
      }
      // 2. スライダーによる追加間引き (均等サンプリング)
      if(sliderEl){
        const thinningRate = parseInt(sliderEl.value, 10);
        if(Number.isFinite(thinningRate) && working.length > 50 && thinningRate > 10){
          let targetPoints;
          if(thinningRate <= 50){
            targetPoints = Math.round(100 - (thinningRate - 10) * (20 / 40));
          } else {
            targetPoints = Math.round(80 - (thinningRate - 50) * (70 / 50));
          }
            // mandatoryGammas を維持しつつ均等化
          working = thinEnvelopeUniform(working, targetPoints, mandatoryGammas);
          console.info(`[reapplyDisplayThinning] slider rate=${thinningRate} target=${targetPoints} -> ${working.length} 点`);
        }
      }
      return working.map(p=>({...p}));
    }catch(err){ console.warn('reapplyDisplayThinning エラー', err); return fullEnvelope.map(p=>({...p})); }
  }

  function generateEnvelope(data, side){
    try{
      // 片側抽出
      let filteredData;
      if(side === 'positive'){
        filteredData = data.filter(pt => Number.isFinite(pt.gamma) && Number.isFinite(pt.Load) && pt.gamma >= 0 && pt.Load >= 0);
      } else {
        filteredData = data.filter(pt => Number.isFinite(pt.gamma) && Number.isFinite(pt.Load) && pt.gamma <= 0 && pt.Load <= 0);
      }
      if(!filteredData || filteredData.length === 0) return [];

      const abs = (v)=>Math.abs(v);
      // 最大荷重点 index
      let idxPmax = 0; let maxAbs = -Infinity;
      for(let i=0;i<filteredData.length;i++){
        const a = abs(filteredData[i].Load);
        if(a > maxAbs){ maxAbs = a; idxPmax = i; }
      }

      const env = [];
      const LOAD_TOL = 1e-9;
      let lastG = -Infinity; let lastL = -Infinity;
      // 立ち上がり〜Pmax: γ単調増加 & 荷重非減少（緩い条件）
      for(let i=0;i<=idxPmax;i++){
        const pt = filteredData[i];
        const g = abs(pt.gamma); const L = abs(pt.Load);
        if(g + 1e-12 <= lastG) continue; // γ非減少
        if(L + LOAD_TOL < lastL) continue; // 荷重大きく低下はスキップ
        env.push({...pt});
        lastG = g; lastL = Math.max(lastL, L);
      }
      // Pmax 未含なら追加
      if(!env.some(p => p.gamma === filteredData[idxPmax].gamma && p.Load === filteredData[idxPmax].Load)){
        const p = filteredData[idxPmax]; env.push({...p}); lastG = abs(p.gamma); lastL = abs(p.Load);
      }
      // Pmax以降: γ単調条件のみ緩めて追加（急減少は許容: 解析用には後段で処理）
      for(let i=idxPmax+1;i<filteredData.length;i++){
        const pt = filteredData[i];
        const g = abs(pt.gamma);
        if(g + 1e-12 <= lastG) continue;
        env.push({...pt});
        lastG = g;
      }
      return env;
    }catch(err){ console.warn('generateEnvelope エラー', err); return []; }
  }

  // 原データからループ（反転点）を検出し、各ループの最大荷重点のγを返す
  function detectLoopPeakGammas(data, side){
    try{
      if(!Array.isArray(data) || data.length < 2) return [];
      // generateEnvelope と同じ片側フィルタ
      let filtered = [];
      if(side === 'positive'){
        filtered = data.filter(pt => Number.isFinite(pt.gamma) && Number.isFinite(pt.Load) && pt.gamma >= 0 && pt.Load >= 0);
      } else {
        filtered = data.filter(pt => Number.isFinite(pt.gamma) && Number.isFinite(pt.Load) && pt.gamma <= 0 && pt.Load <= 0);
      }
      if(filtered.length < 2) return [];
      // 絶対値γで並びは元データ順のまま使用
      const absG = (g)=>Math.abs(g);
      let maxAbsGamma = 0; for(const pt of filtered){ const g = absG(pt.gamma); if(g>maxAbsGamma) maxAbsGamma=g; }
      const gTol = Math.max(1e-9, 0.005 * maxAbsGamma); // 0.5% を反転境界の目安に

      const loopPeaks = [];
      let segStart = 0;
      let lastG = absG(filtered[0].gamma);
      // セグメント内の最大荷重（絶対）、そのγ
      let segMaxAbsLoad = Math.abs(filtered[0].Load);
      let segMaxGamma = absG(filtered[0].gamma);

      for(let i=1; i<filtered.length; i++){
        const g = absG(filtered[i].gamma);
        const L = Math.abs(filtered[i].Load);
        // セグメント内のピーク更新
        if(L > segMaxAbsLoad){ segMaxAbsLoad = L; segMaxGamma = absG(filtered[i].gamma); }
        // 反転検出: γ が十分に小さく戻る（前回から gTol 以上の減少）
        if(g + gTol < lastG){
          // セグメント確定
          if(i - 1 >= segStart){ loopPeaks.push(segMaxGamma); }
          // 新セグメント初期化
          segStart = i;
          segMaxAbsLoad = L;
          segMaxGamma = g;
        }
        lastG = g;
      }
      // 最終セグメント
      if(filtered.length - 1 >= segStart){ loopPeaks.push(segMaxGamma); }

      // 近接重複を整理（近すぎるγは一つに）
      loopPeaks.sort((a,b)=>a-b);
      const dedup = [];
      for(const g of loopPeaks){
        if(dedup.length === 0 || Math.abs(g - dedup[dedup.length-1]) > gTol){ dedup.push(g); }
      }
      return dedup;
    }catch(err){ console.warn('detectLoopPeakGammas エラー', err); return []; }
  }

  // === JTCCM Metrics Calculation (Sections III, IV, V) ===
  function calculateJTCCMMetrics(envelope, delta_u_max, alpha){
    const results = {};

    // Determine the sign of the envelope (positive or negative side)
    const envelopeSign = envelope[0] && envelope[0].Load < 0 ? -1 : 1;

    // Find provisional global Pmax (used for yielding & Pu derivation; may lie after δu)
    const Pmax_global_pt = envelope.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), envelope[0]);
    const Pmax_global = Math.abs(Pmax_global_pt.Load);

    // Calculate Py using Line Method (Section III.1) with provisional global Pmax
    const Py_result = calculatePy_LineMethod(envelope, Pmax_global);
    results.Py = Py_result.Py;
    results.Py_gamma = Py_result.Py_gamma;
    results.lineI = Py_result.lineI;
    results.lineII = Py_result.lineII;
    results.lineIII = Py_result.lineIII;

  // Calculate Pu and μ using Perfect Elasto-Plastic Model (Section IV)
  const Pu_result = calculatePu_EnergyEquivalent(envelope, results.Py, Pmax_global, delta_u_max);
  Object.assign(results, Pu_result);

    // Override Pmax with value BEFORE ultimate displacement δu per user requirement
    const delta_u = results.delta_u; // from Pu_result
    if(isFinite(delta_u)){
      const prePts = envelope.filter(pt => Math.abs(pt.gamma) <= delta_u + 1e-12); // tolerance
      if(prePts.length){
        const Pmax_pre_pt = prePts.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), prePts[0]);
        results.Pmax_global = Pmax_global; // store original for reference
        results.Pmax = Math.abs(Pmax_pre_pt.Load);
        results.Pmax_gamma = Math.abs(Pmax_pre_pt.gamma);
      }else{
        // Fallback keep global
        results.Pmax_global = Pmax_global;
        results.Pmax = Pmax_global;
        results.Pmax_gamma = Math.abs(Pmax_global_pt.gamma);
      }
    }else{
      results.Pmax_global = Pmax_global;
      results.Pmax = Pmax_global;
      results.Pmax_gamma = Math.abs(Pmax_global_pt.gamma);
    }

    // === Second pass: Restrict Py (and dependent values) to pre-δu segment ===
    try{
      const du1 = results.delta_u;
      if(isFinite(du1) && du1 > 0){
        const envPre = envelope.filter(pt => Math.abs(pt.gamma) <= du1 + 1e-12);
        if(envPre.length >= 3){
          const pmaxPrePt = envPre.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), envPre[0]);
          const pmaxPre = Math.abs(pmaxPrePt.Load);
          // Recompute Py with pre-δu envelope
          const Py_pre = calculatePy_LineMethod(envPre, pmaxPre);
          results.Py = Py_pre.Py;
          results.Py_gamma = Py_pre.Py_gamma;
          results.lineI = Py_pre.lineI;
          results.lineII = Py_pre.lineII;
          results.lineIII = Py_pre.lineIII;

          // Recompute Pu/μ/δv/δu with updated Py
          // 注意: 面積Sを正しくδuまで補間するため、積分対象は元の full envelope を渡す
          const Pu_pre = calculatePu_EnergyEquivalent(envelope, results.Py, pmaxPre, delta_u_max, du1);
          Object.assign(results, Pu_pre);

          // Recompute Pmax with final δu restriction
          const du2 = results.delta_u;
          const envPre2 = envelope.filter(pt => Math.abs(pt.gamma) <= du2 + 1e-12);
          if(envPre2.length){
            const pmaxPrePt2 = envPre2.reduce((max, pt) => (Math.abs(pt.Load) > Math.abs(max.Load) ? pt : max), envPre2[0]);
            results.Pmax_global = Pmax_global;
            results.Pmax = Math.abs(pmaxPrePt2.Load);
            results.Pmax_gamma = Math.abs(pmaxPrePt2.gamma);
          }
        }
      }
    }catch(e){
      console.warn('Second-pass (pre-δu) Py/Pu 再計算に失敗:', e);
    }

    // Calculate P0 (Section V.1) using final results
    const P0_result = calculateP0(results, envelope);
    Object.assign(results, P0_result);

    // Calculate Pa (Section V.2)
    results.Pa = results.P0 * alpha;

    return results;
  }

  // === Py Calculation (Line Method - Section III.1) ===
  function calculatePy_LineMethod(envelope, Pmax){
    const p_max = Pmax;

    // 上昇区間は「Pmaxに到達するまでの区間」と定義（途中に一時的下降があっても切り捨てない）
    const ascendingEnvelope = (function(){
      if(!envelope || envelope.length < 2) return envelope;
      // Pmaxインデックスを取得
      let idxMax = 0; let maxAbs = -Infinity;
      for(let i=0;i<envelope.length;i++){ const a=Math.abs(envelope[i].Load); if(a>maxAbs){ maxAbs=a; idxMax=i; } }
      const slice = envelope.slice(0, idxMax+1);
      return slice.length>=3 ? slice : envelope; // 最低点数確保
    })();

    // 0.1, 0.4, 0.9 Pmax を上昇区間内で線形補間（失敗時は全体包絡線でフォールバック）
    const p01 = findPointAtLoadStrict(ascendingEnvelope, 0.1 * p_max) || findPointAtLoad(envelope, 0.1 * p_max);
    const p04 = findPointAtLoadStrict(ascendingEnvelope, 0.4 * p_max) || findPointAtLoad(envelope, 0.4 * p_max);
    const p09 = findPointAtLoadStrict(ascendingEnvelope, 0.9 * p_max) || findPointAtLoad(envelope, 0.9 * p_max);

    if(!p01 || !p04 || !p09){
      // 極端にデータが不足する場合はエラーを投げず、最後の点を使う（安全なフォールバック）
      const last = envelope[envelope.length-1];
      const PyFallback = Math.abs(last.Load) * 0.6; // 仮値（荷重増加がない極端ケース）
      return { Py: PyFallback, Py_gamma: Math.abs(last.gamma)*0.6, lineI: {slope:0, intercept:PyFallback}, lineII:{slope:0, intercept:PyFallback}, lineIII:{slope:0, intercept:PyFallback} };
    }

    // Use absolute values for gamma as well
    const gamma01 = Math.abs(p01.gamma);
    const gamma04 = Math.abs(p04.gamma);
    const gamma09 = Math.abs(p09.gamma);
    const load01 = Math.abs(p01.Load);
    const load04 = Math.abs(p04.Load);
    const load09 = Math.abs(p09.Load);

    // Line I: 0.1 Pmax - 0.4 Pmax
    const lineI = {
      slope: (load04 - load01) / (gamma04 - gamma01),
      intercept: load01 - ((load04 - load01) / (gamma04 - gamma01)) * gamma01
    };

    // Line II: 0.4 Pmax - 0.9 Pmax
    const lineII = {
      slope: (load09 - load04) / (gamma09 - gamma04),
      intercept: load04 - ((load09 - load04) / (gamma09 - gamma04)) * gamma04
    };

    // Line III: Parallel to Line II, tangent to envelope
    const lineIII = findTangentLine(envelope, lineII.slope);

    // Intersection of Line I and Line III
    const gamma_py = (lineIII.intercept - lineI.intercept) / (lineI.slope - lineIII.slope);
    const Py = lineI.slope * gamma_py + lineI.intercept;

    return { Py, Py_gamma: gamma_py, lineI, lineII, lineIII };
  }

  function findPointAtLoad(envelope, targetLoad){
    for(let i=0; i<envelope.length-1; i++){
      const p1 = envelope[i];
      const p2 = envelope[i+1];
      const abs1 = Math.abs(p1.Load);
      const abs2 = Math.abs(p2.Load);
      
      if(abs1 <= targetLoad && abs2 >= targetLoad){
        const ratio = (targetLoad - abs1) / (abs2 - abs1);
        return {
          Load: p1.Load + (p2.Load - p1.Load) * ratio,
          gamma: p1.gamma + (p2.gamma - p1.gamma) * ratio,
          gamma0: p1.gamma0 + (p2.gamma0 - p1.gamma0) * ratio
        };
      }
    }
    return envelope[envelope.length - 1]; // Fallback
  }

  // 上昇区間に限定して targetLoad を線形補間（失敗時は null を返す）
  function findPointAtLoadStrict(envelopeAsc, targetLoad){
    for(let i=0; i<envelopeAsc.length-1; i++){
      const p1 = envelopeAsc[i];
      const p2 = envelopeAsc[i+1];
      const abs1 = Math.abs(p1.Load);
      const abs2 = Math.abs(p2.Load);
      if(abs1 <= targetLoad + 1e-12 && abs2 >= targetLoad - 1e-12){
        const denom = (abs2 - abs1);
        const ratio = denom !== 0 ? (targetLoad - abs1) / denom : 0;
        const clamped = Math.max(0, Math.min(1, ratio));
        return {
          Load: p1.Load + (p2.Load - p1.Load) * clamped,
          gamma: p1.gamma + (p2.gamma - p1.gamma) * clamped,
          gamma0: p1.gamma0 + (p2.gamma0 - p1.gamma0) * clamped
        };
      }
    }
    return null;
  }

  function findLoadAtGamma(envelope, targetGamma){
    // γ位置での荷重を線形補間で取得
    const absTarget = Math.abs(targetGamma);
    for(let i=0; i<envelope.length-1; i++){
      const p1 = envelope[i];
      const p2 = envelope[i+1];
      const g1 = Math.abs(p1.gamma);
      const g2 = Math.abs(p2.gamma);
      
      if(g1 <= absTarget && g2 >= absTarget){
        const ratio = (absTarget - g1) / (g2 - g1);
        return Math.abs(p1.Load) + (Math.abs(p2.Load) - Math.abs(p1.Load)) * ratio;
      }
    }
    return Math.abs(envelope[envelope.length - 1].Load); // Fallback
  }

  function findTangentLine(envelope, slope){
    // Find the point where (|Load| - slope * |gamma|) is maximum
    let maxIntercept = -Infinity;
    for(const pt of envelope){
      const intercept = Math.abs(pt.Load) - slope * Math.abs(pt.gamma);
      if(intercept > maxIntercept) maxIntercept = intercept;
    }
    return { slope, intercept: maxIntercept };
  }

  // === Pu and μ Calculation (Energy Equivalent - Section IV) ===
  function calculatePu_EnergyEquivalent(envelope, Py, Pmax, delta_u_max, fixed_delta_u){
    // δy は包絡線とLineIV(Py水平線)の交点の変形
    // 包絡線上で |Load| が Py に最も近い点、または線形補間でPyを横切る点のγ
    let delta_y = 0;
    try{
      // 包絡線を原点から走査し、Pyを跨ぐ区間を線形補間
      let found = false;
      for(let i=0; i<envelope.length-1; i++){
        const p1 = envelope[i];
        const p2 = envelope[i+1];
        const L1 = Math.abs(p1.Load);
        const L2 = Math.abs(p2.Load);
        // Pyを跨ぐ区間を検出
        if((L1 <= Py && L2 >= Py) || (L1 >= Py && L2 <= Py)){
          const denom = (L2 - L1);
          const ratio = denom !== 0 ? (Py - L1) / denom : 0;
          const clamped = Math.max(0, Math.min(1, ratio));
          const g1 = Math.abs(p1.gamma);
          const g2 = Math.abs(p2.gamma);
          delta_y = g1 + (g2 - g1) * clamped;
          found = true;
          break;
        }
      }
      if(!found){
        // 見つからない場合は最も近い点を採用（フォールバック）
        let minDiff = Infinity; let bestGamma = 0;
        for(const pt of envelope){
          const diff = Math.abs(Math.abs(pt.Load) - Py);
          if(diff < minDiff){ minDiff = diff; bestGamma = Math.abs(pt.gamma); }
        }
        delta_y = bestGamma;
      }
    }catch(_){
      // 極端なエラー時は従来ロジック
      const pt_y = findPointAtLoad(envelope, Py);
      delta_y = Math.abs(pt_y.gamma);
    }

    // Initial stiffness K
    const K = Py / delta_y;

    // Find δu (Section IV.1 Step 9)
    let delta_u;
    if(Number.isFinite(fixed_delta_u)){
      // 既定の（補間済み）δuを固定使用
      delta_u = Math.min(Math.abs(fixed_delta_u), Math.abs(delta_u_max));
    }else{
      const delta_u_candidate1 = findDeltaU_08Pmax(envelope, Pmax);
      const delta_u_candidate2 = delta_u_max; // rad from user input
      delta_u = Math.min(delta_u_candidate1, delta_u_candidate2);
    }

    // Calculate area S under envelope up to δu
  const S = calculateAreaUnderEnvelope(envelope, delta_u);
  // 終局変位位置での包絡線荷重（参考値: 終局時実荷重）
  const load_at_delta_u = findLoadAtGamma(envelope, delta_u);

    // Solve for Pu using energy equivalence (Section IV.1 Step 11-12)
    // S = Pu * (δu - δv/2), where δv = Pu/K
    // S = Pu * δu - Pu²/(2K)
    // Pu²/(2K) - Pu*δu + S = 0
    // Pu = K*δu - sqrt((K*δu)² - 2*K*S)
    const discriminant = Math.pow(K * delta_u, 2) - 2 * K * S;
    if(discriminant < 0){
      console.warn('Pu計算で判別式が負: discriminant =', discriminant);
      appendLog('警告: Pu計算 判別式<0 のためPyにフォールバック (discriminant='+discriminant.toFixed(6)+')');
      // Fallback: use Py
      return {
        delta_y, K, delta_u, S,
        Pu: Py,
        delta_v: delta_y,
        mu: delta_u / delta_y,
        lineV: {start: {gamma:0, Load:0}, end: {gamma: delta_y, Load: Py}},
        lineVI: {gamma_start: delta_y, gamma_end: delta_u, Load: Py}
      };
    }

    const Pu = K * delta_u - Math.sqrt(discriminant);
    const delta_v = Pu / K;
    const mu = delta_u / delta_v;

    // Lines for visualization
    const lineV = {start: {gamma:0, Load:0}, end: {gamma: delta_v, Load: Pu}};
    const lineVI = {gamma_start: delta_v, gamma_end: delta_u, Load: Pu};

    return { delta_y, K, delta_u, S, Pu, delta_v, mu, lineV, lineVI, load_at_delta_u };
  }

  function findDeltaU_08Pmax(envelope, Pmax){
    if(!envelope || envelope.length < 2 || !Number.isFinite(Pmax)){
      return 0;
    }
    const threshold = 0.8 * Math.abs(Pmax);

    // Pmax（|Load|最大）のインデックスを取得
    let idxMax = 0;
    let maxAbsLoad = -Infinity;
    for(let i=0; i<envelope.length; i++){
      const a = Math.abs(envelope[i].Load);
      if(a > maxAbsLoad){ maxAbsLoad = a; idxMax = i; }
    }

    // Pmax以降の区間で |Load| が閾値を上から下へ横切る点を線形補間で求める
    for(let i=idxMax; i<envelope.length-1; i++){
      const p1 = envelope[i];
      const p2 = envelope[i+1];
      const y1 = Math.abs(p1.Load);
      const y2 = Math.abs(p2.Load);
      if(y1 >= threshold && y2 <= threshold){
        const dy = (y2 - y1);
        const t = dy !== 0 ? (threshold - y1) / dy : 0; // 0..1 にクリップ
        const tClamped = Math.max(0, Math.min(1, t));
        const g1 = Math.abs(p1.gamma);
        const g2 = Math.abs(p2.gamma);
        return g1 + (g2 - g1) * tClamped;
      }
    }

    // 横切りが見つからない場合は、終点の |γ| を返す（上限扱い）
    return Math.abs(envelope[envelope.length - 1].gamma);
  }

  function calculateAreaUnderEnvelope(envelope, delta_limit){
    try{
      if(!envelope || envelope.length === 0 || !Number.isFinite(delta_limit) || delta_limit <= 0){
        return 0;
      }
      // 0点(0,0)を含む区分台形積分に変更し、先頭三角形の欠落を防止
      const pts = [];
      pts.push({ g: 0, p: 0 });
      for(const pt of envelope){
        const g = Math.abs(pt.gamma);
        const p = Math.abs(pt.Load);
        // 同一点が続く場合はスキップ
        if(pts.length && Math.abs(pts[pts.length-1].g - g) < 1e-15 && Math.abs(pts[pts.length-1].p - p) < 1e-15){
          continue;
        }
        // g が減少するデータはスキップ（生成済み包絡線は単調増加のはずだが念のため）
        if(g < pts[pts.length-1].g - 1e-15){
          continue;
        }
        pts.push({ g, p });
      }

      let area = 0;
      for(let i=0; i<pts.length-1; i++){
        const g1 = pts[i].g;
        const p1 = pts[i].p;
        const g2 = pts[i+1].g;
        const p2 = pts[i+1].p;

        if(g1 >= delta_limit){
          break;
        }

        const upper = Math.min(g2, delta_limit);
        const width = upper - g1;
        if(width <= 0){
          continue;
        }
        let pUpper = p2;
        if(upper < g2 - 1e-15){
          // 線形補間で終端荷重を算出
          const t = (upper - g1) / (g2 - g1);
          pUpper = p1 + (p2 - p1) * t;
        }
        // 台形公式
        area += width * (p1 + pUpper) / 2;
        if(upper >= delta_limit - 1e-15){
          break;
        }
      }
      return area;
    }catch(_){
      // 失敗時は従来ロジックにフォールバック（保険）
      let area = 0; let prev = null;
      for(const pt of envelope){
        const absGamma = Math.abs(pt.gamma);
        if(absGamma > delta_limit){
          if(prev && Math.abs(prev.gamma) < delta_limit){
            const ratio = (delta_limit - Math.abs(prev.gamma)) / (absGamma - Math.abs(prev.gamma));
            const load_at_limit = Math.abs(prev.Load) + (Math.abs(pt.Load) - Math.abs(prev.Load)) * ratio;
            area += (delta_limit - Math.abs(prev.gamma)) * (Math.abs(prev.Load) + load_at_limit) / 2;
          }
          break;
        }
        if(prev){
          const dg = absGamma - Math.abs(prev.gamma);
          const avg_load = (Math.abs(prev.Load) + Math.abs(pt.Load)) / 2;
          area += dg * avg_load;
        }
        prev = pt;
      }
      return area;
    }
  }

  // === P0 Calculation (Section V.1) ===
  function calculateP0(results, envelope){
    const { Py, Pmax } = results;

    // (a) Yield strength
    const p0_a = Py;

  // (b) Max strength criterion (2/3 Pmax)
  const p0_b = Pmax * (2/3);

  const P0 = Math.min(p0_a, p0_b);

  return { p0_a, p0_b, P0 };
  }

  function findPointAtGamma(envelope, targetGamma, key){
    for(let i=0; i<envelope.length-1; i++){
      const p1 = envelope[i];
      const p2 = envelope[i+1];
      const abs1 = Math.abs(p1[key]);
      const abs2 = Math.abs(p2[key]);
      
      if(abs1 <= targetGamma && abs2 >= targetGamma){
        const ratio = (targetGamma - abs1) / (abs2 - abs1);
        return {
          Load: Math.abs(p1.Load) + (Math.abs(p2.Load) - Math.abs(p1.Load)) * ratio,
          gamma: p1.gamma + (p2.gamma - p1.gamma) * ratio,
          gamma0: p1.gamma0 + (p2.gamma0 - p1.gamma0) * ratio
        };
      }
    }
    return envelope[envelope.length - 1];
  }

  // === Rendering ===
  function renderPlot(envelope, results){
    const { Pmax, Py, Py_gamma, lineI, lineII, lineIII, lineV, lineVI, delta_u, delta_v, p0_a, p0_b } = results;

  // Draw evaluation overlays on the selected side explicitly
  const envelopeSign = (envelope_side && envelope_side.value === 'negative') ? -1 : 1;

  // Calculate data range for auto-fitting based on envelope (not raw data)
      // 現在の範囲を保持するか、新規計算するか
      let ranges;
      const isDialogOpen = pointEditDialog && pointEditDialog.style.display !== 'none';
      if(isDialogOpen && plotDiv && plotDiv._fullLayout && plotDiv._fullLayout.xaxis && plotDiv._fullLayout.yaxis){
        // ポップアップ表示中は既存の範囲を保持
        ranges = {
          xRange: [plotDiv._fullLayout.xaxis.range[0], plotDiv._fullLayout.xaxis.range[1]],
          yRange: [plotDiv._fullLayout.yaxis.range[0], plotDiv._fullLayout.yaxis.range[1]]
        };
        console.debug('[renderPlot] ポップアップ表示中 - 描画範囲を保持:', ranges);
      } else {
        // 新規計算
        ranges = computeEnvelopeRanges(envelope);
        console.debug('[renderPlot] 描画範囲を新規計算:', ranges);
      }

    // レンジの健全性チェック（NaN / Infinity を排除）
    function sanitizeRange(arr, defMin, defMax){
      const a0 = Array.isArray(arr) ? arr[0] : undefined;
      const a1 = Array.isArray(arr) ? arr[1] : undefined;
      let minV = Number.isFinite(a0) ? a0 : defMin;
      let maxV = Number.isFinite(a1) ? a1 : defMax;
      if(!Number.isFinite(minV)) minV = -1;
      if(!Number.isFinite(maxV)) maxV = 1;
      if(minV === maxV){ maxV = minV + 1; }
      if(minV > maxV){ const t = minV; minV = maxV; maxV = t; }
      return [minV, maxV];
    }
    const xRangeSafe = sanitizeRange(ranges && ranges.xRange, -1, 1);
    const yRangeSafe = sanitizeRange(ranges && ranges.yRange, -1, 1);
    
    // 包絡線データを編集可能にするための状態管理
    let editableEnvelope = envelope.map(pt => ({...pt}));    // Original raw data (all points) - showing positive and negative loads
    const trace_rawdata = {
      x: rawData.map(pt => pt.gamma), // rad
      y: rawData.map(pt => pt.Load), // Keep original sign
      mode: 'lines+markers',
      name: '実験データ',
      line: {color: 'lightblue', width: 1},
      marker: {color: 'lightblue', size: 4},
      hoverinfo: 'skip' // 試験データのホバーを無効化し、包絡線点を優先
    };

    // Envelope line - keep original sign
    const trace_env = {
      x: editableEnvelope.map(pt => pt.gamma),
      y: editableEnvelope.map(pt => pt.Load), // Keep original sign from filtered data
      mode: 'lines',
      name: '包絡線',
      line: {color: 'blue', width: 2},
      hoverinfo: 'skip' // 包絡線ラインのホバーを無効化
    };

    // Envelope points (editable) - click to edit
    const trace_env_points = {
      x: editableEnvelope.map(pt => pt.gamma),
      y: editableEnvelope.map(pt => pt.Load),
      mode: 'markers',
      name: '包絡線点',
      marker: {
        color: editableEnvelope.map((pt, idx) => idx === (window._selectedEnvelopePoint || -1) ? 'red' : 'blue'),
        size: editableEnvelope.map((pt, idx) => idx === (window._selectedEnvelopePoint || -1) ? 14 : 10),
        symbol: 'circle', 
        line: {color: 'white', width: 2}
      },
      hovertemplate: '<b>変形:</b> %{x:.6f}<br><b>荷重:</b> %{y:.3f}<br><i>クリックで編集、Delキーで削除</i><extra></extra>'
    };

    // Line I, II, III (Py determination)
    const gamma_range = [0, Math.max(...envelope.map(pt => Math.abs(pt.gamma)))]
    ;
    const trace_lineI = makeLine(lineI, gamma_range, 'Line I (0.1-0.4Pmax)', 'orange', envelopeSign);
    const trace_lineII = makeLine(lineII, gamma_range, 'Line II (0.4-0.9Pmax)', 'darkorange', envelopeSign);
    const trace_lineIII = makeLine(lineIII, gamma_range, 'Line III (接線)', 'red', envelopeSign);
    // Line IV: Py を通る水平線（X軸と平行）
    const lineIV_x = [0, Math.max(...envelope.map(pt => Math.abs(pt.gamma))) * envelopeSign];
    const lineIV_y = [Py * envelopeSign, Py * envelopeSign];
    const trace_lineIV = {
      x: lineIV_x,
      y: lineIV_y,
      mode: 'lines',
      name: 'Line IV (Py水平)',
      line: {color: 'green', width: 1, dash: 'dot'}
    };

    // Py point
    const trace_py = {
      x: [Py_gamma * envelopeSign],
      y: [Py * envelopeSign],
      mode: 'markers',
      name: 'Py (降伏耐力)',
      marker: {color: 'green', size: 12, symbol: 'circle'}
    };

    // Perfect elasto-plastic model (Line V, VI)
    const trace_lineV = {
      x: [0, lineV.end.gamma * envelopeSign],
      y: [0, lineV.end.Load * envelopeSign],
      mode: 'lines',
      name: 'Line V (初期剛性)',
      line: {color: 'purple', width: 2, dash: 'dash'}
    };

    const trace_lineVI = {
      x: [lineVI.gamma_start * envelopeSign, lineVI.gamma_end * envelopeSign],
      y: [lineVI.Load * envelopeSign, lineVI.Load * envelopeSign],
      mode: 'lines',
      name: 'Line VI (Pu)',
      line: {color: 'purple', width: 2, dash: 'dash'}
    };

    // Pmax
    const trace_pmax = {
      x: [results.Pmax_gamma * envelopeSign],
      y: [Pmax * envelopeSign],
      mode: 'markers',
      name: 'Pmax',
      marker: {color: 'red', size: 12, symbol: 'star'}
    };

    // P0 criteria lines
  let gamma_max = Math.max(...envelope.map(pt => Math.abs(pt.gamma)));
    if(!Number.isFinite(gamma_max) || gamma_max <= 0){
      // 範囲から安全な最大を推定
      gamma_max = Math.max(Math.abs(xRangeSafe[0]), Math.abs(xRangeSafe[1]));
      if(!Number.isFinite(gamma_max) || gamma_max <= 0) gamma_max = 1;
    }
  const trace_p0_lines = {
      x: [0, gamma_max * envelopeSign, NaN, 0, gamma_max * envelopeSign],
  y: [p0_a * envelopeSign, p0_a * envelopeSign, NaN, p0_b * envelopeSign, p0_b * envelopeSign],
      mode: 'lines',
      name: 'P0基準 (a,b)',
      line: {color: 'gray', width: 1, dash: 'dot'}
    };

    // δu 縦補助線（終局変位角）: layout.shapes で描画（常に全高に伸ばす）
    const shapes = [];
    if(Number.isFinite(delta_u)){
      const x_du = delta_u * envelopeSign;
      shapes.push({
        type: 'line',
        xref: 'x',
        yref: 'paper',
        x0: x_du,
        x1: x_du,
        y0: 0,
        y1: 1,
        line: {color: 'purple', width: 1.5, dash: 'dot'}
      });
    }

    const layout = {
      title: '荷重-変形関係と評価直線',
      xaxis: {
        title: '変形 δ (mm)',
        range: xRangeSafe,
        autorange: false
      },
      yaxis: {
        title: '荷重 P (kN)',
        range: yRangeSafe,
        autorange: false
      },
      hovermode: 'closest',
      dragmode: 'pan', // デフォルトはパン操作
      showlegend: true,
      legend: {
        orientation: 'h',
        x: 0.5,
        xanchor: 'center',
        y: -0.15,
        yanchor: 'top'
      },
      height: 600,
      uirevision: 'fixed',
      shapes: shapes,
      annotations: (function(){
        const annMode = show_annotations ? show_annotations.value : 'all';
        const baseAnnotations = [
        // 終局変位 δu (mm) → Line VI の終点（delta_u の位置）に表示
        {
          x: (lineVI.gamma_end) * envelopeSign,
          y: (lineVI.Load) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `δu=${delta_u.toFixed(2)} mm`,
          showarrow: true,
          ax: 20, ay: -20,
          font: {size: 12, color: 'purple'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'purple', borderwidth: 1
        },
        
        // 降伏変位 δy (mm) → 包絡線とLineIV交点に表示
        {
          x: (results.delta_y) * envelopeSign,
          y: (Py) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `δy=${results.delta_y.toFixed(2)} mm`,
          showarrow: true,
          ax: 20, ay: 20,
          font: {size: 12, color: 'green'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'green', borderwidth: 1
        },
        // 降伏耐力 Py (kN) → Line I/IIIの交点(Py_gamma)に表示
        {
          x: (Py_gamma) * envelopeSign,
          y: (Py) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `Py=${Py.toFixed(1)} kN`,
          showarrow: true,
          ax: 20, ay: -40,
          font: {size: 12, color: 'green'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'green', borderwidth: 1
        },
        // 終局耐力 Pu (kN) → Line V の終点（delta_v の位置）に表示
        {
          x: (lineV.end.gamma) * envelopeSign,
          y: (lineV.end.Load) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `Pu=${(lineVI.Load).toFixed(1)} kN`,
          showarrow: true,
          ax: 20, ay: -20,
          font: {size: 12, color: 'purple'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'purple', borderwidth: 1
        },
        // 降伏点変位 δv (mm) → Line V の終点（delta_v の位置）に表示
        {
          x: (lineV.end.gamma) * envelopeSign,
          y: (lineV.end.Load) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `δv=${delta_v.toFixed(2)} mm`,
          showarrow: true,
          ax: -30, ay: 20,
          font: {size: 12, color: 'purple'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'purple', borderwidth: 1
        },
        // 最大耐力 Pmax (kN) → Pmax点に表示
        {
          x: (results.Pmax_gamma) * envelopeSign,
          y: (Pmax) * envelopeSign,
          xref: 'x', yref: 'y',
          text: `Pmax=${Pmax.toFixed(1)} kN`,
          showarrow: true,
          ax: 20, ay: -20,
          font: {size: 12, color: 'red'},
          bgcolor: 'rgba(255,255,255,0.7)',
          bordercolor: 'red', borderwidth: 1
        }
        ];

        // P0基準注釈を重ならないよう縦方向にスタック配置
        const p0Values = [
          {label: 'P0(a)=Py', value: p0_a},
          {label: 'P0(b)=2/3·Pmax', value: p0_b}
        ];
        // y座標順でソート
        p0Values.sort((a,b) => b.value - a.value);
        
        // 重なり検出と縦オフセット適用
        const yRangePx = 600; // グラフ高さ（px）
        const yDataRange = Math.abs(yRangeSafe[1] - yRangeSafe[0]);
        const minGapPx = 18; // 最小間隔（px）
        const minGapData = (minGapPx / yRangePx) * yDataRange;
        
        let prevY = Infinity;
        const p0Annotations = [];
        for(let i=0; i<p0Values.length; i++){
          let yPos = p0Values[i].value;
          if(prevY - yPos < minGapData){
            yPos = prevY - minGapData; // 前の注釈から一定間隔空ける
          }
          prevY = yPos;
          p0Annotations.push({
            x: gamma_max * 0.05 * envelopeSign,
            y: yPos * envelopeSign,
            xref: 'x', yref: 'y',
            text: `${p0Values[i].label}=${p0Values[i].value.toFixed(2)} kN`,
            showarrow: false,
            xanchor: 'left',
            yanchor: 'middle',
            font: {size: 10, color: 'gray'},
            bgcolor: 'rgba(255,255,255,0.85)',
            bordercolor: 'gray', borderwidth: 0.5
          });
        }

        // 注釈フィルタリング
        if(annMode === 'none') return [];
        if(annMode === 'p0only') return p0Annotations; // P0(a)～(b)のみ
        if(annMode === 'main') return baseAnnotations; // 主要のみ(P0除外)
        return baseAnnotations.concat(p0Annotations); // 'all' or default
      })()
    };

  const plotConfig = {
    editable: false,
    displayModeBar: true,
    // Box select / Lasso select を有効化
    displaylogo: false,
    // デフォルトのAutoscale/Resetを削除（全データへのフィットを防止）
    modeBarButtonsToRemove: ['autoScale2d', 'resetScale2d'],
    // 包絡線範囲へのフィット専用ボタンを追加
    modeBarButtonsToAdd: [
      {
        name: '包絡線にフィット',
        icon: (Plotly && Plotly.Icons && Plotly.Icons.autoscale) ? Plotly.Icons.autoscale : undefined,
        click: function(gd){
          if(envelopeData && envelopeData.length){
            fitEnvelopeRanges('モードバー');
          }
        }
      }
    ]
  };

  Plotly.newPlot(plotDiv, [
      trace_rawdata,
      trace_env,
      trace_env_points,
      trace_lineI,
  trace_lineII,
  trace_lineIII,
  trace_lineIV,
      trace_py,
      trace_lineV,
      trace_lineVI,
      trace_pmax,
      trace_p0_lines
    ], layout, plotConfig)
    .then(function(){
      // 包絡線点の編集機能を実装
      setupEnvelopeEditing(editableEnvelope);
      
      // Autoscale（モードバーやダブルクリック）が発火した場合も包絡線範囲へ調整
      if(!relayoutHandlerAttached){
        plotDiv.on('plotly_relayout', function(e){
          try{
            if(pointEditDialog && pointEditDialog.style.display !== 'none') return;
            if(!e) return;
            // 何らかの理由でautorangeがtrueになった場合、即キャッシュ適用
            if((e['xaxis.autorange'] === true || e['yaxis.autorange'] === true) && cachedEnvelopeRange){
              requestAnimationFrame(()=>{
                safeRelayout(plotDiv, {
                  'xaxis.autorange': false,
                  'yaxis.autorange': false,
                  'xaxis.range': cachedEnvelopeRange.xRange,
                  'yaxis.range': cachedEnvelopeRange.yRange
                });
              });
            }
          }catch(err){ console.warn('autoscale再調整エラー', err); }
        });
        // ダブルクリックのリセットでも同様にフィット
        plotDiv.on('plotly_doubleclick', function(){
          try{
            // ポップアップ表示中はダブルクリックリセットをスキップ
            if(pointEditDialog && pointEditDialog.style.display !== 'none'){
              console.debug('[ダブルクリック] ポップアップ表示中のためスキップ');
              return false;
            }
            if(envelopeData && envelopeData.length){
              fitEnvelopeRanges('ダブルクリック');
            }
          }catch(err){ console.warn('doubleclick再調整エラー', err); }
          return false; // 既存のデフォルト動作抑制
        });
        relayoutHandlerAttached = true;
      }
    })
    .catch(function(err){
      // 初期 newPlot の undefined 拒否などを握りつぶしてノイズ抑制
      console.info('[plot.init suppressed]', err);
    });
  }

  // === 包絡線点の編集機能 ===
  function setupEnvelopeEditing(editableEnvelope){
    let isDragging = false;
    let dragPointIndex = -1;
    let selectedPointIndex = -1; // Del キー用の選択状態
    let selectedPoints = []; // Box/Lasso select用の複数選択状態
    // window._selectedEnvelopePoint の初期化をコメントアウト（既存の選択状態を保持）
    if(typeof window._selectedEnvelopePoint === 'undefined'){
      window._selectedEnvelopePoint = -1; // 初回のみ初期化
    }
    
    // 既存クリックハンドラを解除
    if(_plotClickHandler && typeof plotDiv.removeListener === 'function'){
      plotDiv.removeListener('plotly_click', _plotClickHandler);
    }
    // 包絡線点のクリック処理（クリックで即座に数値編集ダイアログを開く）
    _plotClickHandler = function(data){
      console.debug('[plotly_click] event points=', data && data.points ? data.points.length : 0);
      if(!data.points || data.points.length === 0) return;
      const pt = data.points[0];
      console.debug('[plotly_click] curveNumber='+pt.curveNumber+' pointIndex='+pt.pointIndex);
      if(pt.curveNumber === 2){
        selectedPointIndex = pt.pointIndex;
        window._selectedEnvelopePoint = pt.pointIndex;
        try{ window._selectedEnvelopePoints = []; }catch(_){ /* noop */ }
        // 視覚的に選択反映
        highlightSelectedPoint(editableEnvelope);
        
        // 解析未実行時のフォールバック: envelopeData が存在しなければ自動解析
        if(!envelopeData && rawData && rawData.length >= 3){
          console.info('[plotly_click] 解析前クリック検出 → 自動解析実行');
          processDataDirect(); // 自動解析
          // 解析後に再選択してダイアログ開く（非同期対応）
          setTimeout(function(){
            if(envelopeData && window._selectedEnvelopePoint >= 0){
              openPointEditDialog();
            }
          }, 100);
        } else {
          // ダイアログを開く
          openPointEditDialog();
        }
        console.debug('[plotly_click] ダイアログ表示要求');
        return;
      }
      // 他のトレースをクリック：選択解除
      selectedPointIndex = -1;
      window._selectedEnvelopePoint = -1;
      highlightSelectedPoint(editableEnvelope);
    };
    plotDiv.on('plotly_click', _plotClickHandler);
    
    // Box Select / Lasso Select で複数点選択を検出
    plotDiv.on('plotly_selected', function(eventData){
      if(!eventData || !eventData.points) {
        selectedPoints = [];
        try{ window._selectedEnvelopePoints = []; }catch(_){ /* noop */ }
        return;
      }
      // 包絡線点トレース（curveNumber === 2）のみを抽出
      selectedPoints = eventData.points
        .filter(pt => pt.curveNumber === 2)
        .map(pt => pt.pointIndex);
      console.debug('[plotly_selected] 選択された包絡線点:', selectedPoints);
      try{ window._selectedEnvelopePoints = selectedPoints.slice(); }catch(_){ /* noop */ }
      // 範囲選択ONなら、選択完了後もselectモード維持
      if(rangeSelectEnabled){
        try{ if(window.Plotly && plotDiv){ safeRelayout(plotDiv, {'dragmode':'select'}); } }catch(_){/* noop */}
      }
    });
    
    // 選択解除
    plotDiv.on('plotly_deselect', function(){
      selectedPoints = [];
      try{ window._selectedEnvelopePoints = []; }catch(_){ /* noop */ }
      console.debug('[plotly_deselect] 選択解除');
      if(rangeSelectEnabled){
        try{ if(window.Plotly && plotDiv){ safeRelayout(plotDiv, {'dragmode':'select'}); } }catch(_){/* noop */}
      }
    });
    
    // ダブルクリックによる点追加やデフォルト操作は許容（別処理はしない）
    
    // ダブルクリックによる新規点追加機能は廃止（仕様変更）
    
    // Delキーで選択中の点を削除
    // 既存のキーリスナーを解除
    if(_keydownHandler){
      document.removeEventListener('keydown', _keydownHandler);
    }
    const handleKeydown = function(e){
      // Delキーで選択中の点を削除（単一点 or 複数点）
      if(e.key === 'Delete' || e.key === 'Del'){
        // Box/Lasso selectで複数選択がある場合
        if(selectedPoints.length > 0){
          // 最小2点は残すためのチェック
          const remainingCount = editableEnvelope.length - selectedPoints.length;
          if(remainingCount < 2){
            alert('包絡線には最低2点が必要です。削除できません。');
            return;
          }
          // 降順ソートして後ろから削除（インデックスずれ防止）
          const sortedIndices = [...selectedPoints].sort((a,b) => b - a);
          sortedIndices.forEach(idx => {
            if(idx >= 0 && idx < editableEnvelope.length){
              editableEnvelope.splice(idx, 1);
            }
          });
          pushHistory(editableEnvelope);
          appendLog(`包絡線点 ${selectedPoints.length}個 を一括削除しました`);
          selectedPoints = [];
          try{ window._selectedEnvelopePoints = []; }catch(_){ /* noop */ }
          selectedPointIndex = -1;
          window._selectedEnvelopePoint = -1;
          recalculateFromEnvelope(editableEnvelope);
          return;
        }
        // 単一点選択の場合（既存の動作）
        if(selectedPointIndex >= 0 && selectedPointIndex < editableEnvelope.length){
          deleteEnvelopePoint(selectedPointIndex, editableEnvelope);
          selectedPointIndex = -1;
          window._selectedEnvelopePoint = -1;
        }
        return;
      }
      // Undo/Redo ショートカット
      if((e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z')){
        e.preventDefault();
        performUndo();
        return;
      }
      if((e.ctrlKey || e.metaKey) && (e.key === 'y' || e.key === 'Y')){
        e.preventDefault();
        performRedo();
        return;
      }
      // 'E' で数値編集ダイアログ
      if(!e.ctrlKey && !e.metaKey && (e.key === 'e' || e.key === 'E')){
        e.preventDefault();
        openPointEditDialog();
        return;
      }
    };
    _keydownHandler = handleKeydown;
    document.addEventListener('keydown', _keydownHandler);
    
  // ドラッグ移動状態
  let shiftDragging = false; // 互換のため名称は維持（実際はトグルでも使用）
  let shiftDragIndex = -1;
  let shiftDragStartX = 0;
  let shiftDragStartY = 0;
    
    plotDiv.on('plotly_hover', function(data){
      if(!data.points || data.points.length === 0) return;
      const pt = data.points[0];
      // 包絡線点（curveNumber === 2）にホバー時、カーソルをポインタに
      if(pt.curveNumber === 2){
        // ドラッグ移動可能条件: トグルON かつ "選択中の点" の上にホバーしている場合のみ hand 表示
        const isSelectedHovered = (typeof window._selectedEnvelopePoint === 'number' && pt.pointIndex === window._selectedEnvelopePoint);
        if(dragMoveEnabled && isSelectedHovered){
          plotDiv.classList.add('drag-point-hover');
          plotDiv.classList.remove('pointer-mode');
        } else {
          plotDiv.classList.add('pointer-mode');
          plotDiv.classList.remove('drag-point-hover');
        }
      } else {
        plotDiv.classList.remove('drag-point-hover');
        plotDiv.classList.remove('pointer-mode');
      }
    });
    
    plotDiv.on('plotly_unhover', function(){
      if(!shiftDragging){
        // ホバー離脱時はクラスをクリア
        plotDiv.classList.remove('drag-point-hover');
        plotDiv.classList.remove('pointer-mode');
      }
    });
    
  // マウスダウン: トグルONかつ包絡線点上ならドラッグ開始（Ctrl+Shift 操作は廃止）
    let mousedownHandler = function(e){
      if(!dragMoveEnabled){
        return; // ドラッグ移動モードがOFFなら何もしない
      }
      
      // 選択中の点が存在しない場合はドラッグ不可
      if(window._selectedEnvelopePoint < 0 || !editableEnvelope || window._selectedEnvelopePoint >= editableEnvelope.length){
        return;
      }
      
      // Plotlyのイベントから選択点の座標を取得
      const xaxis = plotDiv._fullLayout.xaxis;
      const yaxis = plotDiv._fullLayout.yaxis;
      if(!xaxis || !yaxis) return;
      
      const bbox = (dragLayer || plotDiv).getBoundingClientRect();
      const clickX = e.clientX - bbox.left;
      const clickY = e.clientY - bbox.top;
      
      // 選択点のピクセル座標を計算
      const selectedIdx = window._selectedEnvelopePoint;
      const selectedPt = editableEnvelope[selectedIdx];
      const px = xaxis.l2p(selectedPt.gamma);
      const py = yaxis.l2p(selectedPt.Load);
      const dist = Math.sqrt((clickX - px)**2 + (clickY - py)**2);
      
      // 35px以内なら選択点と判定（ヒット領域）
  if(dist < 35){
        // Plotlyのデフォルトドラッグを無効化（先に実行）
        e.stopImmediatePropagation();
        e.preventDefault();
        
        shiftDragging = true;
        shiftDragIndex = selectedIdx;
        shiftDragStartX = clickX;
        shiftDragStartY = clickY;
  plotDiv.classList.add('drag-point-active');
  plotDiv.classList.remove('drag-point-hover');
        
        // Plotlyデフォルトのズーム/パンはイベント抑止で無効化（dragmodeの変更は行わない）
        
        // ツールチップ表示
        if(pointTooltip){
          pointTooltip.textContent = `γ: ${selectedPt.gamma.toFixed(6)}, P: ${selectedPt.Load.toFixed(3)}`;
          pointTooltip.style.left = e.clientX + 10 + 'px';
          pointTooltip.style.top = e.clientY + 10 + 'px';
          pointTooltip.style.display = 'block';
        }
      }
    };
    
    // マウスムーブ: ドラッグ中なら座標更新
    let mousemoveHandler = function(e){
      if(!shiftDragging || shiftDragIndex < 0) return;
      
      const xaxis = plotDiv._fullLayout.xaxis;
      const yaxis = plotDiv._fullLayout.yaxis;
      if(!xaxis || !yaxis) return;
      
      const bbox = (dragLayer || plotDiv).getBoundingClientRect();
      const moveX = e.clientX - bbox.left;
      const moveY = e.clientY - bbox.top;
      
  // データ座標に変換（pixel → linear）
  const newGamma = xaxis.p2l(moveX);
  const newLoad = yaxis.p2l(moveY);

      // ---- スナップ(吸着)処理: 実験データ折れ線に近い場合はその最近傍点へ吸着 ----
      // ピクセル距離閾値（適宜調整可能）
      const SNAP_PX_THRESHOLD = 12;
      let appliedGamma = newGamma;
      let appliedLoad = newLoad;
      let snapped = false;
      if(window.rawData !== undefined){ /* グローバル rawData 参照 */ }
      try{
        const snapCandidate = findNearestRawDataSnap(newGamma, newLoad, xaxis, yaxis, rawData, SNAP_PX_THRESHOLD);
        if(snapCandidate){
          appliedGamma = snapCandidate.gamma;
          appliedLoad = snapCandidate.Load;
          snapped = true;
        }
      }catch(err){ /* 失敗しても通常ドラッグ継続 */ }

      // 包絡線点を更新（スナップ後座標）
      editableEnvelope[shiftDragIndex].gamma = appliedGamma;
      editableEnvelope[shiftDragIndex].Load = appliedLoad;
      
      // プロット更新
      updateEnvelopePlot(editableEnvelope);
      
      // 点編集ダイアログが対象点を編集中なら、入力欄をリアルタイム更新（ユーザーのリクエスト対応）
      if(pointEditDialog && pointEditDialog.style.display !== 'none' && window._selectedEnvelopePoint === shiftDragIndex){
        // 表示精度は既存UIに合わせる（γ: 4～6桁, P: 3桁）必要に応じて後で統一可能
        if(editGammaInput){ editGammaInput.value = appliedGamma.toFixed(4); }
        if(editLoadInput){ editLoadInput.value = appliedLoad.toFixed(1); }
      }

      // ツールチップ更新
      if(pointTooltip){
        pointTooltip.textContent = `γ: ${appliedGamma.toFixed(6)}, P: ${appliedLoad.toFixed(3)}${snapped ? ' [吸着]' : ''}`;
        pointTooltip.style.left = e.clientX + 10 + 'px';
        pointTooltip.style.top = e.clientY + 10 + 'px';
        if(snapped){
          pointTooltip.style.background = 'rgba(255,165,0,0.85)'; // 吸着時はオレンジ強調
        }else{
          pointTooltip.style.background = 'rgba(0,0,0,0.6)';
        }
      }
      
      e.preventDefault();
    };
    
    // マウスアップ: ドラッグ終了
    let mouseupHandler = function(e){
      if(shiftDragging && shiftDragIndex >= 0){
        // 履歴に保存
        pushHistory(editableEnvelope);
        
        // 再計算
        recalculateFromEnvelope(editableEnvelope);
        appendLog(`包絡線点 #${shiftDragIndex} をドラッグ移動しました`);
        
        // ツールチップ非表示
        if(pointTooltip){
          pointTooltip.style.display = 'none';
        }
        
  shiftDragging = false;
  shiftDragIndex = -1;
  // マウスアップ時はクラスをクリア（再ホバーで再設定）
  plotDiv.classList.remove('drag-point-active');
  plotDiv.classList.remove('drag-point-hover');
  plotDiv.classList.remove('pointer-mode');
        
        // Plotlyのドラッグモードを復元（エラーは握りつぶし）
        if(window.Plotly && plotDiv){
          try{
            safeRelayout(plotDiv, {'dragmode': 'pan'});
          }catch(_){/* noop */}
        }
      }
    };
    
    // イベントリスナー登録（キャプチャフェーズで先にイベントを取得）
    const dragLayer = plotDiv.querySelector('.draglayer') || plotDiv;
    // マウスイベント
    dragLayer.addEventListener('mousedown', mousedownHandler, true);
    document.addEventListener('mousemove', mousemoveHandler);
    document.addEventListener('mouseup', mouseupHandler);
    document.addEventListener('mouseleave', mouseupHandler);
    // ホイールズーム抑止（トグルON時）
    const wheelSuppressor = function(e){ if(dragMoveEnabled){ e.preventDefault(); e.stopImmediatePropagation(); } };
    dragLayer.addEventListener('wheel', wheelSuppressor, {passive:false, capture:true});
    // ポインタイベント（Zoom優先を抑止するため早期にハンドリング）
    const pointerdownHandler = function(e){
      if(e && (e.pointerType === 'mouse' || e.pointerType === 'pen')){
        mousedownHandler(e);
      }
    };
    const pointermoveHandler = function(e){
      if(e && (e.pointerType === 'mouse' || e.pointerType === 'pen')){
        mousemoveHandler(e);
      }
    };
    const pointerupHandler = function(e){
      if(e && (e.pointerType === 'mouse' || e.pointerType === 'pen')){
        mouseupHandler(e);
      }
    };
    dragLayer.addEventListener('pointerdown', pointerdownHandler, true);
    document.addEventListener('pointermove', pointermoveHandler, true);
    document.addEventListener('pointerup', pointerupHandler, true);
  }
  
  function highlightSelectedPoint(editableEnvelope){
    // 選択された点を赤色で強調表示
    const colors = editableEnvelope.map((pt, idx) => idx === window._selectedEnvelopePoint ? 'red' : 'blue');
    const sizes = editableEnvelope.map((pt, idx) => idx === window._selectedEnvelopePoint ? 14 : 10);
    
    Plotly.restyle(plotDiv, {
      'marker.color': [colors],
      'marker.size': [sizes]
    }, [2]); // trace 2: 包絡線点
    if(openPointEditButton) openPointEditButton.disabled = (window._selectedEnvelopePoint < 0);
    // 前/次ボタンの活性制御
    if(selectPrevPointButton){
      selectPrevPointButton.disabled = !(window._selectedEnvelopePoint > 0);
    }
    if(selectNextPointButton){
      selectNextPointButton.disabled = !(window._selectedEnvelopePoint >= 0 && window._selectedEnvelopePoint < editableEnvelope.length - 1);
    }
  }
  
  function updateEnvelopePlot(editableEnvelope){
    // 包絡線トレース（trace 1）と包絡線点トレース（trace 2）を更新
    Plotly.restyle(plotDiv, {
      x: [editableEnvelope.map(pt => pt.gamma)],
      y: [editableEnvelope.map(pt => pt.Load)]
    }, [1]); // trace 1: 包絡線
    
    const colors = editableEnvelope.map((pt, idx) => idx === window._selectedEnvelopePoint ? 'red' : 'blue');
    const sizes = editableEnvelope.map((pt, idx) => idx === window._selectedEnvelopePoint ? 14 : 10);
    
    Plotly.restyle(plotDiv, {
      x: [editableEnvelope.map(pt => pt.gamma)],
      y: [editableEnvelope.map(pt => pt.Load)],
      'marker.color': [colors],
      'marker.size': [sizes]
    }, [2]); // trace 2: 包絡線点
    // 旧ドラッグ用ツールチップは廃止
    if(pointTooltip){ pointTooltip.style.display = 'none'; }
  }
  
  function deleteEnvelopePoint(pointIndex, editableEnvelope){
    if(editableEnvelope.length <= 2){
      alert('包絡線には最低2点が必要です');
      return;
    }
    // 履歴: 変更前を保存
    pushHistory(editableEnvelope);
    editableEnvelope.splice(pointIndex, 1);
    window._selectedEnvelopePoint = -1; // 選択解除
    updateEnvelopePlot(editableEnvelope);
    recalculateFromEnvelope(editableEnvelope);
    appendLog('包絡線点を削除しました（残り' + editableEnvelope.length + '点）');
    envelopeData = editableEnvelope.map(p=>({...p}));
    updateHistoryButtons();
  }

  // 前/次選択移動
  function moveSelectedPoint(direction){
    if(!envelopeData || !Array.isArray(envelopeData) || envelopeData.length === 0) return;
    if(typeof window._selectedEnvelopePoint !== 'number' || window._selectedEnvelopePoint < 0){
      // まだ選択が無いときは先頭を選択
      window._selectedEnvelopePoint = 0;
    } else {
      const next = window._selectedEnvelopePoint + direction;
      if(next < 0 || next >= envelopeData.length) return; // 範囲外
      window._selectedEnvelopePoint = next;
    }
    // 再描画で赤丸更新
    renderPlot(envelopeData, analysisResults);
    // ダイアログ開いている場合は内容更新
    if(pointEditDialog && pointEditDialog.style.display !== 'none'){
      openPointEditDialog();
    }
  }

  // ===== Excelインポート機能 =====
  async function handleImportExcelFile(ev){
    const file = importExcelInput.files && importExcelInput.files[0];
    if(!file){ return; }
    if(!window.ExcelJS){ alert('ExcelJSライブラリが読み込まれていません'); return; }
    try{
      const buf = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      // Envelopeシート読み込み
      const wsEnv = wb.getWorksheet('Envelope');
      if(!wsEnv){ alert('Envelope シートが見つかりません'); return; }
      const newEnv = [];
      wsEnv.eachRow((row, rowNumber) => {
        if(rowNumber === 1) return; // header
        const g = row.getCell(1).value;
        const p = row.getCell(2).value;
        if(typeof g === 'number' && typeof p === 'number'){
          newEnv.push({gamma: g, Load: p, gamma0: g});
        }
      });
      if(newEnv.length < 2){ alert('Envelope シートに有効なデータが不足しています'); return; }
      envelopeData = newEnv;
      envelopeData = newEnv;
      // フル包絡線を側別キャッシュへ保持
      const inferredSide = (function(){
        for(const pt of newEnv){
          if(Number.isFinite(pt.Load) && pt.Load !== 0) return pt.Load > 0 ? 'positive' : 'negative';
          if(Number.isFinite(pt.gamma) && pt.gamma !== 0) return pt.gamma > 0 ? 'positive' : 'negative';
        }
        return getCurrentSide();
      })();
      originalEnvelopeBySide[inferredSide] = newEnv.map(p=>({...p}));
      // 取り込んだ包絡線の側を推定してキャッシュ
      try{
        const sign = (function(){
          for(const pt of newEnv){
            if(Number.isFinite(pt.Load) && pt.Load !== 0){ return pt.Load > 0 ? 'positive' : 'negative'; }
            if(Number.isFinite(pt.gamma) && pt.gamma !== 0){ return pt.gamma > 0 ? 'positive' : 'negative'; }
          }
          return getCurrentSide();
        })();
        // ドロップダウンと不一致なら、UI側を推定側へ合わせる
        if(envelope_side){
          const target = (sign === 'negative') ? 'negative' : 'positive';
          if(envelope_side.value !== target){ envelope_side.value = target; }
        }
        // キャッシュ保存
        editedEnvelopeBySide[sign] = newEnv.map(p=>({ ...p }));
        editedDirtyBySide[sign] = true;
      }catch(_){ /* ignore */ }
      // InputData シートがあれば rawData も復元（包絡線再計算に使用するため）
      const wsInput = wb.getWorksheet('InputData');
      if(wsInput){
        const newRaw = [];
        wsInput.eachRow((row,rowNumber)=>{
          if(rowNumber===1) return;
          const g = row.getCell(1).value;
          const p = row.getCell(2).value;
          if(typeof g === 'number' && typeof p === 'number'){
            newRaw.push({gamma: g, Load: p, gamma0: g});
          }
        });
        if(newRaw.length >= 3){ rawData = newRaw; }
      }
      // Summary シートから設定値復元（日本語含）
      const wsSummary = wb.getWorksheet('Summary');
      if(wsSummary){
        // マッピング: 表示名 → 値
        const map = {};
        wsSummary.eachRow((row,rowNumber)=>{
          if(rowNumber===1) return;
          const key = String(row.getCell(1).value||'').trim();
          const val = row.getCell(2).value;
          map[key] = val;
        });
        // 可能なら試験体名称復元
        if(specimen_name && typeof map['試験体名称'] !== 'undefined'){
          specimen_name.value = String(map['試験体名称']);
        }
        // α / 最大変位δmax など
        if(alpha_factor && typeof map['耐力低減係数 α'] !== 'undefined'){
          alpha_factor.value = Number(map['耐力低減係数 α']) || alpha_factor.value;
        }
        if(max_ultimate_deformation && typeof map['最大変位 δmax'] !== 'undefined'){
          max_ultimate_deformation.value = Number(map['最大変位 δmax']) || max_ultimate_deformation.value;
        }
      }
      // 解析再計算
      // 解析再計算（指標は必ずフル包絡線で）
      if(envelopeData.length){
        const alpha = parseFloat(alpha_factor.value);
        const delta_u_max = parseFloat(max_ultimate_deformation.value);
        const sideNow = getCurrentSide();
        const fullForMetrics = originalEnvelopeBySide[sideNow] ? originalEnvelopeBySide[sideNow] : envelopeData;
        analysisResults = calculateJTCCMMetrics(fullForMetrics, delta_u_max, alpha);
        // 表示用間引き
        envelopeData = reapplyDisplayThinning(fullForMetrics, sideNow, analysisResults);
        renderPlot(envelopeData, analysisResults);
        renderResults(analysisResults);
        // 履歴初期化
        historyStack = [cloneEnvelope(envelopeData)];
        redoStack = [];
        updateHistoryButtons();
        appendLog('Excelインポート完了: Envelope/入力/設定を読み込みました');
      }
    }catch(err){
      console.error('Excelインポートエラー:', err);
      alert('Excelインポートに失敗しました');
      appendLog('Excelインポートエラー: '+(err && err.message ? err.message: err));
    }finally{
      importExcelInput.value = '';
    }
  }
  
  function addEnvelopePoint(gamma, load, editableEnvelope){
    // 新しい点を適切な位置に挿入（gamma順）
    let insertIdx = editableEnvelope.findIndex(pt => pt.gamma > gamma);
    if(insertIdx < 0) insertIdx = editableEnvelope.length;
    
    editableEnvelope.splice(insertIdx, 0, {
      gamma: gamma,
      Load: load,
      gamma0: gamma // 簡易的に同値
    });
    
    updateEnvelopePlot(editableEnvelope);
    recalculateFromEnvelope(editableEnvelope);
    appendLog('包絡線点を追加しました（γ=' + gamma.toFixed(6) + ', P=' + load.toFixed(3) + '）');
  }
  
  function addEnvelopePointAtNearestSegment(clickX, clickY, xData, yData, editableEnvelope, xaxis, yaxis){
    // クリック位置から最も近い包絡線セグメント（2点間）を見つける
    let minDist = Infinity;
    let nearestSegmentIdx = 0;
    
    for(let i = 0; i < editableEnvelope.length - 1; i++){
      const p1 = editableEnvelope[i];
      const p2 = editableEnvelope[i + 1];
      
  const x1 = xaxis.l2p(p1.gamma);
  const y1 = yaxis.l2p(p1.Load);
  const x2 = xaxis.l2p(p2.gamma);
  const y2 = yaxis.l2p(p2.Load);
      
      // 線分への最短距離を計算
      const dist = pointToSegmentDistance(clickX, clickY, x1, y1, x2, y2);
      
      if(dist < minDist){
        minDist = dist;
        nearestSegmentIdx = i;
      }
    }
    
    // 最寄りセグメントの中点に新しい点を追加
    const p1 = editableEnvelope[nearestSegmentIdx];
    const p2 = editableEnvelope[nearestSegmentIdx + 1];
    const midGamma = (p1.gamma + p2.gamma) / 2;
    const midLoad = (p1.Load + p2.Load) / 2;
    
    // 履歴: 変更前を保存
    pushHistory(editableEnvelope);
    editableEnvelope.splice(nearestSegmentIdx + 1, 0, {
      gamma: midGamma,
      Load: midLoad,
      gamma0: midGamma
    });
    
    updateEnvelopePlot(editableEnvelope);
    recalculateFromEnvelope(editableEnvelope);
    appendLog('包絡線点を追加しました（γ=' + midGamma.toFixed(6) + ', P=' + midLoad.toFixed(3) + '）');
    envelopeData = editableEnvelope.map(p=>({...p}));
    updateHistoryButtons();
  }
  
  function pointToSegmentDistance(px, py, x1, y1, x2, y2){
    // 点(px, py)から線分(x1,y1)-(x2,y2)への最短距離
    const dx = x2 - x1;
    const dy = y2 - y1;
    const lengthSq = dx * dx + dy * dy;
    
    if(lengthSq === 0){
      // 線分が点の場合
      return Math.sqrt((px - x1) * (px - x1) + (py - y1) * (py - y1));
    }
    
    // 線分上の最近点のパラメータt (0 <= t <= 1)
    let t = ((px - x1) * dx + (py - y1) * dy) / lengthSq;
    t = Math.max(0, Math.min(1, t));
    
    const closestX = x1 + t * dx;
    const closestY = y1 + t * dy;
    
    return Math.sqrt((px - closestX) * (px - closestX) + (py - closestY) * (py - closestY));
  }

  // データ座標系での実験データ線分への吸着処理
  function snapToNearestRawDataSegment(gamma, load, rawData, side){
    try{
      if(!rawData || rawData.length < 2) return null;
      
      // 側面でフィルタリング（positive: gamma >= 0, negative: gamma <= 0）
      const filteredData = side === 'positive' 
        ? rawData.filter(pt => pt.gamma >= 0 && pt.Load >= 0)
        : rawData.filter(pt => pt.gamma <= 0 && pt.Load <= 0);
      
      if(filteredData.length < 2) return null;
      
      let bestSegmentIndex = -1;
      let bestDistance = Infinity;
      let bestT = 0;
      
      // 全線分を走査して最小距離のものを探す

      for(let i = 0; i < filteredData.length - 1; i++){
        const p1 = filteredData[i];
        const p2 = filteredData[i + 1];
        
        const g1 = p1.gamma, l1 = p1.Load;
        const g2 = p2.gamma, l2 = p2.Load;
        
        if(!Number.isFinite(g1) || !Number.isFinite(l1) || !Number.isFinite(g2) || !Number.isFinite(l2)) continue;
        
        // 線分のベクトル
        const dx = g2 - g1;
        const dy = l2 - l1;
        const lengthSq = dx * dx + dy * dy;
        
        if(lengthSq === 0) continue; // 長さ0の線分はスキップ
        
        // 点から線分への投影パラメータt (0 <= t <= 1)
        let t = ((gamma - g1) * dx + (load - l1) * dy) / lengthSq;
        t = Math.max(0, Math.min(1, t));
        
        // 線分上の最近点
        const closestGamma = g1 + t * dx;
        const closestLoad = l1 + t * dy;
        
        // 距離を計算
        const distance = Math.sqrt(
          (gamma - closestGamma) * (gamma - closestGamma) + 
          (load - closestLoad) * (load - closestLoad)
        );
        
        if(distance < bestDistance){
          bestDistance = distance;
          bestSegmentIndex = i;
          bestT = t;
        }
      }
      
      // 最も近い線分が見つかった場合、その線分上の点を返す
      if(bestSegmentIndex >= 0){
        const p1 = filteredData[bestSegmentIndex];
        const p2 = filteredData[bestSegmentIndex + 1];

        
        return {
          gamma: p1.gamma + bestT * (p2.gamma - p1.gamma),
          Load: p1.Load + bestT * (p2.Load - p1.Load)
        };
      }
      
      return null; // 吸着に失敗
    }catch(err){
      console.warn('snapToNearestRawDataSegment エラー:', err);
      return null;
    }
  }


  // 実験データ折れ線に対する最近傍点を探索し、ピクセル距離がthresholdPx以下ならデータ座標を返す
  function findNearestRawDataSnap(gamma, load, xaxis, yaxis, raw, thresholdPx){
    try{
      if(!raw || raw.length < 2 || !xaxis || !yaxis) return null;
      const px = xaxis.l2p(gamma);
      const py = yaxis.l2p(load);
      let best = null;
      let bestDist = Number.isFinite(thresholdPx) ? thresholdPx : 10;
      for(let i=0; i<raw.length-1; i++){
        const r1 = raw[i];
        const r2 = raw[i+1];
        if(!r1 || !r2) continue;
        const g1 = r1.gamma, l1 = r1.Load;
        const g2 = r2.gamma, l2 = r2.Load;
        if(!Number.isFinite(g1) || !Number.isFinite(l1) || !Number.isFinite(g2) || !Number.isFinite(l2)) continue;
        const x1 = xaxis.l2p(g1); const y1 = yaxis.l2p(l1);
        const x2 = xaxis.l2p(g2); const y2 = yaxis.l2p(l2);
        const dx = x2 - x1; const dy = y2 - y1;
        const lenSq = dx*dx + dy*dy;
        let cx = x1, cy = y1; // 最近傍のピクセル座標
        if(lenSq > 0){
          let t = ((px - x1)*dx + (py - y1)*dy) / lenSq;
          t = Math.max(0, Math.min(1, t));
          cx = x1 + t*dx;
          cy = y1 + t*dy;
        }
        const dist = Math.hypot(px - cx, py - cy);
        if(dist <= bestDist){
          bestDist = dist;
          best = {
            gamma: xaxis.p2l(cx),
            Load:  yaxis.p2l(cy),
            distPx: dist,
            segmentIndex: i
          };
        }
      }
      return best;
    }catch(err){ return null; }
  }
  
  function recalculateFromEnvelope(editableEnvelope){
    try{
      // 編集後の包絡線から特性値を再計算
      envelopeData = editableEnvelope.map(pt => ({...pt}));
      // 現在側の編集データとして保存
      setEditedEnvelopeForCurrentSide(envelopeData);
      
      const alpha = parseFloat(alpha_factor.value);
      const delta_u_max = parseFloat(max_ultimate_deformation.value);
      
      if(!isFinite(alpha) || !isFinite(delta_u_max) || delta_u_max <= 0) return;
      
      analysisResults = calculateJTCCMMetrics(envelopeData, delta_u_max, alpha);
      renderResults(analysisResults);
      
      // 評価直線などを再描画
      renderPlot(envelopeData, analysisResults);
      
      appendLog('包絡線編集に基づき特性値を再計算しました');
    }catch(err){
      console.error('再計算エラー:', err);
      appendLog('再計算エラー: ' + (err && err.message ? err.message : err));
    }
  }

  function makeLine(lineObj, gamma_range, name, color, sign = 1){
    const x = gamma_range.map(g => g * sign);
    const y = gamma_range.map(g => (lineObj.slope * g + lineObj.intercept) * sign);
    return {
      x, y,
      mode: 'lines',
      name,
      line: {color, width: 1, dash: 'dash'}
    };
  }

  function renderResults(r){
    document.getElementById('val_pmax').textContent = r.Pmax.toFixed(3);
    document.getElementById('val_py').textContent = r.Py.toFixed(3);
    document.getElementById('val_dy').textContent = r.delta_y.toFixed(3);
    document.getElementById('val_K').textContent = r.K.toFixed(2);
    document.getElementById('val_pu').textContent = r.Pu.toFixed(3);
    document.getElementById('val_dv').textContent = r.delta_v.toFixed(3);
    document.getElementById('val_du').textContent = r.delta_u.toFixed(3);
    document.getElementById('val_mu').textContent = r.mu.toFixed(3);
    // 構造特性係数 Ds = 1 / sqrt(2μ - 1)
    let Ds = '-';
    if(r.mu && r.mu > 0.5){ // 2μ-1 > 0 の領域のみ算出（μ>0.5）
      const denom = Math.sqrt(2 * r.mu - 1);
      if(denom > 0){ Ds = (1 / denom).toFixed(3); }
    }
    const dsEl = document.getElementById('val_ds');
    if(dsEl) dsEl.textContent = Ds;

  document.getElementById('val_p0_a').textContent = r.p0_a.toFixed(3);
  document.getElementById('val_p0_b').textContent = r.p0_b.toFixed(3);
    document.getElementById('val_p0').textContent = r.P0.toFixed(3);

    document.getElementById('val_pa').textContent = r.Pa.toFixed(3);
  }

  async function downloadExcel(){
    if(!window.ExcelJS){
      alert('ExcelJSライブラリが読み込まれていません');
      return;
    }
    try{
      const specimen = (specimen_name && specimen_name.value ? specimen_name.value.trim() : 'testname');
      let wb = null;
      // ネイティブチャート対応: 明示的に有効化された場合のみ template.xlsx を読み込み
      if(window.APP_CONFIG && window.APP_CONFIG.useExcelTemplate){
        try{
          const resp = await fetch('template.xlsx', {cache:'no-cache'});
          if(resp.ok){
            const buf = await resp.arrayBuffer();
            wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buf);
            appendLog('情報: template.xlsx を使用してExcelを生成');
          }else{
            appendLog('情報: template.xlsx が見つかりません (resp='+resp.status+')');
          }
        }catch(e){
          appendLog('情報: template.xlsx 読込不可 (' + (e && e.message ? e.message : e) + ')');
        }
      }
      if(!wb){
        wb = new ExcelJS.Workbook();
        wb.creator = 'hyouka-app';
        wb.created = new Date();
      }

      // 1) 解析結果シート
      let wsSummary = wb.getWorksheet('Summary');
      if(!wsSummary) wsSummary = wb.addWorksheet('Summary');
      const r = analysisResults;
      wsSummary.addRow(['項目','値','単位']);
      wsSummary.addRow(['試験体名称', specimen, '']);
      const rows = [
        ['最大耐力 Pmax', r.Pmax, 'kN'],
        ['降伏耐力 Py', r.Py, 'kN'],
        ['降伏変位 δy', r.delta_y, 'mm'],
        ['初期剛性 K', r.K, 'kN/mm'],
        ['終局耐力 Pu', r.Pu, 'kN'],
        ['終局変位 δu', r.delta_u, 'mm'],
        ['塑性率 μ', r.mu, ''],
        ['P0(a) 降伏耐力', r.p0_a, 'kN'],
  ['P0(b) 最大耐力基準', r.p0_b, 'kN'],
        ['短期基準せん断耐力 P0', r.P0, 'kN'],
        ['短期許容せん断耐力 Pa', r.Pa, 'kN']
      ];
      rows.forEach(row => wsSummary.addRow(row));
      wsSummary.columns.forEach(col => { col.width = 22; });
      // 数値フォーマット適用
      for(let i=2;i<=wsSummary.rowCount;i++){
        const label = wsSummary.getCell(i,1).value;
        const cell = wsSummary.getCell(i,2);
        if(typeof cell.value !== 'number') continue;
        if(wsSummary.getCell(i,3).value === '1/n') continue; // reciprocalは文字列のまま
        else if(label === '初期剛性 K') cell.numFmt = '#,##0.00';
        else cell.numFmt = '#,##0.000';
      }

      // 2) 入力データシート
      let wsInput = wb.getWorksheet('InputData');
      if(!wsInput) wsInput = wb.addWorksheet('InputData');
      // ヘッダ再設定と既存データクリア
      wsInput.spliceRows(1, wsInput.rowCount, ['gamma','Load']);
      rawData.forEach(pt => wsInput.addRow([pt.gamma, pt.Load]));
      wsInput.columns.forEach(c=> c.width = 18);
      for(let i=2;i<=wsInput.rowCount;i++){
        const cg = wsInput.getCell(i,1); if(typeof cg.value==='number') cg.numFmt='0.000000';
        const cp = wsInput.getCell(i,2); if(typeof cp.value==='number') cp.numFmt='0.000';
      }

      // 3) 包絡線シート
      let wsEnv = wb.getWorksheet('Envelope');
      if(!wsEnv) wsEnv = wb.addWorksheet('Envelope');
      wsEnv.spliceRows(1, wsEnv.rowCount, ['gamma','Load']);
      (envelopeData||[]).forEach(pt => wsEnv.addRow([pt.gamma, pt.Load]));
      wsEnv.columns.forEach(c=> c.width = 18);
      for(let i=2;i<=wsEnv.rowCount;i++){
        const cg = wsEnv.getCell(i,1); if(typeof cg.value==='number') cg.numFmt='0.000000';
        const cp = wsEnv.getCell(i,2); if(typeof cp.value==='number') cp.numFmt='0.000';
      }

      // 5) 設定シート（日本語項目付き）
      let wsSettings = wb.getWorksheet('Settings');
      if(!wsSettings) wsSettings = wb.addWorksheet('Settings');
      wsSettings.spliceRows(1, wsSettings.rowCount, ['項目','値','単位']);
      const envSideLabel = (envelope_side && envelope_side.value === 'negative') ? '負側' : '正側';
      const maxDu = Number(max_ultimate_deformation.value);
      wsSettings.addRow(['試験体名称', specimen, '']);
      wsSettings.addRow(['耐力低減係数 α', Number(alpha_factor.value), '']);
      wsSettings.addRow(['最大変位 δmax', maxDu, 'mm']);
      wsSettings.addRow(['評価対象包絡線', envSideLabel, '']);
      // 体裁
      wsSettings.columns.forEach(c=> c.width = 22);
      for(let i=2;i<=wsSettings.rowCount;i++){
        const label = wsSettings.getCell(i,1).value;
        const cell = wsSettings.getCell(i,2);
        if(typeof cell.value === 'number'){
          if(label === '最大変位 δmax') cell.numFmt = '#,##0'; else cell.numFmt = '#,##0.00';
        }
      }

      // 4) グラフシート (画像埋込み)
      // Chartシート: テンプレートがあれば既存を活用。無ければ画像埋め込みの代替。
      let wsChart = wb.getWorksheet('Chart');
      if(!wsChart){
        wsChart = wb.addWorksheet('Chart');
        wsChart.getCell('A1').value = '荷重-変形関係グラフ';
        wsChart.getRow(1).font = {bold:true};
        const pngDataUrl = await Plotly.toImage(plotDiv, {format:'png', width:1200, height:700});
        const base64 = pngDataUrl.replace(/^data:image\/png;base64,/, '');
        const imageId = wb.addImage({base64, extension:'png'});
        wsChart.addImage(imageId, { tl: {col:0, row:2}, ext: {width: 900, height: 520} });
      }

      // 仕上げ: 自動フィルタやスタイル軽微調整
      wsSummary.getRow(1).font = {bold:true};
      wsInput.getRow(1).font = {bold:true};
      wsEnv.getRow(1).font = {bold:true};

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
  const excelFileName = `Results_${specimen.replace(/[^a-zA-Z0-9_\-一-龥ぁ-んァ-ヶ]/g,'_')}.xlsx`;
  a.download = excelFileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }catch(err){
      console.error('Excel出力エラー:', err);
      alert('Excelの生成に失敗しました。');
      appendLog('Excel出力エラー: ' + (err && err.stack ? err.stack : err.message));
    }
  }

  function appendLog(message){
    // ログ機能は無効化（コンソールのみに出力）
    console.log('[LOG]', message);
  }
})();
