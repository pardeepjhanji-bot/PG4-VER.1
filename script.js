
  async function saveInspectorAssignment(){
    const statusEl = document.getElementById('ai_status');
    const mandiSel     = document.getElementById('ai_mandi');
    const inspSel      = document.getElementById('ai_inspector_select');
    const inspManualEl = document.getElementById('ai_inspector_manual');
    const signEl       = document.getElementById('ai_sign');

    const mandi = (mandiSel && mandiSel.value || '').trim();
    const picked = (inspSel && inspSel.value || '').trim();
    const inspector = (picked === '__manual__'
                        ? ((inspManualEl && inspManualEl.value) || '').trim()
                        : picked);
    const signUrl = (signEl && signEl.value || '').trim();

    if(!mandi){ if(statusEl) statusEl.textContent = "❌ Select a mandi."; return; }
    if(!inspector){ if(statusEl) statusEl.textContent = "❌ Select or type inspector name."; return; }

    if(statusEl) statusEl.textContent = "Saving…";
    try{
      const fd = new FormData();
      fd.append('action','assignInspector');
      fd.append('mandi', mandi);
      fd.append('inspector', inspector);
      fd.append('signUrl', signUrl);

      const res = await fetch(SCRIPT_URL, { method:'POST', body: fd });
      let js = null;
      try { js = await res.json(); } catch(e){ js = null; }

      if(js && (js.ok === true || js.status === 'ok')){
        // Update local cache/UI
        if(!window.inspectorMap) window.inspectorMap = {};
        window.inspectorMap[mandi] = { inspector: inspector, signUrl: signUrl };

        if (typeof populateInspectorMandiDropdown === 'function') populateInspectorMandiDropdown();
        if (typeof populateInspectorDropdown === 'function')       populateInspectorDropdown();
        if (typeof applyInspectorLockUI === 'function')            applyInspectorLockUI();

        if (mandiSel) mandiSel.value = '';
        if (inspSel)  inspSel.value  = '';
        if (inspManualEl){ inspManualEl.value=''; inspManualEl.style.display='none'; }
        if (signEl)   signEl.value   = '';

        if(statusEl) statusEl.textContent = "✅ MANDI LINKED TO INSPECTOR SUCCESSFULLY";
      } else {
        const msg = js && (js.msg || js.message) ? js.msg || js.message : "Save failed";
        // Treat "already assigned" as idempotent success
        if (/already assigned/i.test(msg)) {
          if(statusEl) statusEl.textContent = "✅ MANDI LINKED TO INSPECTOR SUCCESSFULLY";
          // still refresh section
          if (typeof populateInspectorMandiDropdown === 'function') populateInspectorMandiDropdown();
          if (typeof populateInspectorDropdown === 'function')       populateInspectorDropdown();
          if (typeof applyInspectorLockUI === 'function')            applyInspectorLockUI();
          if (mandiSel) mandiSel.value = '';
          if (inspSel)  inspSel.value  = '';
          if (inspManualEl){ inspManualEl.value=''; inspManualEl.style.display='none'; }
          if (signEl)   signEl.value   = '';
        } else {
          if(statusEl) statusEl.textContent = "❌ " + msg;
        }
      }
    } catch(err){
      if(statusEl) statusEl.textContent = "❌ Network / script error";
    }
  }




  /* ===== CONFIG ===== */
  const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx6-uRQ6kIhamci5jCp6LEhGK59yn-Nn34Aru4qGZbJwu2HEgg6QzlCBiughSpEYNFt/exec";

  /* ===== Helpers ===== */
  const $ = (sel, root) => (root||document).querySelector(sel);
  const $all = (sel, root) => Array.prototype.slice.call((root||document).querySelectorAll(sel));
  const formatDate = d => { if(!d) return ""; const p=String(d).split("-"); return p.length===3? (p[2]+"-"+p[1]+"-"+p[0]) : d; };
  const maxIsoDate = arr => { if(!arr||!arr.length) return null; let m=null; for (let i=0;i<arr.length;i++){ const d=arr[i]; if(m===null||d>m) m=d; } return m; };
  const toIntNonNeg = v => { const n=parseInt(v,10); return (isNaN(n)||n<0)?0:n; };
  const normalizeType = s => (s||'').trim().toUpperCase().replace(/\s+/g,' ');
  const EMAIL_RX = /^[^@\s]+@gmail\.com$/i;

  /* ===== Client auth keys ===== */
  const AUTH_KEYS = { email:'auth.email', role:'auth.role', verified:'auth.verified' };

  function setAuth(email, role, verified){
    try{
      localStorage.setItem(AUTH_KEYS.email, email||'');
      localStorage.setItem(AUTH_KEYS.role, role||'');
      localStorage.setItem(AUTH_KEYS.verified, JSON.stringify(!!verified));
    }catch(e){}
  }
  function getAuth(){
    try{
      return {
        email: localStorage.getItem(AUTH_KEYS.email) || '',
        role: localStorage.getItem(AUTH_KEYS.role) || '',
        verified: JSON.parse(localStorage.getItem(AUTH_KEYS.verified) || 'false')
      };
    }catch(e){ return { email:'', role:'', verified:false }; }
  }

  async function clearAllCaches(){
    try { if ('caches' in window) { const keys = await caches.keys(); await Promise.all(keys.map(k => caches.delete(k))); } } catch(e){}
  }

  async function logout(){
    // wipe everything we can in the browser
    try{ localStorage.clear(); 
  // --- ADDED: nuke service workers and force a fresh load of the login screen ---
  try {
    if ('serviceWorker' in navigator) {
      const regs = await navigator.serviceWorker.getRegistrations();
      regs.forEach(r => r.unregister());
    }
  } catch(e){}
  // Use replace so Back won't restore
  setTimeout(() => location.replace(location.href.split('#')[0]), 50);
}catch(e){}
    try{ sessionStorage.clear(); }catch(e){}
    await clearAllCaches();
    // UI reset
    updateAuthUI();
        // ADDED: save profile after OTP activation
        await saveUserProfileClient({ Role: (document.getElementById('authRole_reg')?.value || 'DEO').toUpperCase() });
    resetPurchaseSection();
    resetReportSection();
  }

  function setAuthStatus(msg, ok, which='login'){
    const el = which === 'reg' ? document.getElementById('authStatus_reg') : document.getElementById('authStatus');
    el.textContent = msg || '';
    el.style.color = ok ? '#2e7d32' : '#b00020';
  }

  async function ensureUserSheet(){ try { await fetch(SCRIPT_URL+'?action=ensureUserSheet'); } catch(e){} }

  /* ===== Register: request OTP ===== */
  async function requestOtp(){
    const email = (document.getElementById('authEmail_reg').value || document.getElementById('authEmail').value || '').trim();
    const role  = (document.getElementById('authRole_reg')?.value || 'DEO').trim().toUpperCase();
    if(!EMAIL_RX.test(email)){ setAuthStatus('Enter a valid Gmail (e.g., name@gmail.com).', false, 'reg'); return; }
    setAuthStatus('Requesting OTP…', true, 'reg');
    try{
      const fd = new FormData();
      fd.append('action','requestOtp'); fd.append('email', email); fd.append('role', role);
      const res = await fetch(SCRIPT_URL, { method:'POST', body:fd });
      const js = await res.json();
      if(js && js.ok){
        document.getElementById('otpRow').style.display = 'flex';
        setAuthStatus('OTP sent. Enter OTP and set your password below.', true, 'reg');
      } else setAuthStatus(js && js.msg ? js.msg : 'Could not send OTP.', false, 'reg');
    }catch(err){ setAuthStatus('Network error while requesting OTP.', false, 'reg'); }
  }

  /* ===== Verify & set password ===== */
  async function verifyOtp(){
    const email = (document.getElementById('authEmail_reg').value || document.getElementById('authEmail').value || '').trim();
    const role  = (document.getElementById('authRole_reg')?.value || 'DEO').trim().toUpperCase();
    const otp   = (document.getElementById('authOtp').value||'').trim();
    const pass1 = (document.getElementById('authPasswordNew').value||'').trim();
    const pass2 = (document.getElementById('authPasswordConfirm').value||'').trim();
    if(!EMAIL_RX.test(email)){ setAuthStatus('Enter a valid Gmail.', false, 'reg'); return; }
    if(!otp){ setAuthStatus('Enter the OTP.', false, 'reg'); return; }
    if(!pass1 || pass1.length<6){ setAuthStatus('Password must be at least 6 characters.', false, 'reg'); return; }
    if(pass1 !== pass2){ setAuthStatus('Passwords do not match.', false, 'reg'); return; }
    setAuthStatus('Verifying OTP & setting password…', true, 'reg');
    try{
      const fd = new FormData();
      fd.append('action','verifyOtp'); fd.append('email', email); fd.append('otp', otp); fd.append('password', pass1);
      const res = await fetch(SCRIPT_URL, { method:'POST', body:fd }); const js = await res.json();
      if(js && js.ok){
        setAuth(email, role, false); // stay blocked until approval
        setAuthStatus('Activation successful. You can now login with your password.', true, 'reg');
        updateAuthUI();
        // ADDED: save profile after OTP activation
        await saveUserProfileClient({ Role: (document.getElementById('authRole_reg')?.value || 'DEO').toUpperCase() });
      } else setAuthStatus(js && js.msg ? js.msg : 'Invalid/expired OTP.', false, 'reg');
    }catch(err){ setAuthStatus('Network error while verifying OTP.', false, 'reg'); }
  }

  /* ===== Login (with password, no role dropdown) ===== */
  async function login(){
    const email = (document.getElementById('authEmail').value||'').trim();
    const pwd   = (document.getElementById('authPasswordLogin').value||'').trim();
    if(!EMAIL_RX.test(email)){ setAuthStatus('Enter a valid Gmail (e.g., name@gmail.com).', false); return; }
    if(!pwd){ setAuthStatus('Enter your password.', false); return; }
    setAuthStatus('Checking credentials…', true);
    try{
      const fd = new FormData();
      fd.append('action','loginUser'); fd.append('email', email); fd.append('password', pwd);
      const res = await fetch(SCRIPT_URL, { method:'POST', body:fd });
      const js = await res.json();
      if(js && js.ok && js.verified){
        const finalRole = (js.role || 'DEO').toUpperCase();
        setAuth(email, finalRole, true);
        setAuthStatus('Login successful.', true);
        updateAuthUI();
        
        // ADDED: save profile after login
        await saveUserProfileClient({ Role: (js.role || 'DEO').toUpperCase() });// ADDED: save profile after OTP activation
        await saveUserProfileClient({ Role: (document.getElementById('authRole_reg')?.value || 'DEO').toUpperCase() });
      } else setAuthStatus(js && js.msg ? js.msg : 'Login failed.', false);
    }catch(err){
      setAuthStatus('Network error while logging in.', false);
    }
  }

  /* ===== Role-based UI ===== */
  let viewRole = 'DEO';
  function applyRoleView(){
    const isAdminView = (viewRole === 'ADMIN');
    document.getElementById('assignInspectorCard').style.display = isAdminView ? 'block' : 'none';
    document.getElementById('adminOnlyToggles').style.display    = isAdminView ? 'flex'  : 'none';
    if (!isAdminView) {
      document.getElementById('seasonRow').style.display   = 'none';
      document.getElementById('bardanaSetup').style.display= 'none';
    }
    document.getElementById('cardPartA').style.display = 'block';
    document.getElementById('cardPartB').style.display = 'block';
  }

  function updateAuthUI(){
    const auth = getAuth();
    const gate = document.getElementById('authCard');
    const app  = document.getElementById('appWrap');
    const badge= document.getElementById('authUserBadge');
    const adminBar = document.getElementById('adminBar');
    const toggle = document.getElementById('viewRoleToggle');

    if(auth && auth.email && auth.verified){
      // show identity
      document.getElementById('topUserEmail').textContent = auth.email;
      document.getElementById('topUserRole').textContent  = auth.role || 'DEO';
      document.getElementById('appTopBar').classList.remove('hide');

      document.getElementById('authUserEmail').textContent = auth.email;
      document.getElementById('authUserRole').textContent  = auth.role || 'DEO';
      badge.classList.remove('hide');

      gate.classList.add('hide');
      app.classList.remove('hide');

      const isAdmin = (String(auth.role||'').toUpperCase() === 'ADMIN');
      adminBar.classList.toggle('hide', !isAdmin);
      if(isAdmin){
        if(toggle.getAttribute('data-init') !== '1'){
          toggle.checked = true;
          toggle.setAttribute('data-init','1');
        }
        viewRole = toggle.checked ? 'ADMIN' : 'DEO';
        toggle.onchange = function(){ viewRole = this.checked ? 'ADMIN' : 'DEO'; applyRoleView(); };
      } else {
        viewRole = 'DEO';
      }
      applyRoleView();
    } else {
      document.getElementById('appTopBar').classList.add('hide');
      badge.classList.add('hide');
      gate.classList.remove('hide');
      app.classList.add('hide');
    }
  }

  /* ===== Toggle Section util ===== */
  function toggleSection(id) {
    const el = document.getElementById(id);
    if (!el) return;
    const curr = (el.style.display || window.getComputedStyle(el).display);
    el.style.display = (curr === 'none') ? 'flex' : 'none';
  }

  /* ===== Bardana (Server-synced) ===== */
  var bardanaTypes=[], bardanaLocked=true;
  async function loadBardanaFromServer(){
    try {
      let res = await fetch(SCRIPT_URL+"?action=getBardana");
      let js = await res.json();
      if(js && js.ok){
        bardanaTypes = Array.isArray(js.types) && js.types.length ? js.types : ['JUTE NEW','JUTE PREVIOUS YEAR'];
        bardanaLocked = true;
        renderBardanaUI();
      }
    } catch(e){ console.error("Error loading bardana", e); }
  }
  function renderBardanaUI(){
    var bd1=document.getElementById('bd1'), bd2=document.getElementById('bd2'), bd3=document.getElementById('bd3');
    if (bd1) bd1.value=bardanaTypes[0]||'';
    if (bd2) bd2.value=bardanaTypes[1]||'';
    if (bd3) bd3.value=bardanaTypes[2]||'';
    if (bd1) bd1.disabled=bardanaLocked;
    if (bd2) bd2.disabled=bardanaLocked;
    if (bd3) bd3.disabled=bardanaLocked;
    if (document.getElementById('applyBardanaBtn')) document.getElementById('applyBardanaBtn').style.display=bardanaLocked?'none':'inline-block';
    if (document.getElementById('changeBardanaBtn')) document.getElementById('changeBardanaBtn').style.display=bardanaLocked?'inline-block':'none';
    var chips=bardanaTypes.map((t,i)=>'<span class="chip"><strong>'+t+'</strong>&nbsp;<small>#'+(i+1)+'</small></span>').join('');
    var chipsEl=document.getElementById('bardanaChips'); if (chipsEl) chipsEl.innerHTML=chips;
    rebuildEntryTableForBardana(); updateSubtotal();
  }
  async function applyBardana(){
    var vals=[(document.getElementById('bd1').value||'').trim().toUpperCase(),
              (document.getElementById('bd2').value||'').trim().toUpperCase(),
              (document.getElementById('bd3').value||'').trim().toUpperCase()];
    vals=vals.filter(v=>!!v);
    var unique=[]; for (var i=0;i<vals.length;i++){ if(unique.indexOf(vals[i])===-1) unique.push(vals[i]); }
    unique=unique.slice(0,3);
    if(!unique.length){ alert('Enter at least one Bardana type.'); return; }

    try {
      var fd = new FormData();
      fd.append("action","saveBardana");
      fd.append("types", JSON.stringify(unique));
      let res = await fetch(SCRIPT_URL, { method:"POST", body:fd });
      let js = await res.json();
      if(js && js.ok){
        bardanaTypes = unique;
        bardanaLocked = true;
        renderBardanaUI();
        alert("✅ Bardana saved for all users.");
      } else {
        alert("❌ Could not save Bardana.");
      }
    } catch(err){
      alert("❌ Error saving Bardana.");
    }
  }
  function changeBardana(){ if(!confirm('Unlock Bardana to edit types?')) return; bardanaLocked=false; renderBardanaUI(); }

  /* ===== Mandi helpers ===== */
  function cleanMandis(list){
    const seen = new Set();
    const out = [];
    (list||[]).forEach(v=>{
      const s = String(v||'').trim();
      if(!s) return;
      if(/^mandi\s*name$/i.test(s)) return;
      if(seen.has(s)) return;
      seen.add(s); out.push(s);
    });
    return out;
  }
  function populateMandiDatalist(mandis){
    var dl=document.getElementById("mandiList"); if (!dl) return;
    dl.innerHTML="";
    cleanMandis(mandis).forEach(m=>{
      var o=document.createElement('option'); o.value=m; dl.appendChild(o);
    });
  }

  /* ===== Part A: Entry ===== */
  var mandiListData=[], firmMapData={};
  function firmsForCurrentMandi(){ var mandi=document.getElementById("mandiName").value.trim(); var arr = firmMapData[mandi] || []; return arr.slice().sort(); }
  function refreshFirmDatalists(){ var firms=firmsForCurrentMandi(); $all("#entryBody tr").forEach(tr=>{ var listId=tr.querySelector(".firm").getAttribute("list"); var dl=document.getElementById(listId); if(!dl) return; dl.innerHTML=''; firms.forEach(f=>{ var o=document.createElement('option'); o.value=f; dl.appendChild(o); }); }); }
  function rebuildEntryTableForBardana(){
    var head=document.getElementById('entryHeadRow'); if (!head) return;
    var dynamic=bardanaTypes.map(bt=>'<th>'+bt+' (Bags)</th>').join('');
    head.innerHTML='<th>S.No.</th><th>Firm Name</th>'+dynamic+'<th>Total Bags</th><th>Total Weight</th><th>Action</th>';
    document.getElementById('subtotalLabel').colSpan = 2 + bardanaTypes.length;

    var old=$all('#entryBody tr').map(tr=>{
      var firm=(tr.querySelector('.firm')||{}).value||'';
      var perType={}; $all('input.bags',tr).forEach(inp=>{ perType[inp.getAttribute('data-bardana')] = toIntNonNeg(inp.value||0); });
      return { firm:firm, perType:perType };
    });
    document.getElementById('entryBody').innerHTML=''; (old.length?old:[{firm:'',perType:{}}]).forEach(r=>addRow(r));
  }
  function addRow(prefill){
    var tbody=document.getElementById("entryBody"); var sn=tbody.children.length+1; var tr=document.createElement("tr"); var cells='';
    bardanaTypes.forEach(bt=>{
      var v=(prefill && prefill.perType && prefill.perType[bt]!=null)? prefill.perType[bt]: 0;
      cells+= '<td><input type="number" class="bags" data-bardana="'+bt+'" value="'+v+'" min="0" step="1" oninput="sanitizeInt(this); updateSubtotal()" /></td>';
    });
    tr.innerHTML=
        '<td class="fixed-cell">'+sn+'</td>'
      + '<td><input type="text" class="firm" list="firmList_'+sn+'" placeholder="Type or select firm" required />'
      + '<datalist id="firmList_'+sn+'"></datalist></td>'
      + cells
      + '<td class="fixed-cell"><input type="text" class="rowTotalBags" value="0" readonly /></td>'
      + '<td class="fixed-cell"><input type="text" class="rowTotalWeight" value="0" readonly /></td>'
      + '<td><button type="button" class="btn-del" onclick="removeRow(this)">Delete</button></td>';
    tbody.appendChild(tr);
    if(prefill && prefill.firm) tr.querySelector('.firm').value=prefill.firm;
    refreshFirmDatalists(); updateSubtotal();
  }
  function sanitizeInt(el){ el.value=toIntNonNeg(el.value); }
  function renumberRows(){ $all("#entryBody tr").forEach((tr,i)=>{ tr.children[0].textContent=i+1; }); }
  function removeRow(btn){ btn.closest("tr").remove(); renumberRows(); updateSubtotal(); }
  function updateSubtotal(){
    var tb=0, tw=0; var seasonType=document.getElementById("seasonType").value;
    var wpb = seasonType==="Paddy Season" ? 0.375 : 0.50;
    var dec = seasonType==="Paddy Season" ? 3 : 2;
    $all("#entryBody tr").forEach(tr=>{
      var bags=$all('input.bags',tr).reduce((s,inp)=> s+toIntNonNeg(inp.value||0),0);
      var w=+(bags*wpb).toFixed(dec);
      tr.querySelector(".rowTotalBags").value=bags;
      tr.querySelector(".rowTotalWeight").value=w.toFixed(dec);
      tb+=bags; tw+=w;
    });
    document.getElementById("subtotalBags").textContent=tb; document.getElementById("subtotalWeight").textContent=tw.toFixed(dec);
  }

  function resetPurchaseSection(){
    document.getElementById("status").textContent = "";
    document.getElementById("purchaseDate").value = "";
    document.getElementById("mandiName").value = "";
    document.getElementById("entryBody").innerHTML = "";
    addRow();
    updateSubtotal();
  }

  async function validatePurchaseDateStrictlyIncreasing(mandi, purchaseIsoDate){
    var urlAll=SCRIPT_URL+'?action=getReport&date=ALL&mandi='+encodeURIComponent(mandi);
    try{
      var res=await fetch(urlAll); var data=await res.json();
      if(!data.success) return {ok:true};
      var rows=Array.isArray(data.rows)?data.rows:[]; if(!rows.length) return {ok:true};
      var dates=rows.map(r=>r.date).filter(Boolean); var latest=maxIsoDate(dates);
      if(!latest) return {ok:true};
      if(purchaseIsoDate <= latest) return {ok:false, msg:'❌ Date must be after last entry for this mandi ('+formatDate(latest)+').'};
      return {ok:true};
    }catch(e){ return {ok:false, msg:"❌ Could not verify date vs previous entries."}; }
  }

  function updatePartAHeadline(){
    var season = (document.getElementById("seasonType").value || "").toUpperCase();
    var year   = document.getElementById("seasonYear").value || "";
    var headline = "PG-4 REPORT";
    if(season && year){ headline += " ("+season+" "+year+")"; }
    document.getElementById("partAHeadline").textContent = headline;
    document.getElementById("partAHeadlineSub").textContent = "PUNGRAIN";
  }

  document.addEventListener('input', function (e) {
    if (e.target && e.target.classList && e.target.classList.contains('firm')) {
      e.target.classList.remove('field-error');
    }
  });

  async function submitData(){
    const auth = getAuth();
    if(!auth.verified){ alert("Please login first."); return; }

    var date=document.getElementById("purchaseDate").value, mandi=document.getElementById("mandiName").value.trim();
    if(!date || !mandi){ document.getElementById("status").textContent="❌ Please fill Date and Mandi Name."; return; }
    if(!bardanaTypes.length){ document.getElementById("status").textContent='❌ Please configure Bardana types first.'; return; }

    let invalidRows = 0;
    $all("#entryBody tr").forEach(tr=>{
      const firmEl = tr.querySelector(".firm");
      const firm   = (firmEl?.value || "").trim();
      const bagsSum = $all("input.bags", tr).reduce((s,inp)=> s + toIntNonNeg(inp.value||0), 0);
      if (bagsSum > 0 && !firm) { invalidRows++; if (firmEl) firmEl.classList.add("field-error"); }
    });
    if (invalidRows > 0) { document.getElementById("status").textContent = "❌ Please fill Firm for rows with bags or delete those rows."; return; }

    document.getElementById("status").textContent="Checking date…";
    var check=await validatePurchaseDateStrictlyIncreasing(mandi, date);
    if(!check.ok){ document.getElementById("status").textContent=check.msg||"❌ Invalid date"; return; }

    var firmSeen={}; var duplicateFirm=null;
    $all("#entryBody tr").forEach(tr=>{
      const firm=(tr.querySelector(".firm")?.value||"").trim().toUpperCase();
      const bagsSum=$all("input.bags",tr).reduce((s,inp)=> s+toIntNonNeg(inp.value||0),0);
      if(!firm || bagsSum===0) return;
      if(firmSeen[firm]) duplicateFirm=firm; else firmSeen[firm]=true;
    });
    if(duplicateFirm){ document.getElementById("status").textContent='❌ Firm "'+duplicateFirm+'" repeated for this date & mandi.'; return; }

    // Save season if admin unlocked it
    if(!document.getElementById("seasonType").disabled){
      var season=document.getElementById("seasonType").value, year=document.getElementById("seasonYear").value, rate=document.getElementById("seasonRate").value;
      if(!season || !year || !rate){ document.getElementById("status").textContent="❌ Fill Season, Year, Rate."; return; }
      var fd2=new FormData(); fd2.append("action","saveSeason"); fd2.append("season",season); fd2.append("year",year); fd2.append("rate",rate);
      try{ await fetch(SCRIPT_URL,{method:"POST", body:fd2}); }catch(e){}
    }

    var seasonType=document.getElementById("seasonType").value;
    var wpb=seasonType==="Paddy Season" ? 0.375 : 0.50;
    var dec=seasonType==="Paddy Season" ? 3 : 2;

    var rows=[];
    $all("#entryBody tr").forEach(tr=>{
      var firm=(tr.querySelector(".firm")?.value||"").trim();
      var rowBags=0;
      $all("input.bags",tr).forEach(inp=>{ rowBags += toIntNonNeg(inp.value||0); });
      if(rowBags===0) return;
      $all("input.bags",tr).forEach(inp=>{
        var bags=toIntNonNeg(inp.value||0);
        if(bags>0) rows.push({ firm:firm, bardana:inp.getAttribute("data-bardana"), bags:bags, weight:+(bags*wpb).toFixed(dec) });
      });
    });
    if(!rows.length){ document.getElementById("status").textContent="❌ Enter at least one non-zero bags value."; return; }

    var fd=new FormData(); fd.append("payload", JSON.stringify({ date:date, mandi:mandi, rows:rows, user:getAuth().email || '' }));
    document.getElementById("status").textContent="Saving…";
    try{
      var res=await fetch(SCRIPT_URL,{method:"POST", body:fd});
      let okmsg="✅ Saved successfully!";
      try{ var d=await res.json(); if(d && d.success) okmsg='✅ Saved '+d.inserted+' rows!'; }catch(e){}
      document.getElementById("status").textContent=okmsg;

      await loadDataForReport();
      resetPurchaseSection();
    }catch(err){
      document.getElementById("status").textContent="❌ Error saving data: "+err;
    }
  }

  /* ===== Part B: Report ===== */
  var mandiDates={}, currentSeason={season:'',year:'',rate:0}, allMandis=[];
  var inspectorMap={}, allMandisCached=[];

  async function loadMandisAndFirms(){
    try{
      var res=await fetch(SCRIPT_URL); var data=await res.json();
      if(data.success){
        var yearSelect=document.getElementById("seasonYear"); yearSelect.innerHTML=""; var curr=new Date().getFullYear();
        for(var i=0;i<3;i++){ var y1=curr+i, y2=curr+i+1, val=(y1+'-'+y2); var opt=document.createElement("option"); opt.value=val; opt.textContent=val; yearSelect.appendChild(opt); }
        populateMandiDatalist(data.mandis||[]);
        if(data.needsSeason){ document.getElementById("seasonType").disabled=false; document.getElementById("seasonYear").disabled=false; document.getElementById("seasonRate").disabled=false; }
        else { document.getElementById("seasonType").value=data.seasonData.season||""; document.getElementById("seasonYear").value=data.seasonData.year||""; document.getElementById("seasonRate").value=data.seasonData.rate||""; document.getElementById("seasonType").disabled=true; document.getElementById("seasonYear").disabled=true; document.getElementById("seasonRate").disabled=true; }
        updatePartAHeadline();

        mandiListData = data.mandis||[];
        firmMapData   = data.mandiFirms||{};
      } else { document.getElementById("status").textContent="❌ Failed to load data."; }
    }catch(e){ document.getElementById("status").textContent="❌ Error loading Mandi list."; }
  }

  async function loadDataForReport(){
    try{
      var res=await fetch(SCRIPT_URL); var data=await res.json();
      if(data.success){
        mandiDates   = data.mandiDates || {};
        allMandis    = (data.mandis || []);
        inspectorMap = data.inspectorMap || {};
        allMandisCached = allMandis.slice();

        var rm=document.getElementById("reportMandi");
        rm.innerHTML = '<option value="ALL">ALL</option>' + allMandis.map(m=>'<option value="'+m+'">'+m+'</option>').join('');

        if(data.seasonData){
          currentSeason.season = data.seasonData.season || '';
          currentSeason.year   = data.seasonData.year   || '';
          currentSeason.rate   = Number(data.seasonData.rate || 0);
        }
        updateReportDates();
        populateInspectorMandiDropdown();
        populateInspectorDropdown();
        applyInspectorLockUI();

        document.getElementById("seasonType").addEventListener("change", updatePartAHeadline);
        document.getElementById("seasonYear").addEventListener("change", updatePartAHeadline);
        updatePartAHeadline();
      } else {
        document.getElementById("reportStatus").textContent="❌ Could not load report metadata.";
      }
    }catch(err){
      document.getElementById("reportStatus").textContent="❌ Error loading report metadata.";
    }
  }

  function updateReportDates(){
    var mandi=document.getElementById("reportMandi").value;
    var dd=document.getElementById("reportDate");
    dd.innerHTML="";

    var dates=[];
    if(mandi==="ALL"){
      var keys=Object.keys(mandiDates||{});
      for (var i=0;i<keys.length;i++){
        var arr=mandiDates[keys[i]]||[];
        for (var j=0;j<arr.length;j++) dates.push(arr[j]);
      }
    } else {
      dates = mandiDates[mandi] || [];
    }

    var uniq={}, isoDates=[];
    for (var i=0;i<dates.length;i++){ uniq[dates[i]]=true; }
    isoDates = Object.keys(uniq).sort();

    var optAll=document.createElement("option");
    optAll.value="ALL"; optAll.textContent="ALL";
    dd.appendChild(optAll);

    if(!isoDates.length){
      var o=document.createElement("option");
      o.value=""; o.textContent="No dates available";
      dd.appendChild(o);
      return;
    }

    isoDates.forEach(function(iso){
      var o=document.createElement("option");
      o.value=iso;
      o.textContent=formatDate(iso);
      dd.appendChild(o);
    });
  }
  document.addEventListener('change', function(evt){
    if (evt.target && evt.target.id === 'reportMandi') updateReportDates();
  });

  async function fetchAllRowsForMandi(mandi){
    var url=SCRIPT_URL+'?action=getReport&date=ALL&mandi='+encodeURIComponent(mandi);
    var res=await fetch(url); var data=await res.json();
    if(!data.success) return { rows:[], rate:currentSeason.rate||0 };
    return { rows:data.rows||[], rate:Number(data.rate||currentSeason.rate||0) };
  }

  var useRangeEl=document.getElementById("useRange"), rangeWrap=document.getElementById("rangeWrap"), rangeWrap2=document.getElementById("rangeWrap2"), fromDateEl=document.getElementById("fromDate"), toDateEl=document.getElementById("toDate");
  const ymd=d=>(d||'').trim(); const ymdCmp=(a,b)=> a===b?0: (a<b?-1:1);
  const isWithinInclusive=(d,from,to)=>{ if(from && ymdCmp(d,from)<0) return false; if(to && ymdCmp(d,to)>0) return false; return true; };
  useRangeEl.addEventListener("change", function(){
    var on=useRangeEl.checked; rangeWrap.style.display=on?"block":"none"; rangeWrap2.style.display=on?"block":"none";
    if(on){
      var mandi=document.getElementById("reportMandi").value; var dates=[];
      if(mandi==="ALL"){
        var keys=Object.keys(mandiDates||{});
        for (var i=0;i<keys.length;i++){
          var arr=mandiDates[keys[i]]||[];
          for (var j=0;j<arr.length;j++) dates.push(arr[j]);
        }
      } else { dates=mandiDates[mandi]||[]; }
      var uniq={}, iso=[];
      for(var k=0;k<dates.length;k++) uniq[dates[k]]=true;
      iso = Object.keys(uniq).sort();
      if(iso.length){ fromDateEl.value=iso[0]; toDateEl.value=iso[iso.length-1]; }
      else { fromDateEl.value=''; toDateEl.value=''; }
    }
  });
  document.getElementById("reportMandi").addEventListener("change", function(){ if(useRangeEl.checked) useRangeEl.dispatchEvent(new Event("change")); });

  function sumPerType(rows){ var m=new Map(); rows.forEach(r=>{ var t=normalizeType(r.bardana||''); var b=Number(r.bags||0); if(!t) return; m.set(t,(m.get(t)||0)+b); }); return m; }
  function mergeTypeMaps(a,b){ var out=new Map(a); b.forEach((v,k)=>{ out.set(k,(out.get(k)||0)+v); }); return out; }
  function orderedTypesUpToDate(totalsMap,date){
    var datesAsc=Array.from(totalsMap.keys()).sort(); var set=new Set();
    datesAsc.forEach(dk=>{ if(dk<=date){ var tm=totalsMap.get(dk); var per=(tm && tm.perType) ? tm.perType : new Map(); Array.from(per.keys()).forEach(t=>{ if(!set.has(t)) set.add(t); }); } });
    return Array.from(set);
  }

  async function generateReport(){
    var selectedDate=document.getElementById("reportDate").value; var mandiSel=document.getElementById("reportMandi").value;
    if(!mandiSel){ document.getElementById("reportStatus").textContent="Select mandi"; return; }
    document.getElementById("reportStatus").textContent="Loading report…";

    try{
      var mandiBlocks=[];
      if(mandiSel==="ALL"){
        var perMandi=[];
        for(var i=0;i<allMandis.length;i++){
          var m=allMandis[i]; var x=await fetchAllRowsForMandi(m);
          x.rows.sort((a,b)=>{ var c=(a.date||"").localeCompare(b.date||""); if(c!==0) return c; return (a.firm||"").localeCompare(b.firm||""); });
          perMandi.push({mandi:m, rows:x.rows, rate:x.rate});
        }
        perMandi.sort((a,b)=> (a.mandi||"").localeCompare(b.mandi||""));
        perMandi.forEach(function(md){
          var byDate=new Map();
          md.rows.forEach(r=>{ var k=r.date||""; if(!byDate.has(k)) byDate.set(k,[]); byDate.get(k).push(r); });
          var datesAsc=Array.from(byDate.keys()).sort();
          var datesToRender;
          if(useRangeEl.checked){
            var from=ymd(fromDateEl.value), to=ymd(toDateEl.value);
            if(from && to && ymdCmp(from,to)>0){ document.getElementById("reportStatus").textContent="❌ From-date is after To-date."; document.getElementById("reportOutput").innerHTML=""; mandiBlocks=[]; return; }
            datesToRender=datesAsc.filter(d=>isWithinInclusive(d,from,to));
          } else { datesToRender = (selectedDate==="ALL") ? datesAsc : datesAsc.filter(d=>d===selectedDate); }
          if(!datesToRender.length) return;
          var blocks=datesToRender.map(dk=>({ date:dk, rows:(byDate.get(dk)||[]), rate:md.rate, totalsMap:byDate }));
          var firstRenderedDate=datesToRender[0];
          var snStart = firstRenderedDate ? (md.rows.filter(r=> (r.date||"")<firstRenderedDate).length + 1) : 1;
          var pageStartIndex = firstRenderedDate ? datesAsc.findIndex(d=> d===firstRenderedDate) : 0;
          mandiBlocks.push({ mandi:md.mandi, blocks:blocks, snStart:snStart, pageStartIndex:pageStartIndex });
        });
      } else {
        var fetched = await fetchAllRowsForMandi(mandiSel);
        var allRows = fetched.rows, rate0 = fetched.rate;
        allRows.sort((a,b)=>{ var c=(a.date||"").localeCompare(b.date||""); if(c!==0) return c; return (a.firm||"").localeCompare(b.firm||""); });
        var byDate=new Map(); allRows.forEach(r=>{ var k=r.date||""; if(!byDate.has(k)) byDate.set(k,[]); byDate.get(k).push(r); });
        var datesAsc=Array.from(byDate.keys()).sort();
        var datesToRender;
        if(useRangeEl.checked){
          var from2=ymd(fromDateEl.value), to2=ymd(toDateEl.value);
          if(from2 && to2 && ymdCmp(from2,to2)>0){ document.getElementById("reportStatus").textContent="❌ From-date is after To-date."; document.getElementById("reportOutput").innerHTML=""; return; }
          datesToRender=datesAsc.filter(d=>isWithinInclusive(d,from2,to2));
        } else { datesToRender = (selectedDate==="ALL") ? datesAsc : datesAsc.filter(d=>d===selectedDate); }
        if(!datesToRender.length){ document.getElementById("reportStatus").textContent="No data to display."; document.getElementById("reportOutput").innerHTML=""; return; }
        var blocks=datesToRender.map(dk=>({ date:dk, rows:(byDate.get(dk)||[]), rate:rate0, totalsMap:byDate }));
        var firstRenderedDate2=datesToRender[0];
        var snStart2 = firstRenderedDate2 ? (allRows.filter(r=> (r.date||"")<firstRenderedDate2).length + 1) : 1;
        var pageStartIndex2 = firstRenderedDate2 ? datesAsc.findIndex(d=> d===firstRenderedDate2) : 0;
        mandiBlocks.push({ mandi:mandiSel, blocks:blocks, snStart:snStart2, pageStartIndex:pageStartIndex2 });
      }

      if(!mandiBlocks.length){ document.getElementById("reportStatus").textContent="No data to display."; document.getElementById("reportOutput").innerHTML=""; return; }

      var totalsByMandi=new Map();
      async function ensureTotalsForMandi(m){
        if(totalsByMandi.has(m)) return;
        var d = await fetchAllRowsForMandi(m);
        var rows=d.rows, rate=d.rate;
        rows.sort((a,b)=> (a.date||"").localeCompare(b.date||""));
        var byDate=new Map(); rows.forEach(r=>{ var k=r.date||""; if(!byDate.has(k)) byDate.set(k,[]); byDate.get(k).push(r); });
        var tMap=new Map();
        Array.from(byDate.keys()).sort().forEach(dk=>{
          var b=0,w=0,a=0; var dayRows=byDate.get(dk)||[]; var perType=sumPerType(dayRows);
          dayRows.forEach(r=>{ var bb=Number(r.bags||0), ww=Number(r.weight||0), aa=ww*rate; b+=bb; w+=ww; a+=aa; });
          tMap.set(dk,{b:b,w:w,a:a,perType:perType});
        });
        totalsByMandi.set(m,{ totals:tMap, rate:rate });
      }

      var html="";
      for(var mi=0; mi<mandiBlocks.length; mi++){
        var mandiName=mandiBlocks[mi].mandi; var isFirstMandi=(mi===0);
        await ensureTotalsForMandi(mandiName);
        var obj=totalsByMandi.get(mandiName); var totals=obj.totals, mandiRate=obj.rate;

        var assigned = inspectorMap[mandiName] || { inspector:'INSPECTOR', signUrl:'' };
        var inspectorName = assigned.inspector || 'INSPECTOR';
        var signUrl = assigned.signUrl || '';

        html += '<div class="mandi-section '+(isFirstMandi?'first':'')+'">';
        var sn=mandiBlocks[mi].snStart||1; var pageNo=mandiBlocks[mi].pageStartIndex||0;

        var prevCarry={B:0,W:0,A:0, perType:new Map()};
        var sortedAllDates=Array.from(totals.keys()).sort();
        var firstRenderedDate=(mandiBlocks[mi].blocks && mandiBlocks[mi].blocks.length)?mandiBlocks[mi].blocks[0].date:null;
        if(firstRenderedDate){
          sortedAllDates.forEach(function(dk){
            if(dk<firstRenderedDate){
              var t=totals.get(dk)||{b:0,w:0,a:0, perType:new Map()};
              prevCarry.B+=Number(t.b||0); prevCarry.W+=Number(t.w||0); prevCarry.A+=Number(t.a||0);
              (t.perType||new Map()).forEach((v,k)=>{ prevCarry.perType.set(k,(prevCarry.perType.get(k)||0)+Number(v||0)); });
            }
          });
        }

        for(var bi=0; bi<mandiBlocks[mi].blocks.length; bi++){
          var blk=mandiBlocks[mi].blocks[bi]; pageNo++;
          var dayRows=blk.rows; var rate=(typeof blk.rate==='number'?blk.rate:mandiRate||0);
          var typesForHeader=orderedTypesUpToDate(totals, blk.date);

          var todayB=0,todayW=0,todayA=0; var todayTypeMap=sumPerType(dayRows);
          dayRows.forEach(r=>{ var b=Number(r.bags||0), w=Number(r.weight||0), a=w*rate; todayB+=b; todayW+=w; todayA+=a; });

          var prevTypeMap=new Map(prevCarry.perType);
          var uptoB=prevCarry.B+todayB, uptoW=prevCarry.W+todayW, uptoA=prevCarry.A+todayA;
          var uptoTypeMap=mergeTypeMaps(prevTypeMap, todayTypeMap);

          html += '<div class="report-block '+(bi===0?'first-in-section':'')+'">'
            + '<div class="report-header">'
            +   '<div class="report-left" style="font-size:28px; line-height:1.6;">'
            +     '<strong>MANDI NAME :- '+mandiName+'</strong><br>'
            +     '<strong>DISTT :- FEROZEPUR</strong>'
            +   '</div>'
            +   '<div class="report-title" style="font-size:36px; line-height:1.3; text-align:center; text-decoration:underline;">'
            +     'PG-4 REPORT<br>('
            +       String((currentSeason.season||"").toUpperCase()) + ' ' + String(currentSeason.year||"")
            + ')'
            +   '</div>'
            +   '<div class="report-right" style="font-size:28px; line-height:1.6; text-align:right;">'
            +     '<div><strong>DATE: '+formatDate(blk.date)+'</strong></div>'
            +     '<div><strong>Page No. '+pageNo+'</strong></div>'
            +     '<div><strong>Rate @ = '+mandiRate+'</strong></div>'
            +   '</div>'
            + '</div>'
            + '<div class="report-gap-row"></div>'
            + '<table class="report-table" style="width:100%; border-collapse:collapse;">'
            +   '<thead><tr>'
            +     '<th class="col-sn">S.No.</th>'
            +     '<th class="col-date">Date</th>'
            +     '<th class="col-firm">Firm Name</th>'
            +     typesForHeader.map(t=>'<th class="col-num">'+t+'<br>(Bags)</th>').join('')
            +     '<th class="col-num">TOTAL<br>(Bags)</th>'
            +     '<th class="col-num">TOTAL<br>(Weight)</th>'
            +     '<th class="col-amt">Amount</th>'
            +   '</tr></thead><tbody>';

          dayRows.forEach(function(r){
            var b=Number(r.bags||0), w=Number(r.weight||0), a=w*rate; var t=normalizeType(r.bardana||'');
            var perTypeCells=typesForHeader.map(h=> (h===t? b: 0));
            html += '<tr>'
              + '<td class="col-sn">'+ (sn++) +'</td>'
              + '<td class="col-date">'+ formatDate(r.date||"") +'</td>'
              + '<td class="cell-firm">'+ (r.firm||"") +'</td>'
              + perTypeCells.map(v=>'<td class="col-num">'+v+'</td>').join('')
              + '<td class="col-num">'+ b +'</td>'
              + '<td class="col-num">'+ (currentSeason.season==="Paddy Season"?Number(w).toFixed(3):Number(w).toFixed(2)) +'</td>'
              + '<td class="col-amt">'+ (a).toFixed(2)+'</td>'
            + '</tr>';
          });

          var todayTypeCells=typesForHeader.map(t=> '<td class="col-num">'+(todayTypeMap.get(t)||0)+'</td>').join('');
          var prevTypeCells =typesForHeader.map(t=> '<td class="col-num">'+(prevTypeMap.get(t)||0)+'</td>').join('');
          var uptoTypeCells =typesForHeader.map(t=> '<td class="col-num">'+(uptoTypeMap.get(t)||0)+'</td>').join('');

          html += '<tr style="background:#f9f9f9;"><td colspan="3"><strong>Today\'s Total</strong></td>'+todayTypeCells+'<td>'+todayB+'</td><td>'+(currentSeason.season==="Paddy Season"?todayW.toFixed(3):todayW.toFixed(2))+'</td><td>'+todayA.toFixed(2)+'</td></tr>'
            + '<tr style="background:#fffbe6;"><td colspan="3"><strong>Total up to Previous Day</strong></td>'+prevTypeCells+'<td>'+prevCarry.B+'</td><td>'+(currentSeason.season==="Paddy Season"?prevCarry.W.toFixed(3):prevCarry.W.toFixed(2))+'</td><td>'+prevCarry.A.toFixed(2)+'</td></tr>'
            + '<tr style="background:#e6ffe6;"><td colspan="3"><strong>Grand Total up to Today</strong></td>'+uptoTypeCells+'<td>'+uptoB+'</td><td>'+(currentSeason.season==="Paddy Season"?uptoW.toFixed(3):uptoW.toFixed(2))+'</td><td>'+uptoA.toFixed(2)+'</td></tr>'
            + '</tbody></table>'
            + '<div class="stamp">'
              + '<img class="sign-img" src="" alt="Inspector Signature" data-sign-url="'+(signUrl||'')+'" />'
              + '<span class="stamp-line inspector-name">'+inspectorName+'</span>'
              + '<span class="stamp-line inspector-title">INSPECTOR, PUNGRAIN</span>'
              + '<span class="stamp-line mandi-name">'+mandiName+'</span>'
            + '</div>'
            + '<div class="after-report-gap"></div>'
            + '<hr class="full-divider">'
          + '</div>';

          prevCarry={ B:uptoB, W:uptoW, A:uptoA, perType:uptoTypeMap };
        }
        html += '</div>';
      }

      document.getElementById("reportOutput").innerHTML=html; 
      document.getElementById("reportStatus").textContent="";

      const imgs = Array.prototype.slice.call(document.querySelectorAll('#reportOutput img.sign-img'));
      await Promise.all(imgs.map(async (img)=>{
        const sharedUrl = img.getAttribute('data-sign-url')||'';
        try{
          if(sharedUrl){
            const id = (function(u){ try{u=new URL(u); const m=u.pathname.match(/\/file\/d\/([^/]+)/); if(m&&m[1]) return m[1]; const q=u.searchParams.get('id'); if(q) return q;}catch(e){} return ''; })(sharedUrl);
            if(id){
              const res = await fetch(SCRIPT_URL + '?action=signData&id=' + encodeURIComponent(id));
              const js = await res.json();
              if(js && js.dataUrl){ img.src = js.dataUrl; }
            }
          }
        }catch(e){ console.warn('Signature load failed', e); }
      }));

    }catch(err){
      console.error(err);
      document.getElementById("reportStatus").textContent="Error fetching report";
      document.getElementById("reportOutput").innerHTML="";
    }
  }

  function resetReportSection(){
    document.getElementById("reportStatus").textContent = "";
    document.getElementById("reportOutput").innerHTML  = "";
    document.getElementById("useRange").checked = false;
    document.getElementById("rangeWrap").style.display = "none";
    document.getElementById("rangeWrap2").style.display= "none";
    document.getElementById("fromDate").value = "";
    document.getElementById("toDate").value   = "";
    updateReportDates();
  }

  async function downloadPDF(){
    var jsPDF = window.jspdf.jsPDF;
    var container=document.getElementById("reportOutput");
    if(!container || !container.innerHTML.trim()){ alert("⚠️ Generate a report first."); return; }
    var pdf=new jsPDF("p","mm","a4");
    var pageWidth=pdf.internal.pageSize.getWidth(), pageHeight=pdf.internal.pageSize.getHeight();
    var margin=10, maxW=pageWidth-margin*2, maxH=pageHeight-margin*2, gapMM=4;
    var blocks=Array.prototype.slice.call(container.querySelectorAll(".report-block")); if(!blocks.length){ alert("⚠️ Nothing to export."); return; }
    var currentY=margin; var lastMandiEl=null;
    for(var i=0;i<blocks.length;i++){
      var el=blocks[i]; var thisMandiEl=el.closest(".mandi-section");
      if(lastMandiEl && thisMandiEl!==lastMandiEl){ pdf.addPage("a4","p"); currentY=margin; }
      lastMandiEl=thisMandiEl;
      var canvas=await html2canvas(el,{ scale:3, useCORS:true, allowTaint:true });
      var imgData=canvas.toDataURL("image/png");
      var imgW=maxW, imgH=canvas.height*(imgW/canvas.width);
      if(imgH>maxH){ imgH=maxH; imgW=canvas.width*(imgH/canvas.height); }
      if(currentY+imgH>pageHeight-margin){ pdf.addPage("a4","p"); currentY=margin; }
      var x=(pageWidth-imgW)/2; pdf.addImage(imgData,"PNG",x,currentY,imgW,imgH); currentY+=imgH+gapMM;
    }
    var mandi=document.getElementById("reportMandi").value || "ALL"; var onRange=document.getElementById("useRange").checked;
    var dateLabel; if(onRange){ dateLabel=((document.getElementById("fromDate").value||'start')+'_'+(document.getElementById("toDate").value||'end')); } else { dateLabel=document.getElementById("reportDate").value || "ALL"; }
    pdf.save('Report_'+mandi+'_'+dateLabel+'.pdf');
  }

  /* ===== Inspector Master ===== */
  function populateInspectorMandiDropdown(){
    var sel=document.getElementById("ai_mandi");
    var assignedKeys = inspectorMap ? Object.keys(inspectorMap) : [];
    var assigned={}; for (var i=0;i<assignedKeys.length;i++){ assigned[assignedKeys[i]]=true; }
    var options=(allMandisCached||[]).filter(m=> !assigned[m]);
    sel.innerHTML = options.length ? options.map(m=> '<option value="'+m+'">'+m+'</option>').join('') : '<option value="">(All mandis are assigned)</option>';
  }
  function buildInspectorMaster(){
    var master = {};
    try{
      var keys = Object.keys(inspectorMap||{});
      for(var i=0;i<keys.length;i++){
        var m = keys[i];
        var obj = inspectorMap[m];
        if(obj && obj.inspector){
          var name = String(obj.inspector).trim();
          if(name && !master[name]) master[name] = obj.signUrl || '';
        }
      }
    }catch(e){}
    return master;
  }
  function populateInspectorDropdown(){
    var sel = document.getElementById("ai_inspector_select");
    if(!sel) return;
    var master = buildInspectorMaster();
    var names = Object.keys(master).sort();
    sel.innerHTML = '<option value="">Select Inspector</option>' +
      names.map(n => '<option value="'+n+'">'+n+'</option>').join('') +
      '<option value="__manual__">— Manual entry —</option>';
  }
  function applyInspectorLockUI(){
    var assignedCount = Object.keys(inspectorMap||{}).length;
    var totalCount = (allMandisCached||[]).length;
    var lock = totalCount>0 && assignedCount >= totalCount;
    document.getElementById("ai_lock_note").style.display = lock ? 'block' : 'none';
    document.getElementById("ai_mandi").disabled = lock;
    document.getElementById("ai_inspector_select").disabled = lock;
    document.getElementById("ai_inspector_manual").disabled = lock;
    document.getElementById("ai_sign").disabled = lock;
    document.getElementById("ai_saveBtn").disabled = lock;
  }
  document.addEventListener('change', function(e){
    if(e.target && e.target.id === 'ai_inspector_select'){
      var v = e.target.value;
      var wrap = document.getElementById("ai_sign_wrap");
      var sign = document.getElementById("ai_sign");
      var manual = document.getElementById("ai_inspector_manual");
      if(v === '__manual__'){
        manual.style.display = 'block';
        manual.value = '';
        wrap.style.display = 'block';
        sign.value = '';
      } else if(v){
        manual.style.display = 'none';
        var master = buildInspectorMaster();
        var link = master[v] || '';
        wrap.style.display = link ? 'none' : 'block';
        sign.value = link || '';
      } else {
        manual.style.display = 'none';
        wrap.style.display = 'block';
        sign.value = '';
      }
    }
  });

  /* ===== New: Login/Register UI toggle & temp persist ===== */
  (function setupAuthToggle(){
    const toggle = document.getElementById('authModeToggle');
    const label  = document.getElementById('authModeLabel');
    const login  = document.getElementById('loginForm');
    const reg    = document.getElementById('registerForm');
    function render(){
      const isReg = !!toggle.checked;
      label.textContent = isReg ? 'Register' : 'Login';
      login.classList.toggle('hide', isReg);
      reg.classList.toggle('hide', !isReg);
    }
    if (toggle) {
      toggle.addEventListener('change', render);
      render();
    }
  })();

  (function enableTempPersist(){
    const els = document.querySelectorAll('[data-persist]');
    els.forEach(el => {
      const key = 'tmp:' + (el.id || el.name || el.dataset.key || '');
      if (!key) return;
      try {
        const val = sessionStorage.getItem(key);
        if (val !== null) el.value = val;
      } catch(e){}
      el.addEventListener('input', () => {
        try { sessionStorage.setItem(key, el.value); } catch(e){}
      });
    });
  })();

  
/* === ADDED: save profile to Users sheet (non-invasive) === */
async function saveUserProfileClient(extra = {}){
  try{
    const email =
      (document.getElementById('authEmail_reg')?.value ||
       document.getElementById('authEmail')?.value || '').trim();
    if(!email) return;
    const fields = {};
    document.querySelectorAll('[data-user-field]').forEach(el=>{
      const key = el.getAttribute('data-user-field');
      if (key) fields[key] = (el.value ?? '').toString().trim();
    });
    Object.assign(fields, extra||{});
    const fd = new FormData();
    fd.append('action','saveUserProfile');
    fd.append('email', email);
    fd.append('fields', JSON.stringify(fields));
    await fetch(SCRIPT_URL, { method:'POST', body: fd });
  }catch(e){}
}

// ===== INIT =====
  (async function init(){
    await ensureUserSheet();
    updateAuthUI();
        // ADDED: save profile after OTP activation
        await saveUserProfileClient({ Role: (document.getElementById('authRole_reg')?.value || 'DEO').toUpperCase() });

    await loadBardanaFromServer();
    if (document.getElementById('applyBardanaBtn')) document.getElementById('applyBardanaBtn').addEventListener('click', applyBardana);
    if (document.getElementById('changeBardanaBtn')) document.getElementById('changeBardanaBtn').addEventListener('click', changeBardana);

    loadMandisAndFirms();
    addRow();
    document.getElementById("mandiName").addEventListener("change", refreshFirmDatalists);

    // report init
    await loadDataForReport();
    document.getElementById("ai_saveBtn").addEventListener("click", saveInspectorAssignment);
  })();

/* ===== AUTH TOGGLE LABEL & FORM SWITCH ===== */
(function(){
  var t=document.getElementById('authModeToggle'),
      l=document.getElementById('authModeLabel'),
      lf=document.getElementById('loginForm'),
      rf=document.getElementById('registerForm');
  function apply(){
    if(!t) return;
    var isReg=!!t.checked;
    if(l) l.textContent=isReg?'Register':'Login';
    if(lf&&rf){
      lf.classList.toggle('hide', isReg);
      rf.classList.toggle('hide', !isReg);
    }
  }
  if(t){
    t.addEventListener('change', apply);
    if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', apply, {once:true});
    else apply();
  }
})();




  if (typeof updateAuthUI === 'function') {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', updateAuthUI, { once: true });
    } else {
      updateAuthUI();
        // ADDED: save profile after OTP activation
        await saveUserProfileClient({ Role: (document.getElementById('authRole_reg')?.value || 'DEO').toUpperCase() });
    }
  }



(function(){
  function ensureStatusEl(){
    var el = document.getElementById('ai_status');
    if(!el){
      el = document.createElement('div');
      el.id = 'ai_status';
      el.style.position='fixed';
      el.style.right='16px';
      el.style.bottom='16px';
      el.style.padding='8px 12px';
      el.style.background='#fff';
      el.style.border='1px solid #ccc';
      el.style.borderRadius='6px';
      el.style.boxShadow='0 2px 8px rgba(0,0,0,.12)';
      el.style.zIndex='99999';
      el.style.maxWidth='60vw';
      el.style.font='14px/1.4 system-ui, sans-serif';
      el.textContent = 'Status';
      document.body.appendChild(el);
    }
    return el;
  }

  function setStatus(msg){
    try{
      console.log('[Inspector Save] ' + msg);
      ensureStatusEl().textContent = msg;
    }catch(e){}
  }

  // Wrap existing saveInspectorAssignment if present
  var originalSave = window.saveInspectorAssignment;
  window.saveInspectorAssignment = async function(){
    setStatus('Clicked Save…');
    if (typeof SCRIPT_URL === 'undefined' || !SCRIPT_URL){
      setStatus('❌ SCRIPT_URL is not defined in this page.');
      alert('SCRIPT_URL is not defined. Please set your Apps Script web app URL.');
      return;
    }
    try{
      if (originalSave){
        console.log('[Inspector Save] Calling existing saveInspectorAssignment()');
        return await originalSave();
      }
      // Minimal compatible implementation
      const statusEl = document.getElementById('ai_status');
      const mandiSel = document.getElementById('ai_mandi');
      const inspSel  = document.getElementById('ai_inspector_select');
      const inspManualEl = document.getElementById('ai_inspector_manual');
      const signEl   = document.getElementById('ai_sign');

      const mandi = (mandiSel && mandiSel.value || '').trim();
      const chosen = (inspSel && inspSel.value || '').trim();
      const inspector = (chosen === '__manual__' ? ((inspManualEl && inspManualEl.value) || '').trim() : chosen);
      const signUrl = (signEl && signEl.value || '').trim();

      if(!mandi){ setStatus('❌ Select a mandi.'); return; }
      if(!inspector){ setStatus('❌ Select or type inspector name.'); return; }

      setStatus('Saving to Apps Script…');
      const fd = new FormData();
      fd.append('action','assignInspector');
      fd.append('mandi',mandi);
      fd.append('inspector',inspector);
      fd.append('signUrl',signUrl);

      const res = await fetch(SCRIPT_URL, { method:'POST', body: fd });
      const text = await res.text();
      console.log('[Inspector Save] Raw response:', text);
      let js = {};
      try{ js = JSON.parse(text) }catch(e){}

      if (js && js.ok){
        window.inspectorMap = window.inspectorMap || {};
        window.inspectorMap[mandi] = { inspector: inspector, signUrl: signUrl };
        if (typeof populateInspectorMandiDropdown === 'function') populateInspectorMandiDropdown();
        if (typeof populateInspectorDropdown === 'function')       populateInspectorDropdown();
        if (typeof applyInspectorLockUI === 'function')            applyInspectorLockUI();
        if (mandiSel) mandiSel.value='';
        if (inspSel)  inspSel.value='';
        if (inspManualEl){ inspManualEl.value=''; inspManualEl.style.display='none'; }
        if (signEl) signEl.value='';
        setStatus('✅ ' + (js.msg || 'Saved'));
      } else {
        setStatus('❌ Save failed: ' + ((js && js.msg) ? js.msg : text || 'Unknown error'));
        alert('Save failed. Server said: ' + ((js && js.msg) ? js.msg : text));
      }
    } catch(err){
      console.error(err);
      setStatus('❌ Error: ' + err);
      alert('Network or script error: ' + err);
    }
  };

  function bind(){
    var btn = document.getElementById('ai_saveBtn');
    if (btn){
      btn.addEventListener('click', function(e){
        e.preventDefault();
        window.saveInspectorAssignment();
      }, { once:false });
      console.log('[Inspector Save] Bound click to #ai_saveBtn');
      setStatus('Ready');
    } else {
      console.warn('[Inspector Save] Button #ai_saveBtn not found at DOMContentLoaded. Will observe DOM.');
      // Observe DOM in case button is rendered later
      const obs = new MutationObserver(function(){
        var b = document.getElementById('ai_saveBtn');
        if (b){
          b.addEventListener('click', function(e){
            e.preventDefault();
            window.saveInspectorAssignment();
          }, { once:false });
          console.log('[Inspector Save] Late-bound click to #ai_saveBtn');
          setStatus('Ready');
          obs.disconnect();
        }
      });
      obs.observe(document.documentElement || document.body, { childList:true, subtree:true });
    }
  }

  if (document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', bind);
  } else {
    bind();
  }
})();

document.addEventListener('DOMContentLoaded', function(){
  var btn = document.getElementById('ai_saveBtn');
  if(btn && !btn.getAttribute('data-bound')){
    btn.addEventListener('click', saveInspectorAssignment);
    btn.setAttribute('data-bound','1');
  }
});



/* ADDED: handle BFCache so logged-out users don't see prior page */
window.addEventListener('pageshow', function (evt) {
  const navEntries = performance.getEntriesByType && performance.getEntriesByType('navigation');
  const nav = navEntries && navEntries[0];
  const isBackForward = evt.persisted || (nav && nav.type === 'back_forward');
  try {
    const auth = (typeof getAuth === 'function') ? getAuth() : { verified:false };
    if (!auth.verified && isBackForward) {
      location.replace(location.href.split('#')[0]);
    }
  } catch(e){}
});
