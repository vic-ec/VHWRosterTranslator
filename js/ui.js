// ═══════════════════════════════════════════════════════════════
// ui.js — UI helpers, state management, file handling, and
//         preview/edit table.
//
//   setStatus, unlock, getMonthYear, hasLeaveInShifts
//   updateLeaveFields, checkDetailsComplete, checkReady
//   normaliseTime, rebuildMonthDropdown, markDirty, markClean
//   saveDetailsToState, restoreDetailsToForm
//   addFiles, removeFile, renderFileList, mergeAndRefresh
//   fullReset, clearDoctorSelection, buildDoctorGrid
//   autoDetectMonth, buildPreview, makeRowInner
//   attachEditHandlers, getFormDetails, readFile
//
// Depends on: config.js, holidays.js, parser.js,
//             generator-excel.js, generator-docx.js
// ═══════════════════════════════════════════════════════════════

// Wire clear-doctor button
document.addEventListener('DOMContentLoaded',()=>{
  const cdb=$('clearDoctorBtn');
  if(cdb) cdb.addEventListener('click',clearDoctorSelection);
});

function setStatus(msg,type='') {
  const el=$('parseStatus'); if(!el) return;
  el.textContent=msg;
  // Apply colour styling via inline style on the span
  if(!msg){ el.style.display='none'; return; }
  el.style.display='inline';
  if(type==='error') el.style.color='var(--warn)';
  else if(type==='success') el.style.color='var(--success)';
  else if(type==='info') el.style.color='var(--accent-mid)';
  else el.style.color='var(--text-muted)';
}
function unlock(el){el.style.opacity='1';el.style.pointerEvents='';}
function getMonthYear(){
  const m=$('monthSelect').value,y=parseInt($('yearInput').value);
  return {month:m===''?null:parseInt(m),year:isNaN(y)?null:y};
}
function hasLeaveInShifts(){
  if(!state.editedShifts) return false;
  return Object.values(state.editedShifts).some(es=>{
    const lbl=es.typeLabel||'';
    return lbl && !lbl.startsWith('WD Shift') && !lbl.startsWith('WE Shift');
  });
}
function updateLeaveFields(){
  const sec=$('leaveFieldsSection');
  if(sec) sec.style.display=hasLeaveInShifts()?'':'none';
  checkDetailsComplete();
}
function checkDetailsComplete() {
  const first=$('detailFirstName').value.trim();
  const surname=$('detailSurname').value.trim();
  const persal=$('detailPersal').value.trim();
  const desSel=$('detailDesignationSel').value;
  const desOther=$('detailDesignationOther')?.value.trim();
  const designation=desSel==='other'?(desOther||''):desSel;
  const date=$('detailSigDate').value.trim();
  const dateValid=date.length===10&&/^\d{2}\/\d{2}\/\d{4}$/.test(date);
  const leaveVisible=$('leaveFieldsSection')?.style.display!=='none';
  const address=$('detailAddress')?.value.trim()||'';
  const leaveOk=!leaveVisible||(address.length>0);
  const show=!!(first&&surname&&persal&&designation&&dateValid&&leaveOk);
  $('proceedDownloadBtn').style.display=show?'':'none';
  $('annexureCBtn').style.display=show?'':'none';
  $('z1aBtn').style.display=(show&&hasLeaveInShifts())?'':'none';
}
function checkReady(){
  const {month,year}=getMonthYear();
  const nameVal=$('employeeName').value.trim();
  const ok=!!state.selectedDoctor&&month!==null&&!!year;
  const allFilled=ok&&!!nameVal;
  const btn=$('previewBtn');
  btn.disabled=!ok;
  // Fix 1: green when all filled, secondary style otherwise
  if(allFilled){
    btn.className='btn btn-success btn-preview';
  } else {
    btn.className='btn btn-secondary btn-preview';
  }
}
function normaliseTime(val) {
  if(!val) return null;
  const m=val.match(/^(\d{1,2})[H:](\d{2})$/i)||val.match(/^(\d{2})(\d{2})$/);
  if(!m) return null;
  const h=parseInt(m[1]),min=parseInt(m[2]);
  if(h>23||min>59) return null;
  return String(h).padStart(2,'0')+'H'+String(min).padStart(2,'0');
}

function rebuildMonthDropdown() {
  const sel=$('monthSelect');
  const current=sel.value;
  while(sel.options.length>1) sel.remove(1);
  const allMonthNames=['January','February','March','April','May','June',
    'July','August','September','October','November','December'];
  const sorted=[...state.availableMonths].sort((a,b)=>a-b);
  for(const m of sorted){
    const opt=document.createElement('option');
    opt.value=m; opt.textContent=allMonthNames[m];
    sel.appendChild(opt);
  }
  if(state.availableMonths.has(parseInt(current))) sel.value=current;
  else if(sorted.length===1) sel.value=sorted[0];
  checkReady();
}

function markDirty(day) {
  // Data already captured in state.editedShifts — auto-clean immediately
  markClean(day);
}
function markClean(day) {
  state.dirtyDays.delete(day);
  const row=document.querySelector(`tr[data-day="${day}"]`);
  if(!row) return;
  const ac=row.querySelector('.action-cell');
  if(ac){ac.innerHTML=`<button class="row-clear" data-day="${day}" title="Remove">&times;</button>${state.originalShifts[day]?`<button class="row-undo" data-day="${day}" title="Undo changes">&#8635;</button>`:''}`; attachEditHandlers();}
}

function saveDetailsToState() {
  state.savedDetails.firstName=$('detailFirstName').value.trim();
  state.savedDetails.surname=$('detailSurname').value.trim();
  state.savedDetails.persal=$('detailPersal').value.trim();
  // Fix 3: supervisor — use dropdown value, or 'other' text input
  const supSel=$('detailSupervisorSel').value;
  state.savedDetails.supervisor=supSel==='other'?$('detailSupervisorOther').value.trim():supSel;
  state.savedDetails.sigDate=$('detailSigDate').value.trim();
  const desSel=$('detailDesignationSel').value;
  state.savedDetails.designation=desSel==='other'?$('detailDesignationOther').value.trim():desSel;
  state.savedDetails.designationOther=desSel==='other'?$('detailDesignationOther').value.trim():'';
  state.savedDetails.address=$('detailAddress')?.value.trim()||'';
  state.savedDetails.shiftWorker=$('detailShiftWorker')?.value||'yes';
  state.savedDetails.casualEmployee=$('detailCasualEmployee')?.value||'no';
}
function restoreDetailsToForm(isNewDoctor) {
  if(isNewDoctor) {
    $('detailFirstName').value=''; $('detailSurname').value=state.selectedDoctor||'';
    $('detailPersal').value=''; $('detailSupervisorSel').value=''; $('detailSupervisorOther').value=''; $('detailSupervisorOther').style.display='none'; $('detailSigDate').value=''; $('detailDesignationSel').value=''; $('detailDesignationOther').value=''; $('detailDesignationOther').style.display='none';
    state.savedDetails={firstName:'',surname:state.selectedDoctor||'',persal:'',supervisor:'',sigDate:''};
  } else {
    $('detailFirstName').value=state.savedDetails.firstName||'';
    $('detailSurname').value=state.savedDetails.surname||state.selectedDoctor||'';
    $('detailPersal').value=state.savedDetails.persal||'';
    // Fix 3: restore supervisor dropdown + other
    const saved=state.savedDetails.supervisor||'';
    const knownSups=['Philip Cloete','Sebastian De Haan','Paul Xafis'];
    if(knownSups.includes(saved)){
      $('detailSupervisorSel').value=saved; $('detailSupervisorOther').style.display='none';
    } else if(saved){
      $('detailSupervisorSel').value='other'; $('detailSupervisorOther').style.display='';
      $('detailSupervisorOther').value=saved;
    } else {
      $('detailSupervisorSel').value=''; $('detailSupervisorOther').style.display='none';
    }
    $('detailSigDate').value=state.savedDetails.sigDate||'';
    const _sd=state.savedDetails.sigDate||'';
    const _sdm=_sd.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if(_sdm) $('detailSigDatePicker').value=`${_sdm[3]}-${_sdm[2]}-${_sdm[1]}`; else $('detailSigDatePicker').value='';
    // Restore designation
    const savedDes=state.savedDetails.designation||'';
    const knownDes=['Intern','Community Service Medical Officer','Medical Officer Grade 1','Medical Officer Grade 2','Medical Officer Grade 3','Registrar','Medical Specialist Grade 1','Medical Specialist Grade 2','Medical Specialist Grade 3'];
    if(knownDes.includes(savedDes)){$('detailDesignationSel').value=savedDes;$('detailDesignationOther').style.display='none';}
    else if(savedDes){$('detailDesignationSel').value='other';$('detailDesignationOther').style.display='';$('detailDesignationOther').value=savedDes;}
    else{$('detailDesignationSel').value='';$('detailDesignationOther').style.display='none';}
  }
  checkDetailsComplete();
}

// === FILE HANDLING ===
function addFiles(files) {
  for(const f of files){ if(!state.pendingFiles.some(p=>p.name===f.name)) state.pendingFiles.push(f); }
  renderFileList();
  if(state.pendingFiles.length>0){
    $('parseBtn').disabled=false;$('clearBtn').style.display='';
    setStatus(state.pendingFiles.length+' file(s) queued \u2014 click Extract Data','info');
  }
}
function removeFile(name){
  state.pendingFiles=state.pendingFiles.filter(f=>f.name!==name);
  state.parsedFiles=state.parsedFiles.filter(f=>f.name!==name);
  renderFileList();mergeAndRefresh();
  if(!state.pendingFiles.length&&!state.parsedFiles.length){
    // Only fully disable if no consultant file is also queued
    if(!state.consultantFiles?.length){
      $('parseBtn').disabled=true;$('clearBtn').style.display='none';
      setStatus('');
    } else {
      setStatus(state.consultantFiles.length+' file(s) queued — click Extract Data','info');
    }
    rosterList.style.display='none';
  }
}
function renderFileList(){
  const all=[...state.parsedFiles.map(f=>({name:f.name,days:f.days.length,parsed:true})),
             ...state.pendingFiles.map(f=>({name:f.name,days:null,parsed:false}))];
  if(!all.length){rosterList.style.display='none';return;}
  rosterList.style.display='';
  rosterList.innerHTML=all.map(f=>`
    <div class="roster-item">
      <span class="ri-name">${f.name}</span>
      ${f.parsed?`<span class="ri-days">${f.days} days</span>`:`<span class="ri-days" style="color:var(--text-faint)">queued</span>`}
      <button class="ri-remove" data-name="${f.name}">&times;</button>
    </div>`).join('');
  rosterList.querySelectorAll('.ri-remove').forEach(btn=>btn.addEventListener('click',()=>removeFile(btn.dataset.name)));
}
function mergeAndRefresh(){
  const allDays=state.parsedFiles.flatMap(f=>f.days);
  const allDocs=new Set(); state.parsedFiles.forEach(f=>f.doctors.forEach(d=>allDocs.add(d)));
  state.rosterData={days:allDays,doctors:allDocs};
  state.availableMonths=new Set(allDays.map(d=>d.month));
  if(allDays.length){buildDoctorGrid(allDocs);autoDetectMonth();rebuildMonthDropdown();step2.style.display='';unlock(step2);}
}

rosterInput.addEventListener('change',e=>{if(e.target.files.length) addFiles(Array.from(e.target.files));rosterInput.value='';});
rosterZone.addEventListener('dragover',e=>{e.preventDefault();rosterZone.classList.add('drag-over');});
rosterZone.addEventListener('dragleave',()=>rosterZone.classList.remove('drag-over'));
rosterZone.addEventListener('drop',e=>{e.preventDefault();rosterZone.classList.remove('drag-over');if(e.dataTransfer.files.length) addFiles(Array.from(e.dataTransfer.files));});

$('clearBtn').addEventListener('click',()=>{
  state.pendingFiles=[];state.parsedFiles=[];state.rosterData=null;state.selectedDoctor=null;
  state.editedShifts={};state.originalShifts={};state.dirtyDays.clear();state.availableMonths=new Set();
  state.savedDetails={firstName:'',surname:'',persal:'',supervisor:'',sigDate:'',designation:'',designationOther:'',shiftWorker:'yes',casualEmployee:'no',addressDuringLeave:''};
  state.consultantFile=null;state.consultantFiles=[];state.consultantData=null;if($('consultantZone')) setConsultantFile(null);
  renderFileList();$('parseBtn').disabled=true;$('clearBtn').style.display='none';
  rosterList.style.display='none';setStatus('');
  $('doctorGrid').innerHTML='<div class="empty">No roster parsed yet</div>';
  $('previewArea').innerHTML='<div class="empty">Select a doctor and click Preview</div>';
  $('employeeName').value='';
  $('detailsSection').style.display='none';
  const sel=$('monthSelect');while(sel.options.length>1) sel.remove(1);
  const mn=['January','February','March','April','May','June','July','August','September','October','November','December'];
  mn.forEach((n,i)=>{const o=document.createElement('option');o.value=i;o.textContent=n;sel.appendChild(o);});
  step2.style.display='none';step2.style.opacity='0.4';step2.style.pointerEvents='none';
  step3.style.opacity='0.4';step3.style.pointerEvents='none';
  $('clearDoctorBtn').style.display='none';
});

// ── Start Over: full form reset (same as clearBtn + also clears step3 details) ──
function fullReset(){
  // Trigger the clearBtn logic
  $('clearBtn').click();
  // Also clear all detail fields
  const fields=['detailFirstName','detailSurname','detailPersal','detailSigDate'];
  fields.forEach(id=>{const el=$(id);if(el)el.value='';});
  const sels=['detailDesignationSel','detailSupervisorSel','detailShiftWorker','detailCasualEmployee'];
  sels.forEach(id=>{const el=$(id);if(el)el.selectedIndex=0;});
  const others=['detailDesignationOther','detailSupervisorOther'];
  others.forEach(id=>{const el=$(id);if(el){el.value='';el.style.display='none';}});
  const addr=$('detailAddress');if(addr)addr.value='';
  $('detailsSection').style.display='none';
  $('leaveFieldsSection').style.display='none';
  ['proceedDownloadBtn','annexureCBtn','z1aBtn'].forEach(id=>{const el=$(id);if(el)el.style.display='none';});
  // Reset year to current
  const yr=$('yearInput');if(yr)yr.value=new Date().getFullYear();
}
$('resetFormBtn')?.addEventListener('click',()=>{
  if(confirm('Clear all data and start over?')) fullReset();
});

$('parseBtn').addEventListener('click',async()=>{
  if(!state.pendingFiles.length && !state.consultantFile) return;
  // Block if both shift AND consultant files queued — they use different preview tables
  if(state.pendingFiles.length && state.consultantFiles?.length){
    setStatus('Please extract the EC shift roster and consultant roster separately.','error');
    return;
  }
  const btn=$('parseBtn');btn.disabled=true;
  btn.innerHTML='<span class="spinner"></span> Extracting\u2026';
  const totalSrc=state.pendingFiles.length+(state.consultantFile?1:0);  setStatus('Extracting from '+totalSrc+' file(s)\u2026','info');
  let errors=0;
  for(const file of [...state.pendingFiles]){
    try{
      const buf=await readFile(file);
      const ext=file.name.split('.').pop().toLowerCase();
      let result;
      if(ext==='pdf') result=await parseRosterPDF(buf);
      else if(ext==='xlsx'||ext==='xls') result=parseRosterExcel(buf);
      else throw new Error('Unsupported format');
      const monthCounts={};
      for(const d of result.days) monthCounts[d.month]=(monthCounts[d.month]||0)+1;
      const dominantMonth=result.days.length>0
        ? parseInt(Object.entries(monthCounts).sort((a,b)=>b[1]-a[1])[0][0]) : -1;
      const filteredDays=result.days.filter(d=>d.month===dominantMonth);
      state.parsedFiles.push({name:file.name,days:filteredDays,doctors:result.doctors});
      state.pendingFiles=state.pendingFiles.filter(f=>f.name!==file.name);
    }catch(err){console.error('Parse error',file.name,err);errors++;state.pendingFiles=state.pendingFiles.filter(f=>f.name!==file.name);}
  }
  if(state.parsedFiles.length) { mergeAndRefresh();renderFileList(); }
  // Parse consultant roster if one is queued
  if(state.consultantFile){
    try{
      await parseAndStoreConsultantRoster();
      const cDocs=state.consultantData?.doctors?.size||0;
      const total2=state.rosterData?.days.length||0,docs2=state.rosterData?.doctors.size||0;
      const nFiles2=state.parsedFiles.length+(state.consultantFile?1:0);
      const shiftOnly2=state.parsedFiles.length>0;
      setStatus(errors
        ?`\u2713 Extracted with ${errors} error(s) \u2014 ${total2} days, ${docs2} staff`
        :shiftOnly2
          ?`\u2713 ${total2} days \u00b7 ${docs2} staff \u00b7 ${cDocs} consultant(s) across ${nFiles2} file(s)`
          :`\u2713 ${total2} days \u00b7 ${cDocs} consultant(s) across 1 file(s)`
        ,errors?'error':'success');
    }catch(cerr){
      console.error('Consultant parse error',cerr);
      const total=state.rosterData?.days.length||0,docs=state.rosterData?.doctors.size||0;
      setStatus(`\u2713 ${total} days \u00b7 ${docs} staff (consultant parse failed)`,'error');
    }
  } else {
    const total=state.rosterData?.days.length||0,docs=state.rosterData?.doctors.size||0;
    setStatus(errors
      ?`\u2713 Extracted with ${errors} error(s) \u2014 ${total} days, ${docs} staff`
      :`\u2713 ${total} days \u00b7 ${docs} staff across ${state.parsedFiles.length} file(s)`
      ,errors?'error':'success');
  }
  btn.disabled=state.pendingFiles.length>0?false:true;
  btn.textContent='Extract Data';
});

function clearDoctorSelection(){
  state.selectedDoctor=null;
  $('employeeName').value='';
  $('doctorGrid').querySelectorAll('.doctor-chip').forEach(c=>c.classList.remove('selected'));
  $('clearDoctorBtn').style.display='none';
  checkReady();
}
function buildDoctorGrid(doctors){
  const sorted=[...doctors].sort();
  if(!sorted.length){$('doctorGrid').innerHTML='<div class="empty">No names detected.</div>';return;}
  $('doctorGrid').innerHTML=sorted.map(d=>`<div class="doctor-chip" data-name="${d}">${d}</div>`).join('');
  $('doctorGrid').querySelectorAll('.doctor-chip').forEach(chip=>{
    chip.addEventListener('click',()=>{
      $('doctorGrid').querySelectorAll('.doctor-chip').forEach(c=>c.classList.remove('selected'));
      chip.classList.add('selected');
      const isNew=state.selectedDoctor!==chip.dataset.name;
      state.selectedDoctor=chip.dataset.name; $('employeeName').value=chip.dataset.name;
      $('clearDoctorBtn').style.display='inline-block';
      if($('detailsSection').style.display!=='none') restoreDetailsToForm(isNew);
      checkReady();
    });
  });
}
function autoDetectMonth(){
  if(!state.rosterData?.days.length) return;
  const counts={};
  for(const d of state.rosterData.days) counts[d.month]=(counts[d.month]||0)+1;
  const dominant=Object.entries(counts).sort((a,b)=>b[1]-a[1])[0][0];
  $('monthSelect').value=dominant;
  let detectedYear=new Date().getFullYear();
  const fyFile=(state.parsedFiles[0]?.name||'').match(/20(\d{2})/);
  if(fyFile) detectedYear='20'+fyFile[1];
  $('yearInput').value=detectedYear;
  checkReady();
}
$('monthSelect').addEventListener('change',checkReady);
$('yearInput').addEventListener('input',checkReady);
$('employeeName').addEventListener('input',checkReady);

// Step 2: Preview button (moved here)
$('previewBtn').addEventListener('click',()=>{
  const {month,year}=getMonthYear();
  if(!state.selectedDoctor||month===null||!year) return;
  if($('detailsSection').style.display!=='none') saveDetailsToState();
  state.previewMonth=month;state.previewYear=year;state.dirtyDays.clear();
  try {
    buildPreview(state.selectedDoctor,month,year);
  } catch(err) {
    console.error('buildPreview error:', err);
    $('previewArea').innerHTML='<div style="color:red;padding:16px;font-family:monospace;font-size:12px;">Preview error: '+err.message+'<br><pre>'+err.stack+'</pre></div>';
  }
  // Fix 2: reveal step 3 now
  unlock(step3);
  $('detailsSection').style.display='';
  restoreDetailsToForm(false);
});

function buildPreview(doctorName,targetMonth,targetYear){
  const holidays=getSAPublicHolidays(targetYear);
  const daysInMonth=new Date(targetYear,targetMonth+1,0).getDate();
  state.editedShifts={};

  // Consultant-type profile: use consultant parser output only
  // Skip getDoctorShifts entirely — consultant days in rosterData use different column semantics
  const isConsultantMode = activeProfile && activeProfile.roster_type === 'consultant' && state.consultantData;

  if (!isConsultantMode) {
    // Standard shift roster path
    const rawShifts=getDoctorShifts(state.rosterData,doctorName,targetMonth);
    for(const [d,shift] of Object.entries(rawShifts)){
      const {nf,nt,of:otF,ot:otT}=splitShift(shift.start,shift.end);
      let typeLabel='WD Shift - 08H00';
      if(shift.isWeekend){
        if(shift.start==='08:00') typeLabel='WE Shift - 08H00';
        else if(shift.start==='13:00') typeLabel='WE Shift - 13H00';
        else typeLabel='WE Shift - 20H00';
      } else {
        if(shift.start==='08:00') typeLabel='WD Shift - 08H00';
        else if(shift.start==='12:00') typeLabel='WD Shift - 12H00';
        else if(shift.start==='15:00') typeLabel='WD Shift - 15H00';
        else typeLabel='WD Shift - 22H00';
      }
      state.editedShifts[parseInt(d)]={nf,nt,of:otF,ot:otT,label:shift.label,typeLabel,isWE:shift.isWeekend};
    }
  }

  state.originalShifts=JSON.parse(JSON.stringify(state.editedShifts));

  // Overlay consultant shifts (fills editedShifts from consultant parser output)
  const consultantAdded = overlayConsultantShifts(doctorName, targetMonth, targetYear);
  if (consultantAdded > 0) {
    for (const [d, s] of Object.entries(state.editedShifts)) {
      if (!state.originalShifts[d]) state.originalShifts[d] = { ...s };
    }
  }
  let sc=0;
  // Build PH letter map: day number -> superscript letter (a,b,c...)
  const phLetterMap={};
  const phFootnotes=[];
  state.phLetterMap={};
  const letters='abcdefghijklmnopqrstuvwxyz';
  const isConsultantMode2 = activeProfile && activeProfile.roster_type === 'consultant' && state.consultantData;
  const cColspan=isConsultantMode2?7:5;
  for(let d=1;d<=daysInMonth;d++){
    const dateObj2=new Date(targetYear,targetMonth,d);
    const ph2=holidays.get(dateKeyLocal(dateObj2));
    if(ph2){
      const letter=letters[phFootnotes.length]||String(phFootnotes.length+1);
      phLetterMap[d]=letter;
      state.phLetterMap[d]=letter;
      phFootnotes.push({letter,name:ph2});
    }
  }
  const activeTypes = isConsultantMode2 ? CONSULTANT_ACTIVITY_TYPES : ACTIVITY_TYPES;
  const typeOpts=activeTypes.map(t=>`<option value="${t}">${t}</option>`).join('');
  let html=`
  <div style="margin-bottom:8px;font-family:var(--sans);font-size:12px;color:var(--text-muted);">
    Edit time fields or change activity type — changes save automatically. Click <strong>+</strong> to add an activity.
  </div>
  <div class="preview-wrapper"><table class="preview-table">
  ${isConsultantMode2
    ? '<thead><tr><th style="width:36px">Date</th><th style="width:70px">Day</th><th style="width:150px">Type</th><th style="width:60px">Norm From</th><th style="width:60px">Norm To</th><th style="width:60px">OT1 From</th><th style="width:60px">OT1 To</th><th style="width:60px">OT2 From</th><th style="width:60px">OT2 To</th><th style="width:46px;text-align:center">Act</th></tr></thead>'
    : '<thead><tr><th style="width:36px">Date</th><th style="width:70px">Day</th><th style="width:140px">Type</th><th style="width:70px">Normal From</th><th style="width:70px">Normal To</th><th style="width:70px">OT From</th><th style="width:70px">OT To</th><th style="width:76px;text-align:center">Actions</th></tr></thead>'
  }<tbody>`;

  for(let d=1;d<=daysInMonth;d++){
    const dateObj=new Date(targetYear,targetMonth,d);
    const dayName=DAY_NAMES[dateObj.getDay()];
    const isWE=dateObj.getDay()===0||dateObj.getDay()===6;
    const phName=holidays.get(dateKeyLocal(dateObj));
    const es=state.editedShifts[d];
    const isSpecial=isWE||!!phName;
    const defaultType = isConsultantMode2
      ? (isSpecial ? 'On Call - Weekend' : 'Normal Hours - Weekday')
      : (isSpecial ? 'WE Shift - 08H00' : 'WD Shift - 08H00');
    const selectedType=es?.typeLabel||defaultType;
    // PH styling: date cell shows "21*" in dark red, day cell also dark red
    const phStyle=phName?'color:#8B1A1A;font-weight:600;':'';
    const phLetter=(state.phLetterMap&&state.phLetterMap[d])||'';
    const dateCell=phName
      ?`<td style="${phStyle}">${d}<sup style="font-size:9px;vertical-align:super">${phLetter}</sup></td>`
      :`<td>${d}</td>`;
    const dayCell=phName
      ?`<td style="${phStyle}">${dayName}</td>`
      :`<td class="${isWE?'we-label':''}">${dayName}</td>`;

    if(es){
      sc++;
      const rowClass=`shift-row${isWE?' we-row':(phName?' ph-wd-row':'')}`;
      if(isConsultantMode2){
        html+=`<tr data-day="${d}" class="${rowClass}">
          ${dateCell}${dayCell}
          <td><select class="type-select" data-day="${d}" data-is-special="${isSpecial?1:0}">${typeOptsFor(isWE,!!phName,selectedType)}</select></td>
          <td><input class="time-edit" data-day="${d}" data-field="nf"   value="${es.nf||''}"   placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
          <td><input class="time-edit" data-day="${d}" data-field="nt"   value="${es.nt||''}"   placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
          <td><input class="time-edit" data-day="${d}" data-field="ot1f" value="${es.ot1f||''}" placeholder="\u2014" maxlength="5" inputmode="numeric" style="color:#2a5a8a;"></td>
          <td><input class="time-edit" data-day="${d}" data-field="ot1t" value="${es.ot1t||''}" placeholder="\u2014" maxlength="5" inputmode="numeric" style="color:#2a5a8a;"></td>
          <td><input class="time-edit" data-day="${d}" data-field="ot2f" value="${es.ot2f||''}" placeholder="\u2014" maxlength="5" inputmode="numeric" style="color:#6b4fa0;"></td>
          <td><input class="time-edit" data-day="${d}" data-field="ot2t" value="${es.ot2t||''}" placeholder="\u2014" maxlength="5" inputmode="numeric" style="color:#6b4fa0;"></td>
          <td class="action-cell"><button class="row-clear" data-day="${d}" title="Remove">&times;</button></td>
        </tr>`;
      } else {
        html+=`<tr data-day="${d}" class="${rowClass}">
          ${dateCell}${dayCell}
          <td><select class="type-select" data-day="${d}" data-is-special="${isSpecial?1:0}">${typeOptsFor(isWE,!!phName,selectedType)}</select></td>
          <td><input class="time-edit" data-day="${d}" data-field="nf" value="${es.nf||''}" maxlength="5" inputmode="numeric"></td>
          <td><input class="time-edit" data-day="${d}" data-field="nt" value="${es.nt||''}" maxlength="5" inputmode="numeric"></td>
          <td><input class="time-edit" data-day="${d}" data-field="of" value="${es.of||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
          <td><input class="time-edit" data-day="${d}" data-field="ot" value="${es.ot||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
          <td class="action-cell"><button class="row-clear" data-day="${d}" title="Remove">&times;</button></td>
        </tr>`;
      }
    } else if(phName){
      html+=`<tr data-day="${d}" class="ph-row${isWE?' we-row':' ph-wd-row'}">
        ${dateCell}
        ${dayCell}
        <td colspan="${cColspan}" style="font-style:italic;color:#7A3B1E">${phName}</td>
        <td class="action-cell"><button class="row-add" title="Add shift" data-day="${d}" data-is-we="1" data-is-special="1">+</button></td>
      </tr>`;
    } else {
      html+=`<tr data-day="${d}" class="empty-row${isWE?' we-row':''}">
        <td>${d}</td>
        <td class="${isWE?'we-label':''}">${dayName}</td>
        <td colspan="${cColspan}"></td>
        <td class="action-cell"><button class="row-add" title="Add shift" data-day="${d}" data-is-we="${isWE?1:0}" data-is-special="${isSpecial?1:0}">+</button></td>
      </tr>`;
    }
  }
  html+=`</tbody></table></div>
  <div style="margin-top:10px;font-family:var(--sans);font-size:12px;color:var(--text-muted);">
    ${sc} activit${sc!==1?'ies':'y'} found &middot; <strong>${doctorName}</strong> &middot; ${MONTH_NAMES[targetMonth]} ${targetYear}
  </div>`;
  if(phFootnotes.length>0){
    html+=`<div style="margin-top:8px;font-family:var(--sans);font-size:12px;color:#8B1A1A;line-height:1.8;">`+
      phFootnotes.map(f=>`<span style="margin-right:16px;"><sup style="font-size:9px;">${f.letter}</sup> ${f.name}</span>`).join('')+
    `</div>`;
  }
  $('previewArea').innerHTML=html;
  attachEditHandlers();
}

function makeRowInner(d,isWE,phName,dayName,es){
  const isSpecial=isWE||!!phName;
  const isConsMode=activeProfile&&activeProfile.roster_type==='consultant'&&state.consultantData;
  const selectedType=es?.typeLabel||(isConsMode?(isSpecial?'On Call - Weekend':'Normal Hours - Weekday'):(isSpecial?'WE Shift - 08H00':'WD Shift - 08H00'));
  const phStyle=phName?'color:#8B1A1A;font-weight:600;':'';
  const phLetter=(state.phLetterMap&&state.phLetterMap[d])||'';
  const dateCell=phName
    ?`<td style="${phStyle}">${d}<sup style="font-size:9px;vertical-align:super">${phLetter}</sup></td>`
    :`<td>${d}</td>`;
  const dayCell=phName?`<td style="${phStyle}">${dayName}</td>`:`<td class="${isWE?'we-label':''}">${dayName}</td>`;
  if(isConsMode){
    return `
    ${dateCell}
    ${dayCell}
    <td><select class="type-select" data-day="${d}" data-is-special="${isSpecial?1:0}">${typeOptsFor(isWE,!!phName,selectedType)}</select></td>
    <td><input class="time-edit" data-day="${d}" data-field="nf"   value="${es?.nf||''}"   placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="nt"   value="${es?.nt||''}"   placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot1f" value="${es?.ot1f||''}" placeholder="\u2014" maxlength="5" inputmode="numeric" style="color:#2a5a8a;"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot1t" value="${es?.ot1t||''}" placeholder="\u2014" maxlength="5" inputmode="numeric" style="color:#2a5a8a;"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot2f" value="${es?.ot2f||''}" placeholder="\u2014" maxlength="5" inputmode="numeric" style="color:#6b4fa0;"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot2t" value="${es?.ot2t||''}" placeholder="\u2014" maxlength="5" inputmode="numeric" style="color:#6b4fa0;"></td>
    <td class="action-cell"><button class="row-clear" data-day="${d}" title="Remove">&times;</button>${state.originalShifts[d]?`<button class="row-undo" data-day="${d}" title="Undo">&#8635;</button>`:''}</td>`;
  }
  return `
    ${dateCell}
    ${dayCell}
    <td><select class="type-select" data-day="${d}" data-is-special="${isSpecial?1:0}">${typeOptsFor(isWE,!!phName,selectedType)}</select></td>
    <td><input class="time-edit" data-day="${d}" data-field="nf" value="${es?.nf||''}" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="nt" value="${es?.nt||''}" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="of" value="${es?.of||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot" value="${es?.ot||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td class="action-cell"><button class="row-clear" data-day="${d}" title="Remove">&times;</button>${state.originalShifts[d]?`<button class="row-undo" data-day="${d}" title="Undo">&#8635;</button>`:''}</td>`;
}

function attachEditHandlers(){
  document.querySelectorAll('.type-select').forEach(sel=>{
    if(sel.dataset.bound) return; // Fix 4: skip if already has listener
    sel.dataset.bound='1';
    sel.addEventListener('change',()=>{
      const d=parseInt(sel.dataset.day);
      if(sel.value===''){
        delete state.editedShifts[d]; state.dirtyDays.delete(d);
        const row=document.querySelector(`tr[data-day="${d}"]`);
        if(row){
          const dateObj=new Date(state.previewYear,state.previewMonth,d);
          const isWE=dateObj.getDay()===0||dateObj.getDay()===6;
          const phEntry=state.phMap&&state.phMap.get(d);
          row.innerHTML=makeRowInner(d,isWE,phEntry?.name||null,DAY_NAMES[dateObj.getDay()],null,phEntry?.num||null);
          attachEditHandlers();
        }
        updateLeaveFields(); checkDetailsComplete(); return;
      }
      if(!state.editedShifts[d]) state.editedShifts[d]={nf:'',nt:'',of:null,ot:null,label:'Custom',typeLabel:'WD Shift - 08H00',isWE:false};
      state.editedShifts[d].typeLabel=sel.value;
      const row=document.querySelector(`tr[data-day="${d}"]`);
      const cTimes=CONSULTANT_SHIFT_TIMES[sel.value];
      const sTimes=SHIFT_TIMES[sel.value];
      if(cTimes){
        // Consultant type auto-fill (6 fields)
        const isC=activeProfile&&activeProfile.roster_type==='consultant';
        Object.assign(state.editedShifts[d],{nf:cTimes.nf,nt:cTimes.nt,ot1f:cTimes.ot1f,ot1t:cTimes.ot1t,ot2f:cTimes.ot2f,ot2t:cTimes.ot2t});
        if(row){
          const fields=['nf','nt','ot1f','ot1t','ot2f','ot2t'];
          row.querySelectorAll('.time-edit').forEach((inp,i)=>{if(fields[i])inp.value=cTimes[fields[i]]||'';inp.style.borderColor='';});
        }
      } else if(sTimes){
        // Standard shift roster auto-fill
        state.editedShifts[d].nf=sTimes.nf; state.editedShifts[d].nt=sTimes.nt;
        state.editedShifts[d].of=sTimes.of; state.editedShifts[d].ot=sTimes.ot;
        if(row){
          const fields=['nf','nt','of','ot'];
          row.querySelectorAll('.time-edit').forEach((inp,i)=>{inp.value=sTimes[fields[i]]||'';inp.style.borderColor='';});
        }
      } else {
        // Non-shift activity (leave/workshop etc): clear all time fields
        ['nf','nt','of','ot','ot1f','ot1t','ot2f','ot2t'].forEach(k=>state.editedShifts[d][k]='');
        if(row) row.querySelectorAll('.time-edit').forEach(inp=>{inp.value='';inp.style.borderColor='';});
      }
      markDirty(d);
      updateLeaveFields();
    });
  });

  document.querySelectorAll('.time-edit').forEach(inp=>{
    const fresh=inp.cloneNode(true);
    inp.parentNode.replaceChild(fresh,inp);
    fresh.addEventListener('keydown',e=>{
      const allowed=['Backspace','Delete','Tab','ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Enter'];
      if(allowed.includes(e.key)||/^\d$/.test(e.key)||e.key==='H'||e.key==='h'||e.key===':') return;
      e.preventDefault();
    });
    fresh.addEventListener('input',()=>{
      let v=fresh.value.replace(/[^0-9H:]/gi,'').toUpperCase();
      if(/^\d{3,4}$/.test(v)) v=v.slice(0,2)+'H'+v.slice(2);
      if(v!==fresh.value) fresh.value=v;
      fresh.style.borderColor=v.length>0&&!normaliseTime(v)?'var(--warn)':'';
    });
    fresh.addEventListener('blur',()=>{
      const d=parseInt(fresh.dataset.day),field=fresh.dataset.field;
      const val=fresh.value.trim().toUpperCase(),normalised=normaliseTime(val);
      if(normalised){
        fresh.value=normalised;fresh.style.borderColor='';fresh.title='';
        if(!state.editedShifts[d]) state.editedShifts[d]={nf:'',nt:'',of:null,ot:null,label:'Custom',typeLabel:'WD Shift - 08H00',isWE:false};
        if(state.editedShifts[d][field]!==normalised){state.editedShifts[d][field]=normalised;markDirty(d);}
      } else if(val===''){
        fresh.style.borderColor='';
        if(state.editedShifts[d]&&state.editedShifts[d][field]!==null){state.editedShifts[d][field]=null;markDirty(d);}
      } else {fresh.style.borderColor='var(--warn)';fresh.title='Format: HHH00 (e.g. 08H00)';}
    });
  });

  document.querySelectorAll('.row-clear').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const d=parseInt(btn.dataset.day);
      delete state.editedShifts[d];state.dirtyDays.delete(d);
      const row=document.querySelector(`tr[data-day="${d}"]`);
      if(!row) return;
      const dateObj=new Date(state.previewYear,state.previewMonth,d);
      const dayName=DAY_NAMES[dateObj.getDay()];
      const isWE=dateObj.getDay()===0||dateObj.getDay()===6;
      const phName=getSAPublicHolidays(state.previewYear).get(dateKeyLocal(dateObj));
      if(phName){
        row.className='ph-row'+(isWE?' we-row':' ph-wd-row');row.style.opacity='';
        const _phStyle='color:#8B1A1A;font-weight:600;';
        const _phLetter=(state.phLetterMap&&state.phLetterMap[d])||'';
        const _isConsCP=activeProfile&&activeProfile.roster_type==='consultant'&&state.consultantData;
        const _colspanPH=_isConsCP?7:5;
        row.innerHTML=`<td style="${_phStyle}">${d}<sup>${_phLetter}</sup></td><td style="${_phStyle}">${dayName}</td>
          <td colspan="${_colspanPH}" style="font-style:italic;color:#7A3B1E">${phName}</td>
          <td class="action-cell"><button class="row-add" title="Add shift" data-day="${d}" data-is-we="1" data-is-special="1">+</button></td>`;
      } else {
        row.className='empty-row'+(isWE?' we-row':'');row.style.opacity='';
        const _isConsC=activeProfile&&activeProfile.roster_type==='consultant'&&state.consultantData;
        const _colspan=_isConsC?7:5;
        row.innerHTML=`<td>${d}</td><td class="${isWE?'we-label':''}">${dayName}</td>
          <td colspan="${_colspan}"></td>
          <td class="action-cell"><button class="row-add" title="Add shift" data-day="${d}" data-is-we="${isWE?1:0}" data-is-special="${isWE||!!phName?1:0}">+</button></td>`;
      }
      attachEditHandlers();
    });
  });

  document.querySelectorAll('.row-undo').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const d=parseInt(btn.dataset.day);
      const orig=state.originalShifts[d];
      if(!orig) return;
      state.editedShifts[d]={...orig};
      state.dirtyDays.delete(d);
      const row=document.querySelector(`tr[data-day="${d}"]`);
      if(!row) return;
      const dateObj=new Date(state.previewYear,state.previewMonth,d);
      const dayName=DAY_NAMES[dateObj.getDay()];
      const isWE=dateObj.getDay()===0||dateObj.getDay()===6;
      const phName=getSAPublicHolidays(state.previewYear).get(dateKeyLocal(dateObj));
      row.className='shift-row'+(isWE?' we-row':(phName?' ph-wd-row':''));
      row.innerHTML=makeRowInner(d,isWE,phName,dayName,state.editedShifts[d]);
      state.dirtyDays.delete(d);
      attachEditHandlers();
      updateLeaveFields();
      checkDetailsComplete();
    });
  });

  document.querySelectorAll('.row-add').forEach(btn=>{
    btn.addEventListener('click',()=>{
      try {
        const d=parseInt(btn.dataset.day),isWE=btn.dataset.isWe==='1';
        const isSpecialNew=btn.dataset.isSpecial==='1';
        const isConsMode=activeProfile&&activeProfile.roster_type==='consultant'&&state.consultantData;
        const defaultLabel=isConsMode
          ?(isSpecialNew?'On Call - Weekend':'Normal Hours - Weekday')
          :(isSpecialNew?'WE Shift - 08H00':'WD Shift - 08H00');
        const defaultTimes=isConsMode?(CONSULTANT_SHIFT_TIMES[defaultLabel]||{}):(SHIFT_TIMES[defaultLabel]||{});
        const def=isConsMode
          ?{nf:defaultTimes.nf||'',nt:defaultTimes.nt||'',ot1f:defaultTimes.ot1f||'',ot1t:defaultTimes.ot1t||'',ot2f:defaultTimes.ot2f||'',ot2t:defaultTimes.ot2t||'',label:'Custom',typeLabel:defaultLabel,isWE:isWE}
          :{...defaultTimes,label:'Custom',typeLabel:defaultLabel,isWE:isWE};
        state.editedShifts[d]=def;
        const row=document.querySelector(`tr[data-day="${d}"]`);
        const dateObj=new Date(state.previewYear,state.previewMonth,d);
        const dayName=DAY_NAMES[dateObj.getDay()];
        const phName=getSAPublicHolidays(state.previewYear).get(dateKeyLocal(dateObj));
        if(row){
          row.className='shift-row'+(isWE?' we-row':(phName?' ph-wd-row':''));
          row.style.opacity='';
          const inner=makeRowInner(d,isWE,phName,dayName,def);
          row.innerHTML=inner;
          attachEditHandlers();
          updateLeaveFields();
        } else { console.error('row-add: row not found for day',d); }
      } catch(err){ console.error('row-add error:',err.message,err.stack); }
    });
  });
}

// Designation dropdown
$('detailDesignationSel').addEventListener('change',()=>{
  const isOther=$('detailDesignationSel').value==='other';
  $('detailDesignationOther').style.display=isOther?'':'none';
  if(!isOther) $('detailDesignationOther').value='';
  saveDetailsToState(); checkDetailsComplete();
});
$('detailDesignationOther').addEventListener('input',()=>{ saveDetailsToState(); checkDetailsComplete(); });
// Date picker
// Fix 3: supervisor dropdown + other listeners
$('detailSupervisorSel').addEventListener('change',()=>{
  const isOther=$('detailSupervisorSel').value==='other';
  $('detailSupervisorOther').style.display=isOther?'':'none';
  if(!isOther) $('detailSupervisorOther').value='';
  saveDetailsToState();
});
$('detailSupervisorOther').addEventListener('input',()=>{ saveDetailsToState(); });
$('detailSigDatePicker').addEventListener('change',e=>{
  const d=e.target.value;
  if(d){const [y,m,day]=d.split('-');$('detailSigDate').value=`${day}/${m}/${y}`;}
  else $('detailSigDate').value='';
  saveDetailsToState();checkDetailsComplete();
});
$('detailSigDate').addEventListener('input',e=>{
  let v=e.target.value.replace(/\D/g,'');
  if(v.length>2) v=v.slice(0,2)+'/'+v.slice(2);
  if(v.length>5) v=v.slice(0,5)+'/'+v.slice(5);
  if(v.length>10) v=v.slice(0,10);
  e.target.value=v;
  const m=v.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if(m) $('detailSigDatePicker').value=`${m[3]}-${m[2]}-${m[1]}`;
  saveDetailsToState();checkDetailsComplete();
});

['detailFirstName','detailSurname','detailPersal'].forEach(id=>{
  const el=$(id); if(!el) return;
  el.addEventListener('input',()=>{saveDetailsToState();checkDetailsComplete();});
});

function getFormDetails(){
  saveDetailsToState();
  const {month,year}=getMonthYear();
  return {
    firstName:state.savedDetails.firstName, surname:state.savedDetails.surname,
    persal:state.savedDetails.persal, supervisorName:state.savedDetails.supervisor,
    designation:state.savedDetails.designation,
    signatureDate:state.savedDetails.sigDate,
    addressDuringLeave:state.savedDetails.address||'',
    shiftWorker:state.savedDetails.shiftWorker||'yes',
    casualEmployee:state.savedDetails.casualEmployee||'no',
    editedShifts:state.editedShifts,
    month, year,
  };
}

$('proceedDownloadBtn').addEventListener('click',async()=>{
  const {month,year}=getMonthYear();
  if(!state.selectedDoctor||month===null||!year) return;
  const btn=$('proceedDownloadBtn');btn.disabled=true;
  btn.innerHTML='<span class="spinner"></span> Generating\u2026';
  try{
    saveDetailsToState();
    const details=getFormDetails();
    const result=await generateExcel(month,year,details);
    const blob=new Blob([result],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    const safe=(details.firstName+' '+details.surname).trim().replace(/\s+/g,'_');
    a.href=url;a.download=`Duty_Roster_${safe}_${MONTH_NAMES[month]}_${year}.xlsx`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  }catch(err){alert('Error: '+err.message);console.error(err);}
  btn.disabled=false;btn.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:-3px;margin-right:5px"><path d="M12 17V3"/><path d="m6 11 6 6 6-6"/><path d="M19 21H5"/></svg>Download Duty Roster';
});

$('annexureCBtn').addEventListener('click',async()=>{
  const d=getFormDetails();
  const btn=$('annexureCBtn');
  btn.disabled=true; btn.textContent='Generating…';
  try{
    const blob=await generateAnnexureCDocx(d);
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    const safe=(d.firstName+' '+d.surname).trim().replace(/\s+/g,'_');
    a.href=url; a.download=`Annexure_C_${safe}_${MONTH_NAMES[d.month]}_${d.year}.docx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  }catch(err){alert('Error generating Annexure C: '+err.message);console.error(err);}
  btn.disabled=false; btn.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:-3px;margin-right:5px"><path d="M12 17V3"/><path d="m6 11 6 6 6-6"/><path d="M19 21H5"/></svg>Download Annexure C (Overtime) Form';
});

$('z1aBtn').addEventListener('click',async()=>{
  const d=getFormDetails();
  const btn=$('z1aBtn');
  btn.disabled=true; btn.textContent='Generating…';
  try{
    const blob=await generateZ1ADocx(d);
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    const safe=(d.firstName+' '+d.surname).trim().replace(/\s+/g,'_');
    a.href=url; a.download=`Z1a_Leave_${safe}_${MONTH_NAMES[d.month]}_${d.year}.docx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  }catch(err){alert('Error generating Z1(a): '+err.message);console.error(err);}
  btn.disabled=false; btn.innerHTML='<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:-3px;margin-right:5px"><path d="M12 17V3"/><path d="m6 11 6 6 6-6"/><path d="M19 21H5"/></svg>Download Z1(a) Leave Form';
});

function readFile(file){
  return new Promise((res,rej)=>{
    const r=new FileReader();
    r.onload=e=>res(e.target.result);
    r.onerror=()=>rej(new Error('Failed to read '+file.name));
    r.readAsArrayBuffer(file);
  });
}
$('yearInput').value=new Date().getFullYear();

// ═══════════════════════════════════════════════════════════════
