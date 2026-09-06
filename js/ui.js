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
  return Object.values(state.editedShifts).some(es=>isLeaveActivity(es.typeLabel));
}
// The Z1(a) is a leave application, so only actual leave unlocks it.
function hasZ1LeaveInShifts(){
  if(!state.editedShifts) return false;
  return Object.values(state.editedShifts).some(es=>isZ1LeaveActivity(es.typeLabel));
}
// Names for the supervisor dropdown, or null to use a plain text box.
function supervisorOptionsFor(){
  if(activeProfile&&Array.isArray(activeProfile.supervisors)&&activeProfile.supervisors.length)
    return activeProfile.supervisors;
  if(activeProfile&&activeProfile.roster_type==='shift') return LEGACY_EC_SUPERVISORS;
  return null;
}
// The free-text box sits under the dropdown when "Other" is picked, and needs
// a gap there. When the profile has no list it is the whole control, and that
// same gap drops it out of line with Designation and Date of signature.
function showSupervisorBox(show, underSelect){
  const o=$('detailSupervisorOther');
  if(!o) return;
  o.style.display=show?'':'none';
  o.style.marginTop=(show&&underSelect)?'6px':'0';
}
function applySupervisorMode(){
  const sel=$('detailSupervisorSel'), other=$('detailSupervisorOther');
  if(!sel||!other) return;
  const list=supervisorOptionsFor();
  if(!list){
    sel.style.display='none';
    showSupervisorBox(true, false);
    return;
  }
  sel.style.display='';
  while(sel.options.length>1) sel.remove(1);
  for(const n of list){
    const o=document.createElement('option'); o.value=n; o.textContent=n; sel.appendChild(o);
  }
  const oth=document.createElement('option');
  oth.value='other'; oth.textContent='Other\u2026'; sel.appendChild(oth);
  if(sel.value!=='other') showSupervisorBox(false, true);
}
// Put a saved name back into whichever control this profile uses. Everything
// that restores the form goes through here: a profile without a supervisor
// list has no dropdown to hide the text box behind, and the old restore code
// hid both, leaving the field with nothing on screen.
function setSupervisorValue(saved){
  const sel=$('detailSupervisorSel'), other=$('detailSupervisorOther');
  if(!sel||!other) return;
  applySupervisorMode();
  const list=supervisorOptionsFor();
  saved=saved||'';
  if(!list){ sel.value=''; other.value=saved; showSupervisorBox(true, false); return; }
  if(saved&&list.includes(saved)){ sel.value=saved; other.value=''; showSupervisorBox(false, true); }
  else if(saved){ sel.value='other'; other.value=saved; showSupervisorBox(true, true); }
  else { sel.value=''; other.value=''; showSupervisorBox(false, true); }
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
  // Read the supervisor the same way saveDetailsToState does: the free-text
  // box when this profile has no list, or when "Other" is picked.
  const supSel=$('detailSupervisorSel').value;
  const supervisor=(!supervisorOptionsFor()||supSel==='other')
    ? ($('detailSupervisorOther')?.value.trim()||'') : supSel;
  const date=$('detailSigDate').value.trim();
  const dateValid=date.length===10&&/^\d{2}\/\d{2}\/\d{4}$/.test(date);
  const leaveVisible=$('leaveFieldsSection')?.style.display!=='none';
  const address=$('detailAddress')?.value.trim()||'';
  const leaveOk=!leaveVisible||(address.length>0);
  const show=!!(first&&surname&&persal&&designation&&supervisor&&dateValid&&leaveOk);
  // The duty roster and Annexure C never depend on leave; the Z1(a) is
  // hidden outright unless some leave was captured.
  const z1=hasZ1LeaveInShifts();
  $('proceedDownloadBtn').disabled=!show;
  $('annexureCBtn').disabled=!show;
  $('z1aBtn').style.display=z1?'':'none';
  $('z1aBtn').disabled=!(show&&z1);
  const lockNote=$('downloadsLocked');
  if(lockNote) lockNote.style.display=show?'none':'';
  if(typeof wizRenderNav==='function') wizRenderNav();
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

// EC staff work shifts; other departments work ordinary hours and do calls.
// The profile decides what the preview section calls them.
function dutyNoun(){ return (activeProfile&&activeProfile.duty_noun)||'shifts'; }
function applyDutyNoun(){
  // The section heading is now duty-neutral ("Review schedule"), so the
  // profile's word is only needed where the copy actually describes them.
  const empty=$('step2Empty');
  if(empty) empty.textContent='Extract a roster to see detected staff and their '+dutyNoun()+'.';
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

// Overtime bands run back to back: OT1 starts where normal hours end, and OT2
// starts where OT1 ends. Editing the end of one band therefore moves the start
// of the next to match — but only where that next band is actually in use, so
// an empty overtime column is never filled in by accident.
function syncFollowingBand(d,field,value){
  const es=state.editedShifts[d];
  if(!es||!value) return [];
  const follows=isExtendedRosterMode()
    ? {nt:{target:'ot1f',pair:['ot1f','ot1t']}, ot1t:{target:'ot2f',pair:['ot2f','ot2t']}}
    : {nt:{target:'of',pair:['of','ot']}};
  const rule=follows[field];
  if(!rule) return [];
  const inUse=rule.pair.some(f=>!!es[f]);
  if(!inUse||es[rule.target]===value) return [];
  es[rule.target]=value;
  return [rule.target];
}
// Show a programmatic band change in the row the user is looking at.
function applyBandSync(d,fields,value){
  if(!fields.length) return;
  const row=document.querySelector('[data-day="'+d+'"]');
  if(!row) return;
  row.querySelectorAll('.time-edit').forEach(inp=>{
    if(fields.includes(inp.dataset.field)){ inp.value=value; inp.style.borderColor=''; inp.title=''; }
  });
}

// Undo is offered only where the day now differs from what was parsed —
// originalShifts holds a snapshot of every day, so its presence proves nothing.
function isDayEdited(day){
  const o=state.originalShifts&&state.originalShifts[day];
  const e=state.editedShifts&&state.editedShifts[day];
  if(!o||!e) return false;
  for(const k of new Set([...Object.keys(o),...Object.keys(e)])){
    if((o[k]||'')!==(e[k]||'')) return true;
  }
  return false;
}

function markDirty(day) {
  if(typeof wizInvalidateReview==='function') wizInvalidateReview();
  // Data already captured in state.editedShifts — auto-clean immediately
  markClean(day);
}
function markClean(day) {
  state.dirtyDays.delete(day);
  const row=document.querySelector(`tr[data-day="${day}"]`);
  if(!row) return;
  const ac=row.querySelector('.action-cell');
  if(ac){ac.innerHTML=`<button class="row-clear" data-day="${day}" title="Remove">&times;</button>${isDayEdited(day)?`<button class="row-undo" data-day="${day}" title="Undo changes">&#8635;</button>`:''}`; attachEditHandlers();}
}

function saveDetailsToState() {
  state.savedDetails.firstName=$('detailFirstName').value.trim();
  state.savedDetails.surname=$('detailSurname').value.trim();
  state.savedDetails.persal=$('detailPersal').value.trim();
  // Fix 3: supervisor — use dropdown value, or 'other' text input
  const supSel=$('detailSupervisorSel').value;
  // With no dropdown for this profile the text box is the only source.
  state.savedDetails.supervisor=(!supervisorOptionsFor()||supSel==='other')
    ? $('detailSupervisorOther').value.trim() : supSel;
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
    $('detailPersal').value=''; setSupervisorValue(''); $('detailSigDate').value=''; $('detailDesignationSel').value=''; $('detailDesignationOther').value=''; $('detailDesignationOther').style.display='none';
    state.savedDetails={firstName:'',surname:state.selectedDoctor||'',persal:'',supervisor:'',sigDate:''};
  } else {
    $('detailFirstName').value=state.savedDetails.firstName||'';
    $('detailSurname').value=state.savedDetails.surname||state.selectedDoctor||'';
    $('detailPersal').value=state.savedDetails.persal||'';
    // The known names are the ones this profile offers, not a fixed EC list.
    setSupervisorValue(state.savedDetails.supervisor||'');
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
      <span class="tag tag-accent">${(f.name.split('.').pop()||'').toUpperCase()}</span>
      <span class="ri-name">${f.name}</span>
      ${f.parsed?`<span class="ri-days">${f.days} days</span>`:`<span class="ri-days">queued</span>`}
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
  const staffCountEl=$('staffCount'); if(staffCountEl) staffCountEl.textContent='0';
  $('previewArea').innerHTML='<div class="empty">Select a name and click Preview schedule</div>';
  if(typeof updatePreviewTotals==='function') updatePreviewTotals();
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
  applySupervisorMode();
  const addr=$('detailAddress');if(addr)addr.value='';
  $('detailsSection').style.display='none';
  $('leaveFieldsSection').style.display='none';
  ['proceedDownloadBtn','annexureCBtn','z1aBtn'].forEach(id=>{const el=$(id);if(el)el.disabled=true;});
  $('z1aBtn').style.display='none';
  const lockNote=$('downloadsLocked'); if(lockNote) lockNote.style.display='';
  // Reset year to current
  const yr=$('yearInput');if(yr)yr.value=new Date().getFullYear();
  // Start over means start over. The saved department profile is the only
  // thing the app keeps between visits, so it goes too and the next visit
  // begins at the picker. The catalogue cache stays — it is the public list
  // of departments, holds nothing about the user, and is what lets the
  // picker still work offline.
  wizReviewed=false;
  const _ackR=$('reviewAck'); if(_ackR) _ackR.checked=false;
  try { localStorage.removeItem(LS_PROFILE_KEY); } catch(e) {}
  activeProfile=null;
  const modeEl=$('ecMode'); if(modeEl){modeEl.textContent='';modeEl.style.display='none';}
  const offNote=$('ecOfflineNote'); if(offNote) offNote.style.display='none';
  if(typeof window.reopenEcPicker==='function') window.reopenEcPicker();
  // Leaves the viewer icons and the Preview-schedule button in step with the
  // now-empty state. The header context itself is hidden by showEcPicker(),
  // which reopenEcPicker() reaches asynchronously — after this line.
  if(typeof wizRefresh==='function') wizRefresh();
}
// The confirmation panel has already been answered by the time this runs.
function confirmFullReset(){ fullReset(); }
// Two entry points, one action: the masthead button before a department is
// chosen, and the Edit panel once the wizard has taken the masthead's place.
$('resetFormBtn')?.addEventListener('click',confirmFullReset);
document.addEventListener('click',e=>{
  const t=e.target.closest&&e.target.closest('#hdrResetBtn,#wizStartOver');
  if(t){ e.preventDefault(); confirmFullReset(); }
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
      // A grid profile reads its roster as a table of rows and columns
      // whatever file carries it, so a workbook goes to the table parser
      // rather than to the EC shift reader.
      const isGrid=!!activeProfile&&activeProfile.roster_type==='table';
      if(ext==='docx'||ext==='doc'||(ext==='xlsx'&&isGrid)){
        if(!isGrid)
          throw new Error('This department profile is not set up for table rosters');
        result=await parseWordRosterTable(buf,activeProfile,file.name);
        state.tableData={days:result.days,doctors:result.doctors};
        state.tableWarnings=(state.tableWarnings||[]).concat(result.warnings||[]);
      }
      else if(ext==='pdf') result=await parseRosterPDF(buf);
      else if(ext==='xls') throw new Error('Legacy .xls is not supported \u2014 open it in Excel and Save As .xlsx');
      else if(ext==='xlsx') result=await parseRosterExcel(buf);
      else throw new Error('Unsupported format');
      const monthCounts={};
      for(const d of result.days) monthCounts[d.month]=(monthCounts[d.month]||0)+1;
      const dominantMonth=result.days.length>0
        ? parseInt(Object.entries(monthCounts).sort((a,b)=>b[1]-a[1])[0][0]) : -1;
      const filteredDays=result.days.filter(d=>d.month===dominantMonth);
      state.parsedFiles.push({name:file.name,days:filteredDays,doctors:result.doctors,file});
      state.pendingFiles=state.pendingFiles.filter(f=>f.name!==file.name);
    }catch(err){console.error('Parse error',file.name,err);errors++;state.pendingFiles=state.pendingFiles.filter(f=>f.name!==file.name);}
  }
  if(state.parsedFiles.length) { mergeAndRefresh();renderFileList(); }
  // Extraction is done — the wizard can now judge whether step 1 is complete.
  if(typeof wizRefresh==='function') wizRefresh();
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
  syncStaffCollapse();
  checkReady();
}
function countDoctorDays(name){
  const nl=String(name||'').toLowerCase();
  let n=0;
  for(const day of (state.rosterData?.days||[])){
    const all=[...(day.allNames||[]),...(day.shifts?.flat()||[])];
    if(all.some(x=>String(x).toLowerCase()===nl)) n++;
  }
  return n;
}
// The list collapses to the selected name and expands again on request. It is
// presentation only: every chip stays in the DOM and stays clickable once shown.
function syncStaffCollapse(){
  const grid=$('doctorGrid'), btn=$('staffExpandBtn');
  if(!grid||!btn) return;
  const hasSelection=!!grid.querySelector('.doctor-chip.selected');
  const others=grid.querySelectorAll('.doctor-chip:not(.selected)').length;
  if(!hasSelection||!others){
    grid.classList.remove('is-collapsed');
    btn.style.display='none';
    btn.setAttribute('aria-expanded','true');
    return;
  }
  btn.style.display='inline-block';
  const expanded=btn.getAttribute('aria-expanded')==='true';
  grid.classList.toggle('is-collapsed',!expanded);
  btn.textContent=expanded?'Show fewer':'Show all';
}
document.addEventListener('click',e=>{
  const b=e.target.closest&&e.target.closest('#staffExpandBtn');
  if(!b) return;
  b.setAttribute('aria-expanded',b.getAttribute('aria-expanded')==='true'?'false':'true');
  syncStaffCollapse();
});

function buildDoctorGrid(doctors){
  const sorted=[...doctors].sort();
  const countEl=$('staffCount'); if(countEl) countEl.textContent=sorted.length;
  if(!sorted.length){$('doctorGrid').innerHTML='<div class="empty">No names detected.</div>';return;}
  $('doctorGrid').innerHTML=sorted.map(d=>{
    const c=countDoctorDays(d);
    return `<button type="button" class="doctor-chip" data-name="${d}"><span class="dc-name">${d}</span><span class="dc-count">${c} ${c===1?'shift':'shifts'}</span></button>`;
  }).join('');
  $('doctorGrid').querySelectorAll('.doctor-chip').forEach(chip=>{
    chip.addEventListener('click',()=>{
      $('doctorGrid').querySelectorAll('.doctor-chip').forEach(c=>c.classList.remove('selected'));
      chip.classList.add('selected');
      const isNew=state.selectedDoctor!==chip.dataset.name;
      state.selectedDoctor=chip.dataset.name; $('employeeName').value=chip.dataset.name;
      $('clearDoctorBtn').style.display='inline-block';
      const xb=$('staffExpandBtn'); if(xb) xb.setAttribute('aria-expanded','false');
      syncStaffCollapse();
      if($('detailsSection').style.display!=='none') restoreDetailsToForm(isNew);
      checkReady();
    });
  });
  syncStaffCollapse();
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
    $('previewArea').innerHTML='<div style="color:red;padding:16px;font-family:var(--font-body);font-size:12px;">Preview error: '+err.message+'<br><pre>'+err.stack+'</pre></div>';
  }
  // Fix 2: reveal step 3 now
  unlock(step3);
  $('detailsSection').style.display='';
  restoreDetailsToForm(false);
  // A new preview is a different month or person, so a prior acknowledgement
  // no longer describes what is on screen.
  wizReviewed=false;
  const _ack=$('reviewAck'); if(_ack) _ack.checked=false;
  if(typeof wizRefresh==='function') wizRefresh();
});

// Consultant and Word-table rosters both produce normal + OT1 + OT2 bands,
// so they share the wider preview layout and their own activity-type list.
function isTableRosterMode(){
  return !!(activeProfile && activeProfile.roster_type==='table' && state.tableData);
}
function isExtendedRosterMode(){
  return !!(activeProfile && ((activeProfile.roster_type==='consultant' && state.consultantData)
                           || (activeProfile.roster_type==='table' && state.tableData)));
}
// Activity types for a table roster come from the profile's own role labels.
function tableActivityTypes(){
  const rules=(activeProfile&&activeProfile.role_rules)||{};
  const out=[];
  // The implied ordinary working day comes first — it is the most common row.
  const dflt=activeProfile&&activeProfile.default_weekday;
  if(dflt) out.push(dflt.label||'Normal Hours - Weekday');
  for(const [role,r] of Object.entries(rules)){
    for(const k of ['label_weekday','label_weekend','label_ph']){
      const v=r[k]||(role+' - '+k.replace('label_',''));
      if(!out.includes(v)) out.push(v);
    }
  }
  // Leave and other non-roster activities come from the profile when it says
  // so — the EC WD/WE shift types never apply to a table roster.
  const leave=(activeProfile&&activeProfile.leave_types)
    || ACTIVITY_TYPES.filter(t=>/^Leave|^Workshop|^Course|^Conference/.test(t));
  for(const t of leave) if(!out.includes(t)) out.push(t);
  return out;
}

// Default activity type for a newly added row, in the active profile's own
// vocabulary — a table roster's types come from its role labels, not the
// consultant list.
function defaultTypeLabel(isSpecial){
  if(isTableRosterMode()){
    const t=tableActivityTypes();
    return t.find(x=>isSpecial?/- (Weekend|Public Holiday)$/.test(x):/- Weekday$/.test(x))||t[0]||'';
  }
  if(isExtendedRosterMode()) return isSpecial?'On Call - Weekend':'Normal Hours - Weekday';
  return isSpecial?'WE Shift - 08H00':'WD Shift - 08H00';
}

function hoursBetween(from,to){
  const p=t=>{const m=/^(\d{1,2})[H:](\d{2})$/i.exec(String(t||'').trim());return m?parseInt(m[1],10)*60+parseInt(m[2],10):null;};
  const s=p(from),e=p(to);
  if(s===null||e===null) return 0;
  let d=e-s; if(d<0) d+=1440;
  return d/60;
}
function updatePreviewTotals(){
  let normal=0,ot=0;
  for(const es of Object.values(state.editedShifts||{})){
    normal+=hoursBetween(es.nf,es.nt);
    // of/ot is a mirror of the OT2 band (or of OT1 where there is no OT2) on
    // consultant and table rosters, so adding all three would count a night
    // of call twice. Shift rosters only ever populate of/ot.
    const banded=hoursBetween(es.ot1f,es.ot1t)+hoursBetween(es.ot2f,es.ot2t);
    ot+=banded>0?banded:hoursBetween(es.of,es.ot);
  }
  const fmt=v=>(Math.round(v*10)/10).toString().replace(/\.0$/,'')+' h';
  // The figures are hidden on a phone, so the bar may not be on screen at all.
  const nEl=$('totalNormal'),oEl=$('totalOt');
  if(nEl) nEl.textContent=fmt(normal);
  if(oEl) oEl.textContent=fmt(ot);
  updateTotalsDetail();
}

// The panel behind the "Breakdown" toggle. It re-splits exactly the hours the
// bar already shows — nothing is stored, and nothing here is a second source
// of truth: change a time and this is rebuilt from state.editedShifts.
function updateTotalsDetail(){
  const box=$('totalsDetail');
  if(!box) return;
  const {month,year}=getMonthYear();
  const hol=(month!==null&&year)?getSAPublicHolidays(year):new Map();
  let dutyDays=0,leaveDays=0,normal=0,overtime=0,weekendHrs=0,phHrs=0;
  for(const [k,es] of Object.entries(state.editedShifts||{})){
    if(!es) continue;
    if(typeof isLeaveActivity==='function'&&isLeaveActivity(es.typeLabel)){leaveDays++;continue;}
    const n=hoursBetween(es.nf,es.nt);
    const a=hoursBetween(es.ot1f,es.ot1t), b=hoursBetween(es.ot2f,es.ot2t);
    // of/ot mirrors the banded overtime, so it only counts where there is none.
    const ot=(a+b)>0?(a+b):hoursBetween(es.of,es.ot);
    if(n+ot<=0) continue;
    dutyDays++; normal+=n; overtime+=ot;
    if(month!==null&&year){
      const d=new Date(year,month,parseInt(k,10));
      // A public holiday that lands on a weekend counts once, as a holiday.
      if(hol.has(dateKeyLocal(d))) phHrs+=n+ot;
      else if(d.getDay()===0||d.getDay()===6) weekendHrs+=n+ot;
    }
  }
  const fmt=v=>(Math.round(v*10)/10).toString().replace(/\.0$/,'')+' h';
  const days=n=>n===1?'1 day':n+' days';
  const rows=[['Days on duty',days(dutyDays)],['Days on leave',days(leaveDays)],
              ['Normal hours',fmt(normal)],['Overtime hours',fmt(overtime)],
              ['Weekend hours',fmt(weekendHrs)],['Public holiday hours',fmt(phHrs)]];
  box.innerHTML=rows.map(([k,v])=>
    `<div class="td-row"><span class="td-k">${k}</span><span class="td-v">${v}</span></div>`).join('');
}

function buildPreview(doctorName,targetMonth,targetYear){
  const holidays=getSAPublicHolidays(targetYear);
  const daysInMonth=new Date(targetYear,targetMonth+1,0).getDate();
  state.editedShifts={};

  // Consultant-type profile: use consultant parser output only
  // Skip getDoctorShifts entirely — consultant days in rosterData use different column semantics
  const isConsultantMode = isExtendedRosterMode();

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
  const consultantAdded = overlayConsultantShifts(doctorName, targetMonth, targetYear)
                        + overlayTableShifts(doctorName, targetMonth, targetYear);
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
  const isConsultantMode2 = isExtendedRosterMode();
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
  const activeTypes = isTableRosterMode() ? tableActivityTypes()
    : isConsultantMode2 ? CONSULTANT_ACTIVITY_TYPES : ACTIVITY_TYPES;
  const typeOpts=activeTypes.map(t=>`<option value="${t}">${t}</option>`).join('');
  let html=`
  <div class="preview-note">
    Edit time fields or change activity type — changes save automatically. Click <strong>+</strong> to add an activity.
  </div>
  <div class="preview-wrapper"><table class="preview-table ${isConsultantMode2?'pt-ext':'pt-std'}">
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
    const phStyle=phName?'color:var(--color-accent-700);font-weight:800;':'';
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
          <td><input class="time-edit" data-day="${d}" data-field="ot1f" value="${es.ot1f||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
          <td><input class="time-edit" data-day="${d}" data-field="ot1t" value="${es.ot1t||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
          <td><input class="time-edit" data-day="${d}" data-field="ot2f" value="${es.ot2f||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
          <td><input class="time-edit" data-day="${d}" data-field="ot2t" value="${es.ot2t||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
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
        <td colspan="${cColspan}" style="font-style:italic;color:var(--color-accent-700)">${phName}</td>
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
  <div class="preview-foot">
    ${sc} activit${sc!==1?'ies':'y'} found &middot; <strong>${doctorName}</strong> &middot; ${MONTH_NAMES[targetMonth]} ${targetYear}
  </div>`;
  if(phFootnotes.length>0){
    html+=`<div class="ph-footnotes">`+
      phFootnotes.map(f=>`<span style="margin-right:16px;"><sup style="font-size:9px;">${f.letter}</sup> ${f.name}</span>`).join('')+
    `</div>`;
  }
  $('previewArea').innerHTML=html;
  attachEditHandlers();
}

function makeRowInner(d,isWE,phName,dayName,es){
  const isSpecial=isWE||!!phName;
  const isConsMode=isExtendedRosterMode();
  const selectedType=es?.typeLabel||defaultTypeLabel(isSpecial);
  const phStyle=phName?'color:var(--color-accent-700);font-weight:800;':'';
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
    <td><input class="time-edit" data-day="${d}" data-field="ot1f" value="${es?.ot1f||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot1t" value="${es?.ot1t||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot2f" value="${es?.ot2f||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot2t" value="${es?.ot2t||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td class="action-cell"><button class="row-clear" data-day="${d}" title="Remove">&times;</button>${isDayEdited(d)?`<button class="row-undo" data-day="${d}" title="Undo">&#8635;</button>`:''}</td>`;
  }
  return `
    ${dateCell}
    ${dayCell}
    <td><select class="type-select" data-day="${d}" data-is-special="${isSpecial?1:0}">${typeOptsFor(isWE,!!phName,selectedType)}</select></td>
    <td><input class="time-edit" data-day="${d}" data-field="nf" value="${es?.nf||''}" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="nt" value="${es?.nt||''}" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="of" value="${es?.of||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td><input class="time-edit" data-day="${d}" data-field="ot" value="${es?.ot||''}" placeholder="\u2014" maxlength="5" inputmode="numeric"></td>
    <td class="action-cell"><button class="row-clear" data-day="${d}" title="Remove">&times;</button>${isDayEdited(d)?`<button class="row-undo" data-day="${d}" title="Undo">&#8635;</button>`:''}</td>`;
}

function attachEditHandlers(){
  updatePreviewTotals();
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
      fresh.style.borderColor=v.length>0&&!normaliseTime(v)?'var(--color-danger)':'';
    });
    fresh.addEventListener('blur',()=>{
      const d=parseInt(fresh.dataset.day),field=fresh.dataset.field;
      const val=fresh.value.trim().toUpperCase(),normalised=normaliseTime(val);
      if(normalised){
        fresh.value=normalised;fresh.style.borderColor='';fresh.title='';
        if(!state.editedShifts[d]) state.editedShifts[d]={nf:'',nt:'',of:null,ot:null,label:'Custom',typeLabel:'WD Shift - 08H00',isWE:false};
        if(state.editedShifts[d][field]!==normalised){
          state.editedShifts[d][field]=normalised;
          applyBandSync(d,syncFollowingBand(d,field,normalised),normalised);
          markDirty(d);
        }
      } else if(val===''){
        fresh.style.borderColor='';
        if(state.editedShifts[d]&&state.editedShifts[d][field]!==null){state.editedShifts[d][field]=null;markDirty(d);}
      } else {fresh.style.borderColor='var(--color-danger)';fresh.title='Format: HHH00 (e.g. 08H00)';}
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
        const _phStyle='color:var(--color-accent-700);font-weight:800;';
        const _phLetter=(state.phLetterMap&&state.phLetterMap[d])||'';
        const _isConsCP=isExtendedRosterMode();
        const _colspanPH=_isConsCP?7:5;
        row.innerHTML=`<td style="${_phStyle}">${d}<sup>${_phLetter}</sup></td><td style="${_phStyle}">${dayName}</td>
          <td colspan="${_colspanPH}" style="font-style:italic;color:var(--color-accent-700)">${phName}</td>
          <td class="action-cell"><button class="row-add" title="Add shift" data-day="${d}" data-is-we="1" data-is-special="1">+</button></td>`;
      } else {
        row.className='empty-row'+(isWE?' we-row':'');row.style.opacity='';
        const _isConsC=isExtendedRosterMode();
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
        const isConsMode=isExtendedRosterMode();
        const defaultLabel=defaultTypeLabel(isSpecialNew);
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
  showSupervisorBox(isOther, true);
  if(!isOther) $('detailSupervisorOther').value='';
  saveDetailsToState();checkDetailsComplete();
});
$('detailSupervisorOther').addEventListener('input',()=>{ saveDetailsToState();checkDetailsComplete(); });
// Chrome only opens the picker from its own calendar glyph, which is the part
// we have hidden, so ask for it explicitly. The gesture is real and the input
// is rendered, so this is allowed; where it is not supported the native tap
// behaviour still applies.
$('detailSigDatePicker').addEventListener('click',e=>{
  if(typeof e.currentTarget.showPicker==='function'){
    try{ e.currentTarget.showPicker(); }catch(_){}
  }
});
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

// The Z1(a)'s Component is filed as the department alone — the hospital is
// already established by the rest of the form. Profiles are named for people
// picking them off a list ("VHW Anaesthetics"), so the hospital is stripped
// off either end, with or without a separator.
const Z1_HOSPITAL='\\b(?:victoria\\s+hospital(?:\\s+wynberg)?|vhw|vh)\\b';
function stripHospitalName(name){
  return String(name||'')
    .replace(new RegExp('^\\s*'+Z1_HOSPITAL+'\\s*[\\u2013\\u2014\\-:,]?\\s*','i'),'')
    .replace(new RegExp('\\s*[\\u2013\\u2014\\-:,]?\\s*'+Z1_HOSPITAL+'\\s*$','i'),'')
    .trim();
}
function z1ComponentFor(){
  const DEFAULT='Emergency Medicine';
  if(!activeProfile) return DEFAULT;
  // The full name first: ec_short is an abbreviation for badges, not a
  // department name the payroll office would recognise.
  const raw=activeProfile.z1_component||activeProfile.ec_name||activeProfile.ec_short||'';
  return stripHospitalName(raw)||DEFAULT;
}
function getFormDetails(){
  saveDetailsToState();
  const {month,year}=getMonthYear();
  return {
    firstName:state.savedDetails.firstName, surname:state.savedDetails.surname,
    persal:state.savedDetails.persal, supervisorName:state.savedDetails.supervisor,
    designation:state.savedDetails.designation,
    signatureDate:state.savedDetails.sigDate,
    addressDuringLeave:state.savedDetails.address||'',
    component:z1ComponentFor(),
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
  const note=btn.querySelector('.dlnote'),prevNote=note?note.textContent:'';
  if(note) note.innerHTML='<span class="spinner"></span> Generating\u2026';
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
  btn.disabled=false;
  if(note) note.textContent=prevNote;
  checkDetailsComplete();
});

$('annexureCBtn').addEventListener('click',async()=>{
  const d=getFormDetails();
  const btn=$('annexureCBtn');
  btn.disabled=true;
  const note=btn.querySelector('.dlnote'),prevNote=note?note.textContent:'';
  if(note) note.innerHTML='<span class="spinner"></span> Generating…';
  try{
    const blob=await generateAnnexureCDocx(d);
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    const safe=(d.firstName+' '+d.surname).trim().replace(/\s+/g,'_');
    a.href=url; a.download=`Annexure_C_${safe}_${MONTH_NAMES[d.month]}_${d.year}.docx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  }catch(err){alert('Error generating Annexure C: '+err.message);console.error(err);}
  btn.disabled=false;
  if(note) note.textContent=prevNote;
  checkDetailsComplete();
});

$('z1aBtn').addEventListener('click',async()=>{
  const d=getFormDetails();
  const btn=$('z1aBtn');
  btn.disabled=true;
  const note=btn.querySelector('.dlnote'),prevNote=note?note.textContent:'';
  if(note) note.innerHTML='<span class="spinner"></span> Generating…';
  try{
    const blob=await generateZ1ADocx(d);
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    const safe=(d.firstName+' '+d.surname).trim().replace(/\s+/g,'_');
    a.href=url; a.download=`Z1a_Leave_${safe}_${MONTH_NAMES[d.month]}_${d.year}.docx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  }catch(err){alert('Error generating Z1(a): '+err.message);console.error(err);}
  btn.disabled=false;
  if(note) note.textContent=prevNote;
  checkDetailsComplete();
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

// ═══════════════════════════════════════════════════════════════
// Section gating hints + "Set up a new EC" entry point.
// Steps 01–03 are revealed by the app as the user progresses
// (display toggled on #step1 / #step2 / #detailsSection); each one
// shows a short placeholder in its section until then.
// ═══════════════════════════════════════════════════════════════
(function(){
  const GATES=[['step1','step1Empty'],['step2','step2Empty'],['detailsSection','sec3Empty']];
  function syncGates(){
    for(const [id,hintId] of GATES){
      const el=document.getElementById(id),hint=document.getElementById(hintId);
      if(!el||!hint) continue;
      hint.style.display=getComputedStyle(el).display==='none'?'':'none';
    }
  }
  function init(){
    syncGates();
    const mo=new MutationObserver(syncGates);
    GATES.forEach(([id])=>{const el=document.getElementById(id);if(el)mo.observe(el,{attributes:true,attributeFilter:['style','class']});});
    const link=document.getElementById('openWizardLink');
    if(link) link.addEventListener('click',e=>{
      e.preventDefault();
      if(typeof window.openWizard==='function') window.openWizard();
    });
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init);
  else init();
})();

// ═══════════════════════════════════════════════════════════════════════════
// WIZARD SHELL
// Four steps, one on screen at a time. All state is the existing in-memory
// `state` object plus the two fields below — nothing is written to storage,
// so closing the tab genuinely discards the roster, the edits and the details.
// ═══════════════════════════════════════════════════════════════════════════
const WIZ_STEPS_DEF = [
  { n: 1, sec: 'sec-1', title: 'Upload roster files' },
  { n: 2, sec: 'sec-2', title: 'Review schedule' },
  { n: 3, sec: 'sec-3', title: 'Your details' },
  { n: 4, sec: 'sec-4', title: 'Generate documents' },
];
// Session-only: deliberately not persisted, and reset by fullReset().
let wizStep = 1;
let wizReviewed = false;

function wizExtracted(){
  return !!(state.parsedFiles && state.parsedFiles.length) || !!state.tableData;
}
function wizPreviewed(){
  const { month, year } = getMonthYear();
  return !!state.selectedDoctor && month !== null && !!year &&
         !!state.editedShifts && Object.keys(state.editedShifts).length > 0;
}
function wizDetailsDone(){
  const b = $('proceedDownloadBtn');
  return !!b && !b.disabled;
}
// Why the user cannot move on yet — shown, never merely implied.
function wizBlockedReason(step){
  if (step === 1) return wizExtracted() ? null
    : 'Add a roster file and choose Extract data before continuing.';
  if (step === 2) {
    if (!wizPreviewed()) return 'Pick a name and a month, then choose Preview schedule.';
    if (!wizReviewed)    return 'Confirm you have reviewed every day in the month.';
    return null;
  }
  if (step === 3) return wizDetailsDone() ? null
    : 'Complete every field marked with an asterisk.';
  return null;
}
function wizCanEnter(step){
  for (let s = 1; s < step; s++) if (wizBlockedReason(s)) return false;
  return true;
}

function wizRenderSteps(){
  const host = $('wizSteps');
  if (host) host.innerHTML = WIZ_STEPS_DEF.map(s => {
    const done = s.n < wizStep && !wizBlockedReason(s.n);
    const can  = s.n !== wizStep && wizCanEnter(s.n);
    return `<button type="button" class="wizstep" data-go="${s.n}"
      ${s.n === wizStep ? 'aria-current="step"' : ''}
      ${done ? 'data-done="1"' : ''} ${can ? 'data-clickable="1"' : 'disabled'}>
      <span class="n">Step ${s.n}</span><span class="t">${s.title}</span></button>`;
  }).join('');
  // Only the counter: the section heading right below it already names the
  // step, and printing it twice was what made the bar feel crowded.
  const m = $('wizStepMobile');
  if (m) m.innerHTML = `<span class="n">Step ${wizStep} of 4</span>`;
}

// Which rows the Edit panel shows, and what it is called while it shows them.
function showEditChoices(which){
  const p = $('wizChangePeriod'), d = $('wizChangeDept'), t = $('wizEditTitle');
  const sv = $('wizStartOver');
  if (p) p.hidden = which === 'dept';
  if (d) d.hidden = which === 'period';
  if (sv) sv.hidden = which !== 'all';
  if (t) t.textContent = which === 'period' ? 'Change period'
                       : which === 'dept'   ? 'Change department' : 'Edit';
}

function wizRenderContext(){
  const { month, year } = getMonthYear();
  const per = $('wizPeriod');
  if (per) per.textContent = (month !== null && year)
    ? `${MONTH_NAMES[month]} ${year}` : 'No month selected';
  const dept = $('wizDept');
  const deptName = (activeProfile && activeProfile.ec_name)
    ? activeProfile.ec_name : 'VHW Emergency Medicine';
  if (dept) dept.textContent = deptName;
  // The same two facts, as the header's controls.
  const hp = $('hdrPeriodBtn'), hd = $('hdrDeptBtn'), hc = $('hdrCtx'), bar = $('wizBar');
  if (hp && per) hp.textContent = per.textContent;
  if (hd) hd.textContent = deptName;
  if (hc && bar) hc.hidden = bar.hidden;
  for (const [id, n] of [['dlPeriod1', 1], ['dlPeriod2', 2], ['dlPeriod3', 3]]) {
    const el = $(id);
    if (el) el.textContent = (month !== null && year) ? `${MONTH_NAMES[month]} ${year}` : '';
  }
}

function wizRenderNav(){
  for (const s of WIZ_STEPS_DEF) {
    const nav = $('wizNav' + s.n);
    if (!nav) continue;
    const why = nav.querySelector('.wiznav-why');
    const fwd = nav.querySelector('[data-wiz="next"]');
    const reason = wizBlockedReason(s.n);
    if (fwd) fwd.disabled = !!reason;
    // Step 3's reason is already printed under the fields it is about, so the
    // button stays disabled there but says nothing a second time.
    const spoken = s.n === 3 ? null : reason;
    if (why) {
      why.hidden = !spoken;
      const t = why.querySelector('.txt');
      if (t) t.textContent = spoken || '';
    }
  }
}

function wizGo(step){
  step = Math.max(1, Math.min(4, step));
  if (step > wizStep && wizBlockedReason(wizStep)) return;
  if (!wizCanEnter(step)) return;
  wizStep = step;
  for (const s of WIZ_STEPS_DEF) {
    const el = document.getElementById(s.sec);
    if (el) el.hidden = s.n !== step;
  }
  wizRefresh();
  const bar = $('wizBar');
  if (bar) window.scrollTo({ top: 0, behavior: 'smooth' });
  const h = document.querySelector('#' + WIZ_STEPS_DEF[step - 1].sec + ' .sec-head h2');
  if (h) { h.setAttribute('tabindex', '-1'); h.focus({ preventScroll: true }); }
}

function wizRefresh(){
  wizRenderSteps(); wizRenderContext(); wizRenderNav();
  const hb = $('totalsToggle2');
  if (hb) hb.disabled = !wizPreviewed();
  // The button beside Preview schedule greys out; the header and bar icons
  // are hidden until there is a file, so they appear when they become useful.
  const anyFile = !!rosterViewFiles().length;
  const vr = $('viewRosterBtn');
  if (vr) vr.disabled = !anyFile;
  for (const id of ['hdrViewBtn','wizViewBtn']) {
    const el = $(id);
    if (el) el.hidden = !anyFile;
  }
  const ack = $('reviewAckWrap');
  if (ack) ack.hidden = !wizPreviewed();
  buildAttentionItems();
}

// ── Attention items ────────────────────────────────────────────────────────
// A reading aid only. It flags what looks odd; it never marks anything
// reviewed, and continuing still requires the acknowledgement below it.
function buildAttentionItems(){
  const panel = $('attentionPanel');
  if (!panel) return;
  if (!wizPreviewed()) { panel.hidden = true; return; }
  const items = [];
  const { month, year } = getMonthYear();
  const seen = new Set();
  for (const w of (state.tableWarnings || [])) {
    const txt = String(w && w.message ? w.message : w);
    if (seen.has(txt)) continue;
    seen.add(txt);
    const m = txt.match(/\b(\d{1,2})\b/);
    items.push({ day: m ? parseInt(m[1], 10) : null, what: txt });
  }
  for (const [k, es] of Object.entries(state.editedShifts || {})) {
    const d = parseInt(k, 10);
    if (!es || !es.typeLabel) continue;
    if (typeof isLeaveActivity === 'function' && isLeaveActivity(es.typeLabel)) continue;
    const pairs = [[es.nf, es.nt, 'normal hours'], [es.ot1f, es.ot1t, 'OT1'],
                   [es.ot2f, es.ot2t, 'OT2'], [es.of, es.ot, 'overtime']];
    for (const [a, b, lbl] of pairs) {
      if ((a && !b) || (!a && b)) items.push({ day: d, what: `${lbl} has only one time` });
    }
    const total = hoursBetween(es.nf, es.nt) + hoursBetween(es.ot1f, es.ot1t) +
                  hoursBetween(es.ot2f, es.ot2t) +
                  (hoursBetween(es.ot1f, es.ot1t) ? 0 : hoursBetween(es.of, es.ot));
    if (total > 24.01) items.push({ day: d, what: `${total.toFixed(1)} h in one day` });
  }
  items.sort((a, b) => (a.day || 99) - (b.day || 99));
  const icon = $('attentionIcon'), title = $('attentionTitle'),
        note = $('attentionNote'), list = $('attentionList');
  panel.hidden = false;
  const monthName = month !== null ? MONTH_NAMES[month] : 'this month';
  if (!items.length) {
    panel.className = 'attention is-ok';
    if (icon) icon.textContent = '✓';
    if (title) title.textContent = 'No extraction issues detected';
    if (note) { note.textContent = ''; note.hidden = true; }
    if (list) list.innerHTML = '';
    return;
  }
  panel.className = 'attention is-warn';
  if (icon) icon.textContent = '⚠';
  if (title) title.textContent =
    `${items.length} ${items.length === 1 ? 'entry' : 'entries'} may need attention`;
  if (note) { note.hidden = false; note.textContent =
    `These are possible problems, not errors. Reviewing them does not replace checking the whole of ${monthName}.`; }
  if (list) list.innerHTML = items.slice(0, 12).map(it => {
    const label = it.day ? `Review ${it.day} ${monthName}` : 'Review the schedule';
    return `<li><button type="button" data-attn-day="${it.day || ''}">${label}</button>
      <span class="what">— ${String(it.what).replace(/</g, '&lt;')}</span></li>`;
  }).join('');
}

document.addEventListener('click', e => {
  const go = e.target.closest && e.target.closest('[data-go]');
  if (go) { e.preventDefault(); wizGo(parseInt(go.dataset.go, 10)); return; }
  const nav = e.target.closest && e.target.closest('[data-wiz]');
  if (nav) {
    e.preventDefault();
    wizGo(nav.dataset.wiz === 'next' ? wizStep + 1 : wizStep - 1);
    return;
  }
  const attn = e.target.closest && e.target.closest('[data-attn-day]');
  if (attn) {
    e.preventDefault();
    const d = attn.dataset.attnDay;
    const row = d && document.querySelector(`.preview-table tr[data-day="${d}"]`);
    if (row) {
      row.scrollIntoView({ block: 'center', behavior: 'smooth' });
      row.style.outline = '2px solid var(--color-accent)';
      setTimeout(() => { row.style.outline = ''; }, 2000);
    }
  }
});

document.addEventListener('click', e => {
  const t = e.target.closest && e.target.closest('button');
  if (!t) return;
  // The period lives on step 2 and the department in the picker overlay, so
  // the context row sends the user to the control rather than duplicating it.
  if (t.id === 'wizChangePeriod') {
    wizGo(2);
    const m = $('monthSelect');
    if (m) { m.focus(); m.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
    return;
  }
  if (t.id === 'wizChangeDept') {
    if (typeof window.reopenEcPicker === 'function') window.reopenEcPicker();
    return;
  }
});

document.addEventListener('change', e => {
  if (e.target && e.target.id === 'reviewAck') { wizReviewed = e.target.checked; wizRefresh(); }
});

// Editing anything in the schedule invalidates the acknowledgement: the user
// confirmed the month as it was, not as it now is.
function wizInvalidateReview(){
  if (!wizReviewed) return;
  wizReviewed = false;
  const cb = $('reviewAck');
  if (cb) cb.checked = false;
  wizRefresh();
}

// Locking the page hides the scrollbar. Where its slot is not already
// reserved by scrollbar-gutter, the layout widens by that much and every
// centred thing jumps right, so put back exactly the width locking took.
function lockPageScroll(){
  const el=document.documentElement;
  const before=el.clientWidth;
  el.style.overflow='hidden';
  // Belt and braces: if a browser still releases the gutter, put the width back.
  const grew=el.clientWidth-before;
  if(grew>0) el.style.paddingRight=grew+'px';
}
function unlockPageScroll(){
  const el=document.documentElement;
  el.style.overflow='';
  el.style.paddingRight='';
}

// ── Roster viewer ──────────────────────────────────────────────────────────
// Looking at what was uploaded, in the page. A PDF is drawn with the PDF.js
// that already ships here; anything grid-shaped shows the rows the reader
// recovered, which is what a disagreement with the parser is usually about.
// The file is read again from the handle held in state — nothing is copied and
// nothing leaves the browser.
function rosterViewFiles(){
  return (state.parsedFiles || []).filter(f => f && f.file);
}
async function renderRosterView(){
  const body = $('rosterViewBody'), note = $('rosterViewNote'), pick = $('rosterViewPick');
  if (!body) return;
  const files = rosterViewFiles();
  const wrap = document.querySelector('.rv-pick');
  if (wrap) wrap.hidden = files.length < 2;
  if (!files.length) {
    body.innerHTML = '';
    if (note) note.textContent = 'No roster file is loaded. Upload one on step 1 and choose Extract data.';
    return;
  }
  const entry = files[Math.min(pick ? pick.selectedIndex : 0, files.length - 1)];
  body.innerHTML = '<p class="rv-note">Reading\u2026</p>';
  rvPdfPages = []; rvGridRows = null; rvHits = []; rvHitIdx = -1;
  try {
    const buf = await readFile(entry.file);
    const ext = String(entry.name).split('.').pop().toLowerCase();
    if (ext === 'pdf') {
      await drawPdfInto(body, buf);
      rvApplyFind();
      if (note) note.textContent = 'The file as uploaded. Compare it with the schedule behind this panel.';
    } else {
      const { rows, tables } = await gridRowsFor(buf, entry.name);
      rvGridRows = rows;
      rvApplyFind();
      if (note) note.textContent =
        `The rows the reader recovered — ${rows.length} row${rows.length===1?'':'s'}` +
        (tables > 1 ? ` from the largest of ${tables} tables in the file` : '') + '. ' +
        'If a name or a date is missing here, the reader did not find it in the file.';
    }
  } catch (err) {
    body.innerHTML = '';
    if (note) note.textContent = 'Could not read this file back: ' + (err && err.message ? err.message : err);
  }
}
async function drawPdfInto(host, buf){
  host.innerHTML = '';
  rvPdfPages = [];
  const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const vp = page.getViewport({ scale: 2 });
    // The canvas is painted once; highlights go in a sibling layer on top of
    // it, positioned in percentages so they survive the canvas being scaled
    // down to the panel width. Re-searching never re-renders a page.
    const wrap = document.createElement('div');
    wrap.className = 'rv-page';
    const c = document.createElement('canvas');
    c.width = vp.width; c.height = vp.height;
    const layer = document.createElement('div');
    layer.className = 'rv-marks';
    wrap.appendChild(c); wrap.appendChild(layer);
    host.appendChild(wrap);
    await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    const items = (await page.getTextContent()).items;
    rvPdfPages.push({ layer, vp, items, index: rvIndexPage(items) });
  }
}
async function gridRowsFor(buf, name){
  // extractWordTables returns every table it can find. A legacy .doc needs the
  // column count to recover its row boundaries at all, so pass the profile's
  // when there is one. The roster is the biggest table; the rest are headers
  // and notes.
  const nCols = (activeProfile && activeProfile.table && activeProfile.table.columns || []).length || undefined;
  const tables = await extractWordTables(buf, name, nCols);
  const rows = tables.reduce((best, t) => (t.length > best.length ? t : best), []);
  return { rows, tables: tables.length };
}
function drawGridInto(host, rows, re){
  const esc = v => String(v == null ? '' : v).replace(/&/g,'&amp;').replace(/</g,'&lt;');
  // A cell is escaped in pieces so the <mark> can be inserted around a match
  // without the escaping swallowing it.
  const cell = v => {
    const raw = String(v == null ? '' : v);
    if (!re) return esc(raw);
    re.lastIndex = 0;
    let out = '', last = 0, m;
    while ((m = re.exec(raw))) {
      if (!m[0].length) { re.lastIndex++; continue; }
      out += esc(raw.slice(last, m.index)) + '<mark class="rv-hit">' + esc(m[0]) + '</mark>';
      last = m.index + m[0].length;
    }
    return out + esc(raw.slice(last));
  };
  const width = rows.reduce((m, r) => Math.max(m, (r || []).length), 0);
  const head = '<tr><th class="rv-rownum">#</th>' +
    Array.from({length: width}, (_, i) => '<th>' + (i + 1) + '</th>').join('') + '</tr>';
  const cells = rows.map((r, i) => '<tr><td class="rv-rownum">' + (i + 1) + '</td>' +
    Array.from({length: width}, (_, c) => '<td>' + cell((r || [])[c]) + '</td>').join('') + '</tr>').join('');
  host.innerHTML = '<table class="rv-grid"><thead>' + head + '</thead><tbody>' + cells + '</tbody></table>';
}
// Opening the viewer is the same job from three places: the button beside
// Preview schedule, and the header/bar icons that stay in reach while the user
// is scrolling the schedule. One preparer for all of them.
function rosterViewOpen(){
  const pick=$('rosterViewPick');
  if(pick){
    const files=rosterViewFiles();
    pick.innerHTML=files.map(f=>`<option>${String(f.name).replace(/</g,'&lt;')}</option>`).join('');
  }
  const find=$('rosterViewFind');
  if(find) find.value=state.selectedDoctor||'';
  renderRosterView();
}

// ── Roster viewer: find and highlight ──────────────────────────────────────
// The viewer exists so a suspect parse can be checked against the file, and on
// a month-wide consultant roster that means following one surname down the
// page. Every occurrence is boxed and the arrows step through them.
//
// PDF text arrives as items, not lines, and a name can be split across two of
// them, so each page is indexed into one string with a map back to the item
// and character a match starts at. Highlighting is geometry over the painted
// canvas; the grid path re-renders its table with <mark> instead.
let rvPdfPages = [];    // rendered pages: mark layer, viewport, text geometry
let rvGridRows = null;  // the grid path re-marks from these, not from the file
let rvHits = [];        // one entry per match; a match can span several boxes
let rvHitIdx = -1;
let rvFindTimer = null;

// Spaces in the query are loosened to \s* because a PDF splits text wherever
// it likes: "De Haan" can arrive as "De" + "Haan" with no space between them.
function rvFindRe(q){
  const esc = String(q || '').trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (!esc) return null;
  return new RegExp(esc.replace(/\s+/g, '\\s*'), 'gi');
}

function rvIndexPage(items){
  let text = ''; const map = [];
  items.forEach((it, i) => {
    const s = it.str || '';
    for (let c = 0; c < s.length; c++) map.push([i, c]);
    text += s;
    if (it.hasEOL) { text += '\n'; map.push(null); }
  });
  return { text, map };
}

// Where characters [from,to) of a text item sit on the page, as percentages of
// the viewport. PDF.js gives a width for the whole item only, so a partial
// match is apportioned by character count — near enough to box a name.
function rvItemRect(it, vp, from, to){
  const tx = window.pdfjsLib.Util.transform(vp.transform, it.transform);
  const h = Math.hypot(tx[2], tx[3]);
  const per = (it.width * vp.scale) / ((it.str || '').length || 1);
  return {
    left:   (tx[4] + from * per) / vp.width * 100,
    top:    (tx[5] - h) / vp.height * 100,
    width:  Math.max(to - from, 1) * per / vp.width * 100,
    height: h / vp.height * 100,
  };
}

function rvMarkPdf(re){
  rvPdfPages.forEach(p => {
    const { text, map } = p.index;
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(text))) {
      if (!m[0].length) { re.lastIndex++; continue; }
      // Consecutive characters landing in the same item share one box.
      const runs = [];
      for (let k = m.index; k < m.index + m[0].length; k++) {
        const at = map[k];
        if (!at) continue;
        const last = runs[runs.length - 1];
        if (last && last.i === at[0] && last.to === at[1]) last.to = at[1] + 1;
        else runs.push({ i: at[0], from: at[1], to: at[1] + 1 });
      }
      const boxes = runs.map(g => {
        const r = rvItemRect(p.items[g.i], p.vp, g.from, g.to);
        const el = document.createElement('i');
        el.style.left = r.left + '%'; el.style.top = r.top + '%';
        el.style.width = r.width + '%'; el.style.height = r.height + '%';
        p.layer.appendChild(el);
        return el;
      });
      if (boxes.length) rvHits.push(boxes);
    }
  });
}

// Applied on every keystroke, so it never re-reads the file or re-paints a
// page. Nothing is current until the user steps: the count reads "12 matches"
// until then, and the panel does not jump while they are still typing.
function rvApplyFind(){
  const input = $('rosterViewFind');
  const re = rvFindRe(input ? input.value : '');
  rvPdfPages.forEach(p => { p.layer.innerHTML = ''; });
  rvHits = []; rvHitIdx = -1;
  if (rvPdfPages.length) {
    if (re) rvMarkPdf(re);
  } else if (rvGridRows) {
    const body = $('rosterViewBody');
    if (body) {
      drawGridInto(body, rvGridRows, re);
      rvHits = Array.from(body.querySelectorAll('mark.rv-hit')).map(el => [el]);
    }
  }
  rvRenderCount();
}

function rvSetHit(i){
  rvHits.forEach(g => g.forEach(el => el.classList.remove('is-current')));
  rvHitIdx = i;
  if (i < 0 || !rvHits[i]) return;
  rvHits[i].forEach(el => el.classList.add('is-current'));
  rvHits[i][0].scrollIntoView({ block: 'center' });
}

function rvStep(d){
  if (!rvHits.length) return;
  const n = rvHits.length;
  rvSetHit(rvHitIdx < 0 ? (d > 0 ? 0 : n - 1) : (rvHitIdx + d + n) % n);
  rvRenderCount();
}

function rvRenderCount(){
  const el = $('rosterViewCount'), input = $('rosterViewFind');
  const q = input ? input.value.trim() : '';
  const n = rvHits.length;
  if (el) el.textContent = !q ? ''
    : !n ? 'No matches'
    : rvHitIdx < 0 ? n + (n === 1 ? ' match' : ' matches')
    : (rvHitIdx + 1) + ' of ' + n;
  const prev = $('rosterViewPrev'), next = $('rosterViewNext');
  if (prev) prev.disabled = !n;
  if (next) next.disabled = !n;
}

document.addEventListener('input', e => {
  if (!e.target || e.target.id !== 'rosterViewFind') return;
  clearTimeout(rvFindTimer);
  rvFindTimer = setTimeout(rvApplyFind, 160);
});
document.addEventListener('keydown', e => {
  if (!e.target || e.target.id !== 'rosterViewFind' || e.key !== 'Enter') return;
  e.preventDefault();
  clearTimeout(rvFindTimer);
  if (!rvHits.length) rvApplyFind();
  rvStep(e.shiftKey ? -1 : 1);
});
document.addEventListener('click', e => {
  if (!e.target || !e.target.closest) return;
  if (e.target.closest('#rosterViewPrev')) rvStep(-1);
  else if (e.target.closest('#rosterViewNext')) rvStep(1);
});

document.addEventListener('change', e => {
  if (e.target && e.target.id === 'rosterViewPick') renderRosterView();
});

// ── Confirmation ───────────────────────────────────────────────────────────
// One panel in front of everything destructive. Nothing is cleared, removed or
// reset until it comes back true; closing it any way at all — the x, No, the
// backdrop, Escape — answers false.
let confirmSettle = null;
function askConfirm(text){
  const ov = $('confirmOverlay');
  if (!ov) return Promise.resolve(window.confirm(text));
  const body = $('confirmText');
  if (body) body.textContent = text;
  const shell = document.querySelector('.shell');
  // A confirmation raised from inside another panel closes that one first, so
  // there is only ever one thing on screen to answer.
  for (const open of document.querySelectorAll('.modal-overlay.open')) {
    open.classList.remove('open');
    for (const b of document.querySelectorAll('[aria-expanded="true"][aria-haspopup="dialog"]'))
      b.setAttribute('aria-expanded','false');
  }
  return new Promise(resolve => {
    confirmSettle = answer => {
      confirmSettle = null;
      ov.classList.remove('open');
      unlockPageScroll();
      if (shell) shell.inert = false;
      resolve(answer);
    };
    ov.classList.add('open');
    lockPageScroll();
    if (shell) shell.inert = true;
    const yes = $('confirmYes');
    if (yes) yes.focus();
  });
}
(function(){
  function wire(){
    const ov = $('confirmOverlay');
    if (!ov) return;
    const settle = a => { if (confirmSettle) confirmSettle(a); };
    $('confirmYes').addEventListener('click', () => settle(true));
    $('confirmNo').addEventListener('click', () => settle(false));
    $('confirmCloseBtn').addEventListener('click', () => settle(false));
    ov.addEventListener('click', e => { if (e.target === ov) settle(false); });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && ov.classList.contains('open')) settle(false);
    });
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',wire);
  else wire();
})();

// Which controls have to be asked about, and what to ask. Matching on the
// selector rather than tagging every button keeps the rows the preview table
// rebuilds — a Remove button is replaced on every edit — inside the net.
const CONFIRM_ACTIONS = [
  ['#clearBtn',        'Are you sure you want to clear all data?'],
  ['#clearDoctorBtn',  'Are you sure you want to clear the selection?'],
  ['.ri-remove',       'Are you sure you want to remove this file?'],
  ['.row-clear',       'Are you sure you want to remove this entry?'],
  ['#resetFormBtn',    'Are you sure you want to clear all data and start over?'],
  ['#hdrResetBtn',     'Are you sure you want to clear all data and start over?'],
  ['#wizStartOver',    'Are you sure you want to clear all data and start over?'],
];
// Capture, so the question is asked before the handlers that do the work. On
// yes the same click is sent again, flagged, and passes straight through.
document.addEventListener('click', e => {
  // Only a real click from a person gets asked about. The app clicks these
  // buttons itself — fullReset() presses Clear all to empty the queue — and so
  // does the re-dispatch below; intercepting those asked a question nobody was
  // there to answer and left the page inert behind it.
  if (!e.isTrusted) return;
  const el = e.target.closest && e.target.closest(
    CONFIRM_ACTIONS.map(([sel]) => sel).join(','));
  if (!el || el.dataset.confirmed === '1') return;
  const hit = CONFIRM_ACTIONS.find(([sel]) => el.matches(sel));
  if (!hit) return;
  e.preventDefault();
  e.stopPropagation();
  e.stopImmediatePropagation();
  askConfirm(hit[1]).then(ok => {
    if (!ok) return;
    el.dataset.confirmed = '1';
    el.click();
    delete el.dataset.confirmed;
  });
}, true);

// ── Modal panels ───────────────────────────────────────────────────────────
// Opening one marks .shell inert, so the page behind is genuinely muted to
// clicks, tabbing and assistive tech rather than just painted over.
(function(){
  // Both overlays are body-level markup that comes after this script, so the
  // wiring waits for the document rather than running at parse time.
  function dialog(btnId, overlayId, closeId, onOpen){
    const btns=[].concat(btnId).map(id=>document.getElementById(id)).filter(Boolean);
    const btn=btns[0];
    const ov=document.getElementById(overlayId);
    const closeBtn=document.getElementById(closeId);
    if(!btn||!ov||!closeBtn) return;
    const shell=document.querySelector('.shell');
    let lastFocus=null;
    function open(){
      if(typeof onOpen==='function') onOpen();
      lastFocus=document.activeElement;
      ov.classList.add('open');
      for(const b of btns) b.setAttribute('aria-expanded','true');
      lockPageScroll();
      if(shell) shell.inert=true;
      closeBtn.focus();
    }
    function close(){
      ov.classList.remove('open');
      for(const b of btns) b.setAttribute('aria-expanded','false');
      unlockPageScroll();
      if(shell) shell.inert=false;
      if(lastFocus&&lastFocus.focus) lastFocus.focus();
    }
    for(const b of btns) b.addEventListener('click',open);
    closeBtn.addEventListener('click',close);
    // Clicking the backdrop, but not the panel, closes it.
    ov.addEventListener('click',e=>{ if(e.target===ov) close(); });
    // Choosing an action inside the panel dismisses it; the action itself is
    // handled by the delegated wizard listener, which needs the page live.
    ov.addEventListener('click',e=>{
      if(e.target.closest && e.target.closest('.modal-choice')) close();
    });
    document.addEventListener('keydown',e=>{
      if(e.key==='Escape'&&ov.classList.contains('open')) close();
    });
  }
  function wire(){
    dialog('privacyBtn','privacyOverlay','privacyCloseBtn');
    // The bar's icon offers both; each header chip offers only its own, which
    // is what makes them read as controls for that one thing.
    dialog('wizEditBtn','wizEditOverlay','wizEditCloseBtn',()=>showEditChoices('all'));
    dialog('hdrPeriodBtn','wizEditOverlay','wizEditCloseBtn',()=>showEditChoices('period'));
    dialog('hdrDeptBtn','wizEditOverlay','wizEditCloseBtn',()=>showEditChoices('dept'));
    dialog(['totalsToggle','totalsToggle2'],'totalsOverlay','totalsCloseBtn',updateTotalsDetail);
    for(const id of ['viewRosterBtn','hdrViewBtn','wizViewBtn'])
      dialog(id,'rosterViewOverlay','rosterViewCloseBtn',rosterViewOpen);
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',wire);
  else wire();
})();

