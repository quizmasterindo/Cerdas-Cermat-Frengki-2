/* ---------- Utilities ---------- */
function shuffle(arr){ const a=(arr||[]).slice(); for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1)); [a[i],a[j]]=[a[j],a[i]];} return a; }
function dedupeById(arr){ const s=new Set(); return (arr||[]).filter(x=>x&&x.id&&!s.has(x.id)&&(s.add(x.id),true)); }

/* ---------- Team Colors (inline override for Bingo) ---------- */
const TEAM_COLOR_MAP = {
  A: { bg:'#2F80ED', fg:'#ffffff' }, // biru
  B: { bg:'#27AE60', fg:'#ffffff' }, // hijau
  C: { bg:'#F2994A', fg:'#ffffff' }, // oranye
  D: { bg:'#EB5757', fg:'#ffffff' }  // merah
};

/* ---------- Event persistence ---------- */
function eventKey(eventId){ return `quiz-event:${eventId}`; }
function roundStateKey(eventId,roundNo){ return `quiz-state:${eventId}:round${roundNo}`; }
function loadEvent(eventId){ const raw=localStorage.getItem(eventKey(eventId)); if(!raw) return null; try{ return JSON.parse(raw);}catch{return null;} }
function saveEvent(ev){ localStorage.setItem(eventKey(ev.eventId),JSON.stringify(ev)); }

/* ---------- State ---------- */
let teams={ A:{name:"Tim A",score:0}, B:{name:"Tim B",score:0}, C:{name:"Tim C",score:0}, D:{name:"Tim D",score:0} };
let activeTeam=null, mode="regular", bingoHappened=false, currentCell=null, reassignMode=false;
let eventId=null, roundNo=null;
let regularQuestions=[], takeoverQuestions=[];
let boardCells=[], takeoverQueue=[], takeoverIndex=0, regIndex=0;
let boardLocked=false;
let selectedTakeoverCell=null;
let takeoverAnsweredCount=0;
let bingoBonus = {
  diagonal: false,
  horizontal: false,
  vertical: false
};

/* ---------- Excel Question Bank ---------- */
let uploadedQuestions = [];
function splitQuestionBank(){
  regularQuestions = uploadedQuestions.filter(q =>
    !String(q.id || '').toUpperCase().includes('TO')
  );

  takeoverQuestions = uploadedQuestions.filter(q =>
    String(q.id || '').toUpperCase().includes('TO')
  );

  regularQuestions = shuffle(regularQuestions);
  takeoverQuestions = shuffle(takeoverQuestions);

  return {
    regular: regularQuestions.length,
    takeover: takeoverQuestions.length
  };
}
function validateQuestionBank(){

  if(uploadedQuestions.length === 0){
    throw new Error("Silakan upload Bank Soal terlebih dahulu.");
  }

  const info = splitQuestionBank();

  if(info.regular < 20){
    throw new Error(
      `Soal REGULAR hanya ${info.regular}. Minimal 20 soal.`
    );
  }

  if(info.takeover < 5){
    throw new Error(
      `Soal TAKEOVER hanya ${info.takeover}. Minimal 5 soal.`
    );
  }

  return true;
}

/* ---------- Excel Upload ---------- */

async function parseExcelFile(file){

  const data = await file.arrayBuffer();

  const workbook = XLSX.read(data,{
    type:'array'
  });

  const sheet =
    workbook.Sheets[
      workbook.SheetNames[0]
    ];

  const rows =
    XLSX.utils.sheet_to_json(sheet,{
      defval:''
    });

  uploadedQuestions = rows.map((r,i)=>({

    id:
      String(
        r.ID ||
        r.Id ||
        r.id ||
        ''
      ).trim(),

    text: `${String(
        r.Pertanyaan ||
        r.Question ||
        r.question ||
        ''
      ).trim()}
      
    A. ${String(r.A || '').trim()}
    B. ${String(r.B || '').trim()}
    C. ${String(r.C || '').trim()}
    D. ${String(r.D || '').trim()}`,

    ans:
      String(
        r.Jawaban ||
        r.Answer ||
        r.answer ||
        ''
      )
      .trim()
      .toLowerCase()      

  }))
  .filter(q => q.id);

  const info = splitQuestionBank();
  const fileNameEl = document.getElementById('fileName');
  if(fileNameEl){
    fileNameEl.textContent =
      `${uploadedQuestions.length} Soal | REG:${info.regular} | TO:${info.takeover}`;
}

}

/* ---------- Persist round ---------- */
function saveState(){
  if(!eventId||!roundNo) return;
  const s={eventId,roundNo,mode,bingoHappened,boardCells,takeoverQueue,takeoverAnsweredCount,takeoverIndex,teams,regularQuestions,takeoverQuestions,regIndex,activeTeam,bingoBonus};
  localStorage.setItem(roundStateKey(eventId,roundNo), JSON.stringify(s));
}
function loadState(eid,rno){
  const raw=localStorage.getItem(roundStateKey(eid,rno)); if(!raw) return false;
  try{
    const s=JSON.parse(raw);
    eventId=s.eventId; roundNo=s.roundNo;
    mode=s.mode; bingoHappened=s.bingoHappened;
    boardCells=s.boardCells; takeoverQueue=s.takeoverQueue; takeoverIndex=s.takeoverIndex; takeoverAnsweredCount=s.takeoverAnsweredCount || 0;
    teams=s.teams; regularQuestions=s.regularQuestions||[]; takeoverQuestions=s.takeoverQuestions||[];
    regIndex=typeof s.regIndex==='number'?s.regIndex:0; activeTeam=s.activeTeam||null;

    bingoBonus = s.bingoBonus || {
      diagonal:false,
      horizontal:false,
      vertical:false
    };

    try{
      if (document.getElementById("teamA")) document.getElementById("teamA").value = teams?.A?.name || "Tim A";
      if (document.getElementById("teamB")) document.getElementById("teamB").value = teams?.B?.name || "Tim B";
      if (document.getElementById("teamC")) document.getElementById("teamC").value = teams?.C?.name || "Tim C";
      if (document.getElementById("teamD")) document.getElementById("teamD").value = teams?.D?.name || "Tim D";
      bindTeamNameInputs();
    }catch{}
    return true;
  }catch{ return false; }
}

/* ---------- Team name live bindings ---------- */
function applyTeamNamesFromInputs(){
  const a = (document.getElementById("teamA")?.value || "").trim();
  const b = (document.getElementById("teamB")?.value || "").trim();
  const c = (document.getElementById("teamC")?.value || "").trim();
  const d = (document.getElementById("teamD")?.value || "").trim();
  teams.A.name = a || "Tim A";
  teams.B.name = b || "Tim B";
  teams.C.name = c || "Tim C";
  teams.D.name = d || "Tim D";
}
function bindTeamNameInputs(){
  const map = [
    { id: "teamA", key: "A" },
    { id: "teamB", key: "B" },
    { id: "teamC", key: "C" },
    { id: "teamD", key: "D" },
  ];
  map.forEach(({id, key}) => {
    const el = document.getElementById(id);
    if(!el) return;
    const v = (el.value || "").trim();
    teams[key].name = v ? v : `Tim ${key}`;
    el.addEventListener("input", () => {
      const nv = (el.value || "").trim();
      teams[key].name = nv ? nv : `Tim ${key}`;
      updateScores(); saveState();
    });
  });
  updateScores();
}

/* ---------- Popup ---------- */
let __popupHideTimer=null;
function showPopup(msg, opts = {}) {
  const el = document.getElementById("popup"); 
  if (!el) return;

  const center = !!opts.center;
  const sticky = !!opts.sticky;

  el.classList.remove('hidden','center-in-box','sticky');

  if (center) el.classList.add('center-in-box');
  else if (sticky) el.classList.add('sticky');

  if (__popupHideTimer) { clearTimeout(__popupHideTimer); __popupHideTimer = null; }
  el.innerText = msg || '';
  el.style.display = "block";

  if (!sticky) {
    __popupHideTimer = setTimeout(() => {
      el.style.display = "none";
      el.classList.remove('center-in-box');
      __popupHideTimer = null;
    }, 1500);
  }
}

/* ---------- Await operator click ---------- */
let __awaitClose = null;
function awaitOperatorClickTo(onContinue){
  if (__awaitClose) return;
  __awaitClose = true;

  const proceed = () => {
    document.removeEventListener('pointerdown', pointerHandler, true);
    document.removeEventListener('keydown', keyHandler, true);
    __awaitClose = null;
    try { onContinue && onContinue(); } catch {}
  };

  const pointerHandler = () => proceed();
  const keyHandler = (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); proceed(); } };

  document.addEventListener('pointerdown', pointerHandler, true);
  document.addEventListener('keydown', keyHandler, true);
}

/* ---------- Timer Dock ---------- */
function ensureTimerDock(){
  if (document.getElementById('timerDock')) return;
  const dock = document.createElement('div'); dock.id = 'timerDock';

  const cardL = document.createElement('div'); cardL.id='timerAnswerOrTO'; cardL.className='timerCard mode-answer';
  cardL.innerHTML = '<div class="timerTitle">Answer 5s</div><div class="timerNumber">00</div>';

  const cardR = document.createElement('div'); cardR.id='timerSteal'; cardR.className='timerCard mode-steal hidden';
  cardR.innerHTML = '<div class="timerTitle">Steal 10s</div><div class="timerNumber">00</div>';

  dock.appendChild(cardL); dock.appendChild(cardR);
  document.body.appendChild(dock);
}
ensureTimerDock();

/* Engine */
let __timer = {
  running:false,
  type:null,          // 'answer' | 'steal' | 'to'
  remaining:0,
  iv:null,
  targetCard:null,
  onDone:null
};

function timerSetUI(card, {type, seconds}){
  if(!card) return;
  const title = card.querySelector('.timerTitle');
  const num   = card.querySelector('.timerNumber');

  card.classList.remove('mode-answer','mode-steal','mode-to','muted','hidden');
  if(type==='answer'){ card.classList.add('mode-answer'); title.innerText='Answer 5s'; }
  else if(type==='steal'){ card.classList.add('mode-steal'); title.innerText='Steal 10s'; }
  else if(type==='to'){ card.classList.add('mode-to'); title.innerText='TO 5s'; }
  else { title.innerText='Timer'; }

  card.classList.toggle('muted', seconds > 2);
  num.innerText = String(seconds).padStart(2,'0');
}

function timerPrepare(type, seconds, onDone){
  const cardL = document.getElementById('timerAnswerOrTO');
  const cardR = document.getElementById('timerSteal');

  let card = null;
  if(type==='answer' || type==='to'){ card = cardL; }
  if(type==='steal'){ card = cardR; }

  if(cardL) cardL.classList.toggle('hidden', !(type==='answer' || type==='to'));
  if(cardR) cardR.classList.toggle('hidden', (type!=='steal'));

  __timer.running=false;
  __timer.type=type;
  __timer.remaining=seconds;
  __timer.targetCard=card;
  __timer.onDone = (typeof onDone==='function')?onDone:null;

  if(__timer.iv){ clearInterval(__timer.iv); __timer.iv=null; }
  timerSetUI(card, {type, seconds});

  if(type==='steal'){
    Promise.resolve().then(()=>{ timerStart(); });
  }
}

function timerStart(){
  if(!__timer.type || !__timer.targetCard) return;
  if(__timer.running){ timerReset(); return; }
  __timer.running = true;

  let lastTick = Date.now();
  __timer.iv = setInterval(()=>{
    const now = Date.now();
    if(now - lastTick >= 1000){
      lastTick = now;
      __timer.remaining = Math.max(0, __timer.remaining - 1);
      timerSetUI(__timer.targetCard, {type:__timer.type, seconds:__timer.remaining});
      if(__timer.remaining === 2 || __timer.remaining === 1){ playTimerTick(); }
      if(__timer.remaining === 0){
        playTimerEnd();
        timerStopCore();
        try{ __timer.onDone && __timer.onDone(); }catch{}
      }
    }
  }, 200);
}
function timerStopCore(){
  __timer.running=false;
  if(__timer.iv){ clearInterval(__timer.iv); __timer.iv=null; }
}
function timerReset(){
  timerStopCore();
  if(__timer.targetCard){ timerSetUI(__timer.targetCard, {type:__timer.type, seconds:0}); }
}
function bindTimerCardInteractions(){
  ['timerAnswerOrTO','timerSteal'].forEach(id=>{
    const card = document.getElementById(id); if(!card) return;
    let clicks=0, last=0;
    card.addEventListener('click', ()=>{
      const now = Date.now();
      clicks = (now-last < 280) ? clicks+1 : 1; last = now;
      if(clicks>=2){ card.classList.toggle('circle'); clicks=0; return; }
      const desiredType = (id==='timerAnswerOrTO') ? (__timer.type==='to' ? 'to' : 'answer') : 'steal';
      if(__timer.type === desiredType){ timerStart(); }
    });
  });

  // Auto behavior:
  document.body.addEventListener('click', (ev)=>{
    const t = ev.target;
    const id = t && t.id || '';
    if(!/^score[ABCD]$/.test(id)) return;

    setTimeout(()=>{
      if(mode==='takeover' && __timer.type==='to'){
        if(!__timer.running && __timer.targetCard){
          timerStart();
        }
        return;
      }

      if(__timer.type!=='steal' || mode!=='regular' || currentCell===null) return;
      const cs = (boardCells && boardCells[currentCell]) || null;
      if(!cs || cs.type!=='regular' || cs.dead || cs.attempts!==1) return;

      timerStopCore();
      timerPrepare('answer', 5, ()=>{
        try{ playSfx('wrong'); }catch{}
        try{ wrongAnswerRegular(activeTeam, __currentAnsKey); }catch{}
      });
      timerStart();
    }, 0);
  }, false);

  const qb = document.getElementById('questionBox');
  if(qb){
    const obs = new MutationObserver(()=>{
      const hidden = qb.classList.contains('hidden') || getComputedStyle(qb).display==='none';
      if(hidden){ timerReset(); }
    });
    obs.observe(qb, { attributes:true, attributeFilter:['class','style'] });
  }
}
bindTimerCardInteractions();

/* ---------- Start REGULAR ---------- */
function startGame(){
  applyTeamNamesFromInputs();
  bindTeamNameInputs();

  hideLegacyTOButtons();

  const sid=document.getElementById("sessionId")?.value?.trim();
  if(!sid){ alert("Mohon isi Session ID (sebagai Event ID)."); return; }
  eventId=sid;
  
   try{

  validateQuestionBank();

  for(let t in teams){
    teams[t].score = 0;
  }

  mode = "regular";
  bingoHappened = false;
  currentCell = null;
  reassignMode = false;
  bingoBonus = {
    diagonal: false,
    horizontal: false,
    vertical: false
  };

  regIndex = 0;
  boardLocked = false;

  activeTeam = "A";

  roundNo = 1;

  generateBoard();

  saveState();

}catch(e){

  alert(e.message);
  return;

}
   
  timerReset();
  updateScores(); renderBoard();
  document.getElementById("hostControls").classList.add("hidden");
  document.getElementById("questionBox").classList.add("hidden");
  document.getElementById("takeoverBox").classList.toggle("hidden", mode!=="takeover");
  if(mode==="takeover") renderTakeover();
}

/* ---------- Board ---------- */
function generateBoard(){
  boardCells=Array.from({length:25},(_,idx)=>({
    idx,type:'open',questionId:undefined,text:"",
    answered:false,attempts:0,dead:false,team:"",points:0,
    bingoWin:false
  }));
  takeoverQueue=[]; takeoverIndex=0;
  window.boardCells = boardCells; // debug
}
function assignedRegularCount(){ return (boardCells||[]).filter(c=>c.type==='regular').length; }
function renderBoard(){
  const board=document.getElementById("gameBoard");
  board.classList.toggle('board-locked', !!boardLocked && mode==='regular');
  board.classList.remove('board-bingo');
  board.innerHTML="";

  for(let idx=0; idx<25; idx++){
    const cs=boardCells[idx], cell=document.createElement("div");
    cell.classList.add("cell"); cell.innerText=idx+1;
    if(boardLocked && mode==='regular' && idx===currentCell) cell.classList.add("active");
    if(cs.dead) cell.classList.add("dead");
    if(cs.team) cell.classList.add("team"+cs.team);
    if(cs.bingoWin) cell.classList.add("bingo-win");
    if(mode==='takeover' && cs.type==='takeover'){
      cell.classList.add("takeoverSlot");
      if(takeoverQueue[takeoverIndex]===idx) cell.classList.add("takeoverActive");
    }

    // Bingo inline color (kuat)
    if (cs.bingoWin && cs.team && TEAM_COLOR_MAP[cs.team]) {
      const col = TEAM_COLOR_MAP[cs.team];
      cell.style.setProperty('background-color', col.bg, 'important');
      cell.style.setProperty('color', col.fg, 'important');
      cell.style.opacity = '1';
      cell.style.borderColor = '#16a085';
    }

    cell.dataset.index=String(idx);

    cell.addEventListener("click", ()=>{
      if(reassignMode){ handleReassignDOM(cell); return; }
      
      if(mode === "takeover"){
        if(cs.type !== "takeover") return;
        if(cs.dead || cs.answered) return;
        takeoverIndex = takeoverQueue.indexOf(idx);
        renderTakeover();
        return;
      }
      
      if(mode!=="regular") return;
      if(cs.dead || cs.answered) return;
      if(boardLocked && idx!==currentCell) return;

      if(cs.type==='open'){
        const already=assignedRegularCount();
        const quota=(regularQuestions?.length||20);
        if(already>=quota){ maybeStartTakeover(); return; }
        const q=regularQuestions[regIndex]; if(!q){ alert("Bank REGULAR habis."); return; }
        cs.type='regular'; cs.questionId=q.id; cs.text=q.text; regIndex+=1; saveState();
        showQuestionIndex(idx); return;
      }
      if(cs.type==='regular'){ showQuestionIndex(idx); }
    });

    board.appendChild(cell);
  }
}

/* ---------- Parser MCQ (FINAL) ---------- */

/* ---- HTML Normalizer ----
   - Ubah <br>, &lt;br&gt;, &amp;lt;br&amp;gt; -> '\n'
   - Decode entity sampai 3x (antisipasi double/triple encoded)
   - Hapus SEMUA tag: asli <...> dan encoded
*/
function htmlize(text){
  let s = (text || '').toString().replace(/\r/g, '');

  // 1) Normalisasi SEMUA variasi <br> ke newline
  s = s
    .replace(/<br\s*\/?>/gi, '\n')                   // <br>
    .replace(/&lt;br\s*\/?&gt;/gi, '\n')               // &lt;br&gt;
    .replace(/&(amp;)+lt;br\s*\/?&(amp;)+gt/gi, '\n'); // &amp;lt;br&amp;gt;

  // 2) Decode entity berulang (handle double/triple encoding)
  for (let i = 0; i < 3; i++) {
    const ta = document.createElement('textarea');
    ta.innerHTML = s;
    s = ta.value;
  }

  // 3) Bersihkan tag HTML asli & encoded
  for (let i = 0; i < 2; i++) {
    s = s
      .replace(/<[^>]+>/g, '')                        // <b>...</b> dkk (ASLI)  ⬅ penting!
      .replace(/&lt;\/?[^&gt;]+&gt;/gi, '')              // &lt;b&gt;...&lt;/b&gt;
      .replace(/&(amp;)+lt;\/?[^&]+&(amp;)+gt/gi,''); // &amp;lt;b&amp;gt;...
  }

  // 4) Rapikan whitespace
  return s
    .replace(/\u00A0/g,' ')
    .replace(/\u200B/g,'')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/* ---- MCQ Extractor ----
   - Ambil prompt & opsi A–D (A. / A) / A: / A - )
   - Fallback inline jika format baris tidak rapi
*/
function extractPromptAndOptions(htmlText){
  const raw = htmlize(htmlText);

  const lines = raw.split(/\n+/)
    .map(t => t.trim())
    .filter(Boolean);

  const opts = {};
  const optPattern = /^([A-Da-d])\s*[.\):\-]?\s*(.+)$/;

  let promptLines = [];
  let enteringOptions = false;

  for (const line of lines) {
    const m = line.match(optPattern);
    if (m) {
      enteringOptions = true;
      const key = m[1].toLowerCase();
      const val = (m[2] || '').trim();
      if (val) opts[key] = val;
    } else if (!enteringOptions) {
      promptLines.push(line);
    }
  }

  // Fallback inline bila opsi belum terbaca minimal 2
  if (Object.keys(opts).length < 2) {
    const inline = raw.replace(/\n+/g, ' ');
    const inlineRe =
      /([A-Da-d])\s*[.\):\-]\s*([^A-Da-d]+?)(?=(?:\s+[A-Da-d]\s*[.\):\-]\s*)|$)/g;

    let im;
    while ((im = inlineRe.exec(inline)) !== null) {
      const key = im[1].toLowerCase();
      const val = (im[2] || '').trim();
      if (!opts[key] && val) opts[key] = val;
    }

    // Prompt fallback: ambil teks sebelum 'A.'
    if (promptLines.length === 0) {
      const fa = inline.search(/[A-Da-d]\s*[.\):\-]\s*/);
      if (fa > 0) promptLines = [ inline.slice(0, fa).trim() ];
    }
  }

  return {
    prompt: (promptLines.join('\n').trim() || ''),
    options: opts
  };
}

/* ---- Prompt Fallback ----
   - Jika tetap gagal, minimal judul prompt tetap terisi
*/
function derivePromptFromText(questionText){
  const h = htmlize(questionText);
  const split = h.split(/\n\s*[A-Da-d]\s*[.\):\-]\s*/);
  return (split[0] || '').trim();
}

/* ---------- REGULAR tampil & nilai ---------- */
let __currentAnsKey = null;

function getQuestionById(id){
  const qReg=(regularQuestions||[]).find(q=>q.id===id);
  const qTO =(takeoverQuestions||[]).find(q=>q.id===id);
  return qReg||qTO||null;
}
function showQuestionIndex(idx){
  const cs=boardCells[idx];
  if(!cs || cs.type!=="regular" || cs.dead || cs.answered) return;
  currentCell=idx;

  const q=getQuestionById(cs.questionId);
  if(!q){
    showPopup("Soal tidak ditemukan!");
    cs.type='open'; cs.questionId=undefined; cs.text="";
    if (regIndex>0) regIndex -= 1;
    boardLocked=false; currentCell=null;
    saveState(); renderBoard();
    return;
  }

  const ansKey=(q&&q.ans)?String(q.ans).toLowerCase():null;
  __currentAnsKey = ansKey;

  const questionText = cs.text || q.text || "";
  const {prompt, options}=extractPromptAndOptions(questionText);

  const qt=document.getElementById("questionText");
  const answersArea=document.getElementById("answersArea");
  answersArea.innerHTML="";

  let finalPrompt = (prompt && prompt.trim()) ? prompt : derivePromptFromText(questionText);
  finalPrompt = finalPrompt || "Pertanyaan";

  const keys=['a','b','c','d'].filter(k => options[k]);
  if(keys.length === 0){
    showPopup("Pilihan ganda tidak ditemukan. Soal dilewati.");
    cs.type='open'; cs.questionId=undefined; cs.text="";
    if (regIndex>0) regIndex -= 1;
    boardLocked=false; currentCell=null;
    saveState(); renderBoard();
    return;
  }

  qt.innerText = finalPrompt;
  keys.forEach(k=>{
    const btn=document.createElement('button');
    btn.className='answer-btn';
    btn.innerText=`${k.toUpperCase()}. ${options[k]}`;
    btn.onclick=()=>handleAnswerChoiceRegular(k, ansKey);
    answersArea.appendChild(btn);
  });

  document.getElementById("questionBox").classList.remove("hidden");
  boardLocked=true; renderBoard();

  // Answer timer 5s REGULAR
  timerPrepare('answer', 5, ()=>{
    if(!activeTeam){ showPopup("Pilih tim aktif dulu."); return; }
    playSfx('wrong');
    wrongAnswerRegular(activeTeam, __currentAnsKey);
  });
}

function handleAnswerChoiceRegular(chosen, ansKey){
  timerStopCore();

  if(!activeTeam){ showPopup("Pilih pemain aktif dulu.", {center:true}); return; }
  if(!ansKey){ showPopup("Soal belum punya kunci jawaban.", {center:true}); return; }

  if(chosen===ansKey){
    playSfx('correct');
    markChosen(chosen, true, "answersArea");
    answerRegular(activeTeam);
  }else{
    playSfx('wrong');
    markChosen(chosen, false, "answersArea");
    wrongAnswerRegular(activeTeam, ansKey);
  }
}
function markChosen(chosen, isCorrect, containerId){
  const area=document.getElementById(containerId);
  const btn=[...area.querySelectorAll('.answer-btn')]
    .find(b=>b.innerText.trim().toLowerCase().startsWith(chosen.toLowerCase()+'.'));
  if(btn) btn.classList.add(isCorrect?'is-correct':'is-wrong');
}
function highlightCorrectButton(containerId, ansKey){
  const area=document.getElementById(containerId);
  if(!area || !ansKey) return;
  const btn=[...area.querySelectorAll('.answer-btn')]
    .find(b=>b.innerText.trim().toLowerCase().startsWith(ansKey.toLowerCase()+'.'));
  if(btn) btn.classList.add('is-correct');
}
function answerRegular(teamKey){
  timerStopCore();

  if(currentCell===null) return;
  const cs=boardCells[currentCell];
  if(!cs || cs.type!=="regular" || cs.dead) return;

  const points=(cs.attempts===0)?10:5;
  cs.team=teamKey; cs.points=points; cs.answered=true;
  teams[teamKey].score+=points;

  updateScores(); saveState(); renderBoard();
  showPopup(`${teams[teamKey].name} +${points} poin`, {center:true});
  checkBingo(teamKey);

  awaitOperatorClickTo(()=>{
    boardLocked=false; 
    currentCell=null; 
    document.getElementById("questionBox").classList.add("hidden");
    renderBoard();
    maybeStartTakeover();
  });
}
function wrongAnswerRegular(teamKey, ansKey){
  if(currentCell===null) return;
  const cs=boardCells[currentCell];
  if(!cs || cs.type!=="regular" || cs.dead) return;

  const penalty=(cs.attempts===0)?-5:-2;
  teams[teamKey].score+=penalty;

  cs.attempts+=1;
  const nowDead = (cs.attempts>=2);
  if(nowDead){ cs.dead=true; }

  updateScores(); saveState(); renderBoard();
  showPopup(`${teams[teamKey].name} ${penalty} poin`, {center:true});

  if(nowDead){
    highlightCorrectButton("answersArea", ansKey);

    awaitOperatorClickTo(()=>{
      boardLocked=false; 
      currentCell=null; 
      document.getElementById("questionBox").classList.add("hidden");
      renderBoard();
      maybeStartTakeover();
    });
    return;
  }

  timerPrepare('steal', 10, ()=>{
    const cs2=boardCells[currentCell];
    if(cs2){ cs2.dead=true; }
    highlightCorrectButton("answersArea", ansKey);
    showPopup("⏳ Waktu rebut habis — soal hangus.");
    awaitOperatorClickTo(()=>{
      boardLocked=false; currentCell=null; document.getElementById("questionBox").classList.add("hidden");
      renderBoard(); maybeStartTakeover();
    });
  });
}

/* ---------- Reassign ---------- */
function reassignAnswerPrompt(){
  reassignMode=true; alert("🔁 Klik pertanyaan REGULAR yang sudah dijawab untuk direassign.");
}
function handleReassignDOM(cellDiv){
  if(!reassignMode) return;
  const idx=parseInt(cellDiv.dataset.index,10);
  const cs=boardCells[idx];
  if(!cs || cs.type!=='regular' || !cs.answered || cs.dead){
    alert("Hanya REGULAR yang sudah dijawab & belum hangus yang bisa direassign.");
    reassignMode=false; return;
  }
  const prev=cs.team, pts=cs.points||0;
  const newTeam=prompt("Masukkan huruf tim (A/B/C/D):")?.toUpperCase();
  if(!teams[newTeam]){ alert("❌ Tim tidak valid."); reassignMode=false; return; }
  if(prev && prev!==newTeam){
    teams[prev].score-=pts; teams[newTeam].score+=pts; cs.team=newTeam;
    updateScores(); saveState(); renderBoard();
    showPopup(`Reassign: ${pts} poin pindah ke ${teams[newTeam].name}`, {center:true});
    checkBingo(newTeam);
  }
  reassignMode=false;
}

/* ---------- Bingo ---------- */
function checkBingo(teamKey){

  const hasTeam = idx =>
    boardCells[idx].team === teamKey;

  const horizontalLines = Array.from(
    {length:5},
    (_,r)=>boardCells.slice(r*5,(r+1)*5).map(c=>c.idx)
  );

  const verticalLines = Array.from(
    {length:5},
    (_,c)=>[0,1,2,3,4].map(r=>r*5+c)
  );
  
  const diagonal1 = [0,6,12,18,24];
  const diagonal2 = [4,8,12,16,20];

  let awarded = false;

  // ===== HORIZONTAL =====
  const hasHorizontal =
    horizontalLines.some(line =>
      line.every(hasTeam)
    );

  if(hasHorizontal && !bingoBonus.horizontal){

    bingoBonus.horizontal = true;

    teams[teamKey].score += 50;

    horizontalLines.forEach(line=>{
      if(line.every(hasTeam)){
        line.forEach(i=>{
          boardCells[i].bingoWin = true;
        });
      }
    });

    awarded = true;
  }

  // ===== VERTICAL =====
  const hasVertical =
    verticalLines.some(line =>
      line.every(hasTeam)
    );

  if(hasVertical && !bingoBonus.vertical){

    bingoBonus.vertical = true;

    teams[teamKey].score += 50;

    verticalLines.forEach(line=>{
      if(line.every(hasTeam)){
        line.forEach(i=>{
          boardCells[i].bingoWin = true;
        });
      }
    });

    awarded = true;
  }

  // ===== DIAGONAL 1 SAJA =====
  const hasDiagonal =
    diagonal1.every(hasTeam) ||
    diagonal2.every(hasTeam);

  if(hasDiagonal && !bingoBonus.diagonal){
  
      bingoBonus.diagonal = true;
  
      teams[teamKey].score += 50;
  
      if(diagonal1.every(hasTeam)){
          diagonal1.forEach(i=>{
              boardCells[i].bingoWin = true;
          });
      }
  
      if(diagonal2.every(hasTeam)){
          diagonal2.forEach(i=>{
              boardCells[i].bingoWin = true;
          });
      }
  
      awarded = true;
  }

  if(!awarded) return;
  
  updateScores();
  saveState();

  const winnerName =
    teams[teamKey]?.name ||
    `Tim ${teamKey}`;

  showPopup(
    `🎉 BINGO! ${winnerName} mendapat BONUS +50 poin!`,
    {center:true}
  );

  renderBoard();
}

/* ---------- Regular → Takeover ---------- */
function maybeStartTakeover(){
  if(bingoHappened||mode!=="regular") return;
  const already=(boardCells||[]).filter(c=>c.type==='regular').length;
  const quota=(regularQuestions?.length||20);
  const openIdx=boardCells.filter(c=>c.type==='open').map(c=>c.idx);
  const shouldSwitch=(already>=quota) && (openIdx.length===5);
  if(!shouldSwitch) return;

  openIdx.forEach((i, j)=>{
    const cs=boardCells[i];
    cs.type='takeover';
    const tq = takeoverQuestions[j];
    if (tq){
      cs.questionId = tq.id;
      cs.text = tq.text || "";
    } else {
      cs.questionId = undefined;
      cs.text = "";
    }
  });

  takeoverQueue=openIdx.slice(0, Math.min(openIdx.length, takeoverQuestions.length||openIdx.length));
  takeoverIndex=0;
  takeoverAnsweredCount=0;
  mode="takeover"; 
  timerReset();
  saveState();

  document.getElementById("questionBox").classList.add("hidden");
  document.getElementById("takeoverBox").classList.remove("hidden");

  hideLegacyTOButtons();

  renderBoard();
}

/* ---------- Takeover (REGULAR & FINAL) ---------- */
function skipTakeover(reason){
  timerStopCore();

  const cellIdx=takeoverQueue[takeoverIndex];
  const cs=boardCells[cellIdx];
  if (cs){ cs.answered=false; cs.dead=true; cs.team=""; cs.points=0; }
  showPopup(reason||"Soal TO dilewati.", {center:true});
  saveState(); renderBoard();

  awaitOperatorClickTo(()=>{
    document.getElementById("takeoverBox")?.classList.add("hidden");
    takeoverAnsweredCount+=1;
    if(takeoverAnsweredCount >= takeoverQueue.length){
      endRoundNoBingo();
      return;
    }
    renderBoard();
    saveState();
  });
}
function renderTakeover(){
  if(mode!=="takeover") return;
  document.getElementById("takeoverBox")?.classList.remove("hidden");
  const total=takeoverQueue.length;
  if (takeoverAnsweredCount >= total){
    endRoundNoBingo();
    return;
  }
  const i=takeoverIndex;

  const cellIdx=takeoverQueue[i];
  const cs=boardCells[cellIdx];

  let q = (cs && cs.questionId) ? getQuestionById(cs.questionId) : null;
  if(!q){ skipTakeover("Soal TO tidak ditemukan."); return; }

  const ansKey=(q&&q.ans)?String(q.ans).toLowerCase():null;
  const qText = (cs && cs.text) ? cs.text : (q.text||"");
  const {prompt,options}=extractPromptAndOptions(qText);

  let finalPrompt = (prompt && prompt.trim()) ? prompt : derivePromptFromText(qText);
  if(!finalPrompt && Object.keys(options).length===0){ skipTakeover("Soal TO kosong."); return; }

  const ttl = document.getElementById('takeoverTitle');
  const prog = document.getElementById('takeoverProgress');
  if (ttl) ttl.innerText = 'Take Over Round';
  if (prog) prog.innerText = `Take Over: ${takeoverAnsweredCount+1}/${total}`;

  document.getElementById("takeoverText").innerText=finalPrompt || `Take Over: #${takeoverAnsweredCount+1}`;

  const answersArea=document.getElementById("toAnswersArea");
  answersArea.innerHTML="";
  const keys=['a','b','c','d'].filter(k=>options[k]);
  if(keys.length===0){ skipTakeover("Pilihan TO tidak ditemukan."); return; }

  hideLegacyTOButtons();

  keys.forEach(k=>{
    const btn=document.createElement('button');
    btn.className='answer-btn';
    btn.innerText=`${k.toUpperCase()}. ${options[k]}`;
    btn.onclick=()=>{
      timerStopCore();

      if(!activeTeam){ showPopup("Pilih pemain aktif dulu.", {center:true}); return; }
      if(!ansKey){ showPopup("Soal TO belum punya kunci.", {center:true}); return; }
      if(k===ansKey){
        playSfx('correct');
        markChosen(k, true, "toAnswersArea");
        answerTakeover(activeTeam);
      }else{
        playSfx('wrong');
        markChosen(k, false, "toAnswersArea");
        highlightCorrectButton("toAnswersArea", ansKey);
        wrongTakeover(activeTeam, ansKey);
      }
    };
    answersArea.appendChild(btn);
  });

  // Timer TO 5s
  timerPrepare('to', 5, ()=>{
    if(!activeTeam){ showPopup("Pilih pemain aktif dulu."); return; }
    const pen=-10; teams[activeTeam].score+=pen;

    const cellIdx2=takeoverQueue[takeoverIndex];
    const cs2=boardCells[cellIdx2];
    if(cs2){ cs2.answered=false; cs2.dead=true; cs2.team=""; cs2.points=0; }

    saveState(); renderBoard();
    highlightCorrectButton("toAnswersArea", ansKey);
    showPopup(`${teams[activeTeam].name} ${pen} poin`, {center:true});

    awaitOperatorClickTo(()=>{
      document.getElementById("takeoverBox")?.classList.add("hidden");
      takeoverAnsweredCount+=1;
      if(takeoverAnsweredCount >= takeoverQueue.length){
      endRoundNoBingo();
      return;
    }
    renderBoard();
    saveState();
    });

  });

  renderBoard();
}

function answerTakeover(teamKey){
  timerStopCore();

  if(mode!=="takeover"||!teams[teamKey]) return;
  const pts=20;
  teams[teamKey].score+=pts;

  const cellIdx=takeoverQueue[takeoverIndex];
  const cs=boardCells[cellIdx];
  if(cs){ cs.answered=true; cs.dead=false; cs.team=teamKey; cs.points=pts; }

  updateScores(); showPopup(`${teams[teamKey].name} +${pts} poin`, {center:true});
  checkBingo(teamKey); // Final TETAP menilai bingo sesuai engine
  saveState(); renderBoard();

  awaitOperatorClickTo(()=>{
    document.getElementById("takeoverBox")?.classList.add("hidden");
    takeoverAnsweredCount+=1;
    if(takeoverAnsweredCount >= takeoverQueue.length){
      endRoundNoBingo();
      return;
    }
    renderBoard();
    saveState();
  });
}
  
function wrongTakeover(teamKey, ansKey){
  timerStopCore();

  if(mode!=="takeover"||!teams[teamKey]) return;
  const pen=-10; teams[teamKey].score+=pen;

  const cellIdx=takeoverQueue[takeoverIndex];
  const cs=boardCells[cellIdx];
  if(cs){ cs.answered=false; cs.dead=true; cs.team=""; cs.points=0; }

  updateScores(); showPopup(`${teams[teamKey].name} ${pen} poin`, {center:true});
  saveState(); renderBoard();

  awaitOperatorClickTo(()=>{
    document.getElementById("takeoverBox")?.classList.add("hidden");
    takeoverAnsweredCount+=1;
    if(takeoverAnsweredCount >= takeoverQueue.length){
      endRoundNoBingo();
      return;
    }
    renderBoard();
    saveState();
  });
}

/* ---------- End round ---------- */
function endRoundNoBingo(){
  mode="ended"; 
  timerReset();
  saveState(); renderBoard();
  document.getElementById("takeoverBox").classList.add("hidden");

  const standings=[
    {k:'A',name:teams.A.name,score:teams.A.score},
    {k:'B',name:teams.B.name,score:teams.B.score},
    {k:'C',name:teams.C.name,score:teams.C.score},
    {k:'D',name:teams.D.name,score:teams.D.score}
  ].sort((a,b)=>b.score-a.score);

  showPopup(`🏁 Selesai! Peringkat 1: ${standings[0].name} skor ${standings[0].score}`, {sticky:true});
}

/* ---------- Scoreboard ---------- */
function setActiveTeam(k){ if(!teams[k]) return; activeTeam=k; updateScores(); saveState(); }
function updateScores(){
  const a=document.getElementById("scoreA"), b=document.getElementById("scoreB"), c=document.getElementById("scoreC"), d=document.getElementById("scoreD");
  if(!a||!b||!c||!d) return;
  a.innerText=`Tim A (${teams.A.name}): ${teams.A.score}`;
  b.innerText=`Tim B (${teams.B.name}): ${teams.B.score}`;
  c.innerText=`Tim C (${teams.C.name}): ${teams.C.score}`;
  d.innerText=`Tim D (${teams.D.name}): ${teams.D.score}`;
  a.classList.add('teamA'); b.classList.add('teamB'); c.classList.add('teamC'); d.classList.add('teamD');
  [['A',a],['B',b],['C',c],['D',d]].forEach(([k,el])=>{ el.classList.toggle('active',activeTeam===k); el.onclick=()=>setActiveTeam(k); el.title="Klik untuk pilih pemain aktif"; });
}

/* ---------- Legacy TO buttons helper ---------- */
function hideLegacyTOButtons(){
  const m = document.getElementById('legacyTOButtons');
  if (m) m.style.display = 'none';
}

/* ---------- SFX ---------- */
let __audioCtx=null;
function _ensureAudioCtx(){ 
  if(!__audioCtx){ 
    const Ctx=window.AudioContext||window.webkitAudioContext; 
    if(!Ctx) return null; 
    __audioCtx=new Ctx(); 
  } 
  if(__audioCtx.state==='suspended'){ __audioCtx.resume().catch(()=>{});} 
  return __audioCtx; 
}
function playSfx(kind){
  const ctx=_ensureAudioCtx(); if(!ctx) return;
  if(kind==='correct'){
    const o=ctx.createOscillator(), g=ctx.createGain();
    o.type='triangle'; o.frequency.setValueAtTime(660,ctx.currentTime); o.frequency.linearRampToValueAtTime(990,ctx.currentTime+0.22);
    g.gain.setValueAtTime(0.0001,ctx.currentTime); g.gain.exponentialRampToValueAtTime(0.25,ctx.currentTime+0.01); g.gain.exponentialRampToValueAtTime(0.0001,ctx.currentTime+0.25);
    o.connect(g).connect(ctx.destination); o.start(); o.stop(ctx.currentTime+0.26);
  }else if(kind==='wrong'){
    const o=ctx.createOscillator(), g=ctx.createGain();
    o.type='square'; o.frequency.setValueAtTime(120,ctx.currentTime);
    g.gain.setValueAtTime(0.25,ctx.currentTime); g.gain.exponentialRampToValueAtTime(0.001,ctx.currentTime+0.35);
    o.connect(g).connect(ctx.destination); o.start(); o.stop(ctx.currentTime+0.36);
  }
}
function playTimerTick(){
  const ctx=_ensureAudioCtx(); if(!ctx) return;
  const o=ctx.createOscillator(), g=ctx.createGain();
  o.type='sine'; o.frequency.setValueAtTime(880, ctx.currentTime);
  g.gain.setValueAtTime(0.0001, ctx.currentTime);
  g.gain.exponentialRampToValueAtTime(0.18, ctx.currentTime+0.01);
  g.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime+0.15);
  o.connect(g).connect(ctx.destination); o.start(); o.stop(ctx.currentTime+0.18);
}
function playTimerEnd(){
  const ctx=_ensureAudioCtx(); if(!ctx) return;
  const o=ctx.createOscillator(), g=ctx.createGain();
  o.type='sawtooth'; o.frequency.setValueAtTime(220, ctx.currentTime);
  g.gain.setValueAtTime(0.0001, ctx.currentTime);
  g.gain.exponentialRampToValueAtTime(0.28, ctx.currentTime+0.01);
  g.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime+0.35);
  o.connect(g).connect(ctx.destination); o.start(); o.stop(ctx.currentTime+0.36);
}

/* ---------- Timer Dock positioning ---------- */
let __timerDockBaseline = null;
function __isHidden(el){ return !el || el.classList.contains('hidden') || getComputedStyle(el).display === 'none'; }
function __computeTimerScaleFrom(cellWidth){
  const base = 80;
  let f = (cellWidth || base) / base;
  return Math.max(0.85, Math.min(1.35, f));
}
function __applyTimerScaleWith(cellWidth){
  const f = __computeTimerScaleFrom(cellWidth);
  const root = document.documentElement;
  const remPx = parseFloat(getComputedStyle(root).fontSize) || 16;

  const baseW = 130, baseH = 92, gap = 12;
  const baseNumRem = 1.8, baseTitleRem = 0.9;

  root.style.setProperty('--timer-card-minw', `${Math.round(baseW * f)}px`);
  root.style.setProperty('--timer-card-minh', `${Math.round(baseH * f)}px`);
  root.style.setProperty('--timer-gap', `${Math.round(gap * f)}px`);
  root.style.setProperty('--timer-number-size', `${(baseNumRem * remPx * f).toFixed(0)}px`);
  root.style.setProperty('--timer-title-size', `${(baseTitleRem * remPx * f).toFixed(0)}px`);
}
function lockTimerDockBaseline(){
  const board = document.getElementById('gameBoard');
  const qb = document.getElementById('questionBox');
  const tob = document.getElementById('takeoverBox');
  if(!board) return;
  if (!__isHidden(qb) || !__isHidden(tob)) return;

  const r = board.getBoundingClientRect();
  const firstCell = document.querySelector('#gameBoard .cell');
  const cw = firstCell ? firstCell.getBoundingClientRect().width : 80;

  __timerDockBaseline = { top: r.top, right: r.right, cellWidth: cw };
}
function positionTimerDock(){
  const dock  = document.getElementById('timerDock');
  const board = document.getElementById('gameBoard');
  if(!dock || !board) return;

  if(!__timerDockBaseline) lockTimerDockBaseline();
  if(!__timerDockBaseline) return;

  const marginRight = 16;

  __applyTimerScaleWith(__timerDockBaseline.cellWidth);

  dock.classList.remove('stack-vert');
  dock.style.visibility = 'hidden';
  dock.style.left = '-9999px';
  dock.style.top  = '-9999px';

  requestAnimationFrame(()=>{
    const dRect2col = dock.getBoundingClientRect();
    const freeRight = window.innerWidth - (__timerDockBaseline.right + marginRight) - 8;
    if (dRect2col.width > freeRight) dock.classList.add('stack-vert');

    const dRect = dock.getBoundingClientRect();
    dock.style.visibility = 'visible';

    let left = __timerDockBaseline.right + marginRight;
    let top  = __timerDockBaseline.top + (board.getBoundingClientRect().height/2) - (dRect.height/2);

    left = Math.min(left, window.innerWidth - dRect.width - 8);
    top  = Math.max(8, Math.min(top, window.innerHeight - dRect.height - 8));

    dock.style.left = `${left}px`;
    dock.style.top  = `${top}px`;
  });
}
window.addEventListener('scroll', positionTimerDock, { passive:true });
window.addEventListener('resize', ()=>{
  __timerDockBaseline = null;
  requestAnimationFrame(()=>{ lockTimerDockBaseline(); positionTimerDock(); });
}, { passive:true });

const __rb_base = renderBoard;
renderBoard = function(){
  __rb_base.apply(this, arguments);
  requestAnimationFrame(()=>{
    lockTimerDockBaseline();
    positionTimerDock();
  });
};

const __rto_base = renderTakeover;
renderTakeover = function(){
  __rto_base.apply(this, arguments);
  requestAnimationFrame(positionTimerDock);
};

const __sqi_base = showQuestionIndex;
showQuestionIndex = function(){
  __sqi_base.apply(this, arguments);
  requestAnimationFrame(positionTimerDock);
};

const __start_base = startGame;
startGame = function(){
  __start_base.apply(this, arguments);
  requestAnimationFrame(()=>{
    __timerDockBaseline = null;
    lockTimerDockBaseline();
    positionTimerDock();
  });
};

requestAnimationFrame(()=>{ lockTimerDockBaseline(); positionTimerDock(); });

document.addEventListener('DOMContentLoaded',()=>{

  const fileInput =
    document.getElementById('excelFile');

  if(!fileInput) return;

  fileInput.addEventListener(
    'change',
    async (e)=>{

      const file = e.target.files?.[0];

      if(!file) return;

      try{

        await parseExcelFile(file);

      }catch(err){

        console.error(err);

        alert(
          'Gagal membaca file Excel.'
        );
      }
    }
  );

});

/* ---------- Expose ---------- */
window.startGame=startGame;
window.answerTakeover=answerTakeover;
window.answertakeOver=(team)=>answerTakeover(team);
window.wrongTakeover=wrongTakeover;
window.wrongtakeOver=(team)=>wrongTakeover(team);
window.reassignAnswerPrompt=reassignAnswerPrompt;

/* =========================
   === FULL ADMIN FEATURES ===
   - Mode Panitia (Ctrl+Shift+P or click .site-logo 5x)
   - Admin action on cell: Anulir, Pulihkan, Edit skor manual, Reassign (REG & TO, custom poin)
   - Auto recalculasi Bingo setiap perubahan
   - Tidak mengubah HTML/CSS
   ========================= */

/* ---- Admin: State & Hotkeys ---- */
let adminMode = false;
let __adminClicks = 0;
let __adminTimer = null;

function showAdminToast(){
  showPopup(adminMode ? '🛠️ MODE PANITIA AKTIF' : '🛠️ Mode Panitia NONAKTIF', {center:true});
}
function toggleAdminMode(force){
  adminMode = (typeof force==='boolean') ? !!force : !adminMode;
  showAdminToast();
}
window.toggleAdminMode = toggleAdminMode;

// Hotkey: Ctrl + Shift + P
(function(){
  document.addEventListener('keydown', (e)=>{
    if(e.ctrlKey && e.shiftKey && (e.key.toLowerCase()==='p')){
      e.preventDefault();
      toggleAdminMode();
    }
  }, true);
})();

// Gesture: klik logo 5x cepat untuk aktifkan admin
(function(){
  document.addEventListener('click', (e)=>{
    const tgt = e.target;
    if(!tgt) return;
    const isLogo = tgt.classList && tgt.classList.contains('site-logo');
    if(!isLogo) return;
    __adminClicks++;
    if(__adminTimer){ clearTimeout(__adminTimer); }
    __adminTimer = setTimeout(()=>{ __adminClicks=0; }, 1000);
    if(__adminClicks>=5){
      __adminClicks=0;
      toggleAdminMode(true);
    }
  }, true);
})();

/* ---- Admin: Bingo Recall ---- */
function _recalcBingoAll(){
  if(!Array.isArray(boardCells)) return;
  boardCells.forEach(c=>{ c.bingoWin = false; });
  const lines = [
    ...Array.from({length:5}, (_,r)=>boardCells.slice(r*5,(r+1)*5).map(c=>c.idx)), // rows
    ...Array.from({length:5}, (_,c)=>[0,1,2,3,4].map(r=>r*5+c)),                  // cols
    [0,6,12,18,24], [4,8,12,16,20]                                               // diags
  ];
  const teamsKeys = ['A','B','C','D'];
  for(const tk of teamsKeys){
    for(const line of lines){
      if(line.every(i => boardCells[i].team === tk)){
        line.forEach(i => boardCells[i].bingoWin = true);
      }
    }
  }
}

/* ---- Admin: Anulir / Pulihkan / Reassign ---- */
function annulCell(idx){
  const cs = boardCells[idx];
  if(!cs) { showPopup('Sel tidak ditemukan', {center:true}); return false; }
  // rollback skor bila ada
  if(cs.answered && cs.team && cs.points){
    const t = cs.team; if(teams[t]) teams[t].score -= (cs.points||0);
  }
  cs.dead = true;
  cs.answered = false;
  cs.team = '';
  cs.points = 0;
  _recalcBingoAll();
  saveState(); updateScores(); renderBoard();
  showPopup('❌ Soal dianulir', {center:true});
  return true;
}
function restoreCell(idx){
  const cs = boardCells[idx];
  if(!cs) { showPopup('Sel tidak ditemukan', {center:true}); return false; }
  cs.dead = false;
  cs.answered = false;
  cs.team = '';
  cs.points = 0;
  _recalcBingoAll();
  saveState(); updateScores(); renderBoard();
  showPopup('🔄 Soal dipulihkan', {center:true});
  return true;
}
function reassignCell(idx, newTeam, newPoints){
  const cs = boardCells[idx];
  if(!cs) { showPopup('Sel tidak ditemukan', {center:true}); return false; }
  if(!cs.answered){ showPopup('Sel belum dijawab, tidak bisa reassign.', {center:true}); return false; }
  if(!teams[newTeam]){ showPopup('Tim tidak valid.', {center:true}); return false; }
  const prevTeam = cs.team || '';
  const prevPts  = cs.points || 0;
  const useCustom = (typeof newPoints==='number' && !Number.isNaN(newPoints));
  const nextPts   = useCustom ? newPoints : prevPts;
  if(prevTeam && teams[prevTeam]) teams[prevTeam].score -= prevPts;
  cs.team = newTeam;
  cs.points = nextPts;
  teams[newTeam].score += nextPts;
  _recalcBingoAll();
  saveState(); updateScores(); renderBoard();
  showPopup(`✅ Reassign ke ${teams[newTeam].name} ${useCustom?`(${nextPts} poin)`: `(+${nextPts} poin)`}`, {center:true});
  return true;
}

// ===== Admin: Force Resume setelah BINGO =====
function adminForceResumeGame() {
  try {
    // 1) Tutup popup sticky BINGO (kalau masih ada)
    const pop = document.getElementById('popup');
    if (pop) {
      pop.style.display = 'none';
      pop.classList.remove('sticky', 'center-in-box');
      pop.innerText = '';
    }

    // 2) Bersihkan highlight bingo di semua sel
    if (Array.isArray(boardCells)) {
      boardCells.forEach(c => { c.bingoWin = false; });
    }

    // 3) Kembalikan engine dari "ended" ke "regular"
    bingoHappened = false;
    mode = 'regular';
    boardLocked = false;
    currentCell = null;

    // 4) Rapikan UI (pastikan box pertanyaan/TO tertutup)
    const qb  = document.getElementById("questionBox");
    const tob = document.getElementById("takeoverBox");
    if (qb)  qb.classList.add("hidden");
    if (tob) tob.classList.add("hidden");

    // 5) Recalculate bingo lines (hilangkan “bingo hantu”)
    if (typeof _recalcBingoAll === 'function') _recalcBingoAll();

    // 6) Simpan & render
    saveState(); 
    updateScores(); 
    renderBoard();

    // 7) Info operator
    showPopup('▶️ Lanjut main: status BINGO dibatalkan.', {center:true});
  } catch (e) {
    console.warn('adminForceResumeGame error', e);
    alert('Gagal force-resume. Lihat console untuk detail.');
  }
}

/* ---- Admin: Interceptor klik cell ---- */
(function(){
  document.addEventListener('click', (e)=>{
    if(!adminMode) return;
    const cell = e.target && (e.target.closest ? e.target.closest('.cell') : null);
    if(!cell) return;
    e.stopPropagation();
    e.preventDefault();
    const idx = parseInt(cell.dataset.index, 10);
    if(Number.isNaN(idx)) return;
    const cs = boardCells[idx];
    const meta = `#${idx+1} | tipe:${cs?.type||'-'} | dead:${cs?.dead?'Y':'N'} | team:${cs?.team||'-'} | poin:${cs?.points||0}`;
    const choice = prompt(
      `== PANEL PANITIA ==\n${meta}\n\n`+
      `Ketik:\n`+
      `1 = Anulir (abu)\n`+
      `2 = Pulihkan\n`+
      `3 = Edit skor tim (manual)\n`+
      `4 = Reassign ke tim lain (REG/TO)\n`+
      `5 = Force Resume / Batal BINGO (lanjutkan ronde)\n`+ 
      `Lainnya = batal`
    );
    if(choice==='1'){
      annulCell(idx);
    }else if(choice==='2'){
      restoreCell(idx);
    }else if(choice==='3'){
      const t = (prompt('Tim mana? (A/B/C/D)')||'').toUpperCase().trim();
      if(!teams[t]){ alert('Tim tidak valid'); return; }
      const delta = parseInt((prompt('Masukkan perubahan skor (boleh negatif). Contoh: 5 atau -5')||'').trim(),10);
      if(Number.isNaN(delta)){ alert('Angka tidak valid'); return; }
      teams[t].score += delta;
      saveState(); updateScores(); renderBoard();
      showPopup(`🧮 Koreksi skor ${teams[t].name}: ${delta>0?'+':''}${delta} poin`, {center:true});
    }else if(choice==='4'){
      if(!cs.answered){ alert('Sel belum dijawab.'); return; }
      const t = (prompt('Tim baru? (A/B/C/D)')||'').toUpperCase().trim();
      if(!teams[t]){ alert('Tim tidak valid'); return; }
      const npStr = prompt(`Poin baru (opsional). ENTER=pakai ${cs.points||0}`);
      let newPts = Number.NaN;
      if(npStr && npStr.trim()!==''){
        const parsed = parseInt(npStr,10);
        if(Number.isNaN(parsed)){ alert('Angka tidak valid'); return; }
        newPts = parsed;
      }
      reassignCell(idx, t, newPts);
       }else if(choice==='5'){
      adminForceResumeGame();
    }
  }, true);
})();




