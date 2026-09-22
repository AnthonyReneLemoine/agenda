    'use strict';

    /* ============================================================
       ⚙️  CONFIG — REMPLACER CLIENT_ID
       Voir README.md pour obtenir votre Client ID OAuth2
    ============================================================ */
    const CLIENT_ID = '1091567193044-n64qqmbd6pu8l2pkkocab6cnf0nj7qui.apps.googleusercontent.com';
    const SCOPES    = 'https://www.googleapis.com/auth/calendar';
    const API_BASE  = 'https://www.googleapis.com/calendar/v3';

    /* ============================================================
       CONSTANTES AFFICHAGE
    ============================================================ */
    const H_START   = 9;
    const H_END     = 23;
    let SLOT_H    = 24;
    let DAY_W     = 192;

    const DAYS_FR   = ['Dim','Lun','Mar','Mer','Jeu','Ven','Sam'];
    const MONTHS_FR = ['janvier','février','mars','avril','mai','juin',
                       'juillet','août','septembre','octobre','novembre','décembre'];
    const MONTHS_SH = ['Jan','Fév','Mar','Avr','Mai','Jun','Jul','Aoû','Sep','Oct','Nov','Déc'];

    /* ============================================================
       ÉTAT
    ============================================================ */
    let accessToken = null;
    let viewDays    = 7;
    let viewStart   = addD(sod(new Date()), -((new Date().getDay()+6)%7));
    let calendars   = [];
    let events      = [];
    let filter      = '';
    let sidebarVis  = true;
    let mcDate      = new Date();
    let debTimer    = null;
    let worldClockTimer = null;
    let listView    = false;
    let editingDescriptionRaw = '';
    let editingDescriptionPlain = '';
    const LIST_RANGE_DAYS = 90;
    const EVENT_CACHE_TTL = 2 * 60 * 1000;
    const EVENT_PREFETCH_BEFORE_DAYS = 14;
    const EVENT_PREFETCH_AFTER_DAYS = 28;
    const eventRangeCache = new Map();
    let eventLoadSequence = 0;

    /* ============================================================
       AUTH — Google Identity Services (token model)
       Persistance : localStorage jusqu'à l'expiration réelle du jeton.
       L'autorisation déjà accordée est mémorisée séparément afin de
       renouveler le jeton sans redemander systématiquement le compte.
    ============================================================ */
    let tokenClient = null;
    const LS_TOKEN  = 'gcal_access_token';
    const LS_EXPIRY = 'gcal_token_expiry';
    const LS_GRANT  = 'gcal_oauth_grant';
    const AUTH_NOTE_DEFAULT = 'Accès sécurisé via Google OAuth2 — aucune donnée stockée sur nos serveurs.';

    function saveToken(token, expiresIn) {
      const expiry = Date.now() + (expiresIn - 60) * 1000;
      localStorage.setItem(LS_TOKEN,  token);
      localStorage.setItem(LS_EXPIRY, String(expiry));
      // Nettoyage des anciennes versions qui utilisaient sessionStorage.
      sessionStorage.removeItem(LS_TOKEN);
      sessionStorage.removeItem(LS_EXPIRY);
    }
    function loadToken() {
      let token  = localStorage.getItem(LS_TOKEN);
      let expiry = parseInt(localStorage.getItem(LS_EXPIRY) || '0', 10);

      // Migration transparente d'une session ouverte avec l'ancienne version.
      if (!token || !Number.isFinite(expiry) || Date.now() >= expiry) {
        const sessionToken  = sessionStorage.getItem(LS_TOKEN);
        const sessionExpiry = parseInt(sessionStorage.getItem(LS_EXPIRY) || '0', 10);
        if (sessionToken && Date.now() < sessionExpiry) {
          token = sessionToken;
          expiry = sessionExpiry;
          localStorage.setItem(LS_TOKEN, token);
          localStorage.setItem(LS_EXPIRY, String(expiry));
          sessionStorage.removeItem(LS_TOKEN);
          sessionStorage.removeItem(LS_EXPIRY);
        }
      }
      if (token && Date.now() < expiry) return token;
      clearToken(); return null;
    }
    function clearToken() {
      localStorage.removeItem(LS_TOKEN);
      localStorage.removeItem(LS_EXPIRY);
      sessionStorage.removeItem(LS_TOKEN);
      sessionStorage.removeItem(LS_EXPIRY);
    }

    function requireReconnection(message = 'Votre session Google a expiré. Cliquez sur « Se connecter avec Google » pour continuer.') {
      accessToken = null;
      clearToken();
      document.getElementById('auth-overlay').classList.add('on');
      document.getElementById('btn-so').style.display = 'none';
      document.getElementById('auth-note').textContent = message;
      setStatus('error', 'Reconnexion requise');
      showToast(message, 'warning', 6000);
    }

    function initGIS() {
      if (!window.google?.accounts?.oauth2) {
        setTimeout(initGIS, 200); return;
      }
      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: (resp) => {
          if (resp.error) {
            // Si Google ne peut plus réutiliser l'autorisation, le prochain clic
            // réaffichera le choix du compte au lieu de boucler sur l'erreur.
            localStorage.removeItem(LS_GRANT);
            const message = 'Autorisation Google à renouveler. Cliquez de nouveau sur « Se connecter avec Google ».';
            document.getElementById('auth-note').textContent = message;
            showToast(message, 'warning', 6000);
            return;
          }
          accessToken = resp.access_token;
          saveToken(resp.access_token, resp.expires_in || 3600);
          localStorage.setItem(LS_GRANT, '1');
          onSignedIn();
        }
      });

      // Restauration silencieuse tant que le jeton Google est encore valide.
      const saved = loadToken();
      if (saved) {
        accessToken = saved;
        onSignedIn();
      }

      // Google Identity Services ne fournit pas de renouvellement réellement
      // silencieux dans le navigateur. On attend donc une action volontaire
      // de l'utilisateur au lieu d'ouvrir périodiquement une fenêtre Google.
    }

    function doSignIn() {
      if (!tokenClient) { showToast('GIS non initialisé, patientez…', 'warn'); return; }
      document.getElementById('auth-note').textContent = AUTH_NOTE_DEFAULT;
      const prompt = localStorage.getItem(LS_GRANT) === '1' ? '' : 'select_account';
      tokenClient.requestAccessToken({ prompt });
    }

    async function doSignOut() {
      if (!confirm('Se déconnecter ?')) return;
      if (accessToken) google.accounts.oauth2.revoke(accessToken);
      accessToken = null;
      clearToken();
      localStorage.removeItem(LS_GRANT);
      calendars = []; events = [];
      eventRangeCache.clear();
      eventLoadSequence++;
      document.getElementById('auth-overlay').classList.add('on');
      document.getElementById('btn-so').style.display = 'none';
    }

    function onSignedIn() {
      document.getElementById('auth-note').textContent = AUTH_NOTE_DEFAULT;
      document.getElementById('auth-overlay').classList.remove('on');
      document.getElementById('btn-so').style.display = 'inline-flex';
      populateTimes(document.getElementById('m-ts'));
      populateTimes(document.getElementById('m-te'));
      renderMiniCal();
      initResponsive();   // ← adapte la vue selon l'écran
      renderAll();
      loadCalendarList();
      setTimeout(scrollToNow, 150);
      setInterval(updateNowBar, 60000);
    }

    /* ============================================================
       API GOOGLE CALENDAR — wrappers fetch
    ============================================================ */
    async function gcalFetch(url, options = {}) {
      if (!accessToken) {
        requireReconnection();
        throw new Error('Session Google expirée. Reconnexion requise.');
      }
      const resp = await fetch(url, {
        ...options,
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json',
          ...(options.headers || {})
        }
      });
      if (resp.status === 401) {
        requireReconnection();
        throw new Error('Session Google expirée. Reconnexion requise.');
      }
      if (!resp.ok) {
        const err = await resp.json().catch(() => ({}));
        throw new Error(err?.error?.message || `Erreur ${resp.status}`);
      }
      if (resp.status === 204) return null;
      return resp.json();
    }

    /* Récupère tous les calendriers de l'utilisateur */
    const LS_CAL_VIS = 'gcal_cal_visibility';
    const DEFAULT_HIDDEN = ["SA L'Hermine", 'Congés',
      'Jours fériés et autres fêtes en France', 'Exposition'];

    function loadCalVisibility() {
      try { return JSON.parse(localStorage.getItem(LS_CAL_VIS) || 'null'); }
      catch(e) { return null; }
    }
    function saveCalVisibility(cals) {
      const map = {};
      cals.forEach(c => { map[c.id] = c.visible; });
      localStorage.setItem(LS_CAL_VIS, JSON.stringify(map));
    }

    async function gcalGetCalendars() {
      const data    = await gcalFetch(`${API_BASE}/users/me/calendarList?maxResults=250`);
      const saved   = loadCalVisibility(); // null = première ouverture

      return (data.items || []).map(c => {
        let visible;
        if (saved && c.id in saved) {
          // Préférence déjà sauvegardée → la respecter
          visible = saved[c.id];
        } else {
          // Première ouverture → appliquer les masquages par défaut
          visible = !DEFAULT_HIDDEN.includes(c.summary);
        }
        return {
          id:      c.id,
          name:    c.summary || c.id,
          color:   c.backgroundColor || '#3b82f6',
          visible
        };
      });
    }

    /* Récupère les événements d'un calendrier sur une plage */
    async function gcalGetEvents(calId, timeMin, timeMax) {
      const params = new URLSearchParams({
        timeMin, timeMax,
        singleEvents: 'true',
        orderBy: 'startTime',
        maxResults: '2500'
      });
      const data = await gcalFetch(`${API_BASE}/calendars/${encodeURIComponent(calId)}/events?${params}`);
      const cal  = calendars.find(c => c.id === calId);
      const color = cal?.color || '#3b82f6';
      return (data.items || []).map(e => {
        const apiAllDay = !e.start.dateTime;
        const start = apiAllDay ? e.start.date + 'T00:00:00' : e.start.dateTime;
        const end   = apiAllDay ? e.end.date   + 'T00:00:00' : e.end.dateTime;
        // Certains événements importés sont enregistrés par Google avec des
        // dateTime de minuit à minuit, bien que son interface les place dans
        // la zone « journée entière ». On reproduit ce comportement.
        const allDay = apiAllDay || isMidnightSpanningEvent(start,end);
        const descriptionRaw = e.description || '';
        return {
          id:              e.id,
          calendarId:      calId,
          title:           e.summary || '(sans titre)',
          description:     plainTextDescription(descriptionRaw),
          descriptionRaw,
          // Les dates d'un événement « toute la journée » sont des dates locales.
          // Ne pas ajouter Z : cela les convertirait en UTC et peut décaler le jour.
          start,
          end,
          allDay,
          apiAllDay,
          backgroundColor: color,
          borderColor:     color,
          textColor:       '#ffffff'
        };
      });
    }

    /* Crée un événement */
    async function gcalCreateEvent(calId, body) {
      return gcalFetch(
        `${API_BASE}/calendars/${encodeURIComponent(calId)}/events`,
        { method: 'POST', body: JSON.stringify(body) }
      );
    }

    /* Met à jour un événement */
    async function gcalUpdateEvent(calId, evId, body) {
      // PATCH fusionne les champs imbriqués : il faut effacer explicitement
      // l'autre représentation lors d'une conversion date <-> dateTime.
      const patch = { ...body };
      for (const key of ['start', 'end']) {
        const value = body[key];
        if (value?.date) {
          patch[key] = { ...value, dateTime: null, timeZone: null };
        } else if (value?.dateTime) {
          patch[key] = { ...value, date: null };
        }
      }
      return gcalFetch(
        `${API_BASE}/calendars/${encodeURIComponent(calId)}/events/${encodeURIComponent(evId)}`,
        { method: 'PATCH', body: JSON.stringify(patch) }
      );
    }

    /* Supprime un événement */
    async function gcalDeleteEvent(calId, evId) {
      return gcalFetch(
        `${API_BASE}/calendars/${encodeURIComponent(calId)}/events/${encodeURIComponent(evId)}`,
        { method: 'DELETE' }
      );
    }

    /* Déplace un événement vers un autre calendrier (copie + suppression) */
    async function gcalMoveEvent(oldCalId, newCalId, evId) {
      return gcalFetch(
        `${API_BASE}/calendars/${encodeURIComponent(oldCalId)}/events/${encodeURIComponent(evId)}/move?destination=${encodeURIComponent(newCalId)}`,
        { method: 'POST' }
      );
    }

    /* Construit le body Google Calendar API pour un événement */
    function buildGCalBody(data) {
      const body = { summary: data.title, description: data.description || '' };
      const startDate = parseLocalDate(data.startDate);
      const endDate   = parseLocalDate(data.endDate);
      if (data.allDay) {
        // Google Calendar attend une date de fin exclusive pour les événements
        // « toute la journée ». Dans l'interface, la date de fin reste inclusive.
        const endExcl = new Date(endDate);
        endExcl.setDate(endExcl.getDate() + 1);
        body.start = { date: fmtD(startDate) };
        body.end   = { date: fmtD(endExcl) };
      } else {
        const st = combineDateAndTime(startDate, data.startTime);
        const et = combineDateAndTime(endDate,   data.endTime);
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
        body.start = { dateTime: st.toISOString(), timeZone: tz };
        body.end   = { dateTime: et.toISOString(), timeZone: tz };
      }
      /* Récurrence RRULE */
      if (data.recurrence && data.recurrence !== 'none' && data.recurrenceEnd) {
        const freqMap = { daily:'DAILY', weekly:'WEEKLY', biweekly:'WEEKLY', monthly:'MONTHLY' };
        const freq    = freqMap[data.recurrence];
        const until   = data.recurrenceEnd.replace(/-/g, '') + 'T000000Z';
        let rule = `RRULE:FREQ=${freq};UNTIL=${until}`;
        if (data.recurrence === 'biweekly') rule += ';INTERVAL=2';
        body.recurrence = [rule];
      }
      return body;
    }

    /* ============================================================
       UTILITAIRES DATES
    ============================================================ */
    function sod(d) { const r=new Date(d); r.setHours(0,0,0,0); return r; }
    function isoWeek(d){ const t=new Date(d); t.setHours(0,0,0,0); t.setDate(t.getDate()+3-(t.getDay()+6)%7); const w1=new Date(t.getFullYear(),0,4); return 1+Math.round(((t-w1)/86400000-3+(w1.getDay()+6)%7)/7); }
    function addD(d,n){ const r=new Date(d); r.setDate(r.getDate()+n); return r; }
    function sameD(a,b){ return a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate(); }
    function isToday(d){ return sameD(d,new Date()); }
    function eventHasEnded(ev,now=new Date()){
      const end=new Date(ev.end||ev.start);
      return !Number.isNaN(end.getTime())&&end<=now;
    }
    function pz(n){ return String(n).padStart(2,'0'); }
    function fmtD(d){ return `${d.getFullYear()}-${pz(d.getMonth()+1)}-${pz(d.getDate())}`; }
    function fmtT(d){ const tot=Math.round((d.getHours()*60+d.getMinutes())/15)*15; return `${pz(Math.floor(tot/60)%24)}:${pz(tot%60)}`; }
    function fmtTd(d){ return `${pz(d.getHours())}:${pz(d.getMinutes())}`; }
    function minToPx(m){ return (m/30)*SLOT_H; }
    function totalH(){ return (H_END-H_START)*2*SLOT_H; }

    function syncDisplayMetrics(){
      const rs=getComputedStyle(document.documentElement);
      const slot=parseFloat(rs.getPropertyValue('--slot-h'));
      const day=parseFloat(rs.getPropertyValue('--day-w'));
      if(!Number.isNaN(slot) && slot>0) SLOT_H=slot;
      if(!Number.isNaN(day) && day>0) DAY_W=day;
    }
    // Mesure l'espace réel restant, barre de défilement horizontale comprise.
    function fitCalendarHeight(){
      const available=document.getElementById('body-scroll').clientHeight;
      if(listView || available<=0) return;
      SLOT_H=Math.max(1,(available-1)/((H_END-H_START)*2));
      document.documentElement.style.setProperty('--slot-h',SLOT_H+'px');
    }

    let calendarResizeFrame=0;
    function scheduleCalendarResize(){
      if(calendarResizeFrame) return;
      calendarResizeFrame=requestAnimationFrame(()=>{
        calendarResizeFrame=0;
        renderAll();
      });
    }

    function parseLocalDate(s){ const p=s.split('-'); return new Date(+p[0],+p[1]-1,+p[2]); }
    function combineDateAndTime(date, ts){ const [h,m]=ts.split(':').map(Number); const r=new Date(date); r.setHours(h,m,0,0); return r; }

    function isMidnightSpanningEvent(startValue,endValue){
      const start=new Date(startValue), end=new Date(endValue);
      if(Number.isNaN(start.getTime())||Number.isNaN(end.getTime())) return false;
      const startsAtMidnight=start.getHours()===0&&start.getMinutes()===0&&start.getSeconds()===0;
      const endsAtMidnight=end.getHours()===0&&end.getMinutes()===0&&end.getSeconds()===0;
      if(!startsAtMidnight||!endsAtMidnight) return false;
      const followingDay=new Date(start);
      followingDay.setDate(followingDay.getDate()+1);
      return end>=followingDay;
    }

    function plainTextDescription(value){
      const source=String(value||'');
      if(!/<[a-z][\s\S]*>/i.test(source)) return source;
      const withBreaks=source
        .replace(/<br\s*\/?>/gi,'\n')
        .replace(/<\/(p|div|li|tr|h[1-6])\s*>/gi,'\n');
      const doc=new DOMParser().parseFromString(withBreaks,'text/html');
      return (doc.body.textContent||'')
        .replace(/\u00a0/g,' ')
        .replace(/[ \t]+\n/g,'\n')
        .replace(/\n{3,}/g,'\n\n')
        .trim();
    }

    function modalDescriptionForSave(){
      const value=document.getElementById('m-desc').value;
      return value===editingDescriptionPlain ? editingDescriptionRaw : value;
    }

    function updateWorldClocks(){
      const now = new Date();
      const format = tz => new Intl.DateTimeFormat('fr-FR', {
        hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false, timeZone: tz
      }).format(now);
      document.getElementById('clock-la').textContent = format('America/Los_Angeles');
      document.getElementById('clock-soledade').textContent = format('America/Sao_Paulo');
      document.getElementById('clock-abidjan').textContent = format('Africa/Abidjan');
    }

    function initWorldClocks(){
      updateWorldClocks();
      if(worldClockTimer) clearInterval(worldClockTimer);
      worldClockTimer = setInterval(updateWorldClocks, 1000);
    }

    /* ============================================================
       NAVIGATION
    ============================================================ */
    function nav(dir){ const step=listView?30:viewDays; viewStart=addD(viewStart,dir*step); renderAll(); loadEvents(); }
    function goToday(){ const t=sod(new Date()); viewStart=(viewDays===7)?addD(t,-((t.getDay()+6)%7)):t; renderAll(); loadEvents(); }
    function setView(n, btn){
      listView=false;
      viewDays=n;
      document.querySelectorAll('.vbtn').forEach(b=>b.classList.remove('on'));
      btn.classList.add('on');
      if(n===7){ const dow=(viewStart.getDay()+6)%7; viewStart=addD(viewStart,-dow); }
      renderAll(); loadEvents();
      if(typeof updateMobileNavActive==='function') updateMobileNavActive();
    }
    function setListView(btn){
      listView=true;
      document.querySelectorAll('.vbtn').forEach(b=>b.classList.remove('on'));
      if(btn && btn.classList.contains('vbtn')) btn.classList.add('on');
      viewStart=sod(new Date());
      renderAll(); loadEvents();
      if(typeof updateMobileNavActive==='function') updateMobileNavActive();
    }

    /* ============================================================
       SIDEBAR / MINI-CAL
    ============================================================ */
    function toggleSidebar(){
      sidebarVis=!sidebarVis;
      document.getElementById('sidebar').classList.toggle('hidden',!sidebarVis);
    }
    function renderMiniCal(){
      const y=mcDate.getFullYear(), m=mcDate.getMonth();
      document.getElementById('mc-title').textContent=`${MONTHS_FR[m][0].toUpperCase()+MONTHS_FR[m].slice(1)} ${y}`;
      const g=document.getElementById('mc-grid'); g.innerHTML='';
      ['L','M','M','J','V','S','D'].forEach(h=>{ const el=document.createElement('div'); el.className='mch'; el.textContent=h; g.appendChild(el); });
      const first=new Date(y,m,1); const off=(first.getDay()+6)%7;
      for(let i=0;i<off;i++) addMcD(g,new Date(y,m,1-(off-i)),true);
      const dim=new Date(y,m+1,0).getDate();
      for(let d=1;d<=dim;d++) addMcD(g,new Date(y,m,d),false);
      const rem=(7-(off+dim)%7)%7;
      for(let i=1;i<=rem;i++) addMcD(g,new Date(y,m+1,i),true);
    }
    function addMcD(g,date,other){
      const el=document.createElement('div');
      el.className='mcd'+(other?' other-mo':'')+(isToday(date)?' today':'');
      el.textContent=date.getDate();
      el.onclick=()=>{ viewStart=sod(date); renderAll(); loadEvents(); };
      g.appendChild(el);
    }
    function mcPrev(){ mcDate.setMonth(mcDate.getMonth()-1); renderMiniCal(); }
    function mcNext(){ mcDate.setMonth(mcDate.getMonth()+1); renderMiniCal(); }

    /* ============================================================
       RENDU PRINCIPAL
    ============================================================ */
    function renderAll(){
      syncDisplayMetrics();
      const calHead=document.getElementById('cal-head'), alldayRow=document.getElementById('allday-row');
      const calBody=document.getElementById('cal-body'), listViewEl=document.getElementById('list-view');
      if(listView){
        calHead.style.display='none'; alldayRow.style.display='none'; calBody.style.display='none';
        listViewEl.style.display='block';
        renderListView(); return;
      }
      calHead.style.display=''; alldayRow.style.display=''; calBody.style.display='';
      listViewEl.style.display='none';
      renderHead(); renderAllday(); fitCalendarHeight(); renderTimeCol(); renderDays(); bindScroll(); updateNowBar();
    }

    function renderHead(){
      const inner=document.getElementById('head-inner');
      inner.innerHTML=''; inner.style.width=(viewDays*DAY_W)+'px';
      for(let i=0;i<viewDays;i++){
        const d=addD(viewStart,i);
        const c=document.createElement('div');
        const wk=isoWeek(d); const wkClass=wk%2===0?'week-even':'week-odd';
        c.className='day-head '+wkClass+(isToday(d)?' is-today':'');
        c.style.width=DAY_W+'px';
        c.innerHTML=`<span class="dh-wd">${DAYS_FR[(d.getDay()+7)%7]}</span><span class="dh-d">${d.getDate()} ${MONTHS_SH[d.getMonth()]}</span>`;
        c.onclick=()=>{ viewStart=sod(d); viewDays=1; document.querySelectorAll('.vbtn').forEach(b=>b.classList.remove('on')); renderAll(); loadEvents(); };
        inner.appendChild(c);
      }
    }

    function renderAllday(){
      const inner=document.getElementById('allday-inner');
      inner.innerHTML=''; inner.style.width=(viewDays*DAY_W)+'px';
      for(let i=0;i<viewDays;i++){
        const d=addD(viewStart,i);
        const cell=document.createElement('div');
        cell.className='allday-cell'; cell.style.width=DAY_W+'px';
        cell.onclick=()=>openModal({start:d,end:d,allDay:true});
        events.filter(ev=>{
          if(!ev.allDay) return false;
          if(filter&&!ev.title.toLowerCase().includes(filter)) return false;
          const s=new Date(ev.start), e=ev.end?new Date(ev.end):s;
          return d>=sod(s)&&d<sod(e);
        }).forEach(ev=>{
          const chip=document.createElement('div');
          chip.className='allday-ev'+(eventHasEnded(ev)?' is-past':''); chip.style.background=ev.backgroundColor||'#64748b';
          chip.textContent=ev.title;
          chip.onclick=e2=>{e2.stopPropagation();openModal(null,true,ev);};
          cell.appendChild(chip);
        });
        inner.appendChild(cell);
      }
    }

    function renderTimeCol(){
      const inner=document.getElementById('time-inner');
      inner.innerHTML=''; inner.style.height=totalH()+'px';
      // On place le texte "Xh:00" sur la ligne du HAUT du slot correspondant.
      // Chaque slot fait SLOT_H px. Le label est positionné en absolu
      // à exactement minToPx((h - H_START)*60) px depuis le haut.
      const container=document.createElement('div');
      container.style.cssText='position:relative;height:'+totalH()+'px;';
      // Slots de fond (pour les bordures)
      for(let h=H_START;h<H_END;h++) for(let half=0;half<2;half++){
        const el=document.createElement('div');
        el.className='t-label'+(half?' half':'');
        el.style.height=SLOT_H+'px';
        container.appendChild(el);
      }
      // Labels positionnés en absolu exactement sur la ligne de l'heure
      for(let h=H_START;h<H_END;h++){
        const lbl=document.createElement('span');
        lbl.className='t-label-txt';
        lbl.textContent=`${pz(h)}:00`;
        lbl.style.top=minToPx((h-H_START)*60)+'px';
        container.appendChild(lbl);
      }
      inner.appendChild(container);
    }

    function renderDays(){
      const wrap=document.getElementById('days-inner');
      wrap.innerHTML=''; wrap.style.width=(viewDays*DAY_W)+'px'; wrap.style.height=totalH()+'px'; wrap.style.position='relative';
      for(let i=0;i<viewDays;i++){
        const d=addD(viewStart,i);
        const col=document.createElement('div');
        col.className='day-col'+(isToday(d)?' is-today':''); col.style.width=DAY_W+'px'; col.style.height=totalH()+'px';
        for(let h=H_START;h<H_END;h++) for(let half=0;half<2;half++){
          const sl=document.createElement('div'); sl.className='slot'+(half?' half':'');
          sl.onclick=()=>{ const start=new Date(d); start.setHours(h,half?30:0,0,0); openModal({start,end:new Date(start.getTime()+3600000),allDay:false}); };
          col.appendChild(sl);
        }
        const dayEvs=events.filter(ev=>{ if(ev.allDay) return false; if(filter&&!ev.title.toLowerCase().includes(filter)) return false; return sameD(new Date(ev.start),d); });
        computeLayout(dayEvs).forEach(({ev,col:c,cols:t})=>placeEvent(col,ev,c,t));
        wrap.appendChild(col);
      }
      for(let i=0;i<viewDays;i++) if(isToday(addD(viewStart,i))){
        const now=new Date(); const mins=(now.getHours()-H_START)*60+now.getMinutes();
        if(mins>=0&&mins<=(H_END-H_START)*60){
          const bar=document.createElement('div'); bar.id='now-bar';
          bar.style.top=minToPx(mins)+'px'; bar.style.left=(i*DAY_W)+'px'; bar.style.right=((viewDays-i-1)*DAY_W)+'px';
          wrap.appendChild(bar);
        }
        break;
      }
    }

    /* ============================================================
       LAYOUT CONFLITS
    ============================================================ */
    function computeLayout(evs){
      const sorted=[...evs].sort((a,b)=>new Date(a.start)-new Date(b.start)||new Date(a.end||a.start)-new Date(b.end||b.start));
      const result=[];
      let cluster=[], clusterEnd=-Infinity;

      function placeCluster(){
        if(!cluster.length)return;
        const columnEnds=[], clusterItems=[];
        cluster.forEach(ev=>{
          const start=new Date(ev.start).getTime();
          const end=(ev.end?new Date(ev.end):new Date(start+1800000)).getTime();
          let column=columnEnds.findIndex(columnEnd=>columnEnd<=start);
          if(column===-1){column=columnEnds.length;columnEnds.push(end);}
          else columnEnds[column]=end;
          const item={ev,col:column,cols:0};
          clusterItems.push(item);result.push(item);
        });
        clusterItems.forEach(item=>item.cols=columnEnds.length);
        cluster=[];clusterEnd=-Infinity;
      }

      sorted.forEach(ev=>{
        const start=new Date(ev.start).getTime();
        const end=(ev.end?new Date(ev.end):new Date(start+1800000)).getTime();
        if(cluster.length&&start>=clusterEnd)placeCluster();
        cluster.push(ev);clusterEnd=Math.max(clusterEnd,end);
      });
      placeCluster();
      return result;
    }

    function overlappingGeometry(colIdx,totalCols){
      const usable=DAY_W-4;
      if(totalCols<=1)return{left:2,width:usable,zIndex:2};
      const overlap=Math.min(22,Math.max(10,Math.round(usable*.12)));
      const width=Math.floor((usable+overlap*(totalCols-1))/totalCols);
      return{left:2+colIdx*(width-overlap),width,zIndex:2+colIdx};
    }

    function placeEvent(col,ev,colIdx=0,totalCols=1){
      const s=new Date(ev.start), e=ev.end?new Date(ev.end):new Date(s.getTime()+3600000);
      const sm=Math.max(0,(s.getHours()-H_START)*60+s.getMinutes());
      const em=Math.min((H_END-H_START)*60,(e.getHours()-H_START)*60+e.getMinutes());
      const dur=Math.max(15,em-sm);
      const geometry=overlappingGeometry(colIdx,totalCols);
      const el=document.createElement('div'); el.className='ev'+(eventHasEnded(ev)?' is-past':'');
      el.style.top=minToPx(sm)+'px'; el.style.height=Math.max(0,Math.min(totalH()-minToPx(sm),Math.max(22,minToPx(dur))))+'px';
      el.style.background=ev.backgroundColor||'#3b82f6'; el.style.left=geometry.left+'px'; el.style.right='auto'; el.style.width=geometry.width+'px'; el.style.zIndex=String(geometry.zIndex);
      el.dataset.evId=ev.id||''; el.dataset.calId=ev.calendarId||'';
      const desc=(ev.description||'').substring(0,60);
      el.innerHTML=`<div class="ev-t">${esc(ev.title)}</div><div class="ev-h">${fmtTd(s)} – ${fmtTd(e)}</div>${desc?`<div class="ev-d">${esc(desc)}</div>`:''}<div class="ev-resize"></div>`;
      el.addEventListener('mouseenter',ev2=>{ if(!dnd.active) showTip(ev2,ev); });
      el.addEventListener('mousemove', ev2=>{ if(!dnd.active) moveTip(ev2); });
      el.addEventListener('mouseleave',()=>{ if(!dnd.active) hideTip(); });
      el.querySelector('.ev-resize').addEventListener('mousedown',mde=>{ mde.stopPropagation(); mde.preventDefault(); hideTip(); startResize(mde,el,ev); });
      el.addEventListener('mousedown',mde=>{ if(mde.target.classList.contains('ev-resize')) return; mde.preventDefault(); hideTip(); startDrag(mde,el,ev); });
      col.appendChild(el);
    }

    /* ============================================================
       DRAG & DROP — DÉPLACEMENT
    ============================================================ */
    const dnd={active:false};

    function startDrag(mde,el,ev){
      const bs=document.getElementById('body-scroll');
      const ds=document.getElementById('days-scroll'), ghost=document.getElementById('drag-ghost');
      const rect=el.getBoundingClientRect(), dsRect=ds.getBoundingClientRect();
      const offsetY=mde.clientY-rect.top, startX=mde.clientX, startY=mde.clientY;
      let moved=false;
      const evDurMin=Math.round((new Date(ev.end||new Date(ev.start).getTime()+3600000)-new Date(ev.start))/60000);
      ghost.style.width=(DAY_W-4)+'px'; ghost.style.height=minToPx(evDurMin)+'px';
      ghost.style.background=ev.backgroundColor||'#3b82f6'; ghost.style.display='none';

      function onMove(mme){
        if(!moved&&Math.abs(mme.clientX-startX)<4&&Math.abs(mme.clientY-startY)<4) return;
        moved=true; hideTip(); el.classList.add('dragging'); ghost.style.display='block'; dnd.active=true;
        const relX=mme.clientX-dsRect.left+bs.scrollLeft, relY=mme.clientY-dsRect.top+bs.scrollTop-offsetY;
        const colIdx=Math.max(0,Math.min(viewDays-1,Math.floor(relX/DAY_W)));
        const snapMin=Math.max(0,Math.min((H_END-H_START)*60-evDurMin,Math.round((relY/SLOT_H)*30/15)*15));
        ghost.style.left=(colIdx*DAY_W+2)+'px'; ghost.style.top=minToPx(snapMin)+'px'; ghost.style.right='auto';
        dnd.colIdx=colIdx; dnd.snapMin=snapMin;
      }

      function onUp(){
        document.removeEventListener('mousemove',onMove); document.removeEventListener('mouseup',onUp);
        el.classList.remove('dragging'); ghost.style.display='none'; dnd.active=false;
        if(!moved){ openModal(null,true,ev); return; }
        const targetDay=addD(viewStart,dnd.colIdx);
        const newStart=new Date(targetDay); newStart.setHours(H_START+Math.floor(dnd.snapMin/60),dnd.snapMin%60,0,0);
        const newEnd=new Date(newStart.getTime()+evDurMin*60000);
        const idx=events.findIndex(x=>x.id===ev.id);
        if(idx!==-1){ events[idx].start=newStart.toISOString(); events[idx].end=newEnd.toISOString(); }
        renderDays(); renderAllday();
        setStatus('syncing','Déplacement…');
        const tz=Intl.DateTimeFormat().resolvedOptions().timeZone;
        gcalUpdateEvent(ev.calendarId, ev.id, {
          start:{ dateTime:newStart.toISOString(), timeZone:tz },
          end:  { dateTime:newEnd.toISOString(),   timeZone:tz }
        }).then(()=>{ eventRangeCache.delete(ev.calendarId); setStatus('ok'); }).catch(err=>{ setStatus('error',err.message); loadEvents({force:true}); });
      }
      document.addEventListener('mousemove',onMove); document.addEventListener('mouseup',onUp);
    }

    /* ============================================================
       DRAG & DROP — REDIMENSIONNEMENT
    ============================================================ */
    function startResize(mde,el,ev){
      const bs=document.getElementById('body-scroll');
      const ds=document.getElementById('days-scroll'), dsRect=ds.getBoundingClientRect();
      let moved=false;
      const evStart=new Date(ev.start), startMin=(evStart.getHours()-H_START)*60+evStart.getMinutes();
      el.classList.add('resizing');

      function onMove(mme){
        moved=true; hideTip();
        const relY=mme.clientY-dsRect.top+bs.scrollTop;
        const endMin=Math.max(startMin+15,Math.min((H_END-H_START)*60,Math.round((relY/SLOT_H)*30/15)*15));
        el.style.height=Math.max(minToPx(15),minToPx(endMin-startMin))+'px';
        const newEnd=new Date(evStart); newEnd.setMinutes(newEnd.getMinutes()+(endMin-startMin));
        const hEl=el.querySelector('.ev-h'); if(hEl) hEl.textContent=`${fmtTd(evStart)} – ${fmtTd(newEnd)}`;
        el._resizeEndMin=endMin;
      }

      function onUp(){
        document.removeEventListener('mousemove',onMove); document.removeEventListener('mouseup',onUp);
        el.classList.remove('resizing');
        if(!moved||el._resizeEndMin===undefined) return;
        const newEnd=new Date(evStart); newEnd.setHours(H_START+Math.floor(el._resizeEndMin/60),el._resizeEndMin%60,0,0);
        const idx=events.findIndex(x=>x.id===ev.id); if(idx!==-1) events[idx].end=newEnd.toISOString();
        renderDays();
        setStatus('syncing','Redimensionnement…');
        const tz=Intl.DateTimeFormat().resolvedOptions().timeZone;
        gcalUpdateEvent(ev.calendarId, ev.id, {
          start:{ dateTime:evStart.toISOString(), timeZone:tz },
          end:  { dateTime:newEnd.toISOString(),   timeZone:tz }
        }).then(()=>{ eventRangeCache.delete(ev.calendarId); setStatus('ok'); }).catch(err=>{ setStatus('error',err.message); loadEvents({force:true}); });
      }
      document.addEventListener('mousemove',onMove); document.addEventListener('mouseup',onUp);
    }

    function updateNowBar(){ const bar=document.getElementById('now-bar'); if(!bar) return; const now=new Date(); bar.style.top=minToPx((now.getHours()-H_START)*60+now.getMinutes())+'px'; }

    /* ============================================================
       SYNC SCROLL
    ============================================================ */
    function bindScroll(){
      const bs=document.getElementById('body-scroll');
      bs.onscroll=null;
      bs.onscroll=()=>{
        // Sync horizontal : en-tête jours + all-day suivent body-scroll en X
        document.getElementById('head-scroll').scrollLeft   = bs.scrollLeft;
        document.getElementById('allday-scroll').scrollLeft = bs.scrollLeft;
        // Pas besoin de sync vertical : time-col est dans body-scroll (sticky left)
      };
      document.getElementById('head-scroll').scrollLeft=bs.scrollLeft;
      document.getElementById('allday-scroll').scrollLeft=bs.scrollLeft;
    }
    function scrollToNow(){
      const now=new Date();
      const mins=(now.getHours()-H_START)*60+now.getMinutes()-60;
      document.getElementById('body-scroll').scrollTop=Math.max(0,minToPx(mins));
    }

    /* ============================================================
       TOOLTIP
    ============================================================ */
    function showTip(e,ev){ const t=document.getElementById('tip'); const s=new Date(ev.start),en=ev.end?new Date(ev.end):null; t.innerHTML=`<strong>${esc(ev.title)}</strong><br>${fmtTd(s)}${en?' – '+fmtTd(en):''}${ev.description?'<br><span style="opacity:.8">'+esc(ev.description.substring(0,90))+'</span>':''}`; t.style.display='block'; moveTip(e); }
    function moveTip(e){ const t=document.getElementById('tip'); t.style.left=(e.clientX+14)+'px'; t.style.top=(e.clientY-10)+'px'; }
    function hideTip(){ document.getElementById('tip').style.display='none'; }

    /* ============================================================
       RECHERCHE
    ============================================================ */
    function onSearch(){ clearTimeout(debTimer); debTimer=setTimeout(()=>{ filter=document.getElementById('srch').value.toLowerCase().trim(); listView?renderListView():(renderDays(),renderAllday()); },250); }

    /* ============================================================
       CHARGEMENT DONNÉES
    ============================================================ */
    async function loadCalendarList(){
      setStatus('syncing','Chargement…');
      try {
        calendars = await gcalGetCalendars();
        saveCalVisibility(calendars);   // ← mémorise l'état initial
        renderCalList();
        await loadEvents();
        setStatus('ok');
      } catch(e){ setStatus('error', e.message||'Erreur'); }
    }

    function cacheCovers(entry,startMs,endMs){
      return entry&&Date.now()-entry.fetchedAt<EVENT_CACHE_TTL&&entry.startMs<=startMs&&entry.endMs>=endMs;
    }

    function renderLoadedEvents(visibleCalendarIds){
      events=visibleCalendarIds.flatMap(id=>eventRangeCache.get(id)?.items||[]);
      listView?renderListView():(renderDays(),renderAllday());
    }

    async function loadEvents({force=false}={}){
      const requestId=++eventLoadSequence;
      const vis=calendars.filter(c=>c.visible).map(c=>c.id);
      if(!vis.length){ events=[]; listView?renderListView():(renderDays(),renderAllday()); setStatus('ok'); return; }
      const rangeStart=new Date(viewStart);
      const rangeEnd=addD(rangeStart,listView?LIST_RANGE_DAYS:viewDays);
      const startMs=rangeStart.getTime(), endMs=rangeEnd.getTime();
      const missing=vis.filter(id=>force||!cacheCovers(eventRangeCache.get(id),startMs,endMs));

      // Une semaine déjà préchargée s'affiche immédiatement, sans voile blanc.
      if(!missing.length){
        renderLoadedEvents(vis);
        setStatus('ok');
        return;
      }

      const fetchStart=listView?rangeStart:addD(rangeStart,-EVENT_PREFETCH_BEFORE_DAYS);
      const fetchEnd=listView?rangeEnd:addD(rangeEnd,EVENT_PREFETCH_AFTER_DAYS);
      const loadStartedAt=Date.now();
      setStatus('syncing','Chargement…');
      try {
        const loaded=await Promise.all(missing.map(async id=>({
          id,
          items:await gcalGetEvents(id,fetchStart.toISOString(),fetchEnd.toISOString())
        })));
        loaded.forEach(({id,items})=>{
          const current=eventRangeCache.get(id);
          if(!current||current.fetchedAt<=loadStartedAt){
            eventRangeCache.set(id,{
              startMs:fetchStart.getTime(),endMs:fetchEnd.getTime(),fetchedAt:loadStartedAt,items
            });
          }
        });
        // Si l'utilisateur a déjà changé de semaine, cette réponse reste utile
        // au cache mais ne doit pas écraser la nouvelle vue.
        if(requestId!==eventLoadSequence) return;
        renderLoadedEvents(vis);
        setStatus('ok');
      } catch(e){ if(requestId===eventLoadSequence) setStatus('error', e.message||'Erreur'); }
    }

    /* ============================================================
       LISTE CALENDRIERS
    ============================================================ */
    function renderCalList(){
      const wrap=document.getElementById('cal-list');
      const sorted=[...calendars].sort((a,b)=>{ if(a.visible!==b.visible) return b.visible?1:-1; return (a.name||'').localeCompare(b.name||'','fr',{sensitivity:'base'}); });
      wrap.innerHTML=sorted.map(c=>`
        <div class="cal-item">
          <input class="cal-cb" type="checkbox" ${c.visible?'checked':''} style="--cc:${c.color||'#3b82f6'};" data-id="${c.id}" onchange="toggleCal('${c.id}',this.checked)">
          <span style="width:9px;height:9px;border-radius:50%;background:${c.color||'#64748b'};flex-shrink:0;"></span>
          <span style="font-size:12px;color:#374151;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${esc(c.name)}</span>
        </div>`).join('');
    }

    function toggleCal(id,vis){
      const c=calendars.find(x=>x.id===id); if(!c) return;
      c.visible=vis;
      saveCalVisibility(calendars);   // ← persiste la préférence
      renderCalList(); loadEvents();
    }

    /* ============================================================
       STATUS / TOASTS
    ============================================================ */
    function showLoading(){ document.getElementById('loading-overlay').classList.add('on'); }
    function hideLoading(){ document.getElementById('loading-overlay').classList.remove('on'); }
    function setStatus(s,msg){
      const dot=document.getElementById('sdot'),txt=document.getElementById('stxt');
      const m={syncing:{c:'sd-sy',t:msg||'Synchronisation…'},ok:{c:'sd-ok',t:msg||'Système OK'},error:{c:'sd-er',t:msg||'Erreur'}};
      const cfg=m[s]||m.ok; dot.className='sdot '+cfg.c; txt.textContent=cfg.t;
      if(s==='syncing') showLoading(); else hideLoading();
    }
    function showToast(msg,type='info',dur=3500){
      const box=document.getElementById('toast-box'), t=document.createElement('div');
      const cls={success:'t-ok',error:'t-er',warning:'t-wa',info:'t-in'}, ico={success:'✓',error:'✕',warning:'⚠',info:'ℹ'};
      t.className=`toast ${cls[type]||'t-in'}`; t.innerHTML=`<span>${ico[type]||''}</span><span>${esc(msg)}</span>`;
      box.appendChild(t); setTimeout(()=>{ t.classList.add('exit'); setTimeout(()=>t.remove(),260); },dur);
    }

    /* ============================================================
       MODAL
    ============================================================ */
    function openModal(info=null,isEdit=false,ev=null){
      const sel=document.getElementById('m-cal');
      sel.innerHTML=calendars.map(c=>`<option value="${c.id}">${esc(c.name)}</option>`).join('');
      resetModal();
      if(ev){
        const s=new Date(ev.start),e=ev.end?new Date(ev.end):null;
        editingDescriptionRaw=ev.descriptionRaw ?? ev.description ?? '';
        editingDescriptionPlain=plainTextDescription(editingDescriptionRaw);
        document.getElementById('m-title').value=ev.title||'';
        document.getElementById('m-desc').value=editingDescriptionPlain;
        document.getElementById('m-id').value=ev.id||'';
        document.getElementById('m-old-cal').value=ev.calendarId||'';
        document.getElementById('m-ds').value=fmtD(s);
        document.getElementById('m-de').value=fmtD(s);
        if(ev.allDay){
          document.getElementById('m-ad').checked=true;
          document.getElementById('m-ts').disabled=true; document.getElementById('m-te').disabled=true;
          if(e){ const ae=new Date(e); ae.setDate(ae.getDate()-1); if(fmtD(ae)!==fmtD(s)){ document.getElementById('m-de').value=fmtD(ae); document.getElementById('m-md').checked=true; document.getElementById('row-de').style.display='flex'; } }
        } else {
          document.getElementById('m-ts').value=fmtT(s);
          if(e){ document.getElementById('m-te').value=fmtT(e); if(fmtD(s)!==fmtD(e)){ document.getElementById('m-de').value=fmtD(e); document.getElementById('m-md').checked=true; document.getElementById('row-de').style.display='flex'; } }
        }
        sel.value=ev.calendarId||'';
        document.getElementById('btn-del').style.display='inline-flex';
        document.getElementById('btn-dup').style.display='inline-flex';
        document.getElementById('modal-title').textContent="Modifier l'événement";
      } else if(info){
        document.getElementById('modal-title').textContent='Nouvel événement';
        if(info.allDay){ document.getElementById('m-ds').value=fmtD(info.start); document.getElementById('m-de').value=fmtD(info.start); document.getElementById('m-ad').checked=true; document.getElementById('m-ts').disabled=true; document.getElementById('m-te').disabled=true; }
        else if(info.start){ document.getElementById('m-ds').value=fmtD(info.start); document.getElementById('m-de').value=fmtD(info.start); document.getElementById('m-ts').value=fmtT(info.start); document.getElementById('m-te').value=fmtT(info.end||new Date(info.start.getTime()+3600000)); }
      } else {
        document.getElementById('modal-title').textContent='Nouvel événement';
        const now=new Date(); document.getElementById('m-ds').value=fmtD(now); document.getElementById('m-de').value=fmtD(now); document.getElementById('m-ts').value=fmtT(now); document.getElementById('m-te').value=fmtT(new Date(now.getTime()+3600000));
      }
      document.getElementById('modal').style.display='block';
      setTimeout(()=>document.getElementById('m-title').focus(),50);
    }

    function closeModal(){ document.getElementById('modal').style.display='none'; }
    function resetModal(){
      ['m-title','m-desc','m-id','m-old-cal'].forEach(id=>{ document.getElementById(id).value=''; });
      document.getElementById('m-ad').checked=false; document.getElementById('m-md').checked=false;
      editingDescriptionRaw=''; editingDescriptionPlain='';
      document.getElementById('row-de').style.display='none'; document.getElementById('m-ts').disabled=false; document.getElementById('m-te').disabled=false;
      document.getElementById('m-rec').value='none'; document.getElementById('row-recend').style.display='none';
      document.getElementById('btn-del').style.display='none'; document.getElementById('btn-dup').style.display='none';
    }

    function addMinutesToTime(timeValue, minutesToAdd){
      const [hours, minutes] = String(timeValue || '00:00').split(':').map(Number);
      const totalMinutes = ((hours * 60) + minutes + minutesToAdd + 1440) % 1440;
      return `${pz(Math.floor(totalMinutes / 60))}:${pz(totalMinutes % 60)}`;
    }
    function syncEndTimeWithStart(){
      const startTime = document.getElementById('m-ts').value;
      document.getElementById('m-te').value = addMinutesToTime(startTime, 60);
    }
    function onAlldayChange(){ const v=document.getElementById('m-ad').checked; document.getElementById('m-ts').disabled=v; document.getElementById('m-te').disabled=v; }
    function onMultidayChange(){ document.getElementById('row-de').style.display=document.getElementById('m-md').checked?'flex':'none'; }
    function onRecChange(){ document.getElementById('row-recend').style.display=document.getElementById('m-rec').value!=='none'?'block':'none'; }

    async function saveEvent(){
      const title=document.getElementById('m-title').value.trim();
      if(!title){ showToast('Le titre est obligatoire','error'); return; }
      const evId=document.getElementById('m-id').value, oldCal=document.getElementById('m-old-cal').value, calId=document.getElementById('m-cal').value;
      const allDay=document.getElementById('m-ad').checked, multi=document.getElementById('m-md').checked;
      const ds=document.getElementById('m-ds').value, de=multi?document.getElementById('m-de').value:ds;
      if(!ds||!de||parseLocalDate(de)<parseLocalDate(ds)){ showToast('La date de fin doit être égale ou postérieure à la date de début','error'); return; }
      const data={ title, description:modalDescriptionForSave(), startDate:ds, endDate:de, startTime:document.getElementById('m-ts').value, endTime:document.getElementById('m-te').value, allDay, recurrence:document.getElementById('m-rec').value, recurrenceEnd:document.getElementById('m-rend').value };
      const body=buildGCalBody(data);
      setStatus('syncing','Enregistrement…');
      try {
        if(evId){
          if(oldCal!==calId){ await gcalMoveEvent(oldCal,calId,evId); await gcalUpdateEvent(calId,evId,body); }
          else { await gcalUpdateEvent(calId,evId,body); }
          showToast("Événement mis à jour",'success');
        } else {
          await gcalCreateEvent(calId,body);
          showToast('Événement créé','success');
        }
        closeModal(); await loadEvents({force:true}); setStatus('ok');
      } catch(e){ showToast('Erreur : '+(e.message||e),'error'); setStatus('error'); }
    }

    async function deleteCurrent(){
      if(!confirm('Supprimer cet événement ?')) return;
      const id=document.getElementById('m-id').value, calId=document.getElementById('m-cal').value;
      setStatus('syncing','Suppression…');
      try { await gcalDeleteEvent(calId,id); showToast('Événement supprimé','success'); closeModal(); await loadEvents({force:true}); setStatus('ok'); }
      catch(e){ showToast('Erreur : '+(e.message||e),'error'); setStatus('error'); }
    }

    async function duplicateCurrent(){
      const calId=document.getElementById('m-cal').value, allDay=document.getElementById('m-ad').checked, multi=document.getElementById('m-md').checked;
      const ds=document.getElementById('m-ds').value, de=multi?document.getElementById('m-de').value:ds;
      if(!ds||!de||parseLocalDate(de)<parseLocalDate(ds)){ showToast('La date de fin doit être égale ou postérieure à la date de début','error'); return; }
      const data={ title:document.getElementById('m-title').value.trim()+' (copie)', description:modalDescriptionForSave(), startDate:ds, endDate:de, startTime:document.getElementById('m-ts').value, endTime:document.getElementById('m-te').value, allDay, recurrence:'none', recurrenceEnd:'' };
      setStatus('syncing','Duplication…');
      try { await gcalCreateEvent(calId,buildGCalBody(data)); showToast('Événement dupliqué','success'); closeModal(); await loadEvents({force:true}); setStatus('ok'); }
      catch(e){ showToast('Erreur : '+(e.message||e),'error'); setStatus('error'); }
    }

    /* ============================================================
       VUE LISTE
    ============================================================ */
    function renderListView(){
      const wrap=document.getElementById('list-view');
      wrap.innerHTML='';
      const now=new Date();
      const filtered=events
        .filter(ev=>!filter||ev.title.toLowerCase().includes(filter))
        .sort((a,b)=>new Date(a.start)-new Date(b.start));
      if(!filtered.length){
        wrap.innerHTML='<div class="lv-empty">Aucun événement dans cette période.</div>';
        return;
      }
      /* Grouper par jour */
      const byDay=new Map();
      filtered.forEach(ev=>{
        const d=sod(new Date(ev.start)), key=fmtD(d);
        if(!byDay.has(key)) byDay.set(key,{date:d,evs:[]});
        byDay.get(key).evs.push(ev);
      });
      const days=[...byDay.values()].sort((a,b)=>a.date-b.date);
      let nowInserted=false;
      days.forEach(({date,evs})=>{
        const todayDay=isToday(date);
        const sorted=[...evs].sort((a,b)=>{
          if(a.allDay!==b.allDay) return a.allDay?-1:1;
          return new Date(a.start)-new Date(b.start);
        });
        const wk=isoWeek(date);
        const dayRow=document.createElement('div'); dayRow.className='lv-day'+(wk%2!==0?' lv-week-odd':'');
        const dateDiv=document.createElement('div');
        dateDiv.className='lv-date'+(todayDay?' is-today':'');
        dateDiv.innerHTML=`<span class="lv-date-lbl">${DAYS_FR[(date.getDay()+7)%7].toUpperCase()}., ${MONTHS_SH[date.getMonth()].toUpperCase()}</span><span class="lv-date-d">${date.getDate()}</span>`;
        const eventsDiv=document.createElement('div'); eventsDiv.className='lv-events';
        sorted.forEach(ev=>{
          if(!nowInserted&&!ev.allDay&&new Date(ev.start)>=now){
            nowInserted=true;
            const nowEl=document.createElement('div'); nowEl.className='lv-now';
            nowEl.innerHTML='<div class="lv-now-dot"></div><span class="lv-now-lbl">Maintenant</span><div class="lv-now-line"></div>';
            eventsDiv.appendChild(nowEl);
          }
          const s=new Date(ev.start), e=ev.end?new Date(ev.end):null;
          const timeStr=ev.allDay?'Jour entier':`De ${fmtTd(s)} à ${fmtTd(e||s)}`;
          const evDiv=document.createElement('div'); evDiv.className='lv-ev'+(eventHasEnded(ev,now)?' is-past':'');
          evDiv.innerHTML=`<div class="lv-dot" style="background:${ev.backgroundColor||'#64748b'};"></div><div class="lv-ev-body"><div class="lv-ev-title">${esc(ev.title)}</div><div class="lv-ev-time">${timeStr}</div>${ev.description?`<div class="lv-ev-desc">${esc(ev.description.substring(0,80))}</div>`:''}</div>`;
          evDiv.addEventListener('click',()=>openModal(null,true,ev));
          eventsDiv.appendChild(evDiv);
        });
        if(todayDay&&!nowInserted){
          nowInserted=true;
          const nowEl=document.createElement('div'); nowEl.className='lv-now';
          nowEl.innerHTML='<div class="lv-now-dot"></div><span class="lv-now-lbl">Maintenant</span><div class="lv-now-line"></div>';
          eventsDiv.appendChild(nowEl);
        }
        dayRow.appendChild(dateDiv); dayRow.appendChild(eventsDiv);
        wrap.appendChild(dayRow);
      });
    }

    /* ============================================================
       UTILITAIRES
    ============================================================ */
    function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
    function populateTimes(sel){ sel.innerHTML=''; for(let h=0;h<24;h++) for(let m=0;m<60;m+=15){ const v=`${pz(h)}:${pz(m)}`; const o=document.createElement('option'); o.value=v; o.textContent=v; sel.appendChild(o); } }

    document.getElementById('m-ts').addEventListener('change', syncEndTimeWithStart);

    /* ============================================================
       INIT
    ============================================================ */
    document.addEventListener('keydown', e=>{ if(e.key==='Escape') closeModal(); });
    document.getElementById('modal').addEventListener('click', e=>{ if(e.target===document.getElementById('modal')) closeModal(); });
    initGIS();
    initWorldClocks();

    /* ============================================================
       RESPONSIVE — MOBILE
    ============================================================ */
    const isMobile = () => window.innerWidth <= 768;

    function initResponsive() {
      const nav = document.getElementById('mobile-nav');
      const fab = document.getElementById('fab');
      if (isMobile()) {
        nav.style.display = 'flex';
        fab.style.display = 'flex';
        // Vue par défaut mobile : 1 jour
        if (viewDays === 7) {
          viewDays = 1;
          viewStart = sod(new Date());
          document.querySelectorAll('.vbtn').forEach(b=>b.classList.remove('on'));
        }
        updateMobileNavActive();
      } else {
        nav.style.display = 'none';
        fab.style.display = 'none';
      }
    }

    function setMobileView(n) {
      listView=false;
      viewDays = n;
      document.querySelectorAll('.vbtn').forEach(b=>b.classList.remove('on'));
      if(n===7){ const dow=(viewStart.getDay()+6)%7; viewStart=addD(viewStart,-dow); }
      renderAll(); loadEvents();
      updateMobileNavActive();
    }

    function updateMobileNavActive() {
      ['mnv-1','mnv-7'].forEach(id=>{
        const el=document.getElementById(id); if(!el) return;
        const n=parseInt(id.split('-')[1]);
        el.classList.toggle('active', !listView&&viewDays===n);
      });
      const listeEl=document.getElementById('mnv-liste');
      if(listeEl) listeEl.classList.toggle('active',listView);
    }

    function closeSidebarMobile() {
      if(!isMobile()) return;
      sidebarVis = false;
      document.getElementById('sidebar').classList.add('hidden');
      document.getElementById('sidebar-scrim').classList.remove('on');
    }

    // Surcharge toggleSidebar pour mobile (scrim)
    const _toggleSidebarOrig = toggleSidebar;
    function toggleSidebar() {
      sidebarVis = !sidebarVis;
      document.getElementById('sidebar').classList.toggle('hidden', !sidebarVis);
      if (isMobile()) {
        document.getElementById('sidebar-scrim').classList.toggle('on', sidebarVis);
      }
    }

    // Touch swipe sur le calendrier (gauche/droite = nav)
    (function initSwipe() {
      const swipeEl = document.getElementById('body-scroll');
      let tx = 0, ty = 0;
      swipeEl.addEventListener('touchstart', e=>{
        tx = e.touches[0].clientX;
        ty = e.touches[0].clientY;
      }, {passive:true});
      swipeEl.addEventListener('touchend', e=>{
        const dx = e.changedTouches[0].clientX - tx;
        const dy = e.changedTouches[0].clientY - ty;
        if(Math.abs(dx) > Math.abs(dy) && Math.abs(dx) > 60) {
          nav(dx < 0 ? 1 : -1);
        }
      }, {passive:true});
    })();

    // Init au chargement et au resize
    window.addEventListener('resize', ()=>{
      initResponsive();
      scheduleCalendarResize();
    });
    // Recalcule aussi après chargement/filtrage des journées entières,
    // changement de vue, ouverture du panneau ou changement de zoom.
    const calendarSizeObserver=new ResizeObserver(scheduleCalendarResize);
    calendarSizeObserver.observe(document.getElementById('body-scroll'));
    // initResponsive() est appelé dans onSignedIn() après le login
