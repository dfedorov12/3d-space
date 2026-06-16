/* ════════════════════════════════════════════════════════════════════════
   DIHAG · 3D-Space — virtueller Konferenzraum
   - MSAL-Login (DIHAG-Tenant)
   - Zugriff nur für freigegebene Mail-Adressen (SharePoint-Liste AppPermissions, App='3d-space')
   - Three.js: prozeduraler Konferenzraum
   - PeerJS: serverloses WebRTC-Mesh (Peer-IDs aus den freigegebenen Mails abgeleitet)
   - Räumliches Audio: näher = lauter
   ════════════════════════════════════════════════════════════════════════ */

// ════════════════════════════════════════════════════════════════
// CONFIG  (gleiche App-Registrierung wie die übrigen DIHAG-Apps)
// ════════════════════════════════════════════════════════════════
const CLIENT_ID    = '75e627e8-2de0-4ec6-bec9-311757b89e08';
const TENANT_ID    = 'fdb70646-023a-403b-a4b9-1f474a935123';
const SCOPES       = ['User.Read', 'Sites.Read.All', 'Mail.Send'];

const PERM_SITE    = 'dihag.sharepoint.com:/sites/ticket';  // hier liegt die AppPermissions-Liste
const PERM_LIST    = 'AppPermissions';
const APP_KEY      = '3d-space';                            // Spalte "App" in AppPermissions
const PEER_PREFIX  = 'dihag3dspace--';                      // Namespace auf dem PeerJS-Broker
const REQUEST_TO   = 'fedorov@dihag.com';                  // Empfänger der Freigabe-Anfrage

// Diese Adressen kommen IMMER rein (Owner/Admins) – auch wenn die Liste leer/nicht lesbar ist.
const SUPER_ADMINS = ['fedorov@dihag.com'];

// ICE-Server für WebRTC. STUN reicht in offenen Netzen. Im Firmennetz (Firewall / symmetrisches
// NAT, UDP geblockt) wird ein TURN-Server gebraucht – sonst sehen sich zwei Leute zwar im Raum,
// hören sich aber nicht. Dann einfach den TURN-Block ausfüllen (Host/User/Passwort eintragen):
const ICE_SERVERS = [
  { urls: 'stun:stun.l.google.com:19302' },
  { urls: 'stun:global.stun.twilio.com:3478' },
  // { urls: 'turn:DEIN-TURN-HOST:3478',  username: 'USER', credential: 'PASS' },
  // { urls: 'turns:DEIN-TURN-HOST:5349', username: 'USER', credential: 'PASS' },  // TURN über TLS/443 (firewall-freundlich)
];

// ════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════
const $id = id => document.getElementById(id);
const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

function peerIdFor(email){ return PEER_PREFIX + String(email).toLowerCase().replace(/[^a-z0-9]/g,''); }
function nameFromEmail(email){
  return String(email).split('@')[0].replace(/[._-]+/g,' ').replace(/\b\w/g, c => c.toUpperCase());
}
function initials(name){
  return (name||'?').split(' ').filter(Boolean).map(n=>n[0]).join('').substring(0,2).toUpperCase() || '?';
}
function hueOf(email){ let h=0; for(const c of String(email)) h=(h*31 + c.charCodeAt(0))>>>0; return h % 360; }
function cssColor(email){ return `hsl(${hueOf(email)} 60% 55%)`; }

let _toastT;
function toast(msg, ms=2600){
  const t = $id('toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(_toastT); _toastT = setTimeout(()=>t.classList.remove('show'), ms);
}

// ════════════════════════════════════════════════════════════════
// AUTH  (MSAL)
// ════════════════════════════════════════════════════════════════
let msalApp, account, myEmail, myName;

async function initAuth(){
  const redirectUri = window.location.href.split('?')[0].split('#')[0];
  msalApp = new msal.PublicClientApplication({
    auth:  { clientId: CLIENT_ID, authority:`https://login.microsoftonline.com/${TENANT_ID}`, redirectUri },
    cache: { cacheLocation:'localStorage', storeAuthStateInCookie:true }
  });
  await msalApp.initialize();
  await msalApp.handleRedirectPromise();
  const accounts = msalApp.getAllAccounts();
  if (accounts.length){ account = accounts[0]; return true; }
  return false;
}

async function doLogin(){
  $id('boot-btn').style.display='none';
  $id('boot-err').style.display='none';
  $id('boot-spinner').style.display='block';
  $id('boot-sub').textContent='Weiterleitung zur Anmeldung…';
  try{
    // Redirect statt Popup → kein neues Fenster, kein Popup-Blocker.
    // Die Rückkehr verarbeitet initAuth() via handleRedirectPromise().
    await msalApp.loginRedirect({ scopes: SCOPES });
  }catch(e){
    $id('boot-err').textContent = e.message;
    $id('boot-err').style.display='block';
    $id('boot-btn').style.display='block';
    $id('boot-btn').textContent='Erneut versuchen';
    $id('boot-sub').textContent='';
  }
}

function doLogout(){
  try{ peer?.destroy(); }catch{}
  try{ localStream?.getTracks().forEach(t=>t.stop()); }catch{}
  // Redirect-Logout (kein Popup); Rückkehr landet wieder auf der Boot-Seite.
  msalApp?.logoutRedirect({ account }).catch(()=> location.reload());
}

async function getToken(){
  if(!account) throw new Error('Nicht angemeldet');
  try{ return (await msalApp.acquireTokenSilent({ scopes:SCOPES, account })).accessToken; }
  catch{ return (await msalApp.acquireTokenPopup({ scopes:SCOPES, account })).accessToken; }
}

async function gGet(path){
  const tok = await getToken();
  const r = await fetch('https://graph.microsoft.com/v1.0'+path, { headers:{ Authorization:'Bearer '+tok } });
  if(!r.ok) throw new Error('Graph '+r.status+' · '+path);
  return r.json();
}
async function gPost(path, body){
  const tok = await getToken();
  const r = await fetch('https://graph.microsoft.com/v1.0'+path, {
    method:'POST', headers:{ Authorization:'Bearer '+tok, 'Content-Type':'application/json' }, body:JSON.stringify(body)
  });
  if(!r.ok) throw new Error('Graph '+r.status);
  return r.status===204 ? null : r.json();
}

// ════════════════════════════════════════════════════════════════
// ZUGRIFF  (AppPermissions = Freigabe + Teilnehmer-Verzeichnis)
// ════════════════════════════════════════════════════════════════
let roster = new Set();   // alle freigegebenen Mail-Adressen für diesen Raum

async function loadRoster(){
  const site = await gGet(`/sites/${PERM_SITE}`);
  const list = await gGet(`/sites/${site.id}/lists/${PERM_LIST}`);
  const data = await gGet(`/sites/${site.id}/lists/${list.id}/items?$expand=fields&$top=999`);
  const out = new Set();
  for(const it of (data.value||[])){
    const f = it.fields || {};
    const em = (f.UserEmail||'').toLowerCase().trim();
    if(!em) continue;
    if((f.App===APP_KEY || f.App==='*') && (f.Role||'').toLowerCase()!=='none') out.add(em);
  }
  return out;
}

async function afterLogin(){
  $id('boot-sub').textContent='Prüfe Freigabe…';
  myEmail = (account?.username || '').toLowerCase();
  myName  = account?.name || nameFromEmail(myEmail);

  let listOk = true;
  try{ roster = await loadRoster(); }
  catch(e){ listOk = false; roster = new Set(); console.warn('[roster]', e.message); }

  const isAdmin = SUPER_ADMINS.includes(myEmail);
  const allowed = isAdmin || roster.has(myEmail);
  // Admins immer ins Verzeichnis, damit sie für andere auffindbar sind
  SUPER_ADMINS.forEach(a => roster.add(a));

  if(!allowed){ showNoAccess(listOk); return; }
  if(isAdmin && !listOk) toast('Hinweis: AppPermissions-Liste nicht lesbar – du bist als Admin allein im Raum.', 5000);

  enterApp();
}

function showNoAccess(listOk){
  $id('boot').style.display='none';
  $id('no-access').style.display='flex';
  $id('nac-msg').textContent = listOk
    ? `Du (${myEmail}) bist für diesen Raum noch nicht freigegeben. Stelle eine Anfrage – die IT schaltet dich frei.`
    : `Die Freigabeliste konnte nicht gelesen werden. Bitte wende dich an die IT.`;
}

async function requestAccess(){
  const btn = $id('nac-req-btn'); btn.disabled=true; btn.textContent='…';
  try{
    await gPost('/me/sendMail', {
      message:{
        subject:`3D-Space · Freigabe-Anfrage – ${myName}`,
        body:{ contentType:'HTML', content:
`<p>Hallo,</p>
<p>folgende Person beantragt Zugriff auf den <strong>DIHAG 3D-Space</strong> (virtueller Konferenzraum):</p>
<table style="border-collapse:collapse;font-size:14px;font-family:sans-serif">
  <tr><td style="padding:4px 12px 4px 0;color:#666">Name</td><td><strong>${esc(myName)}</strong></td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#666">E-Mail</td><td>${esc(myEmail)}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#666">App</td><td>3d-space</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#666">Datum</td><td>${new Date().toLocaleString('de-DE')}</td></tr>
</table>
<p style="margin-top:16px">Freigabe: Eintrag in der Liste <strong>AppPermissions</strong> mit App=<code>3d-space</code>, Role=<code>viewer</code>.</p>` },
        toRecipients:[{ emailAddress:{ address: REQUEST_TO } }]
      },
      saveToSentItems:false
    });
    $id('nac-sent').style.display='block';
    btn.style.display='none';
  }catch(e){
    $id('nac-err').style.display='block';
    $id('nac-err').textContent='Fehler: '+e.message;
    btn.disabled=false; btn.textContent='📧 Freigabe anfragen';
  }
}

// ════════════════════════════════════════════════════════════════
// APP-EINSTIEG → Szene rendern, dann Mikro-Gate
// ════════════════════════════════════════════════════════════════
async function enterApp(){
  $id('boot').style.display='none';
  $id('app').style.display='block';
  $id('me-name').textContent = myName;
  const av = $id('me-av'); av.textContent = initials(myName); av.style.background = cssColor(myEmail);

  await ensureThree();
  buildScene();
  spawnMyAvatar();
  startRenderLoop();

  // Erst auf Nutzergeste hin Mikro + AudioContext aktivieren (Browser-Policy)
  $id('mic-gate').style.display='flex';
}

// ════════════════════════════════════════════════════════════════
// THREE.JS  — Szene & Raum
// ════════════════════════════════════════════════════════════════
let _threeReady = window.THREE ? Promise.resolve() : new Promise(res => window.addEventListener('three-ready', res, { once:true }));
function ensureThree(){ return _threeReady; }

let renderer, scene, camera, controls, clock, listener;
let myAvatar;
let viewMode = 'follow'; // 'follow' | 'overview'
const ROOM = { x:7.2, z:9.2 };          // halbe Innenmaße (Bewegungsgrenzen)
const TABLE = { x:1.45, z:4.9 };        // Sperrzone (Tisch)

function buildScene(){
  const THREE = window.THREE;
  const canvas = $id('scene');

  renderer = new THREE.WebGLRenderer({ canvas, antialias:true, powerPreference:'high-performance' });
  renderer.setPixelRatio(Math.min(devicePixelRatio, 2));
  renderer.setSize(innerWidth, innerHeight);
  renderer.shadowMap.enabled = true;
  renderer.shadowMap.type = THREE.PCFSoftShadowMap;
  renderer.toneMapping = THREE.ACESFilmicToneMapping;
  renderer.toneMappingExposure = 1.05;
  renderer.outputColorSpace = THREE.SRGBColorSpace;

  scene = new THREE.Scene();
  clock = new THREE.Clock();

  camera = new THREE.PerspectiveCamera(55, innerWidth/innerHeight, 0.1, 600);
  camera.position.set(-6, 3.2, 6.5);

  controls = new window.OrbitControls(camera, renderer.domElement);
  controls.enableDamping = true; controls.dampingFactor = 0.08;
  controls.enablePan = false;
  controls.minDistance = 2.6; controls.maxDistance = 9;
  controls.maxPolarAngle = 1.45;
  controls.target.set(-2.4, 1.2, 2);

  // sanftes Umgebungslicht / Reflexionen
  const pmrem = new THREE.PMREMGenerator(renderer);
  scene.environment = pmrem.fromScene(new window.RoomEnvironment(), 0.035).texture;

  buildSky(THREE);
  buildRoom(THREE);
  buildLights(THREE);
  buildFurniture(THREE);

  // Audio-Listener (deine Ohren) – wird ans eigene Avatar gehängt
  listener = new THREE.AudioListener();

  addEventListener('resize', onResize);
  setupInput();
}

function buildSky(THREE){
  const geo = new THREE.SphereGeometry(280, 32, 16);
  const mat = new THREE.ShaderMaterial({
    side: THREE.BackSide, depthWrite:false,
    uniforms:{ top:{value:new THREE.Color(0x6db3f2)}, bot:{value:new THREE.Color(0xeaf4ff)}, off:{value:40.0}, exp:{value:0.65} },
    vertexShader:`varying vec3 vW; void main(){ vW=(modelMatrix*vec4(position,1.0)).xyz; gl_Position=projectionMatrix*modelViewMatrix*vec4(position,1.0);} `,
    fragmentShader:`uniform vec3 top; uniform vec3 bot; uniform float off; uniform float exp; varying vec3 vW;
      void main(){ float h=normalize(vW+vec3(0.0,off,0.0)).y; float t=pow(max(h,0.0),exp); gl_FragColor=vec4(mix(bot,top,clamp(t,0.0,1.0)),1.0);} `
  });
  scene.add(new THREE.Mesh(geo, mat));

  // Skyline-Silhouette hinter der Fensterfront (+X)
  const sky = new THREE.Group();
  const mat1 = new THREE.MeshStandardMaterial({ color:0x9fb3cc, roughness:1, metalness:0 });
  const mat2 = new THREE.MeshStandardMaterial({ color:0x86a0bd, roughness:1, metalness:0 });
  let seed = 7;
  const rnd = ()=> (seed = (seed*9301+49297)%233280)/233280;
  for(let i=0;i<26;i++){
    const w = 1.6 + rnd()*2.6, h = 5 + rnd()*22, d = 1.6 + rnd()*2.6;
    const b = new THREE.Mesh(new THREE.BoxGeometry(w,h,d), rnd()>0.5?mat1:mat2);
    b.position.set(16 + rnd()*30, h/2 - 1, -22 + i*1.9 + rnd()*1.2);
    sky.add(b);
  }
  scene.add(sky);

  // Bodendunst draußen
  const ground = new THREE.Mesh(new THREE.PlaneGeometry(160,160),
    new THREE.MeshStandardMaterial({ color:0xc7d4e3, roughness:1 }));
  ground.rotation.x = -Math.PI/2; ground.position.set(30,-1,0);
  scene.add(ground);
}

function buildRoom(THREE){
  // Boden (Parkett-Look)
  const floorMat = new THREE.MeshStandardMaterial({ color:0x9c7b58, roughness:0.55, metalness:0.05 });
  const floor = new THREE.Mesh(new THREE.PlaneGeometry(ROOM.x*2+1, ROOM.z*2+1), floorMat);
  floor.rotation.x = -Math.PI/2; floor.receiveShadow = true;
  scene.add(floor);

  // Teppich unter dem Tisch
  const rug = new THREE.Mesh(new THREE.PlaneGeometry(5.5, 12),
    new THREE.MeshStandardMaterial({ color:0x26303f, roughness:0.95 }));
  rug.rotation.x = -Math.PI/2; rug.position.y = 0.01; rug.receiveShadow = true;
  scene.add(rug);

  const wallMat = new THREE.MeshStandardMaterial({ color:0xf2efe9, roughness:0.9 });
  const H = 4, T = 0.2;
  const wall = (w,h,d,x,y,z)=>{ const m=new THREE.Mesh(new THREE.BoxGeometry(w,h,d), wallMat); m.position.set(x,y,z); m.receiveShadow=true; scene.add(m); return m; };
  // Rückwand (-X) mit Bildschirm, Seitenwände (±Z), Fensterwand (+X) separat
  wall(T, H, ROOM.z*2+1, -ROOM.x-0.4, H/2, 0);                 // -X (hinten, Screen-Wand)
  wall(ROOM.x*2+1, H, T, 0, H/2, -ROOM.z-0.4);                 // -Z
  wall(ROOM.x*2+1, H, T, 0, H/2,  ROOM.z+0.4);                 // +Z

  // Decke
  const ceil = new THREE.Mesh(new THREE.PlaneGeometry(ROOM.x*2+1, ROOM.z*2+1),
    new THREE.MeshStandardMaterial({ color:0xf6f7fb, roughness:1 }));
  ceil.rotation.x = Math.PI/2; ceil.position.y = H; scene.add(ceil);

  // Deckenleuchten (emissiv)
  const panelMat = new THREE.MeshStandardMaterial({ color:0xffffff, emissive:0xfff4e0, emissiveIntensity:1.4, roughness:1 });
  for(const z of [-5.5,-1.8,1.8,5.5]){
    const p = new THREE.Mesh(new THREE.BoxGeometry(2.6,0.06,1.0), panelMat);
    p.position.set(0, H-0.05, z); scene.add(p);
  }

  // Fensterwand (+X): Rahmen + Glas
  const frameMat = new THREE.MeshStandardMaterial({ color:0x2b3340, roughness:0.4, metalness:0.6 });
  const glassMat = new THREE.MeshPhysicalMaterial({ color:0xbfe0ff, roughness:0.06, metalness:0, transmission:0.0, transparent:true, opacity:0.16, side:THREE.DoubleSide });
  const wx = ROOM.x+0.4;
  // Glasfläche
  const glass = new THREE.Mesh(new THREE.PlaneGeometry(ROOM.z*2-0.6, H-0.3), glassMat);
  glass.rotation.y = -Math.PI/2; glass.position.set(wx, H/2, 0); scene.add(glass);
  // Rahmen: oben/unten + vertikale Pfosten
  const frame = (w,h,d,y,z)=>{ const m=new THREE.Mesh(new THREE.BoxGeometry(w,h,d), frameMat); m.position.set(wx,y,z); scene.add(m); };
  frame(0.25, 0.25, ROOM.z*2+1, H-0.12, 0);
  frame(0.25, 0.4, ROOM.z*2+1, 0.2, 0);
  for(let z=-ROOM.z; z<=ROOM.z; z+=2.4) frame(0.18, H, 0.14, H/2, z);

  // Wandbildschirm an der -X-Wand
  const board = new THREE.Group();
  const bezel = new THREE.Mesh(new THREE.BoxGeometry(0.12, 2.0, 3.4),
    new THREE.MeshStandardMaterial({ color:0x10141c, roughness:0.4, metalness:0.5 }));
  bezel.position.set(-ROOM.x-0.28, 1.9, 0);
  const screen = new THREE.Mesh(new THREE.PlaneGeometry(3.0, 1.7), makeScreenMat(THREE));
  screen.rotation.y = Math.PI/2; screen.position.set(-ROOM.x-0.21, 1.9, 0);
  board.add(bezel, screen); scene.add(board);
  // global merken (für „Bildschirm teilen")
  wallScreen = screen; wallDefaultMap = screen.material.map;
}

function makeScreenMat(THREE){
  const c = document.createElement('canvas'); c.width=1024; c.height=576;
  const x = c.getContext('2d');
  const g = x.createLinearGradient(0,0,1024,576);
  g.addColorStop(0,'#0b1b3a'); g.addColorStop(1,'#10325f');
  x.fillStyle=g; x.fillRect(0,0,1024,576);
  x.fillStyle='rgba(255,255,255,.92)'; x.font='700 92px Inter, sans-serif'; x.textAlign='center';
  x.fillText('DIHAG', 512, 285);
  x.fillStyle='rgba(120,190,255,.9)'; x.font='500 38px Inter, sans-serif';
  x.fillText('3D-Space · Konferenzraum', 512, 350);
  const tex = new THREE.CanvasTexture(c); tex.colorSpace = THREE.SRGBColorSpace;
  return new THREE.MeshStandardMaterial({ map:tex, emissive:0xffffff, emissiveMap:tex, emissiveIntensity:0.5, roughness:0.6 });
}

function buildLights(THREE){
  scene.add(new THREE.HemisphereLight(0xeaf2ff, 0x8a7a66, 0.55));

  const sun = new THREE.DirectionalLight(0xfff2e0, 2.0);
  sun.position.set(16, 14, 6);
  sun.castShadow = true;
  sun.shadow.mapSize.set(2048,2048);
  sun.shadow.camera.left=-14; sun.shadow.camera.right=14;
  sun.shadow.camera.top=14; sun.shadow.camera.bottom=-14;
  sun.shadow.camera.far=60; sun.shadow.bias=-0.0004;
  scene.add(sun);

  for(const z of [-5,0,5]){
    const p = new THREE.PointLight(0xfff0d8, 14, 12, 2);
    p.position.set(0, 3.5, z); scene.add(p);
  }
}

function buildFurniture(THREE){
  // Konferenztisch
  const tableTop = new THREE.Mesh(new THREE.BoxGeometry(2.6, 0.1, 8.4),
    new THREE.MeshStandardMaterial({ color:0x3a2a1f, roughness:0.28, metalness:0.05 }));
  tableTop.position.set(0, 0.75, 0); tableTop.castShadow = tableTop.receiveShadow = true;
  scene.add(tableTop);
  const inlay = new THREE.Mesh(new THREE.BoxGeometry(2.2, 0.02, 8.0),
    new THREE.MeshStandardMaterial({ color:0x4a3727, roughness:0.2, metalness:0.1 }));
  inlay.position.set(0, 0.81, 0); scene.add(inlay);
  // Tischbeine (Chrom)
  const legMat = new THREE.MeshStandardMaterial({ color:0xb9c0c8, roughness:0.25, metalness:0.9 });
  for(const sx of [-1.05,1.05]) for(const sz of [-3.6,3.6]){
    const leg = new THREE.Mesh(new THREE.BoxGeometry(0.12,0.74,0.12), legMat);
    leg.position.set(sx, 0.37, sz); leg.castShadow = true; scene.add(leg);
  }

  // Stühle entlang beider Seiten + Kopfenden
  const seats = [];
  for(const z of [-3,-1,1,3]){ seats.push([-1.95,z, Math.PI/2]); seats.push([1.95,z,-Math.PI/2]); }
  seats.push([0,-4.6,0]); seats.push([0,4.6,Math.PI]);
  for(const [x,z,ry] of seats){ const ch = makeChair(THREE); ch.position.set(x,0,z); ch.rotation.y = ry; scene.add(ch); }

  // Pflanzen in den Ecken
  for(const [x,z] of [[-6.4,-8.4],[-6.4,8.4],[6.4,-8.4],[6.4,8.4]]){
    const pl = makePlant(THREE); pl.position.set(x,0,z); scene.add(pl);
  }

  // Sideboard an der -Z-Wand
  const side = new THREE.Mesh(new THREE.BoxGeometry(4.5,0.9,0.6),
    new THREE.MeshStandardMaterial({ color:0x6b4b34, roughness:0.4 }));
  side.position.set(-3.5,0.45,-8.7); side.castShadow = side.receiveShadow = true; scene.add(side);
}

function makeChair(THREE){
  const g = new THREE.Group();
  const fab = new THREE.MeshStandardMaterial({ color:0x2f3a4a, roughness:0.85 });
  const chrome = new THREE.MeshStandardMaterial({ color:0xc2c8cf, roughness:0.25, metalness:0.9 });
  const seat = new THREE.Mesh(new THREE.BoxGeometry(0.55,0.1,0.55), fab); seat.position.y=0.5; seat.castShadow=true;
  const back = new THREE.Mesh(new THREE.BoxGeometry(0.55,0.62,0.1), fab); back.position.set(0,0.82,-0.24); back.castShadow=true;
  const stem = new THREE.Mesh(new THREE.CylinderGeometry(0.05,0.05,0.45,12), chrome); stem.position.y=0.27;
  const base = new THREE.Mesh(new THREE.CylinderGeometry(0.3,0.34,0.05,18), chrome); base.position.y=0.04;
  g.add(seat,back,stem,base); return g;
}

function makePlant(THREE){
  const g = new THREE.Group();
  const pot = new THREE.Mesh(new THREE.CylinderGeometry(0.28,0.22,0.5,16),
    new THREE.MeshStandardMaterial({ color:0xb9b1a4, roughness:0.8 }));
  pot.position.y=0.25; pot.castShadow=true;
  const leafMat = new THREE.MeshStandardMaterial({ color:0x2f7d43, roughness:0.85 });
  for(let i=0;i<5;i++){
    const blob = new THREE.Mesh(new THREE.IcosahedronGeometry(0.34 - i*0.02, 1), leafMat);
    blob.position.set((Math.random()-0.5)*0.4, 0.7 + i*0.22, (Math.random()-0.5)*0.4);
    blob.scale.set(1,1.3,1); blob.castShadow=true; g.add(blob);
  }
  g.add(pot); return g;
}

// ════════════════════════════════════════════════════════════════
// AVATARE
// ════════════════════════════════════════════════════════════════
function makeAvatar(THREE, email, name){
  const g = new THREE.Group();
  const col = new THREE.Color(`hsl(${hueOf(email)}, 60%, 55%)`);
  const body = new THREE.Mesh(new THREE.CapsuleGeometry(0.28, 0.7, 6, 14),
    new THREE.MeshStandardMaterial({ color:col, roughness:0.6, metalness:0.05 }));
  body.position.y = 0.63; body.castShadow = true;
  const head = new THREE.Mesh(new THREE.SphereGeometry(0.22, 24, 18),
    new THREE.MeshStandardMaterial({ color:col.clone().offsetHSL(0,-0.1,0.12), roughness:0.5 }));
  head.position.y = 1.32; head.castShadow = true;
  // kleine "Nase", damit man die Blickrichtung sieht
  const nose = new THREE.Mesh(new THREE.ConeGeometry(0.06,0.14,10),
    new THREE.MeshStandardMaterial({ color:0xffffff, roughness:0.5 }));
  nose.rotation.x = Math.PI/2; nose.position.set(0,1.32,0.22);

  // Sprech-Ring
  const ring = new THREE.Mesh(new THREE.TorusGeometry(0.42,0.05,10,40),
    new THREE.MeshStandardMaterial({ color:0x38bdf8, emissive:0x38bdf8, emissiveIntensity:1.6, roughness:0.4 }));
  ring.rotation.x = Math.PI/2; ring.position.y = 0.06; ring.visible = false;

  const label = makeLabel(THREE, name, `hsl(${hueOf(email)},70%,60%)`);
  const hand = makeHand(THREE); hand.visible = false;

  g.add(body, head, nose, ring, label, hand);
  g.userData = { ring, label, hand, col };
  return g;
}

function makeLabel(THREE, text, color){
  const c = document.createElement('canvas'); c.width=256; c.height=64;
  const x = c.getContext('2d');
  rr(x,3,3,250,58,16); x.fillStyle='rgba(10,18,32,.82)'; x.fill();
  rr(x,3,3,250,58,16); x.strokeStyle=color; x.lineWidth=2.5; x.stroke();
  x.fillStyle='#fff'; x.font='600 27px Inter, sans-serif'; x.textAlign='center'; x.textBaseline='middle';
  x.fillText(text.length>16 ? text.slice(0,15)+'…' : text, 128, 34);
  const tex = new THREE.CanvasTexture(c); tex.colorSpace = THREE.SRGBColorSpace; tex.anisotropy = 4;
  const spr = new THREE.Sprite(new THREE.SpriteMaterial({ map:tex, transparent:true, depthTest:false, depthWrite:false }));
  spr.scale.set(1.35, 0.34, 1); spr.position.y = 1.95; spr.renderOrder = 999;
  return spr;
}
function makeHand(THREE){
  const c = document.createElement('canvas'); c.width=64; c.height=64;
  c.getContext('2d').font='52px serif'; c.getContext('2d').textAlign='center'; c.getContext('2d').textBaseline='middle';
  c.getContext('2d').fillText('✋',32,38);
  const tex = new THREE.CanvasTexture(c);
  const spr = new THREE.Sprite(new THREE.SpriteMaterial({ map:tex, transparent:true, depthTest:false }));
  spr.scale.set(0.5,0.5,1); spr.position.y = 2.35; spr.renderOrder = 1000;
  return spr;
}
function rr(x,a,b,w,h,r){ x.beginPath(); x.moveTo(a+r,b); x.arcTo(a+w,b,a+w,b+h,r); x.arcTo(a+w,b+h,a,b+h,r); x.arcTo(a,b+h,a,b,r); x.arcTo(a,b,a+w,b,r); x.closePath(); }

function spawnMyAvatar(){
  const THREE = window.THREE;
  myAvatar = makeAvatar(THREE, myEmail, myName);
  myAvatar.position.set(-2.4, 0, 2);
  myAvatar.rotation.y = Math.PI/2; // Blick zum Tisch (+X)
  myAvatar.userData.label.visible = false; // eigenes Schild ausblenden
  myAvatar.add(listener);                  // deine Ohren = deine Position
  scene.add(myAvatar);
  controls.target.copy(myAvatar.position).add(new THREE.Vector3(0,1.2,0));
}

// ════════════════════════════════════════════════════════════════
// EINGABE / BEWEGUNG
// ════════════════════════════════════════════════════════════════
const keys = {};
function isTyping(){ const a=document.activeElement; return !!a && (a.tagName==='INPUT' || a.tagName==='TEXTAREA' || a.isContentEditable); }
function setupInput(){
  addEventListener('keydown', e=>{
    if(e.code==='Space' && pttMode && !isTyping()){ e.preventDefault(); if(!pttActive){ pttActive=true; updateMicLive(); broadcast(); } return; }
    if(isTyping()) return;
    keys[e.key.toLowerCase()] = true;
  });
  addEventListener('keyup', e=>{
    if(e.code==='Space' && pttMode){ e.preventDefault(); pttActive=false; updateMicLive(); broadcast(); return; }
    keys[e.key.toLowerCase()] = false;
  });
  addEventListener('blur', ()=>{ for(const k in keys) keys[k]=false; pttActive=false; updateMicLive(); });
  setupJoystick();
}

// Mobiler Joystick (nur Touch)
function setupJoystick(){
  if(!('ontouchstart' in window) && !navigator.maxTouchPoints) return;
  const base=$id('joystick'), knob=$id('joy-knob');
  if(!base) return;
  base.style.display='block';
  let active=false, cx=0, cy=0; const R=46;
  const start=e=>{ active=true; const r=base.getBoundingClientRect(); cx=r.left+r.width/2; cy=r.top+r.height/2; mv(e); };
  const mv=e=>{ if(!active) return; const t=e.touches?e.touches[0]:e; let dx=t.clientX-cx, dy=t.clientY-cy; const d=Math.hypot(dx,dy)||1; if(d>R){ dx=dx/d*R; dy=dy/d*R; } knob.style.transform=`translate(${dx}px,${dy}px)`; joyStr=dx/R; joyFwd=-dy/R; if(e.cancelable) e.preventDefault(); };
  const end=()=>{ active=false; joyFwd=joyStr=0; knob.style.transform='translate(0,0)'; };
  base.addEventListener('touchstart',start,{passive:false});
  base.addEventListener('touchmove',mv,{passive:false});
  base.addEventListener('touchend',end); base.addEventListener('touchcancel',end);
}

function onResize(){
  camera.aspect = innerWidth/innerHeight; camera.updateProjectionMatrix();
  renderer.setSize(innerWidth, innerHeight);
}

function handleMovement(dt){
  if(!myAvatar) return;
  const THREE = window.THREE;
  const speed = 3.4;
  let fwd = (keys['w']||keys['arrowup']?1:0) - (keys['s']||keys['arrowdown']?1:0) + joyFwd;
  let str = (keys['d']||keys['arrowright']?1:0) - (keys['a']||keys['arrowleft']?1:0) + joyStr;
  fwd = Math.max(-1, Math.min(1, fwd)); str = Math.max(-1, Math.min(1, str));
  if(Math.abs(fwd) < 0.05 && Math.abs(str) < 0.05) return;

  // Richtung relativ zur Kamera (auf XZ-Ebene)
  const camDir = new THREE.Vector3(); camera.getWorldDirection(camDir); camDir.y=0; camDir.normalize();
  const right = new THREE.Vector3().crossVectors(camDir, new THREE.Vector3(0,1,0)).normalize();
  const move = new THREE.Vector3().addScaledVector(camDir, fwd).addScaledVector(right, str);
  if(move.lengthSq()===0) return;
  move.normalize().multiplyScalar(speed*dt);

  const prev = myAvatar.position.clone();
  let nx = prev.x + move.x, nz = prev.z + move.z;
  nx = Math.max(-ROOM.x, Math.min(ROOM.x, nx));
  nz = Math.max(-ROOM.z, Math.min(ROOM.z, nz));
  // Tisch-Sperrzone: nicht hineinlaufen
  if(Math.abs(nx) < TABLE.x && Math.abs(nz) < TABLE.z){ nx = prev.x; nz = prev.z; }

  myAvatar.position.set(nx, 0, nz);
  // Avatar in Bewegungsrichtung drehen
  const yaw = Math.atan2(move.x, move.z);
  myAvatar.rotation.y = lerpAngle(myAvatar.rotation.y, yaw, 0.2);
  scheduleBroadcast();
}
function lerpAngle(a,b,t){ let d=((b-a+Math.PI)%(2*Math.PI))-Math.PI; return a+d*t; }

// ════════════════════════════════════════════════════════════════
// RENDER-LOOP
// ════════════════════════════════════════════════════════════════
function startRenderLoop(){
  const THREE = window.THREE;
  const tmp = new THREE.Vector3();
  renderer.setAnimationLoop(()=>{
    const dt = Math.min(clock.getDelta(), 0.05);
    handleMovement(dt);

    // Kamera-Ziel
    if(viewMode==='follow' && myAvatar){
      tmp.copy(myAvatar.position).add(new THREE.Vector3(0,1.2,0));
      controls.target.lerp(tmp, 0.15);
    }
    controls.update();

    // Remote-Avatare sanft zur Zielposition bewegen
    peers.forEach(p=>{
      if(!p.avatar) return;
      if(p.tx!==undefined){
        p.avatar.position.x += (p.tx - p.avatar.position.x)*0.2;
        p.avatar.position.z += (p.tz - p.avatar.position.z)*0.2;
        p.avatar.rotation.y = lerpAngle(p.avatar.rotation.y, p.tRy||0, 0.2);
      }
      // Sprech-Level
      const lv = p.analyser ? audioLevel(p.analyser) : 0;
      updateSpeaking(p, lv > 0.05);
    });
    // eigenes Sprech-Level
    if(myAvatar){
      const lv = (isMicLive() && localAnalyser) ? audioLevel(localAnalyser) : 0;
      updateSpeaking({ avatar:myAvatar, _spk:mySpeak }, lv > 0.05);
      mySpeak = myAvatar.userData.ring.visible;
    }
    renderer.render(scene, camera);
  });
}
let mySpeak=false;

function updateSpeaking(p, on){
  const ring = p.avatar?.userData?.ring; if(!ring) return;
  ring.visible = on;
  if(on){ const s = 1 + Math.sin(performance.now()/120)*0.06; ring.scale.set(s,s,s); }
}

function audioLevel(an){
  const buf = an._buf || (an._buf = new Uint8Array(an.fftSize));
  an.getByteTimeDomainData(buf);
  let s=0; for(let i=0;i<buf.length;i++){ const v=(buf[i]-128)/128; s+=v*v; }
  return Math.sqrt(s/buf.length);
}

// ════════════════════════════════════════════════════════════════
// AUDIO + WEBRTC-MESH (PeerJS)
// ════════════════════════════════════════════════════════════════
let peer, localStream, hasMic=false, micEnabled=true, handUp=false;
let actx, localAnalyser;
const peers = new Map();   // peerId → { email, name, avatar, dataConn, call, posAudio, audioEl, analyser, tx, tz, tRy, muted, hand, status, screenOut, screenIn, sharing }

// Feature-Zustände
let screenStream=null;            // mein geteilter Bildschirm-Stream
let currentPresenter=null;        // peerId des aktuell Teilenden, oder 'me'
let _presenterName='';
let wallScreen=null, wallDefaultMap=null;   // Wand-Leinwand-Mesh + Standardtextur
let pttMode=false, pttActive=false;         // Push-to-Talk
let lowPerf=false;
let joyFwd=0, joyStr=0;                      // mobiler Joystick
let currentMicId=null;
let _chatOpen=false, _chatUnread=0;

async function enableMicAndJoin(){
  $id('mic-gate').style.display='none';
  // AudioContext der THREE-Listener verwenden (gleicher Graph wie das räumliche Audio)
  actx = window.THREE.AudioContext.getContext();
  if(actx.state==='suspended') await actx.resume();

  try{
    localStream = await navigator.mediaDevices.getUserMedia({ audio:{ echoCancellation:true, noiseSuppression:true, autoGainControl:true }, video:false });
    hasMic = true;
    const src = actx.createMediaStreamSource(localStream);
    localAnalyser = actx.createAnalyser(); localAnalyser.fftSize = 512;
    src.connect(localAnalyser); // nur Analyse, nicht an die Ausgabe (kein Echo)
    toast('🎤 Mikrofon aktiv – du bist im Raum');
  }catch(e){
    hasMic = false; localStream = silentStream();
    setMicUI(false); micEnabled=false;
    toast('Ohne Mikrofon beigetreten (nur zuhören)', 4000);
  }
  initPeer();
}

function silentStream(){
  const dst = actx.createMediaStreamDestination();
  const osc = actx.createOscillator(); const g = actx.createGain(); g.gain.value = 0;
  osc.connect(g).connect(dst); osc.start();
  return dst.stream;
}

function initPeer(){
  const myId = peerIdFor(myEmail);
  peer = new Peer(myId, { debug: 1, config:{ iceServers: ICE_SERVERS } });

  peer.on('open', ()=>{ connectRoster(); setInterval(rescan, 15000); setInterval(heartbeat, 3000); });

  peer.on('call', call=>{
    // Bildschirm-Übertragung getrennt behandeln
    if(call.metadata?.kind === 'screen'){ acceptScreenCall(call); return; }
    // Doppel-(Voice-)Calls vermeiden
    const ex = peers.get(call.peer);
    if(ex?.call){ try{ call.close(); }catch{} return; }
    call.answer(localStream);
    bindCall(call);
  });

  peer.on('connection', conn=> bindDataConn(conn));

  peer.on('error', err=>{
    if(String(err.type)==='peer-unavailable') return; // Gegenstelle offline → normal
    console.warn('[peer]', err.type, err.message);
  });
  peer.on('disconnected', ()=>{ try{ peer.reconnect(); }catch{} });

  addEventListener('beforeunload', ()=>{ broadcast({ t:'bye' }); try{ peer.destroy(); }catch{} });
}

function connectRoster(){
  roster.forEach(email=>{
    if(email===myEmail) return;
    connectPeer(peerIdFor(email), email);
  });
}
function rescan(){
  roster.forEach(email=>{
    if(email===myEmail) return;
    const pid = peerIdFor(email);
    const p = peers.get(pid);
    if(!p || (!p.dataConn?.open && !p.call)) connectPeer(pid, email);
  });
}

function connectPeer(pid, email){
  // Daten-Verbindung (Position/Status)
  const conn = peer.connect(pid, { reliable:false, metadata:{ email:myEmail, name:myName } });
  bindDataConn(conn, email);
  // Audio-Call
  const call = peer.call(pid, localStream, { metadata:{ email:myEmail, name:myName } });
  if(call) bindCall(call, email);
}

function ensurePeer(pid, email, name){
  let p = peers.get(pid);
  if(!p){
    const THREE = window.THREE;
    const av = makeAvatar(THREE, email||pid, name||nameFromEmail(email||'Gast'));
    // Startposition: Platz nach Index
    const i = peers.size;
    av.position.set(i%2 ? 2.4 : -2.4, 0, -3 + (Math.floor(i/2)%4)*2);
    scene.add(av);
    p = { email, name, avatar:av, tx:av.position.x, tz:av.position.z, tRy:0, muted:false, hand:false, status:'connecting' };
    peers.set(pid, p);
    // Wenn ich gerade teile, dem Neuzugang den Bildschirm anbieten
    if(screenStream) callScreen(pid);
    renderPeople();
  }
  if(email) p.email = email;
  if(name){ p.name = name; updateLabel(p, name); }
  return p;
}

function bindDataConn(conn, email){
  conn.on('open', ()=>{
    const p = ensurePeer(conn.peer, email || conn.metadata?.email, conn.metadata?.name);
    p.dataConn = conn;
    sendState(conn); // sofort meinen Zustand schicken
    renderPeople();
  });
  conn.on('data', d=> onData(conn.peer, d));
  conn.on('close', ()=> dropPeer(conn.peer));
  conn.on('error', ()=>{});
}

function bindCall(call, email){
  const p = ensurePeer(call.peer, email || call.metadata?.email, call.metadata?.name);
  p.call = call;
  monitorIce(p, call);
  call.on('stream', stream=> attachRemoteAudio(call.peer, stream));
  call.on('close', ()=> dropPeer(call.peer));
  call.on('error', ()=>{});
}

// Verbindungs-Status (Ampel) aus dem ICE-Zustand der Voice-Verbindung ableiten
function monitorIce(p, call){
  const pc = call.peerConnection; if(!pc || !pc.addEventListener) return;
  const upd = ()=>{
    const s = pc.iceConnectionState;
    if(s==='connected' || s==='completed') p.status='connected';
    else if(s==='failed' || s==='disconnected' || s==='closed') p.status='failed';
    else if(p.status!=='connected') p.status='connecting';
    renderPeople();
  };
  pc.addEventListener('iceconnectionstatechange', upd); upd();
}

function attachRemoteAudio(pid, stream){
  const p = peers.get(pid); if(!p) return;
  const THREE = window.THREE;

  // 1) Pipeline „anstoßen" (Chrome-Workaround: stummes <audio>-Element)
  if(!p.audioEl){
    const a = document.createElement('audio'); a.srcObject = stream; a.muted = true; a.autoplay = true;
    a.play().catch(()=>{}); p.audioEl = a;
  }
  // 2) Räumliches Audio am Avatar
  if(p.posAudio){ try{ p.posAudio.disconnect(); }catch{} p.avatar.remove(p.posAudio); }
  const pa = new THREE.PositionalAudio(listener);
  pa.setMediaStreamSource(stream);
  pa.setRefDistance(2.2); pa.setRolloffFactor(1.6); pa.setDistanceModel('exponential'); pa.setMaxDistance(35);
  pa.position.y = 1.3; p.avatar.add(pa); p.posAudio = pa;
  // 3) Analyser für Sprech-Anzeige
  try{
    const src = actx.createMediaStreamSource(stream);
    const an = actx.createAnalyser(); an.fftSize = 512; src.connect(an);
    p.analyser = an;
  }catch{}
  // Audio empfangen = Medienweg funktioniert → Ampel grün
  p.status = 'connected'; renderPeople();
}

function dropPeer(pid){
  const p = peers.get(pid); if(!p) return;
  try{ p.posAudio?.disconnect(); }catch{}
  try{ p.call?.close(); }catch{}
  try{ p.dataConn?.close(); }catch{}
  try{ p.screenOut?.close(); }catch{}
  try{ p.screenIn?.close(); }catch{}
  if(p.audioEl){ p.audioEl.srcObject = null; }
  if(p.avatar) scene.remove(p.avatar);
  if(currentPresenter===pid) clearWall();
  peers.delete(pid);
  renderPeople();
}

// ── Daten-Protokoll ───────────────────────────────────────────────
function onData(pid, d){
  const p = peers.get(pid); if(!p || !d) return;
  if(d.email) p.email = d.email;
  if(d.name && d.name!==p.name){ p.name = d.name; updateLabel(p, d.name); }
  if(d.t==='bye'){ dropPeer(pid); return; }
  if(d.t==='chat'){ appendChat(d.name || p.name || 'Gast', d.text, false); return; }
  if(d.t==='screen'){ p.sharing = !!d.on; if(!d.on && currentPresenter===pid) clearWall(); updatePresenterBanner(); renderPeople(); return; }
  if(typeof d.x==='number'){ p.tx=d.x; p.tz=d.z; p.tRy=d.ry; }
  if('muted' in d){ p.muted = d.muted; }
  if('hand' in d){ p.hand = d.hand; if(p.avatar) p.avatar.userData.hand.visible = d.hand; }
  renderPeople();
}
function updateLabel(p, name){
  if(!p.avatar) return;
  const old = p.avatar.userData.label;
  const nl = makeLabel(window.THREE, name, `hsl(${hueOf(p.email||name)},70%,60%)`);
  p.avatar.remove(old); p.avatar.add(nl); p.avatar.userData.label = nl;
}

function myState(){
  return { t:'pos', email:myEmail, name:myName,
           x:+myAvatar.position.x.toFixed(2), z:+myAvatar.position.z.toFixed(2),
           ry:+myAvatar.rotation.y.toFixed(2), muted:!isMicLive(), hand:handUp };
}
function sendState(conn){ try{ if(conn?.open) conn.send(myState()); }catch{} }
function broadcast(extra){
  const msg = myAvatar ? Object.assign(myState(), extra||{}) : (extra||{});
  peers.forEach(p=>{ try{ if(p.dataConn?.open) p.dataConn.send(msg); }catch{} });
}
function heartbeat(){ if(myAvatar) broadcast(); }

let _bcT=0;
function scheduleBroadcast(){
  const now = performance.now();
  if(now - _bcT > 90){ _bcT = now; broadcast(); }
}

// ════════════════════════════════════════════════════════════════
// HUD-STEUERUNG
// ════════════════════════════════════════════════════════════════
// ── Mikrofon / Push-to-Talk ──────────────────────────────────────
function isMicLive(){ if(!hasMic) return false; return pttMode ? pttActive : micEnabled; }
function updateMicLive(){ if(localStream && hasMic) localStream.getAudioTracks().forEach(t=> t.enabled = isMicLive()); }

function toggleMic(){
  if(!hasMic){ toast('Kein Mikrofon verfügbar'); return; }
  if(pttMode){ toast('Push-to-Talk aktiv – Leertaste halten zum Sprechen'); return; }
  micEnabled = !micEnabled;
  updateMicLive();
  setMicUI(micEnabled);
  broadcast();
}
function setMicUI(on){
  const b=$id('btn-mic');
  b.classList.remove('ptt');
  b.classList.toggle('on', on); b.classList.toggle('off', !on);
  $id('mic-ic').textContent = on ? '🎤' : '🔇';
  $id('mic-lbl').textContent = on ? 'Mikro an' : 'Stumm';
}
function togglePTT(){
  if(!hasMic){ toast('Kein Mikrofon verfügbar'); return; }
  pttMode = !pttMode;
  $id('btn-ptt').classList.toggle('on', pttMode);
  if(pttMode){
    micEnabled = false;
    const b=$id('btn-mic'); b.classList.remove('on','off'); b.classList.add('ptt');
    $id('mic-ic').textContent='🎙️'; $id('mic-lbl').textContent='PTT';
    toast('Push-to-Talk: Leertaste gedrückt halten zum Sprechen', 4000);
  }else{
    micEnabled = true; setMicUI(true);
  }
  updateMicLive();
  broadcast();
}

// ── Mikrofon-Auswahl ─────────────────────────────────────────────
async function toggleMicMenu(){
  const menu = $id('mic-menu');
  if(menu.style.display==='block'){ menu.style.display='none'; return; }
  let devs=[];
  try{ devs = (await navigator.mediaDevices.enumerateDevices()).filter(d=>d.kind==='audioinput'); }catch{}
  if(!devs.length){ toast('Keine Mikrofone gefunden'); return; }
  menu.innerHTML = devs.map((d,i)=>
    `<div class="mic-opt${d.deviceId===currentMicId?' sel':''}" onclick="selectMic('${d.deviceId}')">${esc(d.label || ('Mikrofon '+(i+1)))}</div>`).join('');
  menu.style.display='block';
}
async function selectMic(id){
  $id('mic-menu').style.display='none';
  try{
    const ns = await navigator.mediaDevices.getUserMedia({ audio:{ deviceId:{ exact:id }, echoCancellation:true, noiseSuppression:true, autoGainControl:true } });
    const newTrack = ns.getAudioTracks()[0];
    // Track in allen Voice-Calls ersetzen (ohne Renegotiation)
    peers.forEach(p=>{
      const pc = p.call?.peerConnection;
      const s = pc?.getSenders?.().find(x=> x.track && x.track.kind==='audio');
      if(s) s.replaceTrack(newTrack).catch(()=>{});
    });
    // localStream + Analyser aktualisieren
    localStream.getAudioTracks().forEach(t=>{ t.stop(); localStream.removeTrack(t); });
    localStream.addTrack(newTrack);
    if(!hasMic){ hasMic = true; micEnabled = true; }   // war vorher „nur zuhören"
    newTrack.enabled = isMicLive();
    currentMicId = id;
    try{
      const src = actx.createMediaStreamSource(new MediaStream([newTrack]));
      localAnalyser = actx.createAnalyser(); localAnalyser.fftSize = 512; src.connect(localAnalyser);
    }catch{}
    if(!pttMode) setMicUI(micEnabled);
    broadcast();
    toast('🎤 Mikrofon gewechselt');
  }catch(e){ toast('Mikrofon-Wechsel fehlgeschlagen'); }
}

// ── Bildschirm teilen ────────────────────────────────────────────
async function toggleScreenShare(){
  if(screenStream){ stopScreenShare(); return; }
  try{
    screenStream = await navigator.mediaDevices.getDisplayMedia({ video:{ frameRate:15 }, audio:true });
  }catch(e){ toast('Bildschirm teilen abgebrochen'); return; }
  screenStream.getVideoTracks()[0].addEventListener('ended', stopScreenShare); // „Stoppen" im Browser-Dialog
  peers.forEach((p,pid)=> callScreen(pid));
  showScreenOnWall('me', myName, screenStream);
  broadcast({ t:'screen', on:true });
  $id('btn-screen').classList.add('on');
  $id('screen-lbl').textContent = 'Stoppen';
  toast('🖥 Du teilst deinen Bildschirm');
}
function stopScreenShare(){
  if(!screenStream) return;
  try{ screenStream.getTracks().forEach(t=>t.stop()); }catch{}
  screenStream = null;
  peers.forEach(p=>{ try{ p.screenOut?.close(); }catch{} p.screenOut=null; });
  if(currentPresenter==='me') clearWall();
  broadcast({ t:'screen', on:false });
  $id('btn-screen').classList.remove('on');
  $id('screen-lbl').textContent = 'Teilen';
}
function callScreen(pid){
  if(!screenStream || !peer) return;
  const p = peers.get(pid); if(!p) return;
  try{ p.screenOut?.close(); }catch{}
  p.screenOut = peer.call(pid, screenStream, { metadata:{ kind:'screen', email:myEmail, name:myName } });
}
function acceptScreenCall(call){
  call.answer();   // wir senden nichts zurück, nur empfangen
  const p = peers.get(call.peer); if(p) p.screenIn = call;
  const meta = call.metadata || {};
  call.on('stream', stream=> showScreenOnWall(call.peer, meta.name || nameFromEmail(meta.email||'Gast'), stream));
  call.on('close', ()=>{ if(currentPresenter===call.peer) clearWall(); });
  call.on('error', ()=>{});
}
function showScreenOnWall(who, name, stream){
  const THREE = window.THREE;
  let v = $id('screen-video');
  if(!v){ v=document.createElement('video'); v.id='screen-video'; v.autoplay=true; v.playsInline=true; v.style.display='none'; document.body.appendChild(v); }
  v.srcObject = stream; v.muted = (who==='me'); v.play().catch(()=>{});
  if(wallScreen){
    const tex = new THREE.VideoTexture(v); tex.colorSpace = THREE.SRGBColorSpace;
    wallScreen.material.map = tex; wallScreen.material.emissiveMap = tex;
    wallScreen.material.emissiveIntensity = 0.95; wallScreen.material.needsUpdate = true;
  }
  currentPresenter = who; _presenterName = name;
  updatePresenterBanner();
}
function clearWall(){
  const v = $id('screen-video'); if(v) v.srcObject = null;
  if(wallScreen){
    wallScreen.material.map = wallDefaultMap; wallScreen.material.emissiveMap = wallDefaultMap;
    wallScreen.material.emissiveIntensity = 0.5; wallScreen.material.needsUpdate = true;
  }
  currentPresenter = null; _presenterName = '';
  updatePresenterBanner();
}
function updatePresenterBanner(){
  const b = $id('presenter-banner'); if(!b) return;
  if(currentPresenter){ b.style.display='flex'; $id('presenter-name').textContent = currentPresenter==='me' ? 'Du' : _presenterName; }
  else b.style.display='none';
}

// ── Text-Chat ────────────────────────────────────────────────────
function toggleChat(){
  _chatOpen = !_chatOpen;
  $id('chat-panel').style.display = _chatOpen ? 'flex' : 'none';
  $id('btn-chat').classList.toggle('on', _chatOpen);
  if(_chatOpen){ _chatUnread=0; updateChatBadge(); $id('chat-input').focus(); scrollChat(); }
}
function chatKey(e){ if(e.key==='Enter'){ e.preventDefault(); sendChat(); } }
function sendChat(){
  const inp = $id('chat-input'); const text = (inp.value||'').trim(); if(!text) return;
  inp.value='';
  appendChat(myName, text, true);
  peers.forEach(p=>{ try{ if(p.dataConn?.open) p.dataConn.send({ t:'chat', name:myName, text }); }catch{} });
}
function appendChat(name, text, isMe){
  const box = $id('chat-msgs'); if(!box) return;
  const el = document.createElement('div'); el.className = 'chat-msg' + (isMe?' me':'');
  el.innerHTML = `<span class="cm-name">${esc(name)}</span><span class="cm-text">${esc(text)}</span>`;
  box.appendChild(el); scrollChat();
  if(!_chatOpen && !isMe){ _chatUnread++; updateChatBadge(); toast(`💬 ${name}: ${text.slice(0,42)}`); }
}
function scrollChat(){ const b=$id('chat-msgs'); if(b) b.scrollTop = b.scrollHeight; }
function updateChatBadge(){ const b=$id('chat-badge'); if(!b) return; if(_chatUnread>0){ b.style.display='flex'; b.textContent = _chatUnread>9?'9+':_chatUnread; } else b.style.display='none'; }

// ── Performance-Schalter ─────────────────────────────────────────
function toggleQuality(){
  lowPerf = !lowPerf;
  $id('btn-quality').classList.toggle('on', lowPerf);
  if(lowPerf){
    renderer.setPixelRatio(1);
    renderer.shadowMap.enabled = false;
    scene.traverse(o=>{ if(o.isLight) o.castShadow = false; });
    $id('quality-lbl').textContent = 'Sparen';
    toast('Performance-Sparmodus an');
  }else{
    renderer.setPixelRatio(Math.min(devicePixelRatio, 2));
    renderer.shadowMap.enabled = true;
    const sun = scene.children.find(o=>o.isDirectionalLight); if(sun) sun.castShadow = true;
    $id('quality-lbl').textContent = 'Qualität';
    toast('Volle Qualität');
  }
  scene.traverse(o=>{ if(o.material){ (Array.isArray(o.material)?o.material:[o.material]).forEach(m=> m.needsUpdate=true); } });
}
function toggleHand(){
  handUp = !handUp;
  $id('btn-hand').classList.toggle('on', handUp);
  if(myAvatar) myAvatar.userData.hand.visible = handUp;
  broadcast();
  if(handUp) toast('✋ Du hast dich gemeldet');
}
function cycleView(){
  viewMode = viewMode==='follow' ? 'overview' : 'follow';
  $id('view-lbl').textContent = viewMode==='follow' ? 'Folgen' : 'Übersicht';
  if(viewMode==='overview'){
    controls.target.set(0,1,0); controls.maxDistance = 16; camera.position.set(-9,8,10);
  }else{
    controls.maxDistance = 9;
  }
}
function leaveRoom(){
  broadcast({ t:'bye' });
  try{ peer?.destroy(); }catch{}
  try{ localStream?.getTracks().forEach(t=>t.stop()); }catch{}
  location.reload();
}
function toggleFullscreen(){
  if(document.fullscreenElement) document.exitFullscreen();
  else document.documentElement.requestFullscreen?.().catch(()=>{});
}

function renderPeople(){
  const list = $id('people-list'); if(!list) return;
  $id('people-count').textContent = peers.size + 1;
  let html = personRow(myEmail, myName, true, isMicLive()||(!pttMode&&micEnabled), handUp, mySpeak, 'connected', currentPresenter==='me');
  peers.forEach((p,pid)=>{
    const spk = p.avatar?.userData?.ring?.visible;
    html += personRow(p.email||p.name, p.name||nameFromEmail(p.email||'Gast'), false, !p.muted, p.hand, spk, p.status||'connecting', currentPresenter===pid);
  });
  list.innerHTML = html;
}
function personRow(email, name, isMe, micOn, hand, speaking, status, sharing){
  const dotClass = isMe ? 'ok' : (status==='connected' ? 'ok' : status==='failed' ? 'bad' : 'wait');
  const statusTxt = (!isMe && status==='failed') ? 'keine Verbindung'
                  : (!isMe && status!=='connected') ? 'verbindet…'
                  : (speaking ? 'spricht…' : (micOn ? 'verbunden' : 'stumm'));
  const rightIc = sharing ? '🖥' : (hand ? '✋' : (micOn ? '🎙️' : '🔇'));
  return `<div class="person">
    <div class="p-av ${speaking?'speaking':''}" style="background:${cssColor(email)}">${esc(initials(name))}<span class="p-dot ${dotClass}"></span></div>
    <div class="p-info">
      <div class="p-name ${isMe?'is-me':''}">${esc(name)}${isMe?' (du)':''}</div>
      <div class="p-status">${statusTxt}</div>
    </div>
    <div class="p-mic ${micOn?'':'muted'}">${rightIc}</div>
  </div>`;
}

// ════════════════════════════════════════════════════════════════
// START
// ════════════════════════════════════════════════════════════════
(async function boot(){
  try{
    const ok = await initAuth();
    if(ok) await afterLogin();
    else { $id('boot-sub').textContent=''; $id('boot-spinner').style.display='none'; $id('boot-btn').style.display='block'; }
  }catch(e){
    $id('boot-spinner').style.display='none';
    $id('boot-err').textContent = e.message; $id('boot-err').style.display='block';
    $id('boot-btn').style.display='block';
  }
})();
