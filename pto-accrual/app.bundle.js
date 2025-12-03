/* Prairie Forge PTO Accrual */
(()=>{function q(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}var Ue="SS_PF_Config";async function ht(e,t=[Ue]){var o;let n=e.workbook.tables;n.load("items/name"),await e.sync();let a=(o=n.items)==null?void 0:o.find(s=>t.includes(s.name));return a?e.workbook.tables.getItem(a.name):(console.warn("Config table not found. Looking for:",t),null)}function yt(e){let t=e.map(n=>String(n||"").trim().toLowerCase());return{field:t.findIndex(n=>n==="field"||n==="field name"||n==="setting"),value:t.findIndex(n=>n==="value"||n==="setting value"),type:t.findIndex(n=>n==="type"||n==="category"),title:t.findIndex(n=>n==="title"||n==="display name"),permanent:t.findIndex(n=>n==="permanent"||n==="persist")}}async function vt(e=[Ue]){if(!q())return{};try{return await Excel.run(async t=>{let n=await ht(t,e);if(!n)return{};let a=n.getDataBodyRange(),o=n.getHeaderRowRange();a.load("values"),o.load("values"),await t.sync();let s=o.values[0]||[],l=yt(s);if(l.field===-1||l.value===-1)return console.warn("Config table missing FIELD or VALUE columns. Headers:",s),{};let c={};return(a.values||[]).forEach(i=>{var f;let u=String(i[l.field]||"").trim();u&&(c[u]=(f=i[l.value])!=null?f:"")}),console.log("Configuration loaded:",Object.keys(c).length,"fields"),c})}catch(t){return console.error("Failed to load configuration:",t),{}}}async function Se(e,t,n=[Ue]){if(!q())return!1;try{return await Excel.run(async a=>{let o=await ht(a,n);if(!o){console.warn("Config table not found for write");return}let s=o.getDataBodyRange(),l=o.getHeaderRowRange();s.load("values"),l.load("values"),await a.sync();let c=l.values[0]||[],d=yt(c);if(d.field===-1||d.value===-1){console.error("Config table missing FIELD or VALUE columns");return}let u=(s.values||[]).findIndex(f=>String(f[d.field]||"").trim()===e);if(u>=0)s.getCell(u,d.value).values=[[t]];else{let f=new Array(c.length).fill("");d.type>=0&&(f[d.type]="Run Settings"),f[d.field]=e,f[d.value]=t,d.permanent>=0&&(f[d.permanent]="N"),d.title>=0&&(f[d.title]=""),o.rows.add(null,[f]),console.log("Added new config row:",e,"=",t)}await a.sync(),console.log("Saved config:",e,"=",t)}),!0}catch(a){return console.error("Failed to save config:",e,a),!1}}var dn="SS_PF_Config",un="module-prefix",Fe="system",ye={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function bt(){if(!q())return{...ye};try{return await Excel.run(async e=>{var u,f;let t=e.workbook.worksheets.getItemOrNullObject(dn);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...ye};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((u=n.values)!=null&&u.length))return{...ye};let a=n.values,o=gn(a[0]),s=o.get("category"),l=o.get("field"),c=o.get("value");if(s===void 0||l===void 0||c===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...ye};let d={},i=!1;for(let p=1;p<a.length;p++){let r=a[p];if(Te(r[s])===un){let m=String((f=r[l])!=null?f:"").trim().toUpperCase(),y=Te(r[c]);m&&y&&(d[m]=y,i=!0)}}return i?(console.log("[Tab Visibility] Loaded prefix config:",d),d):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...ye})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...ye}}}async function Ge(e){if(!q())return;let t=Te(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await bt();await Excel.run(async a=>{let o=a.workbook.worksheets;o.load("items/name,visibility"),await a.sync();let s={};for(let[p,r]of Object.entries(n))s[r]||(s[r]=[]),s[r].push(p);let l=s[t]||[],c=s[Fe]||[],d=[];for(let[p,r]of Object.entries(s))p!==t&&p!==Fe&&d.push(...r);console.log(`[Tab Visibility] Active prefixes: ${l.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${d.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${c.join(", ")}`);let i=[],u=[];o.items.forEach(p=>{let r=p.name,g=r.toUpperCase(),m=l.some(R=>g.startsWith(R)),y=d.some(R=>g.startsWith(R)),h=c.some(R=>g.startsWith(R));m?(i.push(p),console.log(`[Tab Visibility] SHOW: ${r} (matches active module prefix)`)):h?(u.push(p),console.log(`[Tab Visibility] HIDE: ${r} (system sheet)`)):y?(u.push(p),console.log(`[Tab Visibility] HIDE: ${r} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${r} (no prefix match, leaving as-is)`)});for(let p of i)p.visibility=Excel.SheetVisibility.visible;if(await a.sync(),o.items.filter(p=>p.visibility===Excel.SheetVisibility.visible).length>u.length){for(let p of u)try{p.visibility=Excel.SheetVisibility.hidden}catch(r){console.warn(`[Tab Visibility] Could not hide "${p.name}":`,r.message)}await a.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${i.length}, hid ${u.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function fn(){if(!q()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.visibility!==Excel.SheetVisibility.visible&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${a.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function pn(){if(!q()){console.log("Excel not available");return}try{let e=await bt(),t=[];for(let[n,a]of Object.entries(e))a===Fe&&t.push(n);await Excel.run(async n=>{let a=n.workbook.worksheets;a.load("items/name,visibility"),await n.sync(),a.items.forEach(o=>{let s=o.name.toUpperCase();t.some(l=>s.startsWith(l))&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${o.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function gn(e=[]){let t=new Map;return e.forEach((n,a)=>{let o=Te(n);o&&t.set(o,a)}),t}function Te(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=fn,window.PrairieForge.unhideSystemSheets=pn,window.PrairieForge.applyModuleTabVisibility=Ge);var wt={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var St=wt.ADA_IMAGE_URL;async function xe(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async a=>{let o=a.workbook.worksheets.getItemOrNullObject(e);o.load("isNullObject, name"),await a.sync();let s;o.isNullObject?(s=a.workbook.worksheets.add(e),await a.sync(),await Ot(a,s,t,n)):(s=o,await Ot(a,s,t,n)),s.activate(),s.getRange("A1").select(),await a.sync()})}catch(a){console.error(`Error activating homepage sheet ${e}:`,a)}}async function Ot(e,t,n,a){try{let i=t.getUsedRangeOrNullObject();i.load("isNullObject"),await e.sync(),i.isNullObject||(i.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let o=[[n,""],[a,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=o;let l=t.getRange("A1:Z100");l.format.fill.color="#0f0f0f";let c=t.getRange("A1");c.format.font.bold=!0,c.format.font.size=36,c.format.font.color="#ffffff",c.format.font.name="Segoe UI Light",c.format.verticalAlignment="Center";let d=t.getRange("A2");d.format.font.size=14,d.format.font.color="#a0a0a0",d.format.font.name="Segoe UI",d.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var kt={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function Ae(e){return kt[e]||kt["module-selector"]}function Et(){qe();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${St}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",mn),e}function qe(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function mn(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${St}" alt="Ada" />
                <h2 class="pf-ada-modal__title"><span style="font-weight:400;">ask</span><span style="font-weight:700;">ADA</span></h2>
                <p class="pf-ada-modal__subtitle">Your AI-powered assistant</p>
            </div>
            <div class="pf-ada-modal__body">
                <div class="pf-ada-modal__features">
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">\u{1F4AC}</div>
                        <span class="pf-ada-modal__feature-text">Ask questions about your data</span>
                    </div>
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">\u{1F4CA}</div>
                        <span class="pf-ada-modal__feature-text">Get insights and trend analysis</span>
                    </div>
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">\u{1F50D}</div>
                        <span class="pf-ada-modal__feature-text">Troubleshoot issues quickly</span>
                    </div>
                </div>
            </div>
            <div class="pf-ada-modal__footer">
                <span class="pf-ada-modal__powered-by">Powered by ChatGPT</span>
            </div>
        </div>
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",Je),t.addEventListener("click",o=>{o.target===t&&Je()});let a=o=>{o.key==="Escape"&&(Je(),document.removeEventListener("keydown",a))};document.addEventListener("keydown",a)}function Je(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var _t=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M15 21v-8a1 1 0 0 0-1-1h-4a1 1 0 0 0-1 1v8" />
        <path
            d="M3 10a2 2 0 0 1 .709-1.528l7-6a2 2 0 0 1 2.582 0l7 6A2 2 0 0 1 21 10v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"
        />
    </svg>
`.trim(),Ct=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <rect width="7" height="7" x="3" y="3" rx="1" />
        <rect width="7" height="7" x="14" y="3" rx="1" />
        <rect width="7" height="7" x="14" y="14" rx="1" />
        <rect width="7" height="7" x="3" y="14" rx="1" />
    </svg>
`.trim(),It=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="1"/>
        <circle cx="12" cy="5" r="1"/>
        <circle cx="12" cy="19" r="1"/>
    </svg>
`.trim(),We=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M12 3v18"/>
        <rect width="18" height="18" x="3" y="3" rx="2"/>
        <path d="M3 9h18"/>
        <path d="M3 15h18"/>
    </svg>
`.trim(),Rt=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/>
        <circle cx="9" cy="7" r="4"/>
        <path d="M22 21v-2a4 4 0 0 0-3-3.87"/>
        <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
    </svg>
`.trim(),Pt=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M4 19.5v-15A2.5 2.5 0 0 1 6.5 2H20v20H6.5a2.5 2.5 0 0 1 0-5H20"/>
        <path d="M8 7h6"/>
        <path d="M8 11h8"/>
    </svg>
`.trim(),hn={config:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="3" />
            <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.82 2.82l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.82-2.82l.06-.06A1.65 1.65 0 0 0 3 15a1.65 1.65 0 0 0-1.51-1H1a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 3 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.82-2.82l.06.06A1.65 1.65 0 0 0 9 3.6a1.65 1.65 0 0 0 1-1.51V2a2 2 0 0 1 4 0v.09A1.65 1.65 0 0 0 15 3.6a1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.82 2.82l-.06.06A1.65 1.65 0 0 0 21 9c0 .3.09.58.24.82.17.28.43.51.76.68.21.1.44.18.68.19H23a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1Z" />
        </svg>
    `,import:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M12 3v14" />
            <path d="m7 13 5 5 5-5" />
            <path d="M5 21h14" />
        </svg>
    `,headcount:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
            <circle cx="9" cy="7" r="4" />
            <path d="M22 21v-2a4 4 0 0 0-3-3.87" />
            <path d="M16 3.13a4 4 0 0 1 0 7.75" />
        </svg>
    `,validate:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M20 6 9 17l-5-5" />
        </svg>
    `,review:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 12h3l2-5 4 10 2-5h5" />
        </svg>
    `,journal:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M12 7c-3-1-6-1-9 0v12c3-1 6-1 9 0 3-1 6-1 9 0V7c-3-1-6-1-9 0Z" />
            <path d="M12 7v12" />
        </svg>
    `,archive:`
        <svg class="pf-icon pf-step-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <rect x="3" y="3" width="18" height="4" rx="1" />
            <path d="M5 7v11a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7" />
            <path d="M10 12h4" />
        </svg>
    `};function Tt(e){return e&&hn[e]||""}var ze=`
    <svg
        class="pf-icon pf-lock-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
        <path d="M7 11V7a5 5 0 0 1 10 0" />
    </svg>
`.trim(),Ye=`
    <svg
        class="pf-icon pf-lock-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
        <path d="M7 11V7a5 5 0 0 1 10 0v4" />
        <path d="M12 15v2" />
    </svg>
`.trim(),Ee=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M5 12l4 4 10-10" />
    </svg>
`.trim(),Ne=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <rect x="4" y="3" width="16" height="18" rx="2" />
        <rect x="8" y="7" width="8" height="3" />
        <path d="M8 14h.01" />
        <path d="M12 14h.01" />
        <path d="M16 14h.01" />
        <path d="M8 17h.01" />
        <path d="M12 17h.01" />
        <path d="M16 17h.01" />
    </svg>
`.trim(),ba=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M18 5.5 20.5 8 16 12.5 13.5 10 18 5.5Z" />
        <path d="m12 11 6-6" />
        <path d="M3 22 12 13" />
        <path d="m3 18 4 4" />
        <path d="m11 11 3 3" />
    </svg>
`.trim(),Ke=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" />
        <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 1 0 7.07 7.07l1.71-1.71" />
    </svg>
`.trim(),xt=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
        <path d="M7 10l5 5 5-5" />
        <path d="M12 15V3" />
    </svg>
`.trim(),At=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <circle cx="12" cy="12" r="10" />
        <path d="m15 9-6 6" />
        <path d="m9 9 6 6" />
    </svg>
`.trim(),Nt=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m12 19-7-7 7-7" />
        <path d="M19 12H5" />
    </svg>
`.trim(),Dt=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" />
        <path d="M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7" />
        <path d="M7 3v4a1 1 0 0 0 1 1h7" />
    </svg>
`.trim(),$t=`
    <svg
        class="pf-icon pf-nav-icon"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m12 5 7 7-7 7" />
        <path d="M5 12h14" />
    </svg>
`.trim(),wa=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="10"/>
        <path d="m9 12 2 2 4-4"/>
    </svg>
`.trim(),Oa=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="10"/>
        <line x1="12" x2="12" y1="8" y2="12"/>
        <line x1="12" x2="12.01" y1="16" y2="16"/>
    </svg>
`.trim(),ka=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>
        <path d="M12 9v4"/>
        <path d="M12 17h.01"/>
    </svg>
`.trim(),Sa=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="10"/>
        <path d="M12 16v-4"/>
        <path d="M12 8h.01"/>
    </svg>
`.trim(),Ea=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="M21.801 10A10 10 0 1 1 17 3.335"/>
        <path d="m9 11 3 3L22 4"/>
    </svg>
`.trim(),_a=`
    <svg
        class="pf-icon pf-status-icon"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <circle cx="12" cy="12" r="10"/>
        <path d="m15 9-6 6"/>
        <path d="m9 9 6 6"/>
    </svg>
`.trim(),Ca=`
    <svg
        class="pf-icon pf-mismatch-icon-svg"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m18 9-6-6-6 6"/>
        <path d="M12 3v14"/>
        <path d="M5 21h14"/>
    </svg>
`.trim(),Ia=`
    <svg
        class="pf-icon pf-mismatch-icon-svg"
        aria-hidden="true"
        focusable="false"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <path d="m6 15 6 6 6-6"/>
        <path d="M12 21V7"/>
        <path d="M5 3h14"/>
    </svg>
`.trim(),_e=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/>
        <path d="M21 3v5h-5"/>
        <path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/>
        <path d="M3 21v-5h5"/>
    </svg>
`.trim(),jt=`
    <svg
        class="pf-icon pf-action-icon"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
        aria-hidden="true"
        focusable="false"
    >
        <path d="M3 6h18"/>
        <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/>
        <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>
        <line x1="10" x2="10" y1="11" y2="17"/>
        <line x1="14" x2="14" y1="11" y2="17"/>
    </svg>
`.trim();function Ce(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function H(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function re({textareaId:e,value:t,permanentId:n,isPermanent:a,hintId:o,saveButtonId:s,isSaved:l=!1,placeholder:c="Enter notes here..."}){let d=a?Ye:ze,i=s?`<button type="button" class="pf-action-toggle pf-save-btn ${l?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${Dt}</button>`:"",u=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${a?"is-locked":""}" id="${n}" aria-pressed="${a}" title="Lock notes (retain after archive)">${d}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${Ce(c)}">${Ce(t||"")}</textarea>
                ${o?`<p class="pf-signoff-hint" id="${o}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?H(u,"Lock"):""}
                ${s?H(i,"Save"):""}
            </div>
        </article>
    `}function le({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:a,isComplete:o,saveButtonId:s,isSaved:l=!1,completeButtonId:c,subtext:d="Sign-off below. Click checkmark icon. Done."}){let i=`<button type="button" class="pf-action-toggle ${o?"is-active":""}" id="${c}" aria-pressed="${!!o}" title="Mark step complete">${Ee}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${Ce(d)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${e}" value="${Ce(t)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${n}" value="${Ce(a)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${H(i,"Done")}
            </div>
        </article>
    `}function Qe(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?Ye:ze)}function ae(e,t){e&&e.classList.toggle("is-saved",t)}function Xe(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(a=>{let o=a.getAttribute("data-save-input"),s=document.getElementById(o);if(!s)return;let l=()=>{ae(a,!1)};s.addEventListener("input",l),n.push(()=>s.removeEventListener("input",l))}),()=>n.forEach(a=>a())}function Bt(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function Lt(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var Ze={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},De={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function et(e){e.format.fill.color=Ze.fillColor,e.format.font.color=Ze.fontColor,e.format.font.bold=Ze.bold}function ce(e,t,n,a=!1){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a?De.currencyWithNegative:De.currency]]}function ve(e,t,n){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[De.number]]}function Mt(e,t,n,a=De.date){if(n<=0)return;let o=e.getRangeByIndexes(1,t,n,1);o.numberFormat=[[a]]}var yn="1.1.0",Re="pto-accrual";var ue="PTO Accrual",vn="Calculate your PTO liability, compare against last period, and generate a balanced journal entry\u2014all without leaving Excel.",bn="../module-selector/index.html",wn="pf-loader-overlay",de=["SS_PF_Config"],O={payrollProvider:"PTO_Payroll_Provider",payrollDate:"PTO_Analysis_Date",accountingPeriod:"PTO_Accounting_Period",journalEntryId:"PTO_Journal_Entry_ID",companyName:"SS_Company_Name",accountingSoftware:"SS_Accounting_Software",reviewerName:"PTO_Reviewer",validationDataBalance:"PTO_Validation_Data_Balance",validationCleanBalance:"PTO_Validation_Clean_Balance",validationDifference:"PTO_Validation_Difference",headcountRosterCount:"PTO_Headcount_Roster_Count",headcountPayrollCount:"PTO_Headcount_Payroll_Count",headcountDifference:"PTO_Headcount_Difference",journalDebitTotal:"PTO_JE_Debit_Total",journalCreditTotal:"PTO_JE_Credit_Total",journalDifference:"PTO_JE_Difference"},pe="User opted to skip the headcount review this period.",Be={0:{note:"PTO_Notes_Config",reviewer:"PTO_Reviewer_Config",signOff:"PTO_SignOff_Config"},1:{note:"PTO_Notes_Import",reviewer:"PTO_Reviewer_Import",signOff:"PTO_SignOff_Import"},2:{note:"PTO_Notes_Headcount",reviewer:"PTO_Reviewer_Headcount",signOff:"PTO_SignOff_Headcount"},3:{note:"PTO_Notes_Validate",reviewer:"PTO_Reviewer_Validate",signOff:"PTO_SignOff_Validate"},4:{note:"PTO_Notes_Review",reviewer:"PTO_Reviewer_Review",signOff:"PTO_SignOff_Review"},5:{note:"PTO_Notes_JE",reviewer:"PTO_Reviewer_JE",signOff:"PTO_SignOff_JE"},6:{note:"PTO_Notes_Archive",reviewer:"PTO_Reviewer_Archive",signOff:"PTO_SignOff_Archive"}},Zt={0:"PTO_Complete_Config",1:"PTO_Complete_Import",2:"PTO_Complete_Headcount",3:"PTO_Complete_Validate",4:"PTO_Complete_Review",5:"PTO_Complete_JE",6:"PTO_Complete_Archive"};var oe=[{id:0,title:"Configuration",summary:"Set the analysis date, accounting period, and review details for this run.",description:"Complete this step first to ensure all downstream calculations use the correct period settings.",actionLabel:"Configure Workbook",secondaryAction:{sheet:"SS_PF_Config",label:"Open Config Sheet"}},{id:1,title:"Import PTO Data",summary:"Pull your latest PTO export from payroll and paste it into PTO_Data.",description:"Open your payroll provider, download the PTO report, and paste the data into the PTO_Data tab.",actionLabel:"Import Sample Data",secondaryAction:{sheet:"PTO_Data",label:"Open Data Sheet"}},{id:2,title:"Headcount Review",summary:"Quick check to make sure your roster matches your PTO data.",description:"Compare employees in PTO_Data against your employee roster to catch any discrepancies.",actionLabel:"Open Headcount Review",secondaryAction:{sheet:"SS_Employee_Roster",label:"Open Sheet"}},{id:3,title:"Data Quality Review",summary:"Scan your PTO data for potential errors before crunching numbers.",description:"Identify negative balances, overdrawn accounts, and other anomalies that might need attention.",actionLabel:"Click to Run Quality Check"},{id:4,title:"PTO Accrual Review",summary:"Review the calculated liability for each employee and compare to last period.",description:"The analysis enriches your PTO data with pay rates and department info, then calculates the liability.",actionLabel:"Click to Perform Review"},{id:5,title:"Journal Entry Prep",summary:"Generate a balanced journal entry, run validation checks, and export when ready.",description:"Build the JE from your PTO data, verify debits equal credits, and export for upload to your accounting system.",actionLabel:"Open Journal Draft",secondaryAction:{sheet:"PTO_JE_Draft",label:"Open Sheet"}},{id:6,title:"Archive & Reset",summary:"Save this period's results and prepare for the next cycle.",description:"Archive the current analysis so it becomes the 'prior period' for your next review.",actionLabel:"Archive Run"}];var On=oe.reduce((e,t)=>(e[t.id]="pending",e),{}),D={activeView:"home",activeStepId:null,focusedIndex:0,stepStatuses:On},k={loaded:!1,steps:{},permanents:{},completes:{},values:{},overrides:{accountingPeriod:!1,journalId:!1}},Ie=null,tt=null,$e=null,be=new Map,T={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},U={debitTotal:null,creditTotal:null,difference:null,lineAmountSum:null,analysisChangeTotal:null,jeChangeTotal:null,loading:!1,lastError:null,validationRun:!1,issues:[]},J={hasRun:!1,loading:!1,acknowledged:!1,balanceIssues:[],zeroBalances:[],accrualOutliers:[],totalIssues:0,totalEmployees:0},z={cleanDataReady:!1,employeeCount:0,lastRun:null,loading:!1,lastError:null,missingPayRates:[],missingDepartments:[],ignoredMissingPayRates:new Set,completenessCheck:{accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null}};async function kn(){var e;try{Ie=document.getElementById("app"),tt=document.getElementById("loading"),await Sn(),await En(),(e=window.PrairieForge)!=null&&e.loadSharedConfig&&await window.PrairieForge.loadSharedConfig();let t=Ae(Re);await xe(t.sheetName,t.title,t.subtitle),tt&&tt.remove(),Ie&&(Ie.hidden=!1),ee()}catch(t){throw console.error("[PTO] Module initialization failed:",t),t}}async function Sn(){try{await Ge(Re),console.log(`[PTO] Tab visibility applied for ${Re}`)}catch(e){console.warn("[PTO] Could not apply tab visibility:",e)}}async function En(){var e;if(!q()){k.loaded=!0;return}try{let t=await vt(de),n={};(e=window.PrairieForge)!=null&&e.loadSharedConfig&&(await window.PrairieForge.loadSharedConfig(),window.PrairieForge._sharedConfigCache&&window.PrairieForge._sharedConfigCache.forEach((s,l)=>{n[l]=s}));let a={...t},o={SS_Default_Reviewer:O.reviewerName,Default_Reviewer:O.reviewerName,PTO_Reviewer:O.reviewerName,SS_Company_Name:O.companyName,Company_Name:O.companyName,SS_Payroll_Provider:O.payrollProvider,Payroll_Provider_Link:O.payrollProvider,SS_Accounting_Software:O.accountingSoftware,Accounting_Software_Link:O.accountingSoftware};Object.entries(o).forEach(([s,l])=>{n[s]&&!a[l]&&(a[l]=n[s])}),Object.entries(n).forEach(([s,l])=>{s.startsWith("PTO_")&&l&&(a[s]=l)}),k.permanents=await _n(),k.values=a||{},k.overrides.accountingPeriod=!!(a!=null&&a[O.accountingPeriod]),k.overrides.journalId=!!(a!=null&&a[O.journalEntryId]),Object.entries(Be).forEach(([s,l])=>{var c,d,i;k.steps[s]={notes:(c=a[l.note])!=null?c:"",reviewer:(d=a[l.reviewer])!=null?d:"",signOffDate:(i=a[l.signOff])!=null?i:""}}),k.completes=Object.entries(Zt).reduce((s,[l,c])=>{var d;return s[l]=(d=a[c])!=null?d:"",s},{}),k.loaded=!0}catch(t){console.warn("PTO: unable to load configuration fields",t),k.loaded=!0}}async function _n(){let e={};if(!q())return e;let t=new Map;Object.entries(Be).forEach(([n,a])=>{a.note&&t.set(a.note.trim(),Number(n))});try{await Excel.run(async n=>{let a=n.workbook.tables.getItemOrNullObject(de[0]);if(await n.sync(),a.isNullObject)return;let o=a.getDataBodyRange(),s=a.getHeaderRowRange();o.load("values"),s.load("values"),await n.sync();let c=(s.values[0]||[]).map(i=>String(i||"").trim().toLowerCase()),d={field:c.findIndex(i=>i==="field"||i==="field name"||i==="setting"),permanent:c.findIndex(i=>i==="permanent"||i==="persist")};d.field===-1||d.permanent===-1||(o.values||[]).forEach(i=>{let u=String(i[d.field]||"").trim(),f=t.get(u);if(f==null)return;let p=Qn(i[d.permanent]);e[f]=p})})}catch(n){console.warn("PTO: unable to load permanent flags",n)}return e}function ee(){var c;if(!Ie)return;let e=D.focusedIndex<=0?"disabled":"",t=D.focusedIndex>=oe.length-1?"disabled":"",n=D.activeView==="step"&&D.activeStepId!=null,o=D.activeView==="config"?en():n?An(D.activeStepId):`${In()}${Rn()}`;Ie.innerHTML=`
        <div class="pf-root">
            <div class="pf-brand-float" aria-hidden="true">
                <span class="pf-brand-wave"></span>
            </div>
            <header class="pf-banner">
                <div class="pf-nav-bar">
                    <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${e}>
                        ${Nt}
                        <span class="sr-only">Previous step</span>
                    </button>
                    <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                        ${_t}
                        <span class="sr-only">Module Home</span>
                    </button>
                    <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                        ${Ct}
                        <span class="sr-only">Module Selector</span>
                    </button>
                    <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${t}>
                        ${$t}
                        <span class="sr-only">Next step</span>
                    </button>
                    <span class="pf-nav-divider"></span>
                    <div class="pf-quick-access-wrapper">
                        <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                            ${It}
                            <span class="sr-only">Quick Access Menu</span>
                        </button>
                        <div id="quick-access-dropdown" class="pf-quick-dropdown hidden">
                            <div class="pf-quick-dropdown-header">Quick Access</div>
                            <button id="nav-roster" class="pf-quick-item pf-clickable" type="button">
                                ${Rt}
                                <span>Employee Roster</span>
                            </button>
                            <button id="nav-accounts" class="pf-quick-item pf-clickable" type="button">
                                ${Pt}
                                <span>Chart of Accounts</span>
                            </button>
                        </div>
                    </div>
                </div>
            </header>
            ${o}
            <footer class="pf-brand-footer">
                <div class="pf-brand-text">
                    <div class="pf-brand-label">prairie.forge</div>
                    <div class="pf-brand-meta">\xA9 Prairie Forge LLC, 2025. All rights reserved. Version ${yn}</div>
                    <button type="button" class="pf-config-link" id="showConfigSheets">CONFIGURATION</button>
                </div>
            </footer>
        </div>
    `;let s=D.activeView==="home"||D.activeView!=="step"&&D.activeView!=="config",l=document.getElementById("pf-info-fab-pto");if(s)l&&l.remove();else if((c=window.PrairieForge)!=null&&c.mountInfoFab){let d=Cn(D.activeStepId);PrairieForge.mountInfoFab({title:d.title,content:d.content,buttonId:"pf-info-fab-pto"})}Nn(),jn(),s?Et():qe()}function Cn(e){switch(e){case 0:return{title:"Configuration",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Sets up the key parameters for your PTO review. Complete this before importing data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Key Fields</h4>
                        <ul>
                            <li><strong>Analysis Date</strong> \u2014 The period-end date (e.g., 11/30/2024)</li>
                            <li><strong>Accounting Period</strong> \u2014 Shows up in your JE description</li>
                            <li><strong>Journal Entry ID</strong> \u2014 Reference number for your accounting system</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>The accounting period and JE ID auto-generate based on your analysis date, but you can override them if needed.</p>
                    </div>
                `};case 1:return{title:"Data Import",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Gets your PTO data into the workbook. Pull a report from your payroll provider and paste it into PTO_Data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Required Columns</h4>
                        <p>Your payroll export should include:</p>
                        <ul>
                            <li><strong>Employee Name</strong> \u2014 Full name (used to match against roster)</li>
                            <li><strong>Accrual Rate</strong> \u2014 Hours accrued per pay period</li>
                            <li><strong>Carry Over</strong> \u2014 Hours carried from prior year</li>
                            <li><strong>YTD Accrued</strong> \u2014 Total hours accrued this year</li>
                            <li><strong>YTD Used</strong> \u2014 Total hours used this year</li>
                            <li><strong>Balance</strong> \u2014 Current available hours</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Column headers don't need to match exactly\u2014the system is flexible with naming. Just make sure each field is present.</p>
                    </div>
                `};case 2:return{title:"Headcount Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Compares employee counts between your roster and PTO data to catch discrepancies early.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Data Sources</h4>
                        <ul>
                            <li><strong>SS_Employee_Roster</strong> \u2014 Your centralized employee list</li>
                            <li><strong>PTO_Data</strong> \u2014 The data you just imported</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F50D} What to Look For</h4>
                        <ul>
                            <li><strong>In Roster, Not in PTO</strong> \u2014 May need to add PTO records</li>
                            <li><strong>In PTO, Not in Roster</strong> \u2014 Could be terminated employees</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>If discrepancies are expected (e.g., contractors without PTO), you can skip this check.</p>
                    </div>
                `};case 3:return{title:"Data Quality Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Scans your PTO data for anomalies that could cause problems later in the process.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u26A0\uFE0F Balance Issues (Critical)</h4>
                        <p>Flags when:</p>
                        <ul>
                            <li><strong>Negative Balance</strong> \u2014 Balance is less than zero</li>
                            <li><strong>Overdrawn</strong> \u2014 Used more than available (YTD Used > Carry Over + YTD Accrued)</li>
                        </ul>
                        <p class="pf-info-note">Usually indicates missing accrual entries or data errors in payroll.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} High Accrual Rates (Warning)</h4>
                        <p>Employees with Accrual Rate > 8 hours/period may have data entry errors.</p>
                        <p class="pf-info-note">Most bi-weekly accruals are 3-6 hours.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>You can acknowledge issues and proceed, but it's best to fix them in your source system first.</p>
                    </div>
                `};case 4:return{title:"PTO Accrual Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Calculates the PTO liability for each employee and compares it to last period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Data Sources</h4>
                        <ul>
                            <li><strong>PTO_Data</strong> \u2014 Your imported PTO balances</li>
                            <li><strong>SS_Employee_Roster</strong> \u2014 Department assignments</li>
                            <li><strong>PR_Archive_Summary</strong> \u2014 Pay rates from payroll history</li>
                            <li><strong>PTO_Archive_Summary</strong> \u2014 Last period's liability (for comparison)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4B0} How Liability is Calculated</h4>
                        <div class="pf-info-formula">
                            Liability = Balance (hours) \xD7 Hourly Rate
                        </div>
                        <p class="pf-info-note">Hourly rate comes from Regular Earnings \xF7 80 hours in your payroll history.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4C8} How Change is Calculated</h4>
                        <div class="pf-info-formula">
                            Change = Current Liability \u2212 Prior Period Liability
                        </div>
                        <ul>
                            <li><span style="color: #30d158;">Positive</span> = Liability went up (book expense)</li>
                            <li><span style="color: #ff453a;">Negative</span> = Liability went down (reverse expense)</li>
                        </ul>
                    </div>
                `};case 5:return{title:"Journal Entry Prep",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Generates a balanced journal entry from your PTO analysis, ready for upload to your accounting system.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4DD} How the JE Works</h4>
                        <p>Groups the <strong>Change</strong> amounts by department:</p>
                        <ul>
                            <li><span style="color: #30d158;">Positive Change</span> \u2192 Debit expense account</li>
                            <li><span style="color: #ff453a;">Negative Change</span> \u2192 Credit expense account</li>
                        </ul>
                        <p>The offset always goes to <strong>21540</strong> (Accrued PTO liability).</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F3E2} Department \u2192 Account Mapping</h4>
                        <table style="width:100%; font-size: 12px; margin-top: 8px;">
                            <tr><td>General & Admin</td><td style="text-align:right">64110</td></tr>
                            <tr><td>R&D</td><td style="text-align:right">62110</td></tr>
                            <tr><td>Marketing</td><td style="text-align:right">61610</td></tr>
                            <tr><td>Sales & Marketing</td><td style="text-align:right">61110</td></tr>
                            <tr><td>COGS Onboarding</td><td style="text-align:right">53110</td></tr>
                            <tr><td>COGS Prof. Services</td><td style="text-align:right">56110</td></tr>
                            <tr><td>COGS Support</td><td style="text-align:right">52110</td></tr>
                            <tr><td>Client Success</td><td style="text-align:right">61811</td></tr>
                        </table>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u2705 Validation Checks</h4>
                        <ul>
                            <li><strong>Debits = Credits</strong> \u2014 Entry must balance</li>
                            <li><strong>Line Amounts = $0</strong> \u2014 Net change must be zero</li>
                            <li><strong>JE Matches Analysis</strong> \u2014 Totals tie back to your data</li>
                        </ul>
                    </div>
                `};case 6:return{title:"Archive & Reset",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Saves this period's results so they become the "prior period" for your next review.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4E6} What Gets Saved</h4>
                        <ul>
                            <li><strong>PTO_Archive_Summary</strong> \u2014 Employee name, liability amount, and analysis date</li>
                            <li>This data is used to calculate the "Change" column next period</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u26A0\uFE0F Important</h4>
                        <p>Only the <strong>most recent period</strong> is kept in the archive. Running archive again will overwrite the previous data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Make sure your JE has been uploaded to your accounting system before archiving.</p>
                    </div>
                `};default:return{title:"PTO Accrual",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F44B} Welcome to PTO Accrual</h4>
                        <p>This module helps you calculate PTO liabilities and generate journal entries each period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Workflow Overview</h4>
                        <ol style="margin: 8px 0; padding-left: 20px;">
                            <li>Configure your period settings</li>
                            <li>Import PTO data from payroll</li>
                            <li>Review headcount alignment</li>
                            <li>Check data quality</li>
                            <li>Review calculated liabilities</li>
                            <li>Generate and export journal entry</li>
                            <li>Archive for next period</li>
                        </ol>
                    </div>
                    <div class="pf-info-section">
                        <p>Click a step card to get started, or tap the <strong>\u24D8</strong> button on any step for detailed guidance.</p>
                    </div>
                `}}}function In(){return`
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">PTO Accrual</h2>
            <p class="pf-hero-copy">${vn}</p>
        </section>
    `}function Rn(){return`
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${oe.map((e,t)=>Pn(e,t)).join("")}
            </div>
        </section>
    `}function Pn(e,t){let n=D.stepStatuses[e.id]||"pending",a=D.activeView==="step"&&D.focusedIndex===t?"pf-step-card--active":"",o=Tt(qn(e.id));return`
        <article class="pf-step-card pf-clickable ${a}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${e.title}</h3>
        </article>
    `}function Tn(e){let t=oe.filter(o=>o.id!==6).map(o=>({id:o.id,title:o.title,complete:Bn(o.id)})),n=t.every(o=>o.complete),a=t.map(o=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${o.complete?"is-active":""}" aria-pressed="${o.complete}">
                        ${Ee}
                    </span>
                    <div>
                        <h3>${w(o.title)}</h3>
                        <p class="pf-config-subtext">${o.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${a}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Archive & Reset</h3>
                    <p class="pf-config-subtext">Only enabled when all steps above are complete.</p>
                </div>
                <div class="pf-pill-row pf-config-actions">
                    <button type="button" class="pf-pill-btn" id="archive-run-btn" ${n?"":"disabled"}>Archive</button>
                </div>
            </article>
        </section>
    `}function en(){if(!k.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=Wt(Z(O.payrollDate)),t=Wt(Z(O.accountingPeriod)),n=Z(O.journalEntryId),a=Z(O.accountingSoftware),o=Z(O.payrollProvider),s=Z(O.companyName),l=Z(O.reviewerName),c=he(0),d=!!k.permanents[0],i=!!(rn(k.completes[0])||c.signOffDate),u=me(c==null?void 0:c.reviewer),f=(c==null?void 0:c.signOffDate)||"";return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${w(ue)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every PTO run.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Period Data</h3>
                    <p class="pf-config-subtext">Fields in this section may change each period.</p>
                </div>
                <div class="pf-config-grid">
                    <label class="pf-config-field">
                        <span>Your Name (Used for sign-offs)</span>
                        <input type="text" id="config-user-name" value="${w(l)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>PTO Analysis Date</span>
                        <input type="date" id="config-payroll-date" value="${w(e)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${w(t)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-journal-id" value="${w(n)}" placeholder="PTO-AUTO-YYYY-MM-DD">
                    </label>
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Static Data</h3>
                    <p class="pf-config-subtext">Fields rarely change but should be reviewed.</p>
                </div>
                <div class="pf-config-grid">
                    <label class="pf-config-field">
                        <span>Company Name</span>
                        <input type="text" id="config-company-name" value="${w(s)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${w(o)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${w(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${re({textareaId:"config-notes",value:c.notes||"",permanentId:"config-notes-lock",isPermanent:d,hintId:"",saveButtonId:"config-notes-save"})}
            ${le({reviewerInputId:"config-reviewer",reviewerValue:u,signoffInputId:"config-signoff-date",signoffValue:f,isComplete:i,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"})}
        </section>
    `}function xn(e){let t=he(1),n=!!k.permanents[1],a=me(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(we(k.completes[1])||o),l=Z(O.payrollProvider);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Access your payroll provider to download the latest PTO export, then paste into PTO_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${H(l?`<a href="${w(l)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${Ke}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${Ke}</button>`,"Provider")}
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data sheet">${We}</button>`,"PTO_Data")}
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data to start over">${jt}</button>`,"Clear")}
                </div>
            </article>
            ${re({textareaId:"step-notes-1",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-1",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-1"})}
            ${le({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
        </section>
    `}function An(e){let t=oe.find(c=>c.id===e);if(!t)return"";if(e===0)return en();if(e===1)return xn(t);if(e===2)return ta(t);if(e===3)return aa(t);if(e===4)return oa(t);if(e===5)return sa(t);if(t.id===6)return Tn(t);let n=he(e),a=!!k.permanents[e],o=me(n==null?void 0:n.reviewer),s=(n==null?void 0:n.signOffDate)||"",l=!!(we(k.completes[e])||s);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${w(t.title)}</h2>
            <p class="pf-hero-copy">${w(t.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${re({textareaId:`step-notes-${e}`,value:(n==null?void 0:n.notes)||"",permanentId:`step-notes-lock-${e}`,isPermanent:a,hintId:"",saveButtonId:`step-notes-save-${e}`})}
            ${le({reviewerInputId:`step-reviewer-${e}`,reviewerValue:o,signoffInputId:`step-signoff-${e}`,signoffValue:s,isComplete:l,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`})}
        </section>
    `}function Nn(){var n,a,o,s,l,c,d;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",async()=>{var u;let i=Ae(Re);await xe(i.sheetName,i.title,i.subtitle),Pe({activeView:"home",activeStepId:null}),(u=document.getElementById("pf-hero"))==null||u.scrollIntoView({behavior:"smooth",block:"start"})}),(a=document.getElementById("nav-selector"))==null||a.addEventListener("click",()=>{window.location.href=bn}),(o=document.getElementById("nav-prev"))==null||o.addEventListener("click",()=>Vt(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>Vt(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",i=>{i.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",i=>{!(t!=null&&t.contains(i.target))&&!(e!=null&&e.contains(i.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(l=document.getElementById("nav-roster"))==null||l.addEventListener("click",()=>{Jt("SS_Employee_Roster"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(c=document.getElementById("nav-accounts"))==null||c.addEventListener("click",()=>{Jt("SS_Chart_of_Accounts"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(d=document.getElementById("showConfigSheets"))==null||d.addEventListener("click",async()=>{await Jn()}),document.querySelectorAll("[data-step-card]").forEach(i=>{let u=Number(i.getAttribute("data-step-index")),f=Number(i.getAttribute("data-step-id"));i.addEventListener("click",()=>Le(u,f))}),D.activeView==="config"?$n():D.activeView==="step"&&D.activeStepId!=null&&Dn(D.activeStepId)}function Dn(e){var u,f,p,r,g,m,y,h,R,I,P,C,x,j,$,A,v;let t=e===2?document.getElementById("step-notes-input"):document.getElementById(`step-notes-${e}`),n=e===2?document.getElementById("step-reviewer-name"):document.getElementById(`step-reviewer-${e}`),a=e===2?document.getElementById("step-signoff-date"):document.getElementById(`step-signoff-${e}`),o=document.getElementById("step-back-btn"),s=e===2?document.getElementById("step-notes-lock-2"):document.getElementById(`step-notes-lock-${e}`),l=e===2?document.getElementById("step-notes-save-2"):document.getElementById(`step-notes-save-${e}`);l==null||l.addEventListener("click",async()=>{let b=(t==null?void 0:t.value)||"";await K(e,"notes",b),ae(l,!0)});let c=e===2?document.getElementById("headcount-signoff-save"):document.getElementById(`step-signoff-save-${e}`);c==null||c.addEventListener("click",async()=>{let b=(n==null?void 0:n.value)||"";await K(e,"reviewer",b),ae(c,!0)}),Xe();let d=e===2?"headcount-signoff-toggle":`step-signoff-toggle-${e}`,i=e===2?"step-signoff-date":`step-signoff-${e}`;sn(e,{buttonId:d,inputId:i,canActivate:e===2?()=>{var S;return!it()||((S=document.getElementById("step-notes-input"))==null?void 0:S.value.trim())||""?!0:(window.alert("Please enter a brief explanation of the headcount differences before completing this step."),!1)}:null,onComplete:e===2?ua:null}),o==null||o.addEventListener("click",async()=>{let b=Ae(Re);await xe(b.sheetName,b.title,b.subtitle),Pe({activeView:"home",activeStepId:null})}),s==null||s.addEventListener("click",async()=>{let b=!s.classList.contains("is-locked");Qe(s,b),await an(e,b)}),e===6&&((u=document.getElementById("archive-run-btn"))==null||u.addEventListener("click",()=>{})),e===1&&((f=document.getElementById("import-open-data-btn"))==null||f.addEventListener("click",()=>je("PTO_Data")),(p=document.getElementById("import-clear-btn"))==null||p.addEventListener("click",()=>Gn())),e===2&&((r=document.getElementById("headcount-skip-btn"))==null||r.addEventListener("click",()=>{T.skipAnalysis=!T.skipAnalysis;let b=document.getElementById("headcount-skip-btn");b==null||b.classList.toggle("is-active",T.skipAnalysis),T.skipAnalysis&&Xt(),Qt()}),(g=document.getElementById("headcount-run-btn"))==null||g.addEventListener("click",()=>at()),(m=document.getElementById("headcount-refresh-btn"))==null||m.addEventListener("click",()=>at()),da(),T.skipAnalysis&&Xt(),Qt()),e===3&&((y=document.getElementById("quality-run-btn"))==null||y.addEventListener("click",()=>Ut()),(h=document.getElementById("quality-refresh-btn"))==null||h.addEventListener("click",()=>Ut()),(R=document.getElementById("quality-acknowledge-btn"))==null||R.addEventListener("click",()=>Mn())),e===4&&((I=document.getElementById("analysis-refresh-btn"))==null||I.addEventListener("click",()=>Ft()),(P=document.getElementById("analysis-run-btn"))==null||P.addEventListener("click",()=>Ft()),(C=document.getElementById("payrate-save-btn"))==null||C.addEventListener("click",Ht),(x=document.getElementById("payrate-ignore-btn"))==null||x.addEventListener("click",Ln),(j=document.getElementById("payrate-input"))==null||j.addEventListener("keydown",b=>{b.key==="Enter"&&Ht()})),e===5&&(($=document.getElementById("je-create-btn"))==null||$.addEventListener("click",()=>Un()),(A=document.getElementById("je-run-btn"))==null||A.addEventListener("click",()=>tn()),(v=document.getElementById("je-export-btn"))==null||v.addEventListener("click",()=>Fn()))}function $n(){var d,i,u,f,p;let e=document.getElementById("config-payroll-date");e==null||e.addEventListener("change",r=>{let g=r.target.value||"";if(X(O.payrollDate,g),!!g){if(!k.overrides.accountingPeriod){let m=Yn(g);if(m){let y=document.getElementById("config-accounting-period");y&&(y.value=m),X(O.accountingPeriod,m)}}if(!k.overrides.journalId){let m=Kn(g);if(m){let y=document.getElementById("config-journal-id");y&&(y.value=m),X(O.journalEntryId,m)}}}});let t=document.getElementById("config-accounting-period");t==null||t.addEventListener("change",r=>{k.overrides.accountingPeriod=!!r.target.value,X(O.accountingPeriod,r.target.value||"")});let n=document.getElementById("config-journal-id");n==null||n.addEventListener("change",r=>{k.overrides.journalId=!!r.target.value,X(O.journalEntryId,r.target.value.trim())}),(d=document.getElementById("config-company-name"))==null||d.addEventListener("change",r=>{X(O.companyName,r.target.value.trim())}),(i=document.getElementById("config-payroll-provider"))==null||i.addEventListener("change",r=>{X(O.payrollProvider,r.target.value.trim())}),(u=document.getElementById("config-accounting-link"))==null||u.addEventListener("change",r=>{X(O.accountingSoftware,r.target.value.trim())}),(f=document.getElementById("config-user-name"))==null||f.addEventListener("change",r=>{X(O.reviewerName,r.target.value.trim())});let a=document.getElementById("config-notes");a==null||a.addEventListener("input",r=>{K(0,"notes",r.target.value)});let o=document.getElementById("config-notes-lock");o==null||o.addEventListener("click",async()=>{let r=!o.classList.contains("is-locked");Qe(o,r),await an(0,r)});let s=document.getElementById("config-notes-save");s==null||s.addEventListener("click",async()=>{a&&(await K(0,"notes",a.value),ae(s,!0))});let l=document.getElementById("config-reviewer");l==null||l.addEventListener("change",r=>{let g=r.target.value.trim();K(0,"reviewer",g);let m=document.getElementById("config-signoff-date");if(g&&m&&!m.value){let y=st();m.value=y,K(0,"signOffDate",y),on(0,!0)}}),(p=document.getElementById("config-signoff-date"))==null||p.addEventListener("change",r=>{K(0,"signOffDate",r.target.value||"")});let c=document.getElementById("config-signoff-save");c==null||c.addEventListener("click",async()=>{var m,y;let r=((m=l==null?void 0:l.value)==null?void 0:m.trim())||"",g=((y=document.getElementById("config-signoff-date"))==null?void 0:y.value)||"";await K(0,"reviewer",r),await K(0,"signOffDate",g),ae(c,!0)}),Xe(),sn(0,{buttonId:"config-signoff-toggle",inputId:"config-signoff-date",onComplete:ea})}function Le(e,t=null){if(e<0||e>=oe.length)return;$e=e;let n=t!=null?t:oe[e].id;Pe({focusedIndex:e,activeView:n===0?"config":"step",activeStepId:n}),n===1&&je("PTO_Data"),n===2&&!T.hasAnalyzed&&(ln(),at()),n===3&&je("PTO_Data"),n===5&&je("PTO_JE_Draft")}function Vt(e){let t=D.focusedIndex+e,n=Math.max(0,Math.min(oe.length-1,t));Le(n,oe[n].id)}function jn(){if($e===null)return;let e=document.querySelector(`[data-step-index="${$e}"]`);$e=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}function Bn(e){return rn(k.completes[e])}function Pe(e){e.stepStatuses&&(D.stepStatuses={...D.stepStatuses,...e.stepStatuses}),Object.assign(D,{...e,stepStatuses:D.stepStatuses}),ee()}function te(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}async function Ht(){let e=document.getElementById("payrate-input");if(!e)return;let t=parseFloat(e.value),n=e.dataset.employee,a=parseInt(e.dataset.row,10);if(isNaN(t)||t<=0){window.alert("Please enter a valid pay rate greater than 0.");return}if(!n||isNaN(a)){console.error("Missing employee data on input");return}Y(!0,"Updating pay rate...");try{await Excel.run(async o=>{let s=o.workbook.worksheets.getItem("PTO_Analysis"),l=s.getCell(a-1,3);l.values=[[t]];let c=s.getCell(a-1,8);c.load("values"),await o.sync();let i=(Number(c.values[0][0])||0)*t,u=s.getCell(a-1,9);u.values=[[i]];let f=s.getCell(a-1,10);f.load("values"),await o.sync();let p=Number(f.values[0][0])||0,r=i-p,g=s.getCell(a-1,11);g.values=[[r]],await o.sync()}),z.missingPayRates=z.missingPayRates.filter(o=>o.name!==n),Y(!1),Le(3,3)}catch(o){console.error("Failed to save pay rate:",o),window.alert(`Failed to save pay rate: ${o.message}`),Y(!1)}}function Ln(){let e=document.getElementById("payrate-input");if(!e)return;let t=e.dataset.employee;t&&(z.ignoredMissingPayRates.add(t),z.missingPayRates=z.missingPayRates.filter(n=>n.name!==t)),Le(3,3)}async function Ut(){if(!te()){window.alert("Excel is not available. Open this module inside Excel to run quality check.");return}J.loading=!0,Y(!0,"Analyzing data quality..."),ae(document.getElementById("quality-save-btn"),!1);try{await Excel.run(async t=>{var y;let a=t.workbook.worksheets.getItem("PTO_Data").getUsedRangeOrNullObject();a.load("values"),await t.sync();let o=a.isNullObject?[]:a.values||[];if(!o.length||o.length<2)throw new Error("PTO_Data is empty or has no data rows.");let s=(o[0]||[]).map(h=>F(h));console.log("[Data Quality] PTO_Data headers:",o[0]);let l=s.findIndex(h=>h==="employee name"||h==="employeename");l===-1&&(l=s.findIndex(h=>h.includes("employee")&&h.includes("name"))),l===-1&&(l=s.findIndex(h=>h==="name"||h.includes("name")&&!h.includes("company")&&!h.includes("form"))),console.log("[Data Quality] Employee name column index:",l,"Header:",(y=o[0])==null?void 0:y[l]);let c=N(s,["balance"]),d=N(s,["accrual rate","accrualrate"]),i=N(s,["carry over","carryover"]),u=N(s,["ytd accrued","ytdaccrued"]),f=N(s,["ytd used","ytdused"]),p=[],r=[],g=[],m=o.slice(1);m.forEach((h,R)=>{let I=R+2,P=l!==-1?String(h[l]||"").trim():`Row ${I}`;if(!P)return;let C=c!==-1&&Number(h[c])||0,x=d!==-1&&Number(h[d])||0,j=i!==-1&&Number(h[i])||0,$=u!==-1&&Number(h[u])||0,A=f!==-1&&Number(h[f])||0,v=j+$;C<0?p.push({name:P,issue:`Negative balance: ${C.toFixed(2)} hrs`,rowIndex:I}):A>v&&v>0&&p.push({name:P,issue:`Used ${A.toFixed(0)} hrs but only ${v.toFixed(0)} available`,rowIndex:I}),C===0&&(j>0||$>0)&&r.push({name:P,rowIndex:I}),x>8&&g.push({name:P,accrualRate:x,rowIndex:I})}),J.balanceIssues=p,J.zeroBalances=r,J.accrualOutliers=g,J.totalIssues=p.length,J.totalEmployees=m.filter(h=>h.some(R=>R!==null&&R!=="")).length,J.hasRun=!0});let e=J.balanceIssues.length>0;Pe({stepStatuses:{3:e?"blocked":"complete"}})}catch(e){console.error("Data quality check error:",e),window.alert(`Quality check failed: ${e.message}`),J.hasRun=!1}finally{J.loading=!1,Y(!1),ee()}}function Mn(){J.acknowledged=!0,Pe({stepStatuses:{3:"complete"}}),ee()}async function Vn(){if(te())try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem("PTO_Data"),n=e.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),a=t.getUsedRangeOrNullObject();if(a.load("values"),n.load("isNullObject"),await e.sync(),n.isNullObject){z.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let o=n.getUsedRangeOrNullObject();o.load("values"),await e.sync();let s=a.isNullObject?[]:a.values||[],l=o.isNullObject?[]:o.values||[];if(!s.length||!l.length){z.completenessCheck={accrualRate:null,carryOver:null,ytdAccrued:null,ytdUsed:null,balance:null};return}let c=(u,f,p)=>{let r=(u[0]||[]).map(y=>F(y)),g=N(r,f);return g===-1?null:u.slice(1).reduce((y,h)=>y+(Number(h[g])||0),0)},d=[{key:"accrualRate",aliases:["accrual rate","accrualrate"]},{key:"carryOver",aliases:["carry over","carryover","carry_over"]},{key:"ytdAccrued",aliases:["ytd accrued","ytdaccrued","ytd_accrued"]},{key:"ytdUsed",aliases:["ytd used","ytdused","ytd_used"]},{key:"balance",aliases:["balance"]}],i={};for(let u of d){let f=c(s,u.aliases,"PTO_Data"),p=c(l,u.aliases,"PTO_Analysis");if(f===null||p===null)i[u.key]=null;else{let r=Math.abs(f-p)<.01;i[u.key]={match:r,ptoData:f,ptoAnalysis:p}}}z.completenessCheck=i})}catch(e){console.error("Completeness check failed:",e)}}async function Ft(){if(!te()){window.alert("Excel is not available. Open this module inside Excel to run analysis.");return}Y(!0,"Running analysis...");try{await ln(),await Vn(),z.cleanDataReady=!0,ee()}catch(e){console.error("Full analysis error:",e),window.alert(`Analysis failed: ${e.message}`)}finally{Y(!1)}}async function tn(){if(!te()){window.alert("Excel is not available. Open this module inside Excel to run journal checks.");return}U.loading=!0,U.lastError=null,ae(document.getElementById("je-save-btn"),!1),ee();try{let e=await Excel.run(async t=>{let a=t.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();a.load("values");let o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis");o.load("isNullObject"),await t.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty. Generate the JE first.");let l=(s[0]||[]).map(I=>F(I)),c=N(l,["debit"]),d=N(l,["credit"]),i=N(l,["lineamount","line amount"]),u=N(l,["account number","accountnumber"]);if(c===-1||d===-1)throw new Error("Could not find Debit and Credit columns in PTO_JE_Draft.");let f=0,p=0,r=0,g=0;s.slice(1).forEach(I=>{let P=Number(I[c])||0,C=Number(I[d])||0,x=i!==-1&&Number(I[i])||0,j=u!==-1?String(I[u]||"").trim():"";f+=P,p+=C,r+=x,j&&j!=="21540"&&(g+=x)});let m=0;if(!o.isNullObject){let I=o.getUsedRangeOrNullObject();I.load("values"),await t.sync();let P=I.isNullObject?[]:I.values||[];if(P.length>1){let C=(P[0]||[]).map(j=>F(j)),x=N(C,["change"]);x!==-1&&P.slice(1).forEach(j=>{m+=Number(j[x])||0})}}let y=f-p,h=[];Math.abs(y)>=.01?h.push({check:"Debits = Credits",passed:!1,detail:y>0?`Debits exceed credits by $${Math.abs(y).toLocaleString(void 0,{minimumFractionDigits:2})}`:`Credits exceed debits by $${Math.abs(y).toLocaleString(void 0,{minimumFractionDigits:2})}`}):h.push({check:"Debits = Credits",passed:!0,detail:""}),Math.abs(r)>=.01?h.push({check:"Line Amounts Sum to Zero",passed:!1,detail:`Line amounts sum to $${r.toLocaleString(void 0,{minimumFractionDigits:2})} (should be $0.00)`}):h.push({check:"Line Amounts Sum to Zero",passed:!0,detail:""});let R=Math.abs(g-m);return R>=.01?h.push({check:"JE Matches Analysis Total",passed:!1,detail:`JE expense total ($${g.toLocaleString(void 0,{minimumFractionDigits:2})}) differs from PTO_Analysis Change total ($${m.toLocaleString(void 0,{minimumFractionDigits:2})}) by $${R.toLocaleString(void 0,{minimumFractionDigits:2})}`}):h.push({check:"JE Matches Analysis Total",passed:!0,detail:""}),{debitTotal:f,creditTotal:p,difference:y,lineAmountSum:r,jeChangeTotal:g,analysisChangeTotal:m,issues:h,validationRun:!0}});Object.assign(U,e,{lastError:null})}catch(e){console.warn("PTO JE summary:",e),U.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",U.debitTotal=null,U.creditTotal=null,U.difference=null,U.lineAmountSum=null,U.jeChangeTotal=null,U.analysisChangeTotal=null,U.issues=[],U.validationRun=!1}finally{U.loading=!1,ee()}}var Hn={"general & administrative":"64110","general and administrative":"64110","g&a":"64110","research & development":"62110","research and development":"62110","r&d":"62110",marketing:"61610","cogs onboarding":"53110","cogs prof. services":"56110","cogs professional services":"56110","sales & marketing":"61110","sales and marketing":"61110","cogs support":"52110","client success":"61811"},Gt="21540";async function Un(){if(!te()){window.alert("Excel is not available. Open this module inside Excel to create the journal entry.");return}Y(!0,"Creating PTO Journal Entry...");try{await Excel.run(async e=>{let n=e.workbook.tables.getItem(de[0]).getDataBodyRange();n.load("values");let o=e.workbook.worksheets.getItem("PTO_Analysis").getUsedRangeOrNullObject();o.load("values");let s=e.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");s.load("isNullObject"),await e.sync();let l=[];if(!s.isNullObject){let v=s.getUsedRangeOrNullObject();v.load("values"),await e.sync(),l=v.isNullObject?[]:v.values||[]}let c=n.values||[],d=o.isNullObject?[]:o.values||[];if(!d.length)throw new Error("PTO_Analysis is empty. Run the analysis first.");let i={};c.forEach(v=>{let b=String(v[1]||"").trim(),S=v[2];b&&(i[b]=S)}),(!i[O.journalEntryId]||!i[O.payrollDate])&&console.warn("[JE Draft] Missing config values - RefNumber:",i[O.journalEntryId],"TxnDate:",i[O.payrollDate]);let u=i[O.journalEntryId]||"",f=i[O.payrollDate]||"",p=i[O.accountingPeriod]||"",r="";if(f)try{let v;if(typeof f=="number"||/^\d{4,5}$/.test(String(f).trim())){let b=Number(f),S=new Date(1899,11,30);v=new Date(S.getTime()+b*24*60*60*1e3)}else v=new Date(f);if(!isNaN(v.getTime())&&v.getFullYear()>1970){let b=String(v.getMonth()+1).padStart(2,"0"),S=String(v.getDate()).padStart(2,"0"),W=v.getFullYear();r=`${b}/${S}/${W}`}else console.warn("[JE Draft] Date parsing resulted in invalid date:",f,"->",v),r=String(f)}catch(v){console.warn("[JE Draft] Could not parse TxnDate:",f,v),r=String(f)}let g=p?`${p} PTO Accrual`:"PTO Accrual",m={};if(l.length>1){let v=(l[0]||[]).map(W=>F(W)),b=N(v,["account number","accountnumber","account","acct"]),S=N(v,["account name","accountname","name","description"]);b!==-1&&S!==-1&&l.slice(1).forEach(W=>{let ne=String(W[b]||"").trim(),se=String(W[S]||"").trim();ne&&(m[ne]=se)})}let y=(d[0]||[]).map(v=>F(v)),h=N(y,["department"]),R=N(y,["change"]);if(h===-1||R===-1)throw new Error("Could not find Department or Change columns in PTO_Analysis.");let I={};d.slice(1).forEach(v=>{let b=String(v[h]||"").trim(),S=Number(v[R])||0;b&&S!==0&&(I[b]||(I[b]=0),I[b]+=S)});let P=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],C=[P],x=0,j=0;Object.entries(I).forEach(([v,b])=>{if(Math.abs(b)<.01)return;let S=v.toLowerCase().trim(),W=Hn[S]||"",ne=m[W]||"",se=b>0?Math.abs(b):0,G=b<0?Math.abs(b):0;x+=se,j+=G,C.push([u,r,W,ne,b,se,G,g,v])});let $=x-j;if(Math.abs($)>=.01){let v=$<0?Math.abs($):0,b=$>0?Math.abs($):0,S=m[Gt]||"Accrued PTO";C.push([u,r,Gt,S,-$,v,b,g,""])}let A=e.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");if(A.load("isNullObject"),await e.sync(),A.isNullObject)A=e.workbook.worksheets.add("PTO_JE_Draft");else{let v=A.getUsedRangeOrNullObject();v.load("isNullObject"),await e.sync(),v.isNullObject||v.clear()}if(C.length>0){let v=A.getRangeByIndexes(0,0,C.length,P.length);v.values=C;let b=A.getRangeByIndexes(0,0,1,P.length);et(b);let S=C.length-1;S>0&&(ce(A,4,S,!0),ce(A,5,S),ce(A,6,S)),v.format.autofitColumns()}await e.sync(),A.activate(),A.getRange("A1").select(),await e.sync()}),await tn()}catch(e){console.error("Create JE Draft error:",e),window.alert(`Unable to create Journal Entry: ${e.message}`)}finally{Y(!1)}}async function Fn(){if(!te()){window.alert("Excel is not available. Open this module inside Excel to export.");return}Y(!0,"Preparing JE CSV...");try{let{rows:e}=await Excel.run(async n=>{let o=n.workbook.worksheets.getItem("PTO_JE_Draft").getUsedRangeOrNullObject();o.load("values"),await n.sync();let s=o.isNullObject?[]:o.values||[];if(!s.length)throw new Error("PTO_JE_Draft is empty.");return{rows:s}}),t=la(e);ca(`pto-je-draft-${st()}.csv`,t)}catch(e){console.error("PTO JE export:",e),window.alert("Unable to export the JE draft. Confirm the sheet has data.")}finally{Y(!1)}}async function je(e){if(!(!e||!te()))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(e);n.activate(),n.getRange("A1").select(),await t.sync()})}catch(t){console.error(t)}}async function Gn(){if(!(!te()||!window.confirm("This will clear all data in PTO_Data. Are you sure?"))){Y(!0);try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("PTO_Data"),a=n.getUsedRangeOrNullObject();a.load("rowCount"),await t.sync(),!a.isNullObject&&a.rowCount>1&&(n.getRangeByIndexes(1,0,a.rowCount-1,20).clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),window.alert("PTO_Data cleared successfully. You can now paste new data.")}catch(t){console.error("Clear PTO_Data error:",t),window.alert(`Failed to clear PTO_Data: ${t.message}`)}finally{Y(!1)}}}async function Jt(e){if(!e||!te())return;let t={SS_Employee_Roster:["Employee","Department","Pay_Rate","Status","Hire_Date"],SS_Chart_of_Accounts:["Account_Number","Account_Name","Type","Category"]};try{await Excel.run(async n=>{let a=n.workbook.worksheets.getItemOrNullObject(e);if(a.load("isNullObject"),await n.sync(),a.isNullObject){a=n.workbook.worksheets.add(e);let o=t[e]||["Column1","Column2"],s=a.getRange(`A1:${String.fromCharCode(64+o.length)}1`);s.values=[o],s.format.font.bold=!0,s.format.fill.color="#f0f0f0",s.format.autofitColumns(),await n.sync()}a.activate(),a.getRange("A1").select(),await n.sync()})}catch(n){console.error("Error opening reference sheet:",n)}}async function Jn(){if(!te()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(o=>{o.name.toUpperCase().startsWith("SS_")&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[Config] Made visible: ${o.name}`),n++)}),await e.sync();let a=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");a.load("isNullObject"),await e.sync(),a.isNullObject||(a.activate(),a.getRange("A1").select(),await e.sync()),console.log(`[Config] ${n} system sheets now visible`)})}catch(e){console.error("[Config] Error unhiding system sheets:",e)}}function Z(e){var n,a;let t=String(e!=null?e:"").trim();return(a=(n=k.values)==null?void 0:n[t])!=null?a:""}function me(e){var n;if(e)return e;let t=Z(O.reviewerName);if(t)return t;if((n=window.PrairieForge)!=null&&n._sharedConfigCache){let a=window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer")||window.PrairieForge._sharedConfigCache.get("Default_Reviewer");if(a)return a}return""}function X(e,t,n={}){var l;let a=String(e!=null?e:"").trim();if(!a)return;k.values[a]=t!=null?t:"";let o=(l=n.debounceMs)!=null?l:0;if(!o){let c=be.get(a);c&&clearTimeout(c),be.delete(a),Se(a,t!=null?t:"",de);return}be.has(a)&&clearTimeout(be.get(a));let s=setTimeout(()=>{be.delete(a),Se(a,t!=null?t:"",de)},o);be.set(a,s)}function F(e){return String(e!=null?e:"").trim().toLowerCase()}function Y(e,t="Working..."){let n=document.getElementById(wn);n&&(n.style.display="none")}function nt(){kn()}typeof Office!="undefined"&&Office.onReady?Office.onReady(()=>nt()).catch(()=>nt()):nt();function he(e){return k.steps[e]||{notes:"",reviewer:"",signOffDate:""}}function nn(e){return Be[e]||{}}function qn(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}async function K(e,t,n){let a=k.steps[e]||{notes:"",reviewer:"",signOffDate:""};a[t]=n,k.steps[e]=a;let o=nn(e),s=t==="notes"?o.note:t==="reviewer"?o.reviewer:o.signOff;if(s&&q())try{await Se(s,n,de)}catch(l){console.warn("PTO: unable to save field",s,l)}}async function an(e,t){k.permanents[e]=t;let n=nn(e);if(n!=null&&n.note&&q())try{await Excel.run(async a=>{var p;let o=a.workbook.tables.getItemOrNullObject(de[0]);if(await a.sync(),o.isNullObject)return;let s=o.getDataBodyRange(),l=o.getHeaderRowRange();s.load("values"),l.load("values"),await a.sync();let c=l.values[0]||[],d=c.map(r=>String(r||"").trim().toLowerCase()),i={field:d.findIndex(r=>r==="field"||r==="field name"||r==="setting"),permanent:d.findIndex(r=>r==="permanent"||r==="persist"),value:d.findIndex(r=>r==="value"||r==="setting value"),type:d.findIndex(r=>r==="type"||r==="category"),title:d.findIndex(r=>r==="title"||r==="display name")};if(i.field===-1)return;let f=(s.values||[]).findIndex(r=>String(r[i.field]||"").trim()===n.note);if(f>=0)i.permanent>=0&&(s.getCell(f,i.permanent).values=[[t?"Y":"N"]]);else{let r=new Array(c.length).fill("");i.type>=0&&(r[i.type]="Other"),i.title>=0&&(r[i.title]=""),r[i.field]=n.note,i.permanent>=0&&(r[i.permanent]=t?"Y":"N"),i.value>=0&&(r[i.value]=((p=k.steps[e])==null?void 0:p.notes)||""),o.rows.add(null,[r])}await a.sync()})}catch(a){console.warn("PTO: unable to update permanent flag",a)}}async function on(e,t){let n=Zt[e];if(n&&(k.completes[e]=t?"Y":"",!!q()))try{await Se(n,t?"Y":"",de)}catch(a){console.warn("PTO: unable to save completion flag",n,a)}}function qt(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function Wn(){let e={};return Object.keys(Be).forEach(t=>{var s;let n=parseInt(t,10),a=!!((s=k.steps[n])!=null&&s.signOffDate),o=!!k.completes[n];e[n]=a||o}),e}function sn(e,{buttonId:t,inputId:n,canActivate:a=null,onComplete:o=null}){var d;let s=document.getElementById(t);if(!s)return;let l=document.getElementById(n),c=!!((d=k.steps[e])!=null&&d.signOffDate)||!!k.completes[e];qt(s,c),s.addEventListener("click",()=>{if(!s.classList.contains("is-active")&&e>0){let f=Wn(),{canComplete:p,message:r}=Bt(e,f);if(!p){Lt(r);return}}if(typeof a=="function"&&!a())return;let u=!s.classList.contains("is-active");qt(s,u),l&&(l.value=u?st():"",K(e,"signOffDate",l.value)),on(e,u),u&&typeof o=="function"&&o()})}function w(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function zn(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function rn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function we(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function ot(e){if(!e)return null;let t=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(e));if(!t)return null;let n=Number(t[1]),a=Number(t[2]),o=Number(t[3]);return!n||!a||!o?null:{year:n,month:a,day:o}}function Wt(e){if(!e)return"";let t=ot(e);if(!t)return"";let{year:n,month:a,day:o}=t;return`${n}-${String(a).padStart(2,"0")}-${String(o).padStart(2,"0")}`}function Yn(e){let t=ot(e);return t?`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function Kn(e){let t=ot(e);return t?`PTO-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function st(){let e=new Date,t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),a=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${a}`}function Qn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="y"||t==="yes"||t==="true"||t==="t"||t==="1"}function Xn(e){if(e instanceof Date)return e.getTime();if(typeof e=="number"){let n=Zn(e);return n?n.getTime():null}let t=new Date(e);return Number.isNaN(t.getTime())?null:t.getTime()}function Zn(e){if(!Number.isFinite(e))return null;let t=new Date(Date.UTC(1899,11,30));return new Date(t.getTime()+e*24*60*60*1e3)}function ea(){let e=n=>{var a,o;return((o=(a=document.getElementById(n))==null?void 0:a.value)==null?void 0:o.trim())||""};[{id:"config-payroll-date",field:O.payrollDate},{id:"config-accounting-period",field:O.accountingPeriod},{id:"config-journal-id",field:O.journalEntryId},{id:"config-company-name",field:O.companyName},{id:"config-payroll-provider",field:O.payrollProvider},{id:"config-accounting-link",field:O.accountingSoftware},{id:"config-user-name",field:O.reviewerName}].forEach(({id:n,field:a})=>{let o=e(n);a&&X(a,o)})}function N(e,t=[]){let n=t.map(a=>F(a));return e.findIndex(a=>n.some(o=>a.includes(o)))}function ta(e){var I,P,C,x,j,$,A,v,b;let t=he(2),n=(t==null?void 0:t.notes)||"",a=!!k.permanents[2],o=me(t==null?void 0:t.reviewer),s=(t==null?void 0:t.signOffDate)||"",l=!!(we(k.completes[2])||s),c=T.roster||{},d=T.hasAnalyzed,i=(P=(I=T.roster)==null?void 0:I.difference)!=null?P:0,u=!T.skipAnalysis&&Math.abs(i)>0,f=(C=c.rosterCount)!=null?C:0,p=(x=c.payrollCount)!=null?x:0,r=(j=c.difference)!=null?j:p-f,g=Array.isArray(c.mismatches)?c.mismatches.filter(Boolean):[],m="";T.loading?m=((A=($=window.PrairieForge)==null?void 0:$.renderStatusBanner)==null?void 0:A.call($,{type:"info",message:"Analyzing headcount\u2026",escapeHtml:w}))||"":T.lastError&&(m=((b=(v=window.PrairieForge)==null?void 0:v.renderStatusBanner)==null?void 0:b.call(v,{type:"error",message:T.lastError,escapeHtml:w}))||"");let y=(S,W,ne,se)=>{let G=!d,fe;G?fe='<span class="pf-je-check-circle pf-je-circle--pending"></span>':se?fe=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:fe=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let Me=d?` = ${ne}`:"";return`
            <div class="pf-je-check-row">
                ${fe}
                <span class="pf-je-check-desc-pill">${w(S)}${Me}</span>
            </div>
        `},h=`
        ${y("SS_Employee_Roster count","Active employees in roster",f,!0)}
        ${y("PTO_Data count","Unique employees in PTO data",p,!0)}
        ${y("Difference","Should be zero",r,r===0)}
    `,R=g.length&&!T.skipAnalysis&&d?window.PrairieForge.renderMismatchTiles({mismatches:g,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"PTO Data",escapeHtml:w}):"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${T.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
                    ${At}
                    <span>Skip Analysis</span>
                </button>
            </div>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Headcount Check</h3>
                    <p class="pf-config-subtext">Compare employee roster against PTO data to identify discrepancies.</p>
                </div>
                <div class="pf-signoff-action">
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${Ne}</button>`,"Run")}
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${_e}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Headcount Comparison</h3>
                    <p class="pf-config-subtext">Verify roster and payroll data align before proceeding.</p>
                </div>
                ${m}
                <div class="pf-je-checks-container">
                    ${h}
                </div>
                ${R}
            </article>
            ${re({textareaId:"step-notes-input",value:n,permanentId:"step-notes-lock-2",isPermanent:a,hintId:u?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
            ${le({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:l,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
        </section>
    `}function na(){let e=z.completenessCheck||{},t=z.missingPayRates||[],n=[{key:"accrualRate",label:"Accrual Rate",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"carryOver",label:"Carry Over",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdAccrued",label:"YTD Accrued",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"ytdUsed",label:"YTD Used",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"},{key:"balance",label:"Balance",desc:"\u2211 PTO_Data = \u2211 PTO_Analysis"}],o=n.every(i=>e[i.key]!==null&&e[i.key]!==void 0)&&n.every(i=>{var u;return(u=e[i.key])==null?void 0:u.match}),s=t.length>0,l=i=>{let u=e[i.key],f=u==null,p;return f?p='<span class="pf-je-check-circle pf-je-circle--pending"></span>':u.match?p=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:p=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${p}
                <span class="pf-je-check-desc-pill">${w(i.label)}: ${w(i.desc)}</span>
            </div>
        `},c=n.map(i=>l(i)).join(""),d="";if(s){let i=t[0],u=t.length-1;d=`
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${w(i.name)}</strong> to calculate liability
                </p>
                <div class="pf-readiness-input-row">
                    <div class="pf-readiness-input-field">
                        <span class="pf-readiness-input-prefix">$</span>
                        <input type="number" 
                               id="payrate-input" 
                               class="pf-readiness-input" 
                               placeholder="0.00" 
                               step="0.01"
                               min="0"
                               data-employee="${zn(i.name)}"
                               data-row="${i.rowIndex}">
                    </div>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--secondary" id="payrate-ignore-btn">
                        Skip
                    </button>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--primary" id="payrate-save-btn">
                        Save
                    </button>
                </div>
                ${u>0?`<p class="pf-readiness-remaining">${u} more employee${u>1?"s":""} need pay rates</p>`:""}
            </div>
        `}return`
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-readiness-card">
            <div class="pf-config-head">
                <h3>Data Completeness</h3>
                <p class="pf-config-subtext">Quick check that all your data transferred correctly.</p>
            </div>
            <div class="pf-je-checks-container">
                ${c}
            </div>
            ${d}
        </article>
    `}function aa(e){var r,g,m,y,h,R,I,P;let t=he(3),n=!!k.permanents[3],a=me(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(we(k.completes[3])||o),l=J.hasRun,{balanceIssues:c,zeroBalances:d,accrualOutliers:i,totalEmployees:u}=J,f="";if(J.loading)f=((g=(r=window.PrairieForge)==null?void 0:r.renderStatusBanner)==null?void 0:g.call(r,{type:"info",message:"Analyzing data quality...",escapeHtml:w}))||"";else if(l){let C=c.length,x=i.length+d.length;C>0?f=((y=(m=window.PrairieForge)==null?void 0:m.renderStatusBanner)==null?void 0:y.call(m,{type:"error",title:`${C} Balance Issue${C>1?"s":""} Found`,message:"Review the issues below. Fix in PTO_Data and re-run, or acknowledge to continue.",escapeHtml:w}))||"":x>0?f=((R=(h=window.PrairieForge)==null?void 0:h.renderStatusBanner)==null?void 0:R.call(h,{type:"warning",title:"No Critical Issues",message:`${x} informational item${x>1?"s":""} to review (see below).`,escapeHtml:w}))||"":f=((P=(I=window.PrairieForge)==null?void 0:I.renderStatusBanner)==null?void 0:P.call(I,{type:"success",title:"Data Quality Passed",message:`${u} employee${u!==1?"s":""} checked \u2014 no anomalies found.`,escapeHtml:w}))||""}let p=[];return l&&c.length>0&&p.push(`
            <div class="pf-quality-issue pf-quality-issue--critical">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u26A0\uFE0F</span>
                    <span class="pf-quality-issue-title">Balance Issues (${c.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${c.slice(0,5).map(C=>`<li><strong>${w(C.name)}</strong>: ${w(C.issue)}</li>`).join("")}
                    ${c.length>5?`<li class="pf-quality-more">+${c.length-5} more</li>`:""}
                </ul>
            </div>
        `),l&&i.length>0&&p.push(`
            <div class="pf-quality-issue pf-quality-issue--warning">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u{1F4CA}</span>
                    <span class="pf-quality-issue-title">High Accrual Rates (${i.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${i.slice(0,5).map(C=>`<li><strong>${w(C.name)}</strong>: ${C.accrualRate.toFixed(2)} hrs/period</li>`).join("")}
                    ${i.length>5?`<li class="pf-quality-more">+${i.length-5} more</li>`:""}
                </ul>
            </div>
        `),l&&d.length>0&&p.push(`
            <div class="pf-quality-issue pf-quality-issue--info">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">\u2139\uFE0F</span>
                    <span class="pf-quality-issue-title">Zero Balances (${d.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${d.slice(0,5).map(C=>`<li><strong>${w(C.name)}</strong></li>`).join("")}
                    ${d.length>5?`<li class="pf-quality-more">+${d.length-5} more</li>`:""}
                </ul>
            </div>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Quality Check</h3>
                    <p class="pf-config-subtext">Scan your imported data for common errors before proceeding.</p>
                </div>
                ${f}
                <div class="pf-signoff-action">
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${Ne}</button>`,"Run")}
                </div>
            </article>
            ${p.length>0?`
                <article class="pf-step-card pf-step-detail">
                    <div class="pf-config-head">
                        <h3>Issues Found</h3>
                        <p class="pf-config-subtext">Fix issues in PTO_Data and re-run, or acknowledge to continue.</p>
                    </div>
                    <div class="pf-quality-issues-grid">
                        ${p.join("")}
                    </div>
                    <div class="pf-quality-actions-bar">
                        ${J.acknowledged?'<p class="pf-quality-actions-hint"><span class="pf-acknowledged-badge">\u2713 Issues Acknowledged</span></p>':""}
                        <div class="pf-signoff-action">
                            ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${_e}</button>`,"Refresh")}
                            ${J.acknowledged?"":H(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${Ee}</button>`,"Continue")}
                        </div>
                    </div>
                </article>
            `:""}
            ${re({textareaId:"step-notes-3",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-3",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-3"})}
            ${le({reviewerInputId:"step-reviewer-3",reviewerValue:a,signoffInputId:"step-signoff-3",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-3",completeButtonId:"step-signoff-toggle-3"})}
        </section>
    `}function oa(e){let t=he(4),n=!!k.permanents[4],a=me(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(we(k.completes[4])||o);return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Analysis</h3>
                    <p class="pf-config-subtext">Calculate liabilities and compare against last period.</p>
                </div>
                <div class="pf-signoff-action">
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-run-btn" title="Run analysis and checks">${Ne}</button>`,"Run")}
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="analysis-refresh-btn" title="Refresh data from PTO_Data">${_e}</button>`,"Refresh")}
                </div>
            </article>
            ${na()}
            ${re({textareaId:"step-notes-4",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-4",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-4"})}
            ${le({reviewerInputId:"step-reviewer-4",reviewerValue:a,signoffInputId:"step-signoff-4",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"step-signoff-toggle-4"})}
        </section>
    `}function sa(e){let t=he(5),n=!!k.permanents[5],a=me(t==null?void 0:t.reviewer),o=(t==null?void 0:t.signOffDate)||"",s=!!(we(k.completes[5])||o),l=U.lastError?`<p class="pf-step-note">${w(U.lastError)}</p>`:"",c=U.validationRun,d=U.issues||[],i=[{key:"Debits = Credits",desc:"\u2211 Debit column = \u2211 Credit column"},{key:"Line Amounts Sum to Zero",desc:"\u2211 Line Amount = $0.00"},{key:"JE Matches Analysis Total",desc:"\u2211 Expense line amounts = \u2211 PTO_Analysis Change"}],u=g=>{let m=d.find(R=>R.check===g.key),y=!c,h;return y?h='<span class="pf-je-check-circle pf-je-circle--pending"></span>':m!=null&&m.passed?h=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:h=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${h}
                <span class="pf-je-check-desc-pill">${w(g.desc)}</span>
            </div>
        `},f=i.map(g=>u(g)).join(""),p=d.filter(g=>!g.passed),r="";return c&&p.length>0&&(r=`
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>\u26A0\uFE0F Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${p.map(g=>`<li><strong>${w(g.check)}:</strong> ${w(g.detail)}</li>`).join("")}
                </ul>
            </article>
        `),`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${w(ue)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${w(e.title)}</h2>
            <p class="pf-hero-copy">${w(e.summary||"")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Create a balanced JE from your imported PTO data, grouped by department.</p>
                </div>
                <div class="pf-signoff-action">
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${We}</button>`,"Generate")}
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${_e}</button>`,"Refresh")}
                    ${H(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${xt}</button>`,"Export")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">These checks run automatically after generating your JE.</p>
                </div>
                ${l}
                <div class="pf-je-checks-container">
                    ${f}
                </div>
            </article>
            ${r}
            ${re({textareaId:"step-notes-5",value:(t==null?void 0:t.notes)||"",permanentId:"step-notes-lock-5",isPermanent:n,hintId:"",saveButtonId:"step-notes-save-5"})}
            ${le({reviewerInputId:"step-reviewer-5",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:o,isComplete:s,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
        </section>
    `}function ia(){var t,n;return Math.abs((n=(t=T.roster)==null?void 0:t.difference)!=null?n:0)>0}function it(){return!T.skipAnalysis&&ia()}async function at(){if(!q()){T.loading=!1,T.lastError="Excel runtime is unavailable.",ee();return}T.loading=!0,T.lastError=null,ae(document.getElementById("headcount-save-btn"),!1),ee();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),a=t.workbook.worksheets.getItem("PTO_Data"),o=t.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.getUsedRangeOrNullObject(),l=a.getUsedRangeOrNullObject();s.load("values"),l.load("values"),o.load("isNullObject"),await t.sync();let c=null;o.isNullObject||(c=o.getUsedRangeOrNullObject(),c.load("values")),await t.sync();let d=s.isNullObject?[]:s.values||[],i=l.isNullObject?[]:l.values||[],u=c&&!c.isNullObject?c.values||[]:[],f=u.length?u:i;return ra(d,f)});T.roster=e.roster,T.hasAnalyzed=!0,T.lastError=null}catch(e){console.warn("PTO headcount: unable to analyze data",e),T.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{T.loading=!1,ee()}}function zt(e){if(!e)return!0;let t=e.toLowerCase().trim();return t?["total","subtotal","sum","count","grand","average","avg"].some(a=>t.includes(a)):!0}function ra(e,t){let n={rosterCount:0,payrollCount:0,difference:0,mismatches:[]};if(((e==null?void 0:e.length)||0)<2||((t==null?void 0:t.length)||0)<2)return console.warn("Headcount: insufficient data rows",{rosterRows:(e==null?void 0:e.length)||0,payrollRows:(t==null?void 0:t.length)||0}),{roster:n};let a=Yt(e),o=Yt(t),s=a.headers,l=o.headers,c={employee:Kt(s),termination:s.findIndex(r=>r.includes("termination"))},d={employee:Kt(l)};console.log("Headcount column detection:",{rosterEmployeeCol:c.employee,rosterTerminationCol:c.termination,payrollEmployeeCol:d.employee,rosterHeaders:s.slice(0,5),payrollHeaders:l.slice(0,5)});let i=new Set,u=new Set;for(let r=a.startIndex;r<e.length;r+=1){let g=e[r],m=c.employee>=0?ge(g[c.employee]):"";zt(m)||c.termination>=0&&ge(g[c.termination])||i.add(m.toLowerCase())}for(let r=o.startIndex;r<t.length;r+=1){let g=t[r],m=d.employee>=0?ge(g[d.employee]):"";zt(m)||u.add(m.toLowerCase())}n.rosterCount=i.size,n.payrollCount=u.size,n.difference=n.payrollCount-n.rosterCount,console.log("Headcount results:",{rosterCount:n.rosterCount,payrollCount:n.payrollCount,difference:n.difference});let f=[...i].filter(r=>!u.has(r)),p=[...u].filter(r=>!i.has(r));return n.mismatches=[...f.map(r=>`In roster, missing in PTO_Data: ${r}`),...p.map(r=>`In PTO_Data, missing in roster: ${r}`)],{roster:n}}function Yt(e){if(!Array.isArray(e)||!e.length)return{headers:[],startIndex:1};let t=e.findIndex((o=[])=>o.some(s=>ge(s).toLowerCase().includes("employee"))),n=t===-1?0:t;return{headers:(e[n]||[]).map(o=>ge(o).toLowerCase()),startIndex:n+1}}function Kt(e=[]){let t=-1,n=-1;return e.forEach((a,o)=>{let s=a.toLowerCase();if(!s.includes("employee"))return;let l=1;s.includes("name")?l=4:s.includes("id")?l=2:l=3,l>n&&(n=l,t=o)}),t}function ge(e){return e==null?"":String(e).trim()}async function ln(e=null){let t=async n=>{let a=n.workbook.worksheets.getItem("PTO_Data"),o=n.workbook.worksheets.getItemOrNullObject("PTO_Analysis"),s=n.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster"),l=n.workbook.worksheets.getItemOrNullObject("PR_Archive_Summary"),c=n.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary"),d=a.getUsedRangeOrNullObject();d.load("values"),o.load("isNullObject"),s.load("isNullObject"),l.load("isNullObject"),c.load("isNullObject"),await n.sync();let i=d.isNullObject?[]:d.values||[];if(!i.length)return;let u=(i[0]||[]).map(E=>F(E)),f=u.findIndex(E=>E.includes("employee")&&E.includes("name")),p=f>=0?f:0,r=N(u,["accrual rate"]),g=N(u,["carry over","carryover"]),m=u.findIndex(E=>E.includes("ytd")&&(E.includes("accrued")||E.includes("accrual"))),y=u.findIndex(E=>E.includes("ytd")&&E.includes("used")),h=N(u,["balance","current balance","pto balance"]);console.log("[PTO Analysis] PTO_Data headers:",u),console.log("[PTO Analysis] Column indices found:",{employee:p,accrualRate:r,carryOver:g,ytdAccrued:m,ytdUsed:y,balance:h}),y>=0?console.log(`[PTO Analysis] YTD Used column: "${u[y]}" at index ${y}`):console.warn("[PTO Analysis] YTD Used column NOT FOUND. Headers:",u);let R=i.slice(1).map(E=>ge(E[p])).filter(E=>E&&!E.toLowerCase().includes("total")),I=new Map;i.slice(1).forEach(E=>{let V=F(E[p]);!V||V.includes("total")||I.set(V,E)});let P=new Map;if(s.isNullObject)console.warn("[PTO Analysis] SS_Employee_Roster sheet not found");else{let E=s.getUsedRangeOrNullObject();E.load("values"),await n.sync();let V=E.isNullObject?[]:E.values||[];if(V.length){let B=(V[0]||[]).map(_=>F(_));console.log("[PTO Analysis] SS_Employee_Roster headers:",B);let L=B.findIndex(_=>_.includes("employee")&&_.includes("name"));L<0&&(L=B.findIndex(_=>_==="employee"||_==="name"||_==="full name"));let M=B.findIndex(_=>_.includes("department"));console.log(`[PTO Analysis] Roster column indices - Name: ${L}, Dept: ${M}`),L>=0&&M>=0?(V.slice(1).forEach(_=>{let Q=F(_[L]),ie=ge(_[M]);Q&&P.set(Q,ie)}),console.log(`[PTO Analysis] Built roster map with ${P.size} employees`)):console.warn("[PTO Analysis] Could not find Name or Department columns in SS_Employee_Roster")}}let C=new Map;if(!l.isNullObject){let E=l.getUsedRangeOrNullObject();E.load("values"),await n.sync();let V=E.isNullObject?[]:E.values||[];if(V.length){let B=(V[0]||[]).map(M=>F(M)),L={payrollDate:N(B,["payroll date"]),employee:N(B,["employee"]),category:N(B,["payroll category","category"]),amount:N(B,["amount","gross salary","gross_salary","earnings"])};L.employee>=0&&L.category>=0&&L.amount>=0&&V.slice(1).forEach(M=>{let _=F(M[L.employee]);if(!_)return;let Q=F(M[L.category]);if(!Q.includes("regular")||!Q.includes("earn"))return;let ie=Number(M[L.amount])||0;if(!ie)return;let Oe=Xn(M[L.payrollDate]),ke=C.get(_);(!ke||Oe!=null&&Oe>ke.timestamp)&&C.set(_,{payRate:ie/80,timestamp:Oe})})}}let x=new Map;if(!c.isNullObject){let E=c.getUsedRangeOrNullObject();E.load("values"),await n.sync();let V=E.isNullObject?[]:E.values||[];if(V.length>1){let B=(V[0]||[]).map(_=>F(_)),L=B.findIndex(_=>_.includes("employee")&&_.includes("name")),M=N(B,["liability amount","liability","accrued pto"]);L>=0&&M>=0&&V.slice(1).forEach(_=>{let Q=F(_[L]);if(!Q)return;let ie=Number(_[M])||0;x.set(Q,ie)})}}let j=Z(O.payrollDate)||"",$=[],A=[],v=R.map((E,V)=>{var ct,dt,ut,ft,pt,gt,mt;let B=F(E),L=P.get(B)||"",M=(dt=(ct=C.get(B))==null?void 0:ct.payRate)!=null?dt:"",_=I.get(B),Q=_&&r>=0&&(ut=_[r])!=null?ut:"",ie=_&&g>=0&&(ft=_[g])!=null?ft:"",Oe=_&&m>=0&&(pt=_[m])!=null?pt:"",ke=_&&y>=0&&(gt=_[y])!=null?gt:"";(B.includes("avalos")||B.includes("sarah"))&&console.log(`[PTO Debug] ${E}:`,{ytdUsedIdx:y,rawValue:_?_[y]:"no dataRow",ytdUsed:ke,fullRow:_});let Ve=_&&h>=0&&Number(_[h])||0,rt=V+2;!M&&typeof M!="number"&&$.push({name:E,rowIndex:rt}),L||A.push({name:E,rowIndex:rt});let He=typeof M=="number"&&Ve?Ve*M:0,lt=(mt=x.get(B))!=null?mt:0,cn=(typeof He=="number"?He:0)-lt;return[j,E,L,M,Q,ie,Oe,ke,Ve,He,lt,cn]});z.missingPayRates=$.filter(E=>!z.ignoredMissingPayRates.has(E.name)),z.missingDepartments=A,console.log(`[PTO Analysis] Data quality: ${$.length} missing pay rates, ${A.length} missing departments`);let b=[["Analysis Date","Employee Name","Department","Pay Rate","Accrual Rate","Carry Over","YTD Accrued","YTD Used","Balance","Liability Amount","Accrued PTO $ [Prior Period]","Change"],...v],S=o.isNullObject?n.workbook.worksheets.add("PTO_Analysis"):o,W=S.getUsedRangeOrNullObject();W.load("address"),await n.sync(),W.isNullObject||W.clear();let ne=b[0].length,se=b.length,G=v.length,fe=S.getRangeByIndexes(0,0,se,ne);fe.values=b;let Me=S.getRangeByIndexes(0,0,1,ne);et(Me),G>0&&(Mt(S,0,G),ce(S,3,G),ve(S,4,G),ve(S,5,G),ve(S,6,G),ve(S,7,G),ve(S,8,G),ce(S,9,G),ce(S,10,G),ce(S,11,G,!0)),fe.format.autofitColumns(),S.getRange("A1").select(),await n.sync()};q()&&(e?await t(e):await Excel.run(t))}function la(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let a=String(n);return/[",\n]/.test(a)?`"${a.replace(/"/g,'""')}"`:a}).join(",")).join(`
`)}function ca(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),a=URL.createObjectURL(n),o=document.createElement("a");o.href=a,o.download=e,document.body.appendChild(o),o.click(),o.remove(),setTimeout(()=>URL.revokeObjectURL(a),1e3)}function Qt(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=it(),n=document.getElementById("step-notes-input"),a=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!a;let o=document.getElementById("headcount-notes-hint");o&&(o.textContent=t?"Please document outstanding differences before signing off.":"")}function Xt(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(pe)?t.slice(pe.length).replace(/^\s+/,""):t.replace(new RegExp(`^${pe}\\s*`,"i"),"").trimStart(),a=pe+(n?`
${n}`:"");e.value!==a&&(e.value=a),K(2,"notes",e.value)}function da(){let e=document.getElementById("step-notes-input");e&&e.addEventListener("input",()=>{if(!T.skipAnalysis)return;let t=e.value||"";if(!t.startsWith(pe)){let n=t.replace(pe,"").trimStart();e.value=pe+(n?`
${n}`:"")}K(2,"notes",e.value)})}function ua(){var n;let e=it(),t=((n=document.getElementById("step-notes-input"))==null?void 0:n.value.trim())||"";if(e&&!t){window.alert("Please enter a brief explanation of the outstanding differences before completing this step.");return}}})();
//# sourceMappingURL=app.bundle.js.map
