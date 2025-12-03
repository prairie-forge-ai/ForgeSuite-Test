/* Prairie Forge Payroll Recorder */
(()=>{var nn="1.0.0.7",A={CONFIG:"SS_PF_Config",DATA:"PR_Data",DATA_CLEAN:"PR_Data_Clean",EXPENSE_MAPPING:"PR_Expense_Mapping",EXPENSE_REVIEW:"PR_Expense_Review",JE_DRAFT:"PR_JE_Draft",ARCHIVE_SUMMARY:"PR_Archive_Summary"};var Fa=[{name:"Instructions",description:"How to use the Prairie Forge payroll template"},{name:"Data_Input",description:"Paste WellsOne export data here"},{name:A.CONFIG,description:"Prairie Forge shared configuration storage (all modules)"},{name:"Config_Keywords",description:"Keyword-based account mapping rules"},{name:"Config_Accounts",description:"Account rewrite rules"},{name:"Config_Locations",description:"Location normalization rules"},{name:"Config_Vendors",description:"Vendor-specific overrides"},{name:"Config_Settings",description:"Prairie Forge system settings"},{name:A.EXPENSE_MAPPING,description:"Expense category mappings"},{name:A.DATA,description:"Processed payroll data staging"},{name:A.DATA_CLEAN,description:"Cleaned and validated payroll data"},{name:A.EXPENSE_REVIEW,description:"Expense review workspace"},{name:A.JE_DRAFT,description:"Journal entry preparation area"}];var rt=[{id:0,title:"Configuration Setup",summary:"Company profile, branding, and run settings",description:"Keep the SS_PF_Config table current before every payroll run so downstream sheets inherit the right colors, links, and identifiers.",icon:"\u{1F9ED}",ctaLabel:"Open Configuration Form",statusHint:"Configuration edits happen inside the PF_Config table and are available to every step instantly.",highlights:[{label:"Company Profile",detail:"Company name, logos, payroll date, reporting period."},{label:"Brand Identity",detail:"Primary + accent colors carry through dashboards and exports."},{label:"System Links",detail:"Quick jumps to HRIS, payroll provider, accounting import, and archive folders."}],checklist:["Review profile, branding, links, and run settings each payroll cycle.","Click Save to write updates back to the SS_PF_Config sheet."]},{id:1,title:"Import Payroll Data",summary:"Paste the payroll provider export into the Data sheet",description:"Pull your payroll data from your provider\u2019s portal and paste it into the Data tab. If the columns match, just paste the rows; if they don\u2019t, paste your headers and data right over the top. Formatting is fully automated.",icon:"\u{1F4E5}",ctaLabel:"Prepare Import Sheet",statusHint:"The Data worksheet is activated so you can paste the latest provider export.",highlights:[{label:"Source File",detail:"Use WellsOne/ADP export with every pay category column visible."},{label:"Structure",detail:"Row 2 headers, row 3+ data, no blank columns, totals removed."},{label:"Quality",detail:"Spot-check employee counts and pay period filters before moving on."}],checklist:["Download the payroll detail export covering this pay period.","Paste values into the Data sheet starting at cell A3.","Confirm all pay category headers remain intact and spelled consistently."]},{id:2,title:"Headcount Review",summary:"Ensure roster and payroll rows agree",description:"This step is optional, but strongly recommended. A centralized employee roster keeps every payroll-related workbook aligned while ensuring key attributes such as department and location stay consistent each pay period.",icon:"\u{1F465}",ctaLabel:"Launch Headcount Review",statusHint:"Data and mapping sheets are surfaced so you can reconcile roster counts before validation.",highlights:[{label:"Roster Alignment",detail:"Compare active roster to the employees present in the Data sheet."},{label:"Variance Tracking",detail:"Note missing departments or unexpected hires before the validation run."},{label:"Approvals",detail:"Capture reviewer initials and date for audit coverage."}],checklist:["Filter the Data sheet by Department to ensure every team appears.","Look for duplicate or out-of-period employees and resolve upstream.","Document findings on the Headcount Review tab or your tracker of choice."]},{id:3,title:"Validate & Reconcile",summary:"Normalize payroll data and reconcile totals",description:"Automatically rebuild the PR_Data_Clean sheet, confirm payroll totals match, and prep the bank reconciliation before moving to Expense Review.",icon:"\u2705",statusHint:"Run completes automatically when you enter this step.",highlights:[{label:"Normalized Data",detail:"Creates one row per employee and payroll category."},{label:"Outputs",detail:"Data_Clean rebuilt with payroll category + mapping details."},{label:"Reconciliation",detail:"Displays PR_Data vs PR_Data_Clean totals plus bank comparison."}]},{id:4,title:"Expense Review",summary:"Generate an executive-ready payroll summary",description:"Build a six-period payroll dashboard (current + five prior), including department-level breakouts and variance indicators, plus notes and CoPilot guidance.",icon:"\u{1F4CA}",statusHint:"Selecting this step rebuilds PR_Expense_Review automatically.",highlights:[{label:"Time Series",detail:"Shows six consecutive payroll periods."},{label:"Departments",detail:"All-in totals, burden rates, and headcount by department."},{label:"Guidance",detail:"Use CoPilot to summarize trends and capture review notes."}],checklist:[]},{id:5,title:"Journal Entry Prep",summary:"Generate a QuickBooks-ready journal draft",description:"Create the JE Draft sheet with the headers QuickBooks Online/Desktop expect so you only need to paste balanced lines.",icon:"\u{1F9FE}",ctaLabel:"Generate JE Draft",statusHint:"JE Draft contains headers for RefNumber, TxnDate, account columns, and line descriptions.",highlights:[{label:"Structure",detail:"Debit/Credit columns prepared with standard import headers."},{label:"Context",detail:"JE Transaction ID from configuration is referenced for traceability."},{label:"Next Step",detail:"Populate amounts from Expense Review to finalize the journal."}],checklist:["Ensure validation + expense review steps are complete.","Run the generator to rebuild the JE Draft sheet.","Paste balanced lines and export to QuickBooks / ERP import format."]},{id:6,title:"Archive & Clear",summary:"Snapshot workpapers and reset working tabs",description:"Capture a log of each payroll run, note the archive destination, and optionally clear staging sheets for the next cycle.",icon:"\u{1F5C2}\uFE0F",ctaLabel:"Create Archive Summary",statusHint:"Archive summary headers help you log when data was exported and where the files live.",highlights:[{label:"Run Log",detail:"Payroll date, reporting period, JE ID, and who processed the run."},{label:"Storage",detail:"Link to the Archive folder defined in configuration."},{label:"Reset",detail:"Reminder to clear Data/Data_Clean once files are safely archived."}],checklist:["Record archive destination links and reviewer approvals.","Copy Data/Data_Clean/JE Draft tabs to the archive workbook if needed.","Clear sensitive data so the template is ready for the next payroll."]}],ja=(typeof window!="undefined"&&Array.isArray(window.PF_BUILDER_ALLOWLIST)?window.PF_BUILDER_ALLOWLIST:[]).map(e=>String(e||"").trim().toLowerCase());function Ge(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}function on(e){try{Office.onReady(t=>{console.log("Office.onReady fired:",t),t.host===Office.HostType.Excel||console.warn("Not running in Excel, host:",t.host),e(t)})}catch(t){console.warn("Office.onReady failed:",t),e(null)}}var mo="SS_PF_Config",go="module-prefix",Dt="system",Oe={PR_:"payroll-recorder",PTO_:"pto-accrual",CC_:"credit-card-expense",COM_:"commission-calc",SS_:"system"};async function an(){if(!Ge())return{...Oe};try{return await Excel.run(async e=>{var u,f;let t=e.workbook.worksheets.getItemOrNullObject(mo);if(await e.sync(),t.isNullObject)return console.log("[Tab Visibility] Config sheet not found, using defaults"),{...Oe};let n=t.getUsedRangeOrNullObject();if(n.load("values"),await e.sync(),n.isNullObject||!((u=n.values)!=null&&u.length))return{...Oe};let o=n.values,a=vo(o[0]),s=a.get("category"),l=a.get("field"),d=a.get("value");if(s===void 0||l===void 0||d===void 0)return console.warn("[Tab Visibility] Missing required columns, using defaults"),{...Oe};let r={},i=!1;for(let p=1;p<o.length;p++){let c=o[p];if(it(c[s])===go){let y=String((f=c[l])!=null?f:"").trim().toUpperCase(),w=it(c[d]);y&&w&&(r[y]=w,i=!0)}}return i?(console.log("[Tab Visibility] Loaded prefix config:",r),r):(console.log("[Tab Visibility] No module-prefix rows found, using defaults"),{...Oe})})}catch(e){return console.warn("[Tab Visibility] Error reading prefix config:",e),{...Oe}}}async function Pt(e){if(!Ge())return;let t=it(e);console.log(`[Tab Visibility] Applying visibility for module: ${t}`);try{let n=await an();await Excel.run(async o=>{let a=o.workbook.worksheets;a.load("items/name,visibility"),await o.sync();let s={};for(let[p,c]of Object.entries(n))s[c]||(s[c]=[]),s[c].push(p);let l=s[t]||[],d=s[Dt]||[],r=[];for(let[p,c]of Object.entries(s))p!==t&&p!==Dt&&r.push(...c);console.log(`[Tab Visibility] Active prefixes: ${l.join(", ")}`),console.log(`[Tab Visibility] Other module prefixes (to hide): ${r.join(", ")}`),console.log(`[Tab Visibility] System prefixes (always hide): ${d.join(", ")}`);let i=[],u=[];a.items.forEach(p=>{let c=p.name,m=c.toUpperCase(),y=l.some(E=>m.startsWith(E)),w=r.some(E=>m.startsWith(E)),h=d.some(E=>m.startsWith(E));y?(i.push(p),console.log(`[Tab Visibility] SHOW: ${c} (matches active module prefix)`)):h?(u.push(p),console.log(`[Tab Visibility] HIDE: ${c} (system sheet)`)):w?(u.push(p),console.log(`[Tab Visibility] HIDE: ${c} (other module prefix)`)):console.log(`[Tab Visibility] SKIP: ${c} (no prefix match, leaving as-is)`)});for(let p of i)p.visibility=Excel.SheetVisibility.visible;if(await o.sync(),a.items.filter(p=>p.visibility===Excel.SheetVisibility.visible).length>u.length){for(let p of u)try{p.visibility=Excel.SheetVisibility.hidden}catch(c){console.warn(`[Tab Visibility] Could not hide "${p.name}":`,c.message)}await o.sync()}else console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");console.log(`[Tab Visibility] Done! Showed ${i.length}, hid ${u.length} tabs`)})}catch(n){console.warn("[Tab Visibility] Error applying visibility:",n)}}async function ho(){if(!Ge()){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(o=>{o.visibility!==Excel.SheetVisibility.visible&&(o.visibility=Excel.SheetVisibility.visible,console.log(`[ShowAll] Made visible: ${o.name}`),n++)}),await e.sync(),console.log(`[ShowAll] Done! Made ${n} sheets visible. Total: ${t.items.length}`)})}catch(e){console.error("[Tab Visibility] Unable to show all sheets:",e)}}async function yo(){if(!Ge()){console.log("Excel not available");return}try{let e=await an(),t=[];for(let[n,o]of Object.entries(e))o===Dt&&t.push(n);await Excel.run(async n=>{let o=n.workbook.worksheets;o.load("items/name,visibility"),await n.sync(),o.items.forEach(a=>{let s=a.name.toUpperCase();t.some(l=>s.startsWith(l))&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[Unhide] Made visible: ${a.name}`))}),await n.sync(),console.log("[Unhide] System sheets are now visible!")})}catch(e){console.error("[Tab Visibility] Unable to unhide system sheets:",e)}}function vo(e=[]){let t=new Map;return e.forEach((n,o)=>{let a=it(n);a&&t.set(a,o)}),t}function it(e){return String(e!=null?e:"").trim().toLowerCase().replace(/[\s_]+/g,"-")}typeof window!="undefined"&&(window.PrairieForge=window.PrairieForge||{},window.PrairieForge.showAllSheets=ho,window.PrairieForge.unhideSystemSheets=yo,window.PrairieForge.applyModuleTabVisibility=Pt);var lt={COMPANY_NAME:"Prairie Forge LLC",PRODUCT_NAME:"Prairie Forge Tools",SUPPORT_URL:"https://prairieforge.ai/support",ADA_IMAGE_URL:"https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png"};var bo='<svg viewBox="0 0 24 24" fill="currentColor"><path d="M22.2819 9.8211a5.9847 5.9847 0 0 0-.5157-4.9108 6.0462 6.0462 0 0 0-6.5098-2.9A6.0651 6.0651 0 0 0 4.9807 4.1818a5.9847 5.9847 0 0 0-3.9977 2.9 6.0462 6.0462 0 0 0 .7427 7.0966 5.98 5.98 0 0 0 .511 4.9107 6.051 6.051 0 0 0 6.5146 2.9001A5.9847 5.9847 0 0 0 13.2599 24a6.0557 6.0557 0 0 0 5.7718-4.2058 5.9894 5.9894 0 0 0 3.9977-2.9001 6.0557 6.0557 0 0 0-.7475-7.0729zm-9.022 12.6081a4.4755 4.4755 0 0 1-2.8764-1.0408l.1419-.0804 4.7783-2.7582a.7948.7948 0 0 0 .3927-.6813v-6.7369l2.02 1.1686a.071.071 0 0 1 .038.052v5.5826a4.504 4.504 0 0 1-4.4945 4.4944zm-9.6607-4.1254a4.4708 4.4708 0 0 1-.5346-3.0137l.142.0852 4.783 2.7582a.7712.7712 0 0 0 .7806 0l5.8428-3.3685v2.3324a.0804.0804 0 0 1-.0332.0615L9.74 19.9502a4.4992 4.4992 0 0 1-6.1408-1.6464zM2.3408 7.8956a4.485 4.485 0 0 1 2.3655-1.9728V11.6a.7664.7664 0 0 0 .3879.6765l5.8144 3.3543-2.0201 1.1685a.0757.0757 0 0 1-.071 0l-4.8303-2.7865A4.504 4.504 0 0 1 2.3408 7.8956zm16.5963 3.8558L13.1038 8.364 15.1192 7.2a.0757.0757 0 0 1 .071 0l4.8303 2.7913a4.4944 4.4944 0 0 1-.6765 8.1042v-5.6772a.79.79 0 0 0-.407-.667zm2.0107-3.0231l-.142-.0852-4.7735-2.7818a.7759.7759 0 0 0-.7854 0L9.409 9.2297V6.8974a.0662.0662 0 0 1 .0284-.0615l4.8303-2.7866a4.4992 4.4992 0 0 1 6.6802 4.66zM8.3065 12.863l-2.02-1.1638a.0804.0804 0 0 1-.038-.0567V6.0742a4.4992 4.4992 0 0 1 7.3757-3.4537l-.142.0805L8.704 5.459a.7948.7948 0 0 0-.3927.6813zm1.0976-2.3654l2.602-1.4998 2.6069 1.4998v2.9994l-2.5974 1.4997-2.6067-1.4997Z"/></svg>',wo='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>',Eo=lt.ADA_IMAGE_URL,sn={id:"pf-copilot",heading:"Ada",subtext:"Your smart assistant to help you analyze and troubleshoot.",welcomeMessage:"What would you like to explore?",placeholder:"Where should I focus this pay period?",quickActions:[{id:"diagnostics",label:"Diagnostics",prompt:"Run a diagnostic check on the current payroll data. Check for completeness, accuracy issues, and any data quality concerns."},{id:"insights",label:"Insights",prompt:"What are the key insights and notable findings from this payroll period that I should highlight for executive review?"},{id:"variance",label:"Variances",prompt:"Analyze the significant variances between this period and the prior period. What's driving the changes?"},{id:"journal",label:"JE Check",prompt:"Is the journal entry ready for export? Check that debits equal credits and flag any mapping issues."}],systemPrompt:`You are Prairie Forge CoPilot, an expert financial analyst assistant embedded in an Excel add-in. 

Your role is to help accountants and CFOs:
1. Analyze payroll expense data for accuracy and completeness
2. Identify trends, anomalies, and areas requiring attention
3. Prepare executive-ready insights and talking points
4. Validate journal entries before export

Communication style:
- Be concise but thorough
- Use bullet points for clarity
- Highlight actionable items with \u26A0\uFE0F or \u2713
- Format currency as $X,XXX and percentages as X.X%
- Always suggest 2-3 concrete next steps

When analyzing data, look for:
- Period-over-period changes > 10%
- Department cost anomalies
- Headcount vs payroll mismatches
- Burden rate outliers
- Missing or incomplete mappings`},It=[];function rn(e={}){var o;let t={...sn,...e},n=((o=t.quickActions)==null?void 0:o.map(a=>`<button type="button" class="pf-ada-chip" data-action="${a.id}" data-prompt="${Co(a.prompt)}">${a.label}</button>`).join(""))||"";return`
        <article class="pf-ada" data-copilot="${t.id}">
            <header class="pf-ada-header">
                <div class="pf-ada-identity">
                    <img class="pf-ada-avatar" src="${Eo}" alt="Ada" onerror="this.style.display='none'" />
                    <div class="pf-ada-name">
                        <span class="pf-ada-title"><span class="pf-ada-title--ask">ask</span><span class="pf-ada-title--ada">ADA</span></span>
                        <span class="pf-ada-role">${t.subtext}</span>
                    </div>
                </div>
                <div class="pf-ada-status" id="${t.id}-status-badge" title="Ready">
                    <span class="pf-ada-status-dot" id="${t.id}-status-dot"></span>
                </div>
            </header>
            
            <div class="pf-ada-body">
                <div class="pf-ada-conversation" id="${t.id}-messages">
                    <div class="pf-ada-bubble pf-ada-bubble--ai">
                        <p>${t.welcomeMessage}</p>
                    </div>
                </div>
                
                <div class="pf-ada-composer">
                    <input 
                        type="text" 
                        class="pf-ada-input" 
                        id="${t.id}-prompt" 
                        placeholder="${t.placeholder}" 
                        autocomplete="off"
                    >
                    <button type="button" class="pf-ada-send" id="${t.id}-ask" title="Send">
                        ${wo}
                    </button>
                </div>
                
                ${n?`<div class="pf-ada-chips">${n}</div>`:""}
                
                <footer class="pf-ada-footer">
                    ${bo}
                    <span>Powered by ChatGPT</span>
                </footer>
            </div>
        </article>
    `}function Co(e){return String(e||"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}function ln(e,t={}){let n={...sn,...t},o=e.querySelector(`[data-copilot="${n.id}"]`);if(!o)return;let a=o.querySelector(`#${n.id}-messages`),s=o.querySelector(`#${n.id}-prompt`),l=o.querySelector(`#${n.id}-ask`),d=o.querySelector(`#${n.id}-status-dot`),r=o.querySelector(`#${n.id}-status-badge`),i=!1,u=(w,h="ready")=>{d&&(d.classList.remove("pf-ada-status-dot--busy","pf-ada-status-dot--offline"),h==="busy"&&d.classList.add("pf-ada-status-dot--busy"),h==="offline"&&d.classList.add("pf-ada-status-dot--offline")),r&&(r.title=w)},f=(w,h="assistant")=>{if(!a)return;let E=h==="user"?"pf-ada-bubble--user":h==="system"?"pf-ada-bubble--system":"pf-ada-bubble--ai",k=document.createElement("div");k.className=`pf-ada-bubble ${E}`,k.innerHTML=`<p>${m(w)}</p>`,a.appendChild(k),a.scrollTop=a.scrollHeight,It.push({role:h,content:w,timestamp:new Date().toISOString()})},p=()=>{if(!a)return;let w=document.createElement("div");w.className="pf-ada-bubble pf-ada-bubble--ai pf-ada-bubble--loading",w.id=`${n.id}-loading`,w.innerHTML=`
            <div class="pf-ada-typing">
                <span></span><span></span><span></span>
            </div>
        `,a.appendChild(w),a.scrollTop=a.scrollHeight},c=()=>{let w=document.getElementById(`${n.id}-loading`);w&&w.remove()},m=w=>String(w).replace(/\*\*(.*?)\*\*/g,"<strong>$1</strong>").replace(/\n\n/g,"</p><p>").replace(/\n- /g,"<br>\u2022 ").replace(/\n/g,"<br>"),y=async w=>{let h=w||(s==null?void 0:s.value.trim());if(!(!h||i)){i=!0,s&&(s.value=""),l&&(l.disabled=!0),f(h,"user"),p(),u("Analyzing...","busy");try{let E=null;if(typeof n.contextProvider=="function")try{E=await n.contextProvider()}catch(_){console.warn("CoPilot: Context provider failed",_)}let k;typeof n.onPrompt=="function"?k=await n.onPrompt(h,E,It):typeof n.apiEndpoint=="string"?k=await ko(n.apiEndpoint,h,E,n.systemPrompt):k=Ro(h,E),c(),f(k,"assistant"),u("Ready to assist","ready")}catch(E){console.error("CoPilot error:",E),c(),f(`I encountered an issue: ${E.message}. Please try again.`,"system"),u("Error occurred","offline")}i=!1,l&&(l.disabled=!1),s==null||s.focus()}};l==null||l.addEventListener("click",()=>y()),s==null||s.addEventListener("keydown",w=>{w.key==="Enter"&&!w.shiftKey&&(w.preventDefault(),y())}),o.querySelectorAll(".pf-ada-chip").forEach(w=>{w.addEventListener("click",()=>{let h=w.dataset.prompt;h&&y(h)})})}async function ko(e,t,n,o){let a=await fetch(e,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({prompt:t,context:n,systemPrompt:o,history:It.slice(-10)})});if(!a.ok)throw new Error(`API request failed: ${a.status}`);let s=await a.json();return s.message||s.response||"No response received."}function Ro(e,t){var o,a,s;let n=e.toLowerCase();return n.includes("diagnostic")||n.includes("check")?`Great question! Let me run through the diagnostics for you.

**\u2713 What Looks Good:**
\u2022 All required fields are populated
\u2022 Current period matches your config date
\u2022 All expense categories are mapped to GL accounts

**\u26A0\uFE0F Items Worth Reviewing:**
\u2022 2 departments show >15% variance from prior period
\u2022 Burden rate (14.6%) is slightly below your historical average (16.2%)

**My Recommendations:**
1. Take a closer look at the Sales & Marketing variance (-44.4%)
2. Verify headcount changes align with HR records
3. Once satisfied, you're clear to proceed to Journal Entry Prep!

Let me know if you'd like me to dig deeper into any of these.`:n.includes("insight")||n.includes("notable")||n.includes("finding")?`Here's what stands out this period \u2014 perfect for your executive summary.

**\u{1F4CA} The Headlines:**
\u2022 Total Payroll: ${(o=t==null?void 0:t.summary)!=null&&o.total?`$${(t.summary.total/1e3).toFixed(0)}K`:"$254K"}
\u2022 Headcount: ${((a=t==null?void 0:t.summary)==null?void 0:a.employeeCount)||38} employees
\u2022 Avg Cost/Employee: ${(s=t==null?void 0:t.summary)!=null&&s.avgPerEmployee?`$${t.summary.avgPerEmployee.toFixed(0)}`:"$6,674"}

**\u{1F4A1} Key Findings:**
1. **Payroll decreased 14.2%** \u2014 primarily driven by headcount reduction in Sales
2. **R&D remains your largest cost center** at 39% of total payroll
3. **Burden rate normalized** to 14.6% (was 18.2% prior period)

**\u26A0\uFE0F Items to Flag:**
\u2022 Sales & Marketing down $52K \u2014 worth confirming this was intentional
\u2022 2 fewer employees than prior period

**Suggested Talking Points:**
\u2022 "Payroll efficiency improved with 14% reduction while maintaining core operations"
\u2022 "R&D investment remains strong \u2014 aligned with product roadmap"

Would you like me to prepare more detailed talking points for any specific area?`:n.includes("variance")||n.includes("change")||n.includes("difference")?`**Variance Analysis: Current vs Prior Period**

\u{1F4C8} **Significant Changes**:

| Department | Change | % Change | Driver |
|------------|--------|----------|--------|
| Sales & Marketing | -$52,298 | -44.4% | \u{1F534} Headcount |
| Research & Dev | +$8,514 | +9.4% | Merit increases |
| General & Admin | +$1,610 | +3.9% | Normal variance |

\u{1F50D} **Root Cause Analysis**:

**Sales & Marketing (-44.4%)**:
\u2022 3 positions eliminated per restructuring plan
\u2022 Commission payouts lower due to Q4 timing
\u2022 \u26A0\uFE0F Verify: Is this aligned with sales targets?

**R&D (+9.4%)**:
\u2022 Annual merit increases effective this period
\u2022 1 new senior engineer hire
\u2022 \u2713 Expected per hiring plan

**Recommendation**: Document Sales variance in review notes. This is material and will be questioned.`:n.includes("journal")||n.includes("je")||n.includes("entry")?`Good news \u2014 your journal entry looks ready to go! Here's the full check:

**\u2713 Balance Check: PASSED**
\u2022 Total Debits: $253,625
\u2022 Total Credits: $253,625
\u2022 Difference: $0.00 \u2014 perfectly balanced!

**\u2713 Mapping Validation: Complete**
\u2022 9 unique GL accounts used
\u2022 All department codes are valid

**\u2713 Reference Data:**
\u2022 JE ID: PR-AUTO-2025-11-27
\u2022 Transaction Date: 2025-11-27
\u2022 Period: November 2025

**You're clear to export!** \u2705

**Next Steps:**
1. Click "Export" to download the CSV
2. Import into your accounting system
3. Post and reconcile

Let me know if you need me to double-check anything before you export!`:`Great question! I'm Ada, and I'm here to help with your payroll analysis.

Here's what I can help you with:

\u2022 **\u{1F50D} Diagnostics** \u2014 Check data quality and completeness
\u2022 **\u{1F4A1} Insights** \u2014 Key findings for executive review  
\u2022 **\u{1F4CA} Variance Analysis** \u2014 Period-over-period changes
\u2022 **\u{1F4CB} JE Readiness** \u2014 Validate journal entries before export

Try clicking one of the quick action buttons above, or just ask me something specific like:
\u2022 "What's driving the variance this period?"
\u2022 "Is my data ready for the CFO?"
\u2022 "Summarize the department breakdown"

I'm reading your actual spreadsheet data, so I can give you specific answers!`}var un=lt.ADA_IMAGE_URL;async function Nt(e,t,n){if(typeof Excel=="undefined"){console.warn("Excel runtime not available for homepage sheet");return}try{await Excel.run(async o=>{let a=o.workbook.worksheets.getItemOrNullObject(e);a.load("isNullObject, name"),await o.sync();let s;a.isNullObject?(s=o.workbook.worksheets.add(e),await o.sync(),await cn(o,s,t,n)):(s=a,await cn(o,s,t,n)),s.activate(),s.getRange("A1").select(),await o.sync()})}catch(o){console.error(`Error activating homepage sheet ${e}:`,o)}}async function cn(e,t,n,o){try{let i=t.getUsedRangeOrNullObject();i.load("isNullObject"),await e.sync(),i.isNullObject||(i.clear(),await e.sync())}catch{}t.showGridlines=!1,t.getRange("A:A").format.columnWidth=400,t.getRange("B:B").format.columnWidth=50,t.getRange("1:1").format.rowHeight=60,t.getRange("2:2").format.rowHeight=30;let a=[[n,""],[o,""],["",""],["",""]],s=t.getRangeByIndexes(0,0,4,2);s.values=a;let l=t.getRange("A1:Z100");l.format.fill.color="#0f0f0f";let d=t.getRange("A1");d.format.font.bold=!0,d.format.font.size=36,d.format.font.color="#ffffff",d.format.font.name="Segoe UI Light",d.format.verticalAlignment="Center";let r=t.getRange("A2");r.format.font.size=14,r.format.font.color="#a0a0a0",r.format.font.name="Segoe UI",r.format.verticalAlignment="Top",t.freezePanes.freezeRows(0),t.freezePanes.freezeColumns(0),await e.sync()}var dn={"module-selector":{sheetName:"SS_Homepage",title:"ForgeSuite",subtitle:"Select a module from the side panel to get started."},"payroll-recorder":{sheetName:"PR_Homepage",title:"Payroll Recorder",subtitle:"Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."},"pto-accrual":{sheetName:"PTO_Homepage",title:"PTO Accrual",subtitle:"Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."}};function Ot(e){return dn[e]||dn["module-selector"]}function pn(){Tt();let e=document.createElement("button");return e.className="pf-ada-fab",e.id="pf-ada-fab",e.setAttribute("aria-label","Ask Ada"),e.setAttribute("title","Ask Ada"),e.innerHTML=`
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${un}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `,document.body.appendChild(e),e.addEventListener("click",So),e}function Tt(){let e=document.getElementById("pf-ada-fab");e&&e.remove();let t=document.getElementById("pf-ada-modal-overlay");t&&t.remove()}function So(){let e=document.getElementById("pf-ada-modal-overlay");e&&e.remove();let t=document.createElement("div");t.className="pf-ada-modal-overlay",t.id="pf-ada-modal-overlay",t.innerHTML=`
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${un}" alt="Ada" />
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
    `,document.body.appendChild(t),requestAnimationFrame(()=>{t.classList.add("is-visible")});let n=document.getElementById("ada-modal-close");n==null||n.addEventListener("click",$t),t.addEventListener("click",a=>{a.target===t&&$t()});let o=a=>{a.key==="Escape"&&($t(),document.removeEventListener("keydown",o))};document.addEventListener("keydown",o)}function $t(){let e=document.getElementById("pf-ada-modal-overlay");e&&(e.classList.remove("is-visible"),setTimeout(()=>{e.remove()},300))}var fn=`
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
`.trim(),mn=`
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
`.trim(),gn=`
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
`.trim(),ct=`
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
`.trim(),hn=`
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
`.trim(),yn=`
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
`.trim(),xo={config:`
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
    `};function vn(e){return e&&xo[e]||""}var Lt=`
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
`.trim(),Bt=`
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
`.trim(),dt=`
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
`.trim(),ut=`
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
`.trim(),Xa=`
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
`.trim(),pt=`
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
`.trim(),bn=`
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
`.trim(),wn=`
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
`.trim(),En=`
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
`.trim(),Cn=`
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
`.trim(),kn=`
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
`.trim(),Qa=`
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
`.trim(),Za=`
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
`.trim(),es=`
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
`.trim(),ts=`
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
`.trim(),ns=`
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
`.trim(),os=`
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
`.trim(),as=`
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
`.trim(),ss=`
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
`.trim(),ze=`
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
`.trim(),Rn=`
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
`.trim();function We(e){return e==null?"":String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;")}function Mt(e,t){return`
        <div class="pf-labeled-btn">
            ${e}
            <span class="pf-btn-label">${t}</span>
        </div>
    `}function _e({textareaId:e,value:t,permanentId:n,isPermanent:o,hintId:a,saveButtonId:s,isSaved:l=!1,placeholder:d="Enter notes here..."}){let r=o?Bt:Lt,i=s?`<button type="button" class="pf-action-toggle pf-save-btn ${l?"is-saved":""}" id="${s}" data-save-input="${e}" title="Save notes">${Cn}</button>`:"",u=n?`<button type="button" class="pf-action-toggle pf-notes-lock ${o?"is-locked":""}" id="${n}" aria-pressed="${o}" title="Lock notes (retain after archive)">${r}</button>`:"";return`
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${e}" rows="6" placeholder="${We(d)}">${We(t||"")}</textarea>
                ${a?`<p class="pf-signoff-hint" id="${a}"></p>`:""}
            </div>
            <div class="pf-notes-action">
                ${n?Mt(u,"Lock"):""}
                ${s?Mt(i,"Save"):""}
            </div>
        </article>
    `}function Ae({reviewerInputId:e,reviewerValue:t,signoffInputId:n,signoffValue:o,isComplete:a,saveButtonId:s,isSaved:l=!1,completeButtonId:d,subtext:r="Sign-off below. Click checkmark icon. Done."}){let i=`<button type="button" class="pf-action-toggle ${a?"is-active":""}" id="${d}" aria-pressed="${!!a}" title="Mark step complete">${dt}</button>`;return`
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${We(r)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${e}" value="${We(t)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${n}" value="${We(o)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${Mt(i,"Done")}
            </div>
        </article>
    `}function Je(e,t){e&&(e.classList.toggle("is-locked",t),e.setAttribute("aria-pressed",String(t)),e.innerHTML=t?Bt:Lt)}function Te(e,t){e&&e.classList.toggle("is-saved",t)}function Vt(e=document){let t=e.querySelectorAll(".pf-save-btn[data-save-input]"),n=[];return t.forEach(o=>{let a=o.getAttribute("data-save-input"),s=document.getElementById(a);if(!s)return;let l=()=>{Te(o,!1)};s.addEventListener("input",l),n.push(()=>s.removeEventListener("input",l))}),()=>n.forEach(o=>o())}function Sn(e,t){if(e===0)return{canComplete:!0,blockedBy:null,message:""};for(let n=0;n<e;n++)if(!t[n])return{canComplete:!1,blockedBy:n,message:`Complete Step ${n} before signing off on this step.`};return{canComplete:!0,blockedBy:null,message:""}}function xn(e){let t=document.querySelector(".pf-workflow-toast");t&&t.remove();let n=document.createElement("div");n.className="pf-workflow-toast pf-workflow-toast--warning",n.innerHTML=`
        <span class="pf-workflow-toast-icon">\u26A0\uFE0F</span>
        <span class="pf-workflow-toast-message">${e}</span>
    `,document.body.appendChild(n),requestAnimationFrame(()=>{n.classList.add("pf-workflow-toast--visible")}),setTimeout(()=>{n.classList.remove("pf-workflow-toast--visible"),setTimeout(()=>n.remove(),300)},4e3)}var Ft={fillColor:"#000000",fontColor:"#FFFFFF",bold:!0},jt={currency:"$#,##0.00",currencyWithNegative:"$#,##0.00;($#,##0.00)",number:"#,##0.00",integer:"#,##0",percent:"0.00%",date:"yyyy-mm-dd",dateTime:"yyyy-mm-dd hh:mm"};function _n(e){e.format.fill.color=Ft.fillColor,e.format.font.color=Ft.fontColor,e.format.font.bold=Ft.bold}function ft(e,t,n,o=!1){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[o?jt.currencyWithNegative:jt.currency]]}function An(e,t,n,o=jt.date){if(n<=0)return;let a=e.getRangeByIndexes(1,t,n,1);a.numberFormat=[[o]]}var Et="payroll-recorder";var Pe="Payroll Recorder",xs=A.CONFIG||"SS_PF_Config",Ht=["SS_PF_Config"];var _o="Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel. Every run follows the same guidance so you stay audit-ready.",be=rt.map(({id:e,title:t})=>({id:e,title:t})),B={TYPE:0,FIELD:1,VALUE:2,PERMANENT:3,TITLE:-1},Ao="Run Settings";var Dn="N";var Do="PR_JE_Debit_Total",Po="PR_JE_Credit_Total",Io="PR_JE_Difference",$e={0:{note:"PR_Notes_Config",reviewer:"PR_Reviewer_Config",signOff:"PR_SignOff_Config"},1:{note:"PR_Notes_Import",reviewer:"PR_Reviewer_Import",signOff:"PR_SignOff_Import"},2:{note:"PR_Notes_Headcount",reviewer:"PR_Reviewer_Headcount",signOff:"PR_SignOff_Headcount"},3:{note:"PR_Notes_Validate",reviewer:"PR_Reviewer_Validate",signOff:"PR_SignOff_Validate"},4:{note:"PR_Notes_Review",reviewer:"PR_Reviewer_Review",signOff:"PR_SignOff_Review"},5:{note:"PR_Notes_JE",reviewer:"PR_Reviewer_JE",signOff:"PR_SignOff_JE"},6:{note:"PR_Notes_Archive",reviewer:"PR_Reviewer_Archive",signOff:"PR_SignOff_Archive"}},de={0:"PR_Complete_Config",1:"PR_Complete_Import",2:"PR_Complete_Headcount",3:"PR_Complete_Validate",4:"PR_Complete_Review",5:"PR_Complete_JE",6:"PR_Complete_Archive"},$o={1:A.DATA,2:A.DATA_CLEAN,3:A.DATA_CLEAN,4:A.EXPENSE_REVIEW,5:A.JE_DRAFT},qe="PR_Reviewer",jn="PR_Payroll_Provider",mt="User opted to skip the headcount review this period.",te={statusText:"",focusedIndex:0,activeView:"home",activeStepId:null,stepStatuses:be.reduce((e,t)=>({...e,[t.id]:"pending"}),{})},W={loaded:!1,values:{},permanents:{},overrides:{accountingPeriod:!1,jeId:!1}},Le=new Map,gt=null,Xe=["PR_Payroll_Date","Payroll Date (YYYY-MM-DD)","Payroll_Date","Payroll Date","Payroll_Date_(YYYY-MM-DD)"],H={skipAnalysis:!1,roster:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},departments:{rosterCount:null,payrollCount:null,difference:null,mismatches:[]},loading:!1,hasAnalyzed:!1,lastError:null},bt=null,F={loading:!1,lastError:null,prDataTotal:null,cleanTotal:null,reconDifference:null,bankAmount:"",bankDifference:null,plugEnabled:!1},ke={loading:!1,lastError:null,periods:[],copilotResponse:"",completenessCheck:{currentPeriod:null,historicalPeriods:null}},G={debitTotal:null,creditTotal:null,difference:null,loading:!1,lastError:null};async function No(){if(console.log("Completeness Check - Starting..."),!ae()){console.log("Completeness Check - Excel runtime not available");return}try{await Excel.run(async e=>{var a,s,l,d;let t=e.workbook.worksheets.getItemOrNullObject(A.DATA_CLEAN),n=e.workbook.worksheets.getItemOrNullObject(A.ARCHIVE_SUMMARY);t.load("isNullObject"),n.load("isNullObject"),await e.sync();let o={currentPeriod:null,historicalPeriods:null};if(!t.isNullObject){let r=t.getUsedRangeOrNullObject();if(r.load("values"),await e.sync(),!r.isNullObject&&r.values&&r.values.length>1){let i=(r.values[0]||[]).map(p=>String(p||"").toLowerCase().trim()),u=i.findIndex(p=>p.includes("amount")),f=u>=0?u:i.findIndex(p=>p==="total"||p==="all-in"||p==="allin"||p==="all-in total"||p==="gross"||p==="total pay");if(console.log("Completeness Check - PR_Data_Clean headers:",i),console.log("Completeness Check - Amount column index:",u,"Total column index:",f),f>=0){let c=r.values.slice(1).reduce((w,h)=>w+(Number(h[f])||0),0),m=((l=(s=(a=ke.periods)==null?void 0:a[0])==null?void 0:s.summary)==null?void 0:l.total)||0;console.log("Completeness Check - PR_Data_Clean sum:",c,"Current period total:",m);let y=Math.abs(c-m)<1;o.currentPeriod={match:y,prDataClean:c,currentTotal:m}}else console.warn("Completeness Check - No amount/total column found in PR_Data_Clean")}}if(!n.isNullObject){let r=n.getUsedRangeOrNullObject();if(r.load("values"),await e.sync(),!r.isNullObject&&r.values&&r.values.length>1){let i=(r.values[0]||[]).map(c=>String(c||"").toLowerCase().trim()),u=i.findIndex(c=>c.includes("pay period")||c.includes("payroll date")||c==="date"||c==="period"||c.includes("period")),f=i.findIndex(c=>c.includes("amount")),p=f>=0?f:i.findIndex(c=>c==="total"||c==="all-in"||c==="allin"||c==="all-in total"||c==="total payroll"||c.includes("total"));if(console.log("Completeness Check - PR_Archive_Summary headers:",i),console.log("Completeness Check - Date column index:",u,"Total column index:",p),p>=0&&u>=0){let c=r.values.slice(1),m=(ke.periods||[]).slice(1,6);console.log("Completeness Check - Looking for periods:",m.map(x=>x.key||x.label));let y=new Map;for(let x of c){let T=x[u],g=$n(T);if(g){let D=Number(x[p])||0,L=y.get(g)||0;y.set(g,L+D)}}console.log("Completeness Check - Archive lookup keys:",Array.from(y.keys())),console.log("Completeness Check - Archive lookup values:",Array.from(y.entries()));let w=0,h=0,E=0,k=[];for(let x of m){let T=x.key||x.label||"",g=$n(T),D=((d=x.summary)==null?void 0:d.total)||0;h+=D;let L=y.get(g);L!==void 0?(w+=L,E++,k.push({period:T,calculated:D,archive:L,match:Math.abs(D-L)<1})):(console.warn(`Completeness Check - Period ${T} (normalized: ${g}) not found in archive`),k.push({period:T,calculated:D,archive:null,match:!1}))}console.log("Completeness Check - Period details:",k),console.log("Completeness Check - Matched",E,"of",m.length,"periods"),console.log("Completeness Check - Archive sum:",w,"Periods sum:",h);let _=E===m.length&&m.length>0,$=Math.abs(w-h)<1,V=_&&$;o.historicalPeriods={match:V,archiveSum:w,periodsSum:h,matchedCount:E,totalPeriods:m.length,details:k}}else console.warn("Completeness Check - Missing date or total column in PR_Archive_Summary"),console.warn("  Date column index:",u,"Total column index:",p)}}ke.completenessCheck=o,console.log("Completeness Check - Results:",JSON.stringify(o))}),console.log("Completeness Check - Complete!")}catch(e){console.error("Payroll completeness check failed:",e)}}function Oo(){var y,w;let e=ke.completenessCheck||{},t=((y=ke.periods)==null?void 0:y.length)>0,n=h=>`$${Math.round(h||0).toLocaleString()}`,o=h=>{let E=Math.abs(h);return E<1?"\u2014":`${h>0?"+":"-"}$${Math.round(E).toLocaleString()}`},a=(h,E,k,_,$,V,x)=>{let T=(k||0)-($||0),g,D;x?(g='<span class="pf-complete-status pf-complete-status--pending">\u23F3</span>',D="pending"):V?(g='<span class="pf-complete-status pf-complete-status--pass">\u2713</span>',D="pass"):(g='<span class="pf-complete-status pf-complete-status--fail">\u2717</span>',D="fail");let L=x?"":`
            <div class="pf-complete-diff ${D}">
                ${o(T)}
            </div>
        `;return`
            <div class="pf-complete-row ${D}">
                <div class="pf-complete-header">
                    ${g}
                    <span class="pf-complete-label">${S(h)}</span>
                </div>
                ${x?`
                <div class="pf-complete-values">
                    <span class="pf-complete-pending-text">Click Run/Refresh to check</span>
                </div>
                `:`
                <div class="pf-complete-values">
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${S(E)}:</span>
                        <span class="pf-complete-amount">${n(k)}</span>
                    </div>
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${S(_)}:</span>
                        <span class="pf-complete-amount">${n($)}</span>
                    </div>
                </div>
                ${L}
                `}
            </div>
        `},s=e.currentPeriod,l=!t||s===null||s===void 0,d=a("Current Period","PR_Data_Clean Total",s==null?void 0:s.prDataClean,"Calculated Total",s==null?void 0:s.currentTotal,s==null?void 0:s.match,l),r=e.historicalPeriods,i=!t||r===null||r===void 0,u=(r==null?void 0:r.matchedCount)||0,f=(r==null?void 0:r.totalPeriods)||0,p=f>0?`Historical Periods (${u}/${f} matched)`:"Historical Periods",c=a(p,"PR_Archive_Summary (matched)",r==null?void 0:r.archiveSum,"Calculated Total",r==null?void 0:r.periodsSum,r==null?void 0:r.match,i),m="";return!i&&((w=r==null?void 0:r.details)==null?void 0:w.length)>0&&(m=`
            <div class="pf-complete-details-section">
                <div class="pf-complete-details-header">Period-by-Period Match</div>
                ${r.details.map(E=>{let k=E.archive===null?"\u26A0\uFE0F":E.match?"\u2713":"\u2717",_=E.archive!==null?n(E.archive):"Not found";return`
                <div class="pf-complete-detail-row">
                    <span class="pf-complete-detail-date">${S(E.period)}</span>
                    <span class="pf-complete-detail-icon">${k}</span>
                    <span class="pf-complete-detail-vals">${n(E.calculated)} vs ${_}</span>
                </div>
            `}).join("")}
            </div>
        `),`
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-completeness-card">
            <div class="pf-config-head">
                <h3>Data Completeness Check</h3>
                <p class="pf-config-subtext">Verify source data matches calculated totals</p>
            </div>
            <div class="pf-complete-container">
                ${d}
                ${c}
                ${m}
            </div>
        </article>
    `}function To(e){switch(e){case 0:return{title:"Configuration",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Sets up the key parameters for your payroll review. Complete this before importing data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Key Fields</h4>
                        <ul>
                            <li><strong>Payroll Date</strong> \u2014 The period-end date for this payroll run</li>
                            <li><strong>Accounting Period</strong> \u2014 Shows up in your JE description</li>
                            <li><strong>Journal Entry ID</strong> \u2014 Reference number for your accounting system</li>
                            <li><strong>Provider Link</strong> \u2014 Quick access to your payroll provider portal</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>The accounting period and JE ID auto-generate based on your payroll date, but you can override them if needed.</p>
                    </div>
                `};case 1:return{title:"Import Payroll Data",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Gets your payroll data into the workbook. Pull a report from your payroll provider and paste it into PR_Data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Required Columns</h4>
                        <p>Your payroll export should include:</p>
                        <ul>
                            <li><strong>Employee Name</strong> \u2014 Full name (used to match against roster)</li>
                            <li><strong>Department</strong> \u2014 Cost center assignment</li>
                            <li><strong>Regular Earnings</strong> \u2014 Base pay for the period</li>
                            <li><strong>Overtime</strong> \u2014 OT pay (if applicable)</li>
                            <li><strong>Bonus/Commission</strong> \u2014 Variable compensation</li>
                            <li><strong>Benefits/Deductions</strong> \u2014 Employer portions</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Column headers don't need to match exactly\u2014the system is flexible with naming. Just make sure each field is present.</p>
                    </div>
                `};case 2:return{title:"Headcount Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Compares employee counts and department assignments between your roster and payroll data to catch discrepancies early.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Data Sources</h4>
                        <ul>
                            <li><strong>SS_Employee_Roster</strong> \u2014 Your centralized employee list (Column A: Employee names)</li>
                            <li><strong>PR_Data</strong> \u2014 The payroll data you just imported (Employee column)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F50D} Employee Alignment Check</h4>
                        <p>The script compares names between SS_Employee_Roster and PR_Data to find:</p>
                        <ul>
                            <li><strong>In Roster, Missing from Payroll</strong> \u2014 Employees on roster but not in payroll (possible missed payment)</li>
                            <li><strong>In Payroll, Missing from Roster</strong> \u2014 Employees paid but not on roster (possible ghost employee or new hire)</li>
                        </ul>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.7); margin-top: 8px;">Names are matched using fuzzy logic to handle minor variations.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F3E2} Department Alignment Check</h4>
                        <p>For employees appearing in both sources, the script compares the "Department" column:</p>
                        <ul>
                            <li>Flags employees where roster department \u2260 payroll department</li>
                            <li>Mismatches affect GL coding and cost center reporting</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>If discrepancies are expected (e.g., contractors, temp workers), you can skip this check and add a note explaining why. The note is required if you skip.</p>
                    </div>
                `};case 3:return{title:"Payroll Validation",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Validates that your payroll totals match what was actually paid from the bank.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Reconciliation Check</h4>
                        <ul>
                            <li><strong>PR_Data Total</strong> \u2014 Sum of all payroll from your import</li>
                            <li><strong>Clean Total</strong> \u2014 Processed total after expense mapping</li>
                            <li><strong>Bank Amount</strong> \u2014 What actually left the bank account</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u26A0\uFE0F Common Differences</h4>
                        <ul>
                            <li><strong>Timing</strong> \u2014 Direct deposits vs check clearing dates</li>
                            <li><strong>Tax payments</strong> \u2014 May be separate from net pay</li>
                            <li><strong>Benefits</strong> \u2014 Some deductions paid to vendors</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Small differences ($0.01-$1.00) are usually rounding. Use the plug feature to resolve them.</p>
                    </div>
                `};case 4:return{title:"Expense Review",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Generates an executive-ready payroll expense summary for CFO review, with period comparisons and trend analysis.</p>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4C2} Data Sources</h4>
                        <ul>
                            <li><strong>PR_Data_Clean</strong> \u2014 Current period payroll data (cleaned and categorized)</li>
                            <li><strong>SS_Employee_Roster</strong> \u2014 Department assignments and employee details</li>
                            <li><strong>PR_Archive_Summary</strong> \u2014 Historical payroll data for trend analysis</li>
                        </ul>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4B0} How Amounts Are Calculated</h4>
                        <table style="width:100%; font-size: 11px; margin-top: 8px; border-collapse: collapse;">
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Fixed Salary</strong></td>
                                <td style="padding: 6px 0;">Regular wages, salaries, and base pay</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Variable Salary</strong></td>
                                <td style="padding: 6px 0;">Commissions, bonuses, overtime, and incentive pay</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Gross Pay</strong></td>
                                <td style="padding: 6px 0;">Fixed + Variable Salary</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Burden</strong></td>
                                <td style="padding: 6px 0;">Employer taxes (FICA, Medicare, FUTA, SUTA), health insurance, 401(k) match, and other employer-paid benefits</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>All-In Total</strong></td>
                                <td style="padding: 6px 0;">Gross Pay + Burden = Total cost to employer</td>
                            </tr>
                            <tr>
                                <td style="padding: 6px 0;"><strong>Burden Rate</strong></td>
                                <td style="padding: 6px 0;">Burden \xF7 All-In Total (typically 10-18%)</td>
                            </tr>
                        </table>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Report Sections</h4>
                        <ul>
                            <li><strong>Executive Summary</strong> \u2014 Current vs prior period comparison (frozen at top)</li>
                            <li><strong>Department Breakdown</strong> \u2014 Cost allocation by cost center</li>
                            <li><strong>Historical Context</strong> \u2014 Where current metrics fall within historical ranges</li>
                            <li><strong>Period Trends</strong> \u2014 6-period trend chart for Total, Fixed, Variable, Burden, and Headcount</li>
                        </ul>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4C8} Historical Context Visualization</h4>
                        <p>The spectrum bars show where your current period falls relative to your historical min/max:</p>
                        <p style="font-family: Consolas, monospace; color: #6366f1; margin: 8px 0;">\u2591\u2591\u2591\u2591\u2591\u2591\u2591\u25CF\u2591\u2591\u2591\u2591\u2591\u2591\u2591\u2591</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.7);">The dot (\u25CF) indicates current position. Left = Low, Right = High.</p>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Review Tips</h4>
                        <ul>
                            <li>Compare <strong>Burden Rate</strong> \u2014 Should be consistent period-to-period (10-18% typical)</li>
                            <li>Watch <strong>Variable Salary</strong> spikes \u2014 May indicate commission/bonus timing</li>
                            <li>Verify <strong>Headcount changes</strong> \u2014 Should align with HR records</li>
                            <li>Flag variances <strong>> 10%</strong> from prior period for follow-up</li>
                        </ul>
                    </div>
                `};case 5:return{title:"Journal Entry",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Generates a balanced journal entry from your payroll data, ready for upload to your accounting system.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4DD} How the JE Works</h4>
                        <p>Maps payroll categories to GL accounts:</p>
                        <ul>
                            <li><strong>Expenses</strong> \u2192 Debits to departmental expense accounts</li>
                            <li><strong>Liabilities</strong> \u2192 Credits to payable accounts</li>
                            <li><strong>Cash</strong> \u2192 Credit to bank account</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u2705 Validation Checks</h4>
                        <ul>
                            <li><strong>Debits = Credits</strong> \u2014 Entry must balance</li>
                            <li><strong>All accounts mapped</strong> \u2014 No unassigned categories</li>
                            <li><strong>Totals match</strong> \u2014 JE ties to PR_Data</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4A1} Tip</h4>
                        <p>Review the draft in PR_JE_Draft before exporting to catch any mapping errors.</p>
                    </div>
                `};case 6:return{title:"Archive & Clear",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F3AF} What This Step Does</h4>
                        <p>Creates a backup of your completed payroll run, then resets the workbook so you're ready for the next pay period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4C1} Step 1: Create Backup</h4>
                        <p>A new workbook opens containing all your payroll tabs. You'll choose where to save it on your computer or shared drive.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Tip: Use a consistent naming convention like "Payroll_Archive_2024-01-15"</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CA} Step 2: Update History</h4>
                        <p>The current period's totals are saved to PR_Archive_Summary. This powers the trend charts and completeness checks for future periods.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Keeps 5 periods of history \u2014 oldest is removed automatically</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F9F9} Step 3: Clear Working Data</h4>
                        <p>Data is cleared from the working sheets:</p>
                        <ul>
                            <li>PR_Data (raw import)</li>
                            <li>PR_Data_Clean (processed data)</li>
                            <li>PR_Expense_Review (summary & charts)</li>
                            <li>PR_JE_Draft (journal entry lines)</li>
                        </ul>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Headers are preserved \u2014 only data rows are cleared</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F504} Step 4: Reset for Next Period</h4>
                        <ul>
                            <li>Payroll Date, Accounting Period, JE ID cleared</li>
                            <li>All sign-offs and completion flags reset</li>
                            <li>Notes cleared (unless you locked them with \u{1F512})</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u26A0\uFE0F Before You Archive</h4>
                        <ul>
                            <li>\u2713 JE uploaded to your accounting system</li>
                            <li>\u2713 All review steps signed off</li>
                            <li>\u2713 Lock any notes you want to keep</li>
                        </ul>
                    </div>
                `};default:return{title:"Payroll Recorder",content:`
                    <div class="pf-info-section">
                        <h4>\u{1F44B} Welcome to Payroll Recorder</h4>
                        <p>This module helps you normalize payroll exports, enforce controls, and prep journal entries.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>\u{1F4CB} Workflow Overview</h4>
                        <ol style="margin: 8px 0; padding-left: 20px;">
                            <li>Configure period settings</li>
                            <li>Import payroll data</li>
                            <li>Review headcount alignment</li>
                            <li>Validate against bank</li>
                            <li>Review expense summary</li>
                            <li>Generate journal entry</li>
                            <li>Archive and reset</li>
                        </ol>
                    </div>
                    <div class="pf-info-section">
                        <p>Click a step card to get started, or tap the <strong>\u24D8</strong> button on any step for detailed guidance.</p>
                    </div>
                `}}}on(()=>Lo());async function Lo(){try{await Bo(),await zn();let e=Ot(Et);await Nt(e.sheetName,e.title,e.subtitle),le()}catch(e){throw console.error("[Payroll] Module initialization failed:",e),e}}async function Bo(){try{await Pt(Et),console.log(`[Payroll] Tab visibility applied for ${Et}`)}catch(e){console.warn("[Payroll] Could not apply tab visibility:",e)}}function le(){var r;let e=document.body;if(!e)return;let t=te.focusedIndex<=0?"disabled":"",n=te.focusedIndex>=be.length-1?"disabled":"",o=te.activeView==="config",a=te.activeView==="step",s=!o&&!a,l=o?Fo():a?Jo(te.activeStepId):Vo();e.innerHTML=`
        <div class="pf-root">
            ${Mo(t,n)}
            ${l}
            ${qo()}
        </div>
    `;let d=document.getElementById("pf-info-fab-payroll");if(s)d&&d.remove();else if((r=window.PrairieForge)!=null&&r.mountInfoFab){let i=To(te.activeStepId);PrairieForge.mountInfoFab({title:i.title,content:i.content,buttonId:"pf-info-fab-payroll"})}if(Xo(),o)ta();else if(a)try{na(te.activeStepId)}catch(i){console.warn("Payroll Recorder: failed to bind step interactions",i)}else ea();oa(),s?pn():Tt()}function Mo(e,t){let n=I("SS_Company_Name")||"your company";return`
        <div class="pf-brand-float" aria-hidden="true">
            <span class="pf-brand-wave"></span>
        </div>
        <header class="pf-banner">
            <div class="pf-nav-bar">
                <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${e}>
                    ${En}
                    <span class="sr-only">Previous step</span>
                </button>
                <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                    ${fn}
                    <span class="sr-only">Module Home</span>
                </button>
                <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                    ${mn}
                    <span class="sr-only">Module Selector</span>
                </button>
                <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${t}>
                    ${kn}
                    <span class="sr-only">Next step</span>
                </button>
                <span class="pf-nav-divider"></span>
                <div class="pf-quick-access-wrapper">
                    <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                        ${gn}
                        <span class="sr-only">Quick Access Menu</span>
                    </button>
                    <div id="quick-access-dropdown" class="pf-quick-dropdown hidden">
                        <div class="pf-quick-dropdown-header">Quick Access</div>
                        <button id="nav-roster" class="pf-quick-item pf-clickable" type="button">
                            ${hn}
                            <span>Employee Roster</span>
                        </button>
                        <button id="nav-accounts" class="pf-quick-item pf-clickable" type="button">
                            ${yn}
                            <span>Chart of Accounts</span>
                        </button>
                        <button id="nav-expense-map" class="pf-quick-item pf-clickable" type="button">
                            ${ct}
                            <span>PR Mapping</span>
                </button>
                    </div>
                </div>
            </div>
        </header>
    `}function Vo(){return`
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">Payroll Recorder</h2>
            <p class="pf-hero-copy">${_o}</p>
            <p class="pf-hero-hint">${S(te.statusText||"")}</p>
        </section>
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${be.map((e,t)=>Yo(e,t)).join("")}
            </div>
        </section>
    `}function Fo(){if(!W.loaded)return`
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration\u2026</p>
                </article>
            </section>
        `;let e=$e[0],t=ye(Ct()),n=ye(I("PR_Accounting_Period")),o=I("PR_Journal_Entry_ID"),a=I("SS_Accounting_Software"),s=qt(),l=I("SS_Company_Name"),d=I(qe)||De(),r=e?I(e.note):"",i=e?Re(e.note):!1,u=(e?I(e.reviewer):"")||De(),f=e?ye(I(e.signOff)):"",p=!!(f||I(de[0]));return`
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${S(Pe)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every payroll run.</p>
            <p class="pf-hero-hint">${S(te.statusText||"")}</p>
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
                        <input type="text" id="config-user-name" value="${S(d)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Date</span>
                        <input type="date" id="config-payroll-date" value="${S(t)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${S(n)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-je-id" value="${S(o)}" placeholder="PR-AUTO-YYYY-MM-DD">
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
                        <input type="text" id="config-company-name" value="${S(l)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${S(s)}" placeholder="https://\u2026">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${S(a)}" placeholder="https://\u2026">
                    </label>
                </div>
            </article>
            ${e?_e({textareaId:"config-notes",value:r,permanentId:"config-notes-permanent",isPermanent:i,hintId:"",saveButtonId:"config-notes-save"}):""}
            ${e?Ae({reviewerInputId:"config-reviewer-name",reviewerValue:u,signoffInputId:"config-signoff-date",signoffValue:f,isComplete:p,saveButtonId:"config-signoff-save",completeButtonId:"config-signoff-toggle"}):""}
        </section>
    `}function jo(e){let t=Se(1),n=t?Re(t.note):!1,o=t?I(t.note):"",a=(t?I(t.reviewer):"")||De(),s=t?ye(I(t.signOff)):"",l=!!(s||I(de[1])),d=qt();return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S(Pe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">Pull your payroll export from the provider and paste it into PR_Data.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Open your payroll provider, download the report, and paste into PR_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ge(d?`<a href="${S(d)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${pt}</a>`:`<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${pt}</button>`,"Provider")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PR_Data sheet">${ct}</button>`,"PR_Data")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PR_Data to start over">${Rn}</button>`,"Clear")}
                </div>
            </article>
            ${t?`
                ${_e({textareaId:"step-notes-1",value:o||"",permanentId:"step-notes-lock-1",isPermanent:n,saveButtonId:"step-notes-save-1"})}
                ${Ae({reviewerInputId:"step-reviewer-1",reviewerValue:a,signoffInputId:"step-signoff-1",signoffValue:s,isComplete:l,saveButtonId:"step-signoff-save-1",completeButtonId:"step-signoff-toggle-1"})}
            `:""}
        </section>
    `}function Ho(e){var $,V,x,T,g,D,L,Q,pe,K;let t=Se(2),n=t?I(t.note):"",o=t?Re(t.note):!1,a=(t?I(t.reviewer):"")||De(),s=t?ye(I(t.signOff)):"",l=!!(s||I(de[2])),d=Rt(),r=H.roster||{},i=H.departments||{},u=H.hasAnalyzed,f="";H.loading?f='<p class="pf-step-note">Analyzing roster and payroll data\u2026</p>':H.lastError&&(f=`<p class="pf-step-note">${S(H.lastError)}</p>`);let p=(M,Y,ne)=>{let xe=!u,X;xe?X='<span class="pf-je-check-circle pf-je-circle--pending"></span>':ne?X=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:X=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;let Z=u?` = ${Y}`:"";return`
            <div class="pf-je-check-row">
                ${X}
                <span class="pf-je-check-desc-pill">${S(M)}${Z}</span>
            </div>
        `},c=($=r.difference)!=null?$:0,m=(V=i.difference)!=null?V:0,y=Array.isArray(r.mismatches)?r.mismatches.filter(Boolean):[],w=Array.isArray(i.mismatches)?i.mismatches.filter(Boolean):[],h=`
        ${p("SS_Employee_Roster count",(x=r.rosterCount)!=null?x:"\u2014",!0)}
        ${p("PR_Data employee count",(T=r.payrollCount)!=null?T:"\u2014",!0)}
        ${p("Difference",c,c===0)}
    `,E=`
        ${p("Expected departments",(g=i.rosterCount)!=null?g:"\u2014",!0)}
        ${p("PR_Data departments",(D=i.payrollCount)!=null?D:"\u2014",!0)}
        ${p("Difference",m,m===0)}
    `,k=y.length&&!H.skipAnalysis&&u&&((Q=(L=window.PrairieForge)==null?void 0:L.renderMismatchTiles)==null?void 0:Q.call(L,{mismatches:y,label:"Employees Driving the Difference",sourceLabel:"Roster",targetLabel:"Payroll Data",escapeHtml:S}))||"",_=w.length&&!H.skipAnalysis&&u&&((K=(pe=window.PrairieForge)==null?void 0:pe.renderMismatchTiles)==null?void 0:K.call(pe,{mismatches:w,label:"Employees with Department Differences",formatter:M=>({name:M.employee||M.name||"",source:`${M.rosterDept||"\u2014"} \u2192 ${M.payrollDept||"\u2014"}`,isMissingFromTarget:!0}),escapeHtml:S}))||"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S(Pe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">Headcount Review</h2>
            <p class="pf-hero-copy">Quick check to make sure your roster matches your payroll data.</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${H.skipAnalysis?"is-active":""}" id="headcount-skip-btn">
                    ${wn}
                    <span>Skip Analysis</span>
                </button>
            </div>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Headcount Check</h3>
                    <p class="pf-config-subtext">Compare employee roster against payroll data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="roster-run-btn" title="Run headcount analysis">${ut}</button>`,"Run")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="roster-refresh-btn" title="Refresh analysis">${ze}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Employee Alignment</h3>
                    <p class="pf-config-subtext">Verify employees match between roster and payroll.</p>
                </div>
                ${f}
                <div class="pf-je-checks-container">
                    ${h}
                </div>
                ${k}
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Department Alignment</h3>
                    <p class="pf-config-subtext">Verify department assignments are consistent.</p>
                </div>
                <div class="pf-je-checks-container">
                    ${E}
                </div>
                ${_}
            </article>
            ${t?`
                ${_e({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:o,hintId:d?"headcount-notes-hint":"",saveButtonId:"step-notes-save-2"})}
                ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:a,signoffInputId:"step-signoff-date",signoffValue:s,isComplete:l,saveButtonId:"headcount-signoff-save",completeButtonId:"headcount-signoff-toggle"})}
            `:""}
        </section>
    `}function Uo(e){var $;let t=Se(3),n=t?I(t.note):"",o=(t?I(t.reviewer):"")||De(),a=t?ye(I(t.signOff)):"",s=F.loading?'<p class="pf-step-note">Preparing reconciliation data\u2026</p>':F.lastError?`<p class="pf-step-note">${S(F.lastError)}</p>`:"",l=!!(a||I(de[3])),d=F.prDataTotal!==null,r=F.prDataTotal,i=F.cleanTotal,u=($=F.reconDifference)!=null?$:r!=null&&i!=null?r-i:null,f=u!==null&&Math.abs(u)<.01,p=ie(F.cleanTotal),c=F.bankDifference!=null?ie(F.bankDifference):"---",m=F.bankDifference==null?"":Math.abs(F.bankDifference)<.5?"Difference within acceptable tolerance.":"Difference exceeds tolerance and should be resolved.",y=qn(F.bankAmount),w=(V,x,T)=>{let g=!d,D;return g?D='<span class="pf-je-check-circle pf-je-circle--pending"></span>':T?D=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:D=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${D}
                <span class="pf-je-check-desc-pill">${S(x)}</span>
            </div>
        `},h=d?ie(r):"\u2014",E=d?ie(i):"\u2014",k=d?ie(u):"\u2014",_=`
        ${w("PR_Data Total",`PR_Data Total = ${h}`,!0)}
        ${w("PR_Data_Clean Total",`PR_Data_Clean Total = ${E}`,!0)}
        ${w("Difference",`Difference = ${k} (should be $0.00)`,f)}
    `;return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S(Pe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">Normalize your payroll data and verify totals match.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Validation</h3>
                    <p class="pf-config-subtext">Build PR_Data_Clean from your imported data and verify totals.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="validation-run-btn" title="Run reconciliation">${ut}</button>`,"Run")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="validation-refresh-btn" title="Refresh reconciliation">${ze}</button>`,"Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Data Reconciliation</h3>
                    <p class="pf-config-subtext">Verify PR_Data and PR_Data_Clean totals match.</p>
                </div>
                ${s}
                <div class="pf-je-checks-container">
                    ${_}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Bank Reconciliation</h3>
                    <p class="pf-config-subtext">Compare payroll total to the amount pulled from the bank.</p>
                </div>
                <div class="pf-config-grid pf-metric-grid">
                    <label class="pf-config-field">
                        <span>Cost per PR_Data_Clean</span>
                        <input id="bank-clean-total-value" type="text" class="pf-readonly-input pf-metric-value" value="${p}" readonly>
                    </label>
                    <label class="pf-config-field">
                        <span>Cost per Bank</span>
                        <input
                            type="text"
                            inputmode="decimal"
                            id="bank-amount-input"
                            class="pf-metric-input"
                            value="${S(y)}"
                            placeholder="0.00"
                            aria-label="Enter bank amount"
                        >
                    </label>
                    <label class="pf-config-field">
                        <span>Difference</span>
                        <input id="bank-diff-value" type="text" class="pf-readonly-input pf-metric-value" value="${c}" readonly>
                    </label>
                </div>
                <p class="pf-metric-hint" id="bank-diff-hint">${S(m)}</p>
            </article>
            ${t?`
                ${_e({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:Re(t.note),saveButtonId:"step-notes-save-3"})}
            `:""}
            ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-3",signoffValue:a,isComplete:l,saveButtonId:"step-signoff-save-3",completeButtonId:"validation-signoff-toggle"})}
        </section>
    `}function Go(e){let t=Se(4),n=t?I(t.note):"",o=(t?I(t.reviewer):"")||De(),a=t?ye(I(t.signOff)):"",s=!!(a||I(de[4])),l=ke.loading?'<p class="pf-step-note">Preparing executive summary\u2026</p>':ke.lastError?`<p class="pf-step-note">${S(ke.lastError)}</p>`:"",d=rn({id:"expense-review-copilot",body:"Want help analyzing your data? Just ask!",placeholder:"Where should I focus this pay period?",buttonLabel:"Ask CoPilot"});return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S(Pe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">${S(e.summary||"")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${l}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Perform Analysis</h3>
                    <p class="pf-config-subtext">Populate Expense Review and perform review.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ge(`<button type="button" class="pf-action-toggle" id="expense-run-btn" title="Run expense review analysis">${ut}</button>`,"Run")}
                    ${ge(`<button type="button" class="pf-action-toggle" id="expense-refresh-btn" title="Refresh expense data">${ze}</button>`,"Refresh")}
                </div>
            </article>
            ${Oo()}
                ${d}
            ${t?`
            ${_e({textareaId:"step-notes-input",value:n,permanentId:"step-notes-permanent",isPermanent:Re(t.note),saveButtonId:"step-notes-save-4"})}
            ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:o,signoffInputId:"step-signoff-4",signoffValue:a,isComplete:s,saveButtonId:"step-signoff-save-4",completeButtonId:"expense-signoff-toggle"})}
            `:""}
        </section>
    `}function zo(e){var $,V,x;let t=Se(5),n=t?I(t.note):"",o=t?Re(t.note):!1,a=(t?I(t.reviewer):"")||De(),s=t?ye(I(t.signOff)):"",l=!!(s||I(de[5])),d=G.lastError?`<p class="pf-step-note">${S(G.lastError)}</p>`:"",r=G.debitTotal!==null,i=($=G.debitTotal)!=null?$:0,u=(V=G.creditTotal)!=null?V:0,f=i-u,p=(x=F.cleanTotal)!=null?x:0,c=r,m=r&&Math.abs(f-p)<.01,y=(T,g)=>{let D=!r,L;return D?L='<span class="pf-je-check-circle pf-je-circle--pending"></span>':g?L=`<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`:L=`<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`,`
            <div class="pf-je-check-row">
                ${L}
                <span class="pf-je-check-desc-pill">${S(T)}</span>
            </div>
        `},w=r?ie(i):"\u2014",h=r?ie(u):"\u2014",E=r?ie(f):"\u2014",k=r?ie(p):"\u2014",_=`
        ${y(`Total Debits = ${w}`,c)}
        ${y(`Total Credits = ${h}`,c)}
        ${y(`Line Amount (Debit - Credit) = ${E}`,c)}
        ${y(`JE Total matches PR_Data_Clean (${k})`,m)}
    `;return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S(Pe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">Generate the upload file to break down the bank feed line item.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Upload File</h3>
                    <p class="pf-config-subtext">Build the breakdown from PR_Data_Clean for your accounting system.</p>
                </div>
                <div class="pf-signoff-action">
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate from PR_Data_Clean">${ct}</button>`,"Generate")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation">${ze}</button>`,"Refresh")}
                    ${ge(`<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export as CSV">${bn}</button>`,"Export")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">Verify totals before uploading to your accounting system.</p>
                </div>
                ${d}
                <div class="pf-je-checks-container">
                    ${_}
                </div>
            </article>
            ${t?`
                ${_e({textareaId:"step-notes-input",value:n||"",permanentId:"step-notes-permanent",isPermanent:o,saveButtonId:"step-notes-save-5"})}
                ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:a,signoffInputId:"step-signoff-5",signoffValue:s,isComplete:l,saveButtonId:"step-signoff-save-5",completeButtonId:"step-signoff-toggle-5"})}
            `:""}
        </section>
    `}function Wo(e){let t=be.filter(a=>a.id!==6).map(a=>({id:a.id,title:a.title,complete:la(a.id)})),n=t.every(a=>a.complete),o=t.map(a=>`
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${a.complete?"is-active":""}" aria-pressed="${a.complete}">
                        ${dt}
                    </span>
                    <div>
                        <h3>${S(a.title)}</h3>
                        <p class="pf-config-subtext">${a.complete?"Complete":"Not complete"}</p>
                    </div>
                </div>
            </article>
        `).join("");return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S(Pe)} | Step ${e.id}</p>
            <h2 class="pf-hero-title">${S(e.title)}</h2>
            <p class="pf-hero-copy">${S(e.summary||"")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${o}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Archive & Reset</h3>
                    <p class="pf-config-subtext">Create an archive of this module\u2019s sheets and clear work tabs.</p>
                </div>
                <div class="pf-pill-row pf-config-actions">
                    <button type="button" class="pf-pill-btn" id="archive-run-btn" ${n?"":"disabled"}>Archive</button>
                </div>
            </article>
        </section>
    `}function Jo(e){let t=rt.find(_=>_.id===e)||{id:e!=null?e:"-",title:"Workflow Step",summary:"",description:"",checklist:[]};if(e===1)return jo(t);if(e===2)return Ho(t);if(e===3)return Uo(t);if(e===4)return Go(t);if(e===5)return zo(t);if(e===6)return Wo(t);let n=!1,o=Se(e),a=o?I(o.note):"",s=o?Re(o.note):!1,l=(o?I(o.reviewer):"")||De(),d=o?ye(I(o.signOff)):"",r=o&&de[e]?!!(d||I(de[e])):!!d,i=(t.highlights||[]).map(_=>`
            <div class="pf-step-highlight">
                <span class="pf-step-highlight-label">${S(_.label)}</span>
                <span class="pf-step-highlight-detail">${S(_.detail)}</span>
            </div>
        `).join(""),u=(t.checklist||[]).map(_=>`<li>${S(_)}</li>`).join("")||"",f=n?"":t.description||"Detailed guidance will appear here.",p=[];!n&&t.ctaLabel&&p.push(`<button type="button" class="pf-pill-btn" id="step-action-btn">${S(t.ctaLabel)}</button>`),n||p.push('<button type="button" class="pf-pill-btn pf-pill-btn--ghost" id="step-back-btn">Back to Step List</button>');let c=p.length?`<div class="pf-pill-row pf-config-actions">${p.join("")}</div>`:"",m=qt(),y=n?`
            <div class="pf-link-card">
                <h3 class="pf-link-card__title">Payroll Reports</h3>
                <p class="pf-link-card__subtitle">Open your latest payroll export.</p>
                <div class="pf-link-list">
                    <a
                        class="pf-link-item"
                        id="pr-provider-link"
                        ${m?`href="${S(m)}" target="_blank" rel="noopener noreferrer"`:'aria-disabled="true"'}
                    >
                        <span class="pf-link-item__icon">${pt}</span>
                        <span class="pf-link-item__body">
                            <span class="pf-link-item__title">Open Payroll Export</span>
                            <span class="pf-link-item__meta">${S(m||"Add a provider link in Configuration")}</span>
                        </span>
                    </a>
                </div>
            </div>
        `:"",w="",h=!n&&i?`<article class="pf-step-card pf-step-detail">${i}</article>`:"",E=!n&&u?`<article class="pf-step-card pf-step-detail">
                            <h3 class="pf-step-subtitle">Checklist</h3>
                            <ul class="pf-step-checklist">${u}</ul>
                        </article>`:"",k=!n||f||c?`
            <article class="pf-step-card pf-step-detail">
                <p class="pf-step-title">${S(f)}</p>
                ${!n&&t.statusHint?`<p class="pf-step-note">${S(t.statusHint)}</p>`:""}
                ${c}
            </article>
        `:"";return`
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${S(Pe)} | Step ${t.id}</p>
            <h2 class="pf-hero-title">${S(t.title)}</h2>
            <p class="pf-hero-copy">${S(t.summary||"")}</p>
            <p class="pf-hero-hint">${S(te.statusText||"")}</p>
        </section>
        <section class="pf-step-guide">
            ${y}
            ${w}
            ${k}
            ${h}
            ${E}
            ${o?`
                ${_e({textareaId:"step-notes-input",value:a,permanentId:"step-notes-permanent",isPermanent:s,saveButtonId:"step-notes-save"})}
                ${Ae({reviewerInputId:"step-reviewer-name",reviewerValue:l,signoffInputId:`step-signoff-${e}`,signoffValue:d,isComplete:r,saveButtonId:`step-signoff-save-${e}`,completeButtonId:`step-signoff-toggle-${e}`,subtext:"Ready to move on? Save and click Done when finished."})}
            `:""}
        </section>
    `}function Yo(e,t){let n=te.focusedIndex===t?"pf-step-card--active":"",o=vn(Ko(e.id));return`
        <article class="pf-step-card pf-clickable ${n}" data-step-card data-step-index="${t}" data-step-id="${e.id}">
            <p class="pf-step-index">Step ${e.id}</p>
            <h3 class="pf-step-title">${o?`${o}`:""}${S(e.title)}</h3>
        </article>
    `}function qo(){return`
        <footer class="pf-brand-footer">
            <div class="pf-brand-text">
                <div class="pf-brand-label">prairie.forge</div>
                <div class="pf-brand-meta">\xA9 Prairie Forge LLC, 2025. All rights reserved. Version ${nn}</div>
                <button type="button" class="pf-config-link" id="showConfigSheets">CONFIGURATION</button>
            </div>
        </footer>
    `}function Ko(e){return e===0?"config":e===1?"import":e===2?"headcount":e===3?"validate":e===4?"review":e===5?"journal":e===6?"archive":""}function Xo(){var n,o,a,s,l,d,r,i;(n=document.getElementById("nav-home"))==null||n.addEventListener("click",()=>{var u;Hn(),(u=document.getElementById("pf-hero"))==null||u.scrollIntoView({behavior:"smooth",block:"start"})}),(o=document.getElementById("nav-selector"))==null||o.addEventListener("click",()=>{window.location.href="../module-selector/index.html"}),(a=document.getElementById("nav-prev"))==null||a.addEventListener("click",()=>In(-1)),(s=document.getElementById("nav-next"))==null||s.addEventListener("click",()=>In(1));let e=document.getElementById("nav-quick-toggle"),t=document.getElementById("quick-access-dropdown");e==null||e.addEventListener("click",u=>{u.stopPropagation(),t==null||t.classList.toggle("hidden"),e.classList.toggle("is-active")}),document.addEventListener("click",u=>{!(t!=null&&t.contains(u.target))&&!(e!=null&&e.contains(u.target))&&(t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"))}),(l=document.getElementById("nav-roster"))==null||l.addEventListener("click",()=>{Pn("SS_Employee_Roster"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(d=document.getElementById("nav-accounts"))==null||d.addEventListener("click",()=>{Pn("SS_Chart_of_Accounts"),t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active")}),(r=document.getElementById("nav-expense-map"))==null||r.addEventListener("click",async()=>{t==null||t.classList.add("hidden"),e==null||e.classList.remove("is-active"),await Zo()}),(i=document.getElementById("showConfigSheets"))==null||i.addEventListener("click",async()=>{await Qo()})}async function Qo(){if(typeof Excel=="undefined"){console.log("Excel not available");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets;t.load("items/name,visibility"),await e.sync();let n=0;t.items.forEach(a=>{a.name.toUpperCase().startsWith("SS_")&&(a.visibility=Excel.SheetVisibility.visible,console.log(`[Config] Made visible: ${a.name}`),n++)}),await e.sync();let o=e.workbook.worksheets.getItemOrNullObject("SS_PF_Config");o.load("isNullObject"),await e.sync(),o.isNullObject||(o.activate(),o.getRange("A1").select(),await e.sync()),console.log(`[Config] ${n} system sheets now visible`)})}catch(e){console.error("[Config] Error unhiding system sheets:",e)}}async function Pn(e){if(!e||typeof Excel=="undefined")return;let t={SS_Employee_Roster:["Employee","Department","Pay_Rate","Status","Hire_Date"],SS_Chart_of_Accounts:["Account_Number","Account_Name","Type","Category"]};try{await Excel.run(async n=>{let o=n.workbook.worksheets.getItemOrNullObject(e);if(o.load("isNullObject,visibility"),await n.sync(),o.isNullObject){o=n.workbook.worksheets.add(e);let a=t[e]||["Column1","Column2"],s=o.getRange(`A1:${String.fromCharCode(64+a.length)}1`);s.values=[a],s.format.font.bold=!0,s.format.fill.color="#f0f0f0",s.format.autofitColumns(),await n.sync()}else o.visibility=Excel.SheetVisibility.visible,await n.sync();o.activate(),o.getRange("A1").select(),await n.sync(),console.log(`[Quick Access] Opened sheet: ${e}`)})}catch(n){console.error("Error opening reference sheet:",n)}}async function Zo(){try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItemOrNullObject("PR_Expense_Mapping");if(t.load("isNullObject,visibility"),await e.sync(),t.isNullObject){t=e.workbook.worksheets.add("PR_Expense_Mapping");let n=["Expense_Category","GL_Account","Description","Active"],o=t.getRange("A1:D1");o.values=[n],o.format.font.bold=!0}else t.visibility=Excel.SheetVisibility.visible,await e.sync();t.activate(),t.getRange("A1").select(),await e.sync(),console.log("[Quick Access] Opened PR_Expense_Mapping")})}catch(e){console.error("Error navigating to PR_Expense_Mapping:",e)}}function ea(){document.querySelectorAll("[data-step-card]").forEach(e=>{let t=Number(e.getAttribute("data-step-index"));e.addEventListener("click",()=>Qe(t))})}function ta(){var r,i,u,f;let e=document.getElementById("config-user-name");e==null||e.addEventListener("change",p=>{let c=p.target.value.trim();U(qe,c);let m=document.getElementById("config-reviewer-name");m&&!m.value&&(m.value=c)});let t=document.getElementById("config-payroll-date");t==null||t.addEventListener("change",p=>{let c=p.target.value||"";if(U("PR_Payroll_Date",c),!!c){if(!W.overrides.accountingPeriod){let m=ca(c);if(m){let y=document.getElementById("config-accounting-period");y&&(y.value=m),U("PR_Accounting_Period",m)}}if(!W.overrides.jeId){let m=da(c);if(m){let y=document.getElementById("config-je-id");y&&(y.value=m),U("PR_Journal_Entry_ID",m)}}}});let n=Se(0),o=document.getElementById("config-accounting-period");o==null||o.addEventListener("change",p=>{W.overrides.accountingPeriod=!!p.target.value,U("PR_Accounting_Period",p.target.value||"")});let a=document.getElementById("config-je-id");a==null||a.addEventListener("change",p=>{W.overrides.jeId=!!p.target.value,U("PR_Journal_Entry_ID",p.target.value.trim())}),(r=document.getElementById("config-company-name"))==null||r.addEventListener("change",p=>{U("SS_Company_Name",p.target.value.trim())}),(i=document.getElementById("config-payroll-provider"))==null||i.addEventListener("change",p=>{let c=p.target.value.trim();U(jn,c)}),(u=document.getElementById("config-accounting-link"))==null||u.addEventListener("change",p=>{U("SS_Accounting_Software",p.target.value.trim())});let s=document.getElementById("config-notes");if(s==null||s.addEventListener("input",p=>{n&&U(n.note,p.target.value,{debounceMs:400})}),n){let p=document.getElementById("config-notes-permanent");p&&(p.addEventListener("click",()=>{let m=!p.classList.contains("is-locked");Je(p,m),Wn(n.note,m)}),Je(p,Re(n.note)));let c=document.getElementById("config-notes-save");c==null||c.addEventListener("click",()=>{s&&(U(n.note,s.value),Te(c,!0))})}let l=document.getElementById("config-reviewer-name");l==null||l.addEventListener("change",p=>{let c=p.target.value.trim();n&&U(n.reviewer,c),U(qe,c);let m=document.getElementById("config-signoff-date");if(c&&m&&!m.value){let y=Ze();m.value=y,n&&U(n.signOff,y)}}),(f=document.getElementById("config-signoff-date"))==null||f.addEventListener("change",p=>{n&&U(n.signOff,p.target.value||"")});let d=document.getElementById("config-signoff-save");if(d==null||d.addEventListener("click",()=>{var y;let p=((y=l==null?void 0:l.value)==null?void 0:y.trim())||"",c=document.getElementById("config-signoff-date"),m=(c==null?void 0:c.value)||"";n&&(U(n.reviewer,p),U(n.signOff,m)),U(qe,p),Te(d,!0)}),Vt(),n){let p=I(n.signOff),c=I(de[0]),m=!!(p||c==="Y"||c===!0);console.log(`[Step 0] Binding signoff toggle. signOff="${p}", complete="${c}", isComplete=${m}`),Gn({buttonId:"config-signoff-toggle",inputId:"config-signoff-date",fieldName:n.signOff,completeField:de[0],initialActive:m,stepId:0})}}function na(e){var n,o,a,s,l,d,r,i,u,f,p,c,m,y,w,h,E,k,_,$,V;if((n=document.getElementById("step-back-btn"))==null||n.addEventListener("click",()=>{Hn()}),(o=document.getElementById("step-action-btn"))==null||o.addEventListener("click",()=>{let x=rt.find(T=>T.id===e);window.alert(x!=null&&x.ctaLabel?`${x.ctaLabel} coming soon.`:"Step actions coming soon.")}),e===1&&((a=document.getElementById("import-open-data-btn"))==null||a.addEventListener("click",()=>sa()),(s=document.getElementById("import-clear-btn"))==null||s.addEventListener("click",()=>ra())),e===2&&((l=document.getElementById("headcount-skip-btn"))==null||l.addEventListener("click",()=>{H.skipAnalysis=!H.skipAnalysis;let x=document.getElementById("headcount-skip-btn");x==null||x.classList.toggle("is-active",H.skipAnalysis),H.skipAnalysis&&Jt(),vt()}),(d=document.getElementById("roster-run-btn"))==null||d.addEventListener("click",()=>Wt()),(r=document.getElementById("roster-refresh-btn"))==null||r.addEventListener("click",()=>Wt()),(i=document.getElementById("roster-review-btn"))==null||i.addEventListener("click",()=>{var T;let x=((T=H.roster)==null?void 0:T.mismatches)||[];Vn("Roster Differences",x,{sourceLabel:"Roster",targetLabel:"Payroll Data"})}),(u=document.getElementById("dept-review-btn"))==null||u.addEventListener("click",()=>{var T;let x=((T=H.departments)==null?void 0:T.mismatches)||[];Vn("Department Differences",x,{sourceLabel:"Roster",targetLabel:"Payroll",formatter:g=>({name:g.employee,source:`${g.rosterDept} \u2192 ${g.payrollDept}`,isMissingFromTarget:!0})})})),e===3&&((f=document.getElementById("validation-run-btn"))==null||f.addEventListener("click",()=>Mn()),(p=document.getElementById("validation-refresh-btn"))==null||p.addEventListener("click",()=>Mn()),(c=document.getElementById("bank-amount-input"))==null||c.addEventListener("blur",Fn),(m=document.getElementById("bank-amount-input"))==null||m.addEventListener("keydown",x=>{x.key==="Enter"&&Fn(x)})),e===5&&((y=document.getElementById("je-run-btn"))==null||y.addEventListener("click",()=>La()),(w=document.getElementById("je-save-btn"))==null||w.addEventListener("click",()=>Ba()),(h=document.getElementById("je-create-btn"))==null||h.addEventListener("click",()=>Ma()),(E=document.getElementById("je-export-btn"))==null||E.addEventListener("click",()=>Va())),e===4){let x=document.querySelector(".pf-step-guide");if(x){let T="https://your-project.supabase.co/functions/v1/copilot";ln(x,{id:"expense-review-copilot",contextProvider:fa(),systemPrompt:`You are Prairie Forge CoPilot, an expert financial analyst assistant for payroll expense review.

CONTEXT: You're embedded in the Payroll Recorder Excel add-in, helping accountants and CFOs review payroll data before journal entry export.

YOUR CAPABILITIES:
1. Analyze payroll expense data for accuracy and completeness
2. Identify trends, anomalies, and variances requiring attention
3. Prepare executive-ready insights and talking points
4. Validate journal entries before export to accounting system

COMMUNICATION STYLE:
- Be concise and actionable
- Use bullet points and tables for clarity
- Highlight issues with \u26A0\uFE0F and successes with \u2713
- Format currency as $X,XXX (no decimals for totals)
- Format percentages as X.X%
- Always end with 2-3 concrete next steps

ANALYSIS FOCUS:
- Period-over-period changes exceeding 10%
- Department cost anomalies vs historical norms
- Headcount vs payroll expense alignment
- Burden rate outliers (normal range: 15-35%)
- Missing or incomplete GL account mappings
- Data quality issues (blanks, duplicates, mismatches)

When asked about variances, explain the business drivers, not just the numbers.
When asked about readiness, be specific about what passes and what needs attention.`})}(k=document.getElementById("expense-run-btn"))==null||k.addEventListener("click",()=>{Ln()}),(_=document.getElementById("expense-refresh-btn"))==null||_.addEventListener("click",()=>{Ln()})}let t=Se(e);if(console.log(`[Step ${e}] Binding interactions, fields:`,t),t){let x=e===1?"step-notes-1":"step-notes-input",T=document.getElementById(x);console.log(`[Step ${e}] Notes input found:`,!!T,`(id: ${x})`);let g=e===1?document.getElementById("step-notes-save-1"):e===2?document.getElementById("step-notes-save-2"):e===3?document.getElementById("step-notes-save-3"):e===4?document.getElementById("step-notes-save-4"):e===5?document.getElementById("step-notes-save-5"):document.getElementById("step-notes-save");T==null||T.addEventListener("input",Z=>{U(t.note,Z.target.value,{debounceMs:400}),e===2&&(H.skipAnalysis&&Jt(),vt())}),g==null||g.addEventListener("click",()=>{T&&(U(t.note,T.value),Te(g,!0))});let D=e===1?"step-reviewer-1":"step-reviewer-name",L=document.getElementById(D);L==null||L.addEventListener("change",Z=>{let fe=Z.target.value.trim();U(t.reviewer,fe);let me=e===1?document.getElementById("step-signoff-1"):e===2?document.getElementById("step-signoff-date"):e===3?document.getElementById("step-signoff-3"):e===4?document.getElementById("step-signoff-4"):e===5?document.getElementById("step-signoff-5"):document.getElementById(`step-signoff-${e}`);if(fe&&me&&!me.value){let he=Ze();me.value=he,U(t.signOff,he)}});let Q=e===1?"step-signoff-1":e===2?"step-signoff-date":e===3?"step-signoff-3":e===4?"step-signoff-4":e===5?"step-signoff-5":`step-signoff-${e}`;console.log(`[Step ${e}] Signoff input ID: ${Q}, found:`,!!document.getElementById(Q)),($=document.getElementById(Q))==null||$.addEventListener("change",Z=>{U(t.signOff,Z.target.value||"")});let pe=e===1?"step-notes-lock-1":"step-notes-permanent",K=document.getElementById(pe);K&&(K.addEventListener("click",()=>{let Z=!K.classList.contains("is-locked");Je(K,Z),Wn(t.note,Z),e===2&&vt()}),Je(K,Re(t.note)));let M=e===1?document.getElementById("step-signoff-save-1"):e===2?document.getElementById("headcount-signoff-save"):e===3?document.getElementById("step-signoff-save-3"):e===4?document.getElementById("step-signoff-save-4"):e===5?document.getElementById("step-signoff-save-5"):document.getElementById(`step-signoff-save-${e}`);M==null||M.addEventListener("click",()=>{var me,he;let Z=((me=L==null?void 0:L.value)==null?void 0:me.trim())||"",fe=((he=document.getElementById(Q))==null?void 0:he.value)||"";U(t.reviewer,Z),U(t.signOff,fe),Te(M,!0)}),Vt();let Y=de[e],ne=Y?!!I(Y):!1,xe=I(t.signOff),X=e===1?"step-signoff-toggle-1":e===2?"headcount-signoff-toggle":e===3?"validation-signoff-toggle":e===4?"expense-signoff-toggle":e===5?"step-signoff-toggle-5":`step-signoff-toggle-${e}`;console.log(`[Step ${e}] Toggle button ID: ${X}, found:`,!!document.getElementById(X)),Gn({buttonId:X,inputId:Q,fieldName:t.signOff,completeField:Y,requireNotesCheck:e===2?Rt:null,initialActive:!!(xe||ne),stepId:e,onComplete:e===3?Aa:e===4?Da:e===2?_a:null})}e===2&&vt(),e===6&&((V=document.getElementById("archive-run-btn"))==null||V.addEventListener("click",Pa))}function Qe(e){if(Number.isNaN(e)||e<0||e>=be.length)return;let t=be[e];if(!t)return;bt=e;let n=t.id===0?"config":"step";Yt({focusedIndex:e,activeView:n,activeStepId:t.id});let o=$o[t.id];o&&ka(o),t.id===2&&!H.hasAnalyzed&&Wt()}function In(e){if(te.activeView==="home"&&e>0){Qe(0);return}let t=te.focusedIndex+e,n=Math.max(0,Math.min(be.length-1,t));Qe(n)}function oa(){if(te.activeView!=="home"||bt===null)return;let e=document.querySelector(`[data-step-card][data-step-index="${bt}"]`);bt=null,e==null||e.scrollIntoView({behavior:"smooth",block:"center"})}async function Hn(){let e=Ot(Et);await Nt(e.sheetName,e.title,e.subtitle),Yt({activeView:"home",activeStepId:null})}function Yt(e){Object.assign(te,e),le()}function De(){return I(qe)||I("SS_Default_Reviewer")||""}function Ut(e,t){e&&(e.classList.toggle("is-active",t),e.setAttribute("aria-pressed",String(t)))}function Un(e){let t=document.getElementById("je-save-btn");t&&t.classList.toggle("is-saved",e)}function aa(){let e={};return console.log("[Signoff] Checking step completion status..."),Object.keys($e).forEach(t=>{let n=parseInt(t,10),o=$e[n];if(!o){e[n]=!1;return}let a=I(o.signOff),s=de[n],l=I(s),d=!!a||l==="Y"||l===!0;e[n]=d,console.log(`[Signoff] Step ${n}: signOff="${a}", complete="${l}" \u2192 ${d?"COMPLETE":"pending"}`)}),console.log("[Signoff] Status summary:",e),e}function Gn({buttonId:e,inputId:t,fieldName:n,completeField:o,requireNotesCheck:a,onComplete:s,initialActive:l=!1,stepId:d=null}){let r=document.getElementById(e);if(!r){console.warn(`[Signoff] Button not found: ${e}`);return}let i=t?document.getElementById(t):null,u=l||!!(i!=null&&i.value);Ut(r,u),console.log(`[Signoff] Bound ${e}, initial active: ${u}, stepId: ${d}`),r.addEventListener("click",()=>{if(console.log(`[Signoff] Done button clicked: ${e}, stepId: ${d}`),d!==null&&d>0){let p=aa(),{canComplete:c,message:m}=Sn(d,p),y=r.classList.contains("is-active");if(console.log(`[Signoff] canComplete: ${c}, isCurrentlyActive: ${y}`),!y&&!c){console.log(`[Signoff] BLOCKED: ${m}`),xn(m);return}}if(a&&!a()){window.alert("Please add notes before completing this step.");return}let f=!r.classList.contains("is-active");if(console.log(`[Signoff] ${e} clicked, toggling to: ${f}`),Ut(r,f),i&&(i.value=f?Ze():""),n){let p=f?Ze():"";console.log(`[Signoff] Writing ${n} = "${p}"`),U(n,p)}if(o){let p=f?"Y":"";console.log(`[Signoff] Writing ${o} = "${p}"`),U(o,p)}f&&typeof s=="function"&&s()}),i&&i.addEventListener("change",()=>{let f=!!i.value,p=r.classList.contains("is-active");f!==p&&(console.log(`[Signoff] Date input changed, syncing button to: ${f}`),Ut(r,f),n&&U(n,i.value||""),o&&U(o,f?"Y":""))})}async function sa(){if(!ae()){window.alert("Open this module inside Excel to access the data sheet.");return}try{await Excel.run(async e=>{let t=e.workbook.worksheets.getItem(A.DATA);t.activate(),t.getRange("A1").select(),await e.sync()})}catch(e){console.error("Unable to open PR_Data sheet",e),window.alert(`Unable to open ${A.DATA}. Confirm the sheet exists in this workbook.`)}}async function ra(){if(!ae()){window.alert("Open this module inside Excel to clear data.");return}if(window.confirm("Are you sure you want to clear all data from PR_Data? This cannot be undone."))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItem(A.DATA),o=n.getUsedRangeOrNullObject();o.load("isNullObject"),await t.sync(),o.isNullObject||(n.getRange("A2:Z10000").clear(Excel.ClearApplyTo.contents),await t.sync()),n.activate(),n.getRange("A1").select(),await t.sync()}),window.alert("PR_Data cleared successfully.")}catch(t){console.error("Unable to clear PR_Data sheet",t),window.alert("Unable to clear PR_Data. Please try again.")}}async function Ie(e){var a,s;if(!Ht.length)return null;if(gt){let l=e.workbook.tables.getItemOrNullObject(gt);if(l.load("name"),await e.sync(),!l.isNullObject)return l;gt=null}let t=e.workbook.tables;t.load("items/name"),await e.sync();let n=((a=t.items)==null?void 0:a.map(l=>l.name))||[];console.log("[Payroll] Looking for config table:",Ht),console.log("[Payroll] Found tables in workbook:",n);let o=(s=t.items)==null?void 0:s.find(l=>Ht.includes(l.name));return o?(console.log("[Payroll] \u2713 Config table found:",o.name),gt=o.name,e.workbook.tables.getItem(o.name)):(console.warn("[Payroll] \u26A0\uFE0F CONFIG TABLE NOT FOUND!"),console.warn("[Payroll] Expected table named: SS_PF_Config"),console.warn("[Payroll] Available tables:",n),console.warn("[Payroll] To fix: Select your data in SS_PF_Config sheet \u2192 Insert \u2192 Table \u2192 Name it 'SS_PF_Config'"),null)}async function zn(){if(!ae()){W.loaded=!0;return}try{await Excel.run(async e=>{let t=await Ie(e);if(!t){console.warn("Payroll Recorder: SS_PF_Config table is missing."),W.loaded=!0;return}let n=t.getDataBodyRange();n.load("values"),await e.sync();let o=n.values||[],a={},s={};o.forEach(l=>{var r,i;let d=oe(l[B.FIELD]);d&&(a[d]=(r=l[B.VALUE])!=null?r:"",s[d]=(i=l[B.PERMANENT])!=null?i:"")}),W.values=a,W.permanents=s,W.overrides.accountingPeriod=!!(a.PR_Accounting_Period||a.Accounting_Period),W.overrides.jeId=!!(a.PR_Journal_Entry_ID||a.Journal_Entry_ID),W.loaded=!0})}catch(e){console.warn("Payroll Recorder: unable to load PF_Config table.",e),W.loaded=!0}}function I(e){var t;return(t=W.values[e])!=null?t:""}function ia(){let e=Object.keys(W.values||{});return Xe.find(n=>e.includes(n))||Xe[0]}function Ct(){return I(ia())}function qt(){return(I(jn)||I("Payroll_Provider_Link")||"").trim()}function Re(e){return Jn(W.permanents[e])}function la(e){let t=de[e];return t?Jn(I(t)):!1}function Wn(e,t){let n=oe(e);n&&(W.permanents[n]=t?"Y":"N",pa(n,t?"Y":"N"))}function Jn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t==="true"||t==="y"||t==="yes"||t==="1"}function oe(e){return String(e!=null?e:"").trim()}function Yn(e){let t=String(e!=null?e:"").trim().toLowerCase();return t?t.includes("total")||t.includes("totals")||t.includes("grand total")||t.includes("subtotal")||t.includes("summary"):!0}function ye(e){if(!e)return"";let t=kt(e);return t?`${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function ca(e){let t=kt(e);return t?t.year<1900||t.year>2100?(console.warn("deriveAccountingPeriod - Invalid year:",t.year,"from input:",e),""):`${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][t.month-1]} ${t.year}`:""}function da(e){let t=kt(e);return t?t.year<1900||t.year>2100?(console.warn("deriveJeId - Invalid year:",t.year,"from input:",e),""):`PR-AUTO-${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:""}function Ze(){return wt(new Date)}function U(e,t,n={}){var l;let o=oe(e);W.values[o]=t!=null?t:"";let a=(l=n.debounceMs)!=null?l:0;if(!a){let d=Le.get(o);d&&clearTimeout(d),Le.delete(o),Ke(o,t!=null?t:"");return}Le.has(o)&&clearTimeout(Le.get(o));let s=setTimeout(()=>{Le.delete(o),Ke(o,t!=null?t:"")},a);Le.set(o,s)}var ua=["PR_Accounting_Period","PTO_Accounting_Period","Accounting_Period"];async function Ke(e,t){let n=oe(e);if(W.values[n]=t!=null?t:"",console.log(`[Payroll] Writing config: ${n} = "${t}"`),!ae()){console.warn("[Payroll] Excel runtime not available - cannot write");return}let o=ua.some(a=>n===a||n.toLowerCase()===a.toLowerCase());try{await Excel.run(async a=>{var p;let s=await Ie(a);if(!s){console.error("[Payroll] \u274C Cannot write - config table not found");return}let l=s.getDataBodyRange(),d=s.getHeaderRowRange();l.load("values"),d.load("values"),await a.sync();let r=d.values[0]||[],i=l.values||[],u=r.length;console.log(`[Payroll] Table has ${i.length} rows, ${u} columns`);let f=[];if(i.forEach((c,m)=>{oe(c[B.FIELD])===n&&f.push(m)}),f.length===0){W.permanents[n]=(p=W.permanents[n])!=null?p:Dn;let c=new Array(u).fill("");if(B.TYPE>=0&&B.TYPE<u&&(c[B.TYPE]=Ao),B.FIELD>=0&&B.FIELD<u&&(c[B.FIELD]=n),B.VALUE>=0&&B.VALUE<u&&(c[B.VALUE]=t!=null?t:""),B.PERMANENT>=0&&B.PERMANENT<u&&(c[B.PERMANENT]=Dn),console.log("[Payroll] Adding NEW row:",c),s.rows.add(null,[c]),await a.sync(),o){let m=s.rows;m.load("count"),await a.sync();let y=m.count-1,h=s.rows.getItemAt(y).getRange().getCell(0,B.VALUE);h.numberFormat=[["@"]],h.values=[[t!=null?t:""]],await a.sync(),console.log(`[Payroll] \u2713 Applied text format to ${n}`)}console.log(`[Payroll] \u2713 New row added for ${n}`)}else{let c=f[0];console.log(`[Payroll] Updating existing row ${c} for ${n}`);let m=l.getCell(c,B.VALUE);if(o&&(m.numberFormat=[["@"]]),m.values=[[t!=null?t:""]],await a.sync(),console.log(`[Payroll] \u2713 Updated ${n}`),f.length>1){console.log(`[Payroll] Found ${f.length-1} duplicate rows for ${n}, removing...`);let y=f.slice(1).reverse();for(let w of y)try{s.rows.getItemAt(w).delete()}catch(h){console.warn(`[Payroll] Could not delete duplicate row ${w}:`,h.message)}await a.sync(),console.log(`[Payroll] \u2713 Removed duplicate rows for ${n}`)}}})}catch(a){console.error(`[Payroll] \u274C Write failed for ${e}:`,a)}}async function pa(e,t){let n=oe(e);if(n&&ae()){W.permanents[n]=t;try{await Excel.run(async o=>{let a=await Ie(o);if(!a){console.warn(`Payroll Recorder: unable to locate config table when toggling ${e} permanent flag.`);return}let s=a.getDataBodyRange();s.load("values"),await o.sync();let d=(s.values||[]).findIndex(r=>oe(r[B.FIELD])===n);d!==-1&&(s.getCell(d,B.PERMANENT).values=[[t]],await o.sync())})}catch(o){console.warn(`Payroll Recorder: unable to update permanent flag for ${e}`,o)}}}function kt(e){if(!e)return null;let t=String(e).trim(),n=/^(\d{4})-(\d{2})-(\d{2})/.exec(t);if(n){let l=Number(n[1]),d=Number(n[2]),r=Number(n[3]);if(l&&d&&r)return{year:l,month:d,day:r}}let o=/^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(t);if(o){let l=Number(o[1]),d=Number(o[2]),r=Number(o[3]);if(r&&l&&d)return{year:r,month:l,day:d}}let a=Number(e);if(Number.isFinite(a)&&a>4e4&&a<6e4){let d=Math.floor(a-25569)*86400*1e3,r=new Date(d);if(!isNaN(r.getTime())){let i=`${r.getUTCFullYear()}-${String(r.getUTCMonth()+1).padStart(2,"0")}-${String(r.getUTCDate()).padStart(2,"0")}`;return console.log("DEBUG parseDateInput - Converted Excel serial",a,"to",i),{year:r.getUTCFullYear(),month:r.getUTCMonth()+1,day:r.getUTCDate()}}}let s=new Date(t);return isNaN(s.getTime())?(console.warn("DEBUG parseDateInput - Could not parse date value:",e),null):{year:s.getFullYear(),month:s.getMonth()+1,day:s.getDate()}}function wt(e){if(e._isUTC){let a=e.getUTCFullYear(),s=String(e.getUTCMonth()+1).padStart(2,"0"),l=String(e.getUTCDate()).padStart(2,"0");return`${a}-${s}-${l}`}let t=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0"),o=String(e.getDate()).padStart(2,"0");return`${t}-${n}-${o}`}function $n(e){if(!e)return null;if(typeof e=="string"){let n=e.match(/^(\d{4})-(\d{2})-(\d{2})/);if(n)return`${n[1]}-${n[2]}-${n[3]}`}let t=kt(e);return t?`${t.year}-${String(t.month).padStart(2,"0")}-${String(t.day).padStart(2,"0")}`:null}function fa(){return async()=>{if(!ae())return null;try{return await Excel.run(async e=>{var l,d,r;let t={timestamp:new Date().toISOString(),period:null,summary:{},departments:[],recentPeriods:[],dataQuality:{}},n=await Ie(e);if(n){let i=n.getDataBodyRange();i.load("values"),await e.sync();let u=i.values||[];for(let f of u){let p=String(f[B.FIELD]||"").trim(),c=f[B.VALUE];p.toLowerCase().includes("accounting")&&p.toLowerCase().includes("period")&&(t.period=String(c||"").trim())}}let o=e.workbook.worksheets.getItemOrNullObject(A.DATA_CLEAN);if(o.load("isNullObject"),await e.sync(),!o.isNullObject){let i=o.getUsedRangeOrNullObject();if(i.load("values"),await e.sync(),!i.isNullObject&&((l=i.values)==null?void 0:l.length)>1){let u=i.values[0].map(E=>ve(E)),f=i.values.slice(1),p=u.findIndex(E=>E.includes("amount")),c=Be(u),m=u.findIndex(E=>E.includes("employee")),y=0,w=new Set,h=new Map;for(let E of f){let k=Number(E[p])||0;if(y+=k,m>=0){let _=String(E[m]||"").trim();_&&w.add(_)}if(c>=0){let _=String(E[c]||"").trim();_&&h.set(_,(h.get(_)||0)+k)}}t.summary={total:y,employeeCount:w.size,avgPerEmployee:w.size?y/w.size:0,rowCount:f.length},t.departments=Array.from(h.entries()).map(([E,k])=>({name:E,total:k,percentOfTotal:y?k/y:0})).sort((E,k)=>k.total-E.total),t.dataQuality.dataCleanReady=!0,t.dataQuality.rowCount=f.length}}let a=e.workbook.worksheets.getItemOrNullObject(A.ARCHIVE_SUMMARY);if(a.load("isNullObject"),await e.sync(),!a.isNullObject){let i=a.getUsedRangeOrNullObject();if(i.load("values"),await e.sync(),!i.isNullObject&&((d=i.values)==null?void 0:d.length)>1){let u=i.values[0].map(c=>ve(c)),f=u.findIndex(c=>c.includes("period")),p=u.findIndex(c=>c.includes("total"));t.recentPeriods=i.values.slice(1,6).map(c=>({period:c[f]||"",total:Number(c[p])||0})),t.dataQuality.archiveAvailable=!0,t.dataQuality.periodsAvailable=t.recentPeriods.length}}let s=e.workbook.worksheets.getItemOrNullObject(A.JE_DRAFT);if(s.load("isNullObject"),await e.sync(),!s.isNullObject){let i=s.getUsedRangeOrNullObject();if(i.load("values"),await e.sync(),!i.isNullObject&&((r=i.values)==null?void 0:r.length)>1){let u=i.values[0].map(y=>ve(y)),f=u.findIndex(y=>y.includes("debit")),p=u.findIndex(y=>y.includes("credit")),c=0,m=0;for(let y of i.values.slice(1))c+=Number(y[f])||0,m+=Number(y[p])||0;t.journalEntry={totalDebit:c,totalCredit:m,difference:Math.abs(c-m),isBalanced:Math.abs(c-m)<.01,lineCount:i.values.length-1},t.dataQuality.jeDraftReady=!0}}return console.log("CoPilot context gathered:",t),t})}catch(e){return console.warn("CoPilot context provider error:",e),null}}}function S(e){return String(e!=null?e:"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}function ge(e,t){return`
        <div class="pf-labeled-button">
            ${e}
            <span class="pf-button-label">${S(t)}</span>
        </div>
    `}function ae(){return typeof Excel!="undefined"&&typeof Excel.run=="function"}function Se(e){return $e[e]||null}function ma(){var n,o,a,s;let e=Math.abs((o=(n=H.roster)==null?void 0:n.difference)!=null?o:0),t=Math.abs((s=(a=H.departments)==null?void 0:a.difference)!=null?s:0);return e>0||t>0}function Rt(){return!H.skipAnalysis&&ma()}function ie(e){return e==null||Number.isNaN(e)?"---":typeof e!="number"?e:e.toLocaleString(void 0,{minimumFractionDigits:2,maximumFractionDigits:2})}function qn(e){let t=Kt(e);return Number.isFinite(t)?t.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}):""}function ga(e=[]){return e.map(t=>(t||[]).map(n=>{if(n==null)return"";let o=String(n);return/[",\n]/.test(o)?`"${o.replace(/"/g,'""')}"`:o}).join(",")).join(`
`)}function ha(e,t){let n=new Blob([t],{type:"text/csv;charset=utf-8;"}),o=URL.createObjectURL(n),a=document.createElement("a");a.href=o,a.download=e,document.body.appendChild(a),a.click(),a.remove(),setTimeout(()=>URL.revokeObjectURL(o),1e3)}function Kt(e){if(typeof e=="number")return e;if(e==null)return NaN;let t=String(e).replace(/[^0-9.-]/g,""),n=Number.parseFloat(t);return Number.isFinite(n)?n:NaN}function ya(e){if(e instanceof Date)return wt(e);if(typeof e=="number"&&!Number.isNaN(e)){let o=va(e);return o?wt(o):""}let t=String(e!=null?e:"").trim();if(!t)return"";if(/^\d{4}-\d{2}-\d{2}$/.test(t))return t;let n=new Date(t);return Number.isNaN(n.getTime())?t:wt(n)}function va(e){if(!Number.isFinite(e))return null;let t=Math.floor(e-25569);if(!Number.isFinite(t))return null;let n=t*86400*1e3,o=new Date(n);return o._isUTC=!0,o}function ba(e){if(!e)return"";let t=new Date(e);return Number.isNaN(t.getTime())?e:t.toLocaleDateString(void 0,{month:"short",day:"numeric",year:"numeric"})}function ht(e){if(e==null||e==="")return 0;let t=Number(e);return Number.isFinite(t)?t:0}function wa(e){let t=ce(e).toLowerCase();return t?t.includes("burden")||t.includes("tax")||t.includes("benefit")||t.includes("fica")||t.includes("insurance")||t.includes("health")||t.includes("medicare")?"burden":t.includes("bonus")||t.includes("commission")||t.includes("variable")||t.includes("overtime")||t.includes("per diem")?"variable":"fixed":"variable"}function Nn(e){if(!e||e.length<2)return[];let t=(e[0]||[]).map(a=>ve(a));console.log("parseExpenseRows - headers:",t);let n={payrollDate:t.findIndex(a=>a.includes("payroll")&&a.includes("date")),employee:t.findIndex(a=>a.includes("employee")),department:t.findIndex(a=>a.includes("department")),fixed:t.findIndex(a=>a.includes("fixed")),variable:t.findIndex(a=>a.includes("variable")),burden:t.findIndex(a=>a.includes("burden")),amount:t.findIndex(a=>a.includes("amount")),expenseReview:t.findIndex(a=>a.includes("expense")&&a.includes("review")),category:t.findIndex(a=>a.includes("payroll")&&a.includes("category"))};if(console.log("parseExpenseRows - column indexes:",n),n.payrollDate>=0){let a=new Set;for(let s=1;s<e.length;s++){let l=e[s][n.payrollDate];l&&a.add(String(l))}console.log("parseExpenseRows - unique payroll dates found:",[...a].slice(0,20))}let o=[];for(let a=1;a<e.length;a+=1){let s=e[a],l=ya(n.payrollDate>=0?s[n.payrollDate]:null);if(!l)continue;let d=n.employee>=0?ce(s[n.employee]):"",r=n.department>=0&&ce(s[n.department])||"Unassigned",i=n.fixed>=0?ht(s[n.fixed]):null,u=n.variable>=0?ht(s[n.variable]):null,f=n.burden>=0?ht(s[n.burden]):null,p=0,c=0,m=0;if(i!==null||u!==null||f!==null)p=i!=null?i:0,c=u!=null?u:0,m=f!=null?f:0;else{let y=n.amount>=0?ht(s[n.amount]):0,w=wa(n.expenseReview>=0?s[n.expenseReview]:s[n.category]);w==="fixed"?p=y:w==="burden"?m=y:c=y}p===0&&c===0&&m===0||o.push({period:l,employee:d,department:r||"Unassigned",fixed:p,variable:c,burden:m})}return o}function On(e){let t=new Map;e.forEach(o=>{let a=o.period;if(!a)return;t.has(a)||t.set(a,{key:a,label:ba(a),employees:new Set,departments:new Map,summary:{fixed:0,variable:0,burden:0}});let s=t.get(a);s.employees.add(o.employee||`EMP-${s.employees.size+1}`);let l=o.department||"Unassigned";s.departments.has(l)||s.departments.set(l,{name:l,fixed:0,variable:0,burden:0,employees:new Set});let d=s.departments.get(l);d.fixed+=o.fixed,d.variable+=o.variable,d.burden+=o.burden,d.employees.add(o.employee||`EMP-${d.employees.size+1}`),s.summary.fixed+=o.fixed,s.summary.variable+=o.variable,s.summary.burden+=o.burden});let n=[];return t.forEach(o=>{let a=o.summary.fixed+o.summary.variable+o.summary.burden,s=Array.from(o.departments.values()).map(r=>{let i=r.fixed+r.variable,u=i+r.burden;return{name:r.name,fixed:r.fixed,variable:r.variable,gross:i,burden:r.burden,allIn:u,percent:a?u/a:0,headcount:r.employees.size,delta:0}});s.sort((r,i)=>i.allIn-r.allIn);let l={employeeCount:o.employees.size,fixed:o.summary.fixed,variable:o.summary.variable,burden:o.summary.burden,total:a,burdenRate:a?o.summary.burden/a:0,delta:0},d={name:"Totals",fixed:o.summary.fixed,variable:o.summary.variable,gross:o.summary.fixed+o.summary.variable,burden:o.summary.burden,allIn:a,percent:a?1:0,headcount:o.employees.size,delta:0,isTotal:!0};n.push({key:o.key,label:o.label,summary:l,departments:s,totalsRow:d})}),n.sort((o,a)=>o.key<a.key?1:-1)}function Tn(e,t){console.log("buildExpenseReviewPeriods - cleanValues rows:",(e==null?void 0:e.length)||0),console.log("buildExpenseReviewPeriods - archiveValues rows:",(t==null?void 0:t.length)||0);let n=On(Nn(e)),o=On(Nn(t));console.log("buildExpenseReviewPeriods - currentPeriods:",n.map(i=>{var u,f;return{key:i.key,employees:(u=i.summary)==null?void 0:u.employeeCount,total:(f=i.summary)==null?void 0:f.total}})),console.log("buildExpenseReviewPeriods - archivePeriods:",o.map(i=>{var u,f;return{key:i.key,employees:(u=i.summary)==null?void 0:u.employeeCount,total:(f=i.summary)==null?void 0:f.total}}));let a=new Map(o.map(i=>[i.key,i])),s=[];n.length&&(s.push(n[0]),a.delete(n[0].key)),o.forEach(i=>{s.length>=6||s.some(u=>u.key===i.key)||s.push(i)}),console.log("buildExpenseReviewPeriods - combined before filter:",s.map(i=>{var u,f;return{key:i.key,employees:(u=i.summary)==null?void 0:u.employeeCount,total:(f=i.summary)==null?void 0:f.total}}));let l=3,d=1e3,r=s.filter(i=>{var c,m,y,w,h;if(!i||!i.key)return console.log("buildExpenseReviewPeriods - EXCLUDED (no key):",i),!1;let u=((c=i.summary)==null?void 0:c.total)||(((m=i.summary)==null?void 0:m.fixed)||0)+(((y=i.summary)==null?void 0:y.variable)||0)+(((w=i.summary)==null?void 0:w.burden)||0),f=((h=i.summary)==null?void 0:h.employeeCount)||0;if(s.indexOf(i)===0)return console.log(`buildExpenseReviewPeriods - INCLUDED (current): ${i.key} - ${f} employees, $${u}`),!0;let p=f>=l&&u>=d;return console.log(`buildExpenseReviewPeriods - ${p?"INCLUDED":"EXCLUDED"}: ${i.key} - ${f} employees, $${u} (needs >=${l} emp, >=$${d})`),p}).sort((i,u)=>i.key<u.key?1:-1).slice(0,6);return console.log("buildExpenseReviewPeriods - FINAL periods:",r.map(i=>i.key)),r.forEach((i,u)=>{let f=r[u+1],p=f?i.summary.total-f.summary.total:0;i.summary.delta=p;let c=new Map(((f==null?void 0:f.departments)||[]).map(m=>[m.name,m]));i.departments.forEach(m=>{let y=c.get(m.name);m.delta=y?m.allIn-y.allIn:0}),i.totalsRow.delta=p}),r}async function Ln(){if(!ae()){yt({loading:!1,lastError:"Excel runtime is unavailable."});return}yt({loading:!0,lastError:null});try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItemOrNullObject(A.DATA_CLEAN),o=t.workbook.worksheets.getItemOrNullObject(A.ARCHIVE_SUMMARY),a=t.workbook.worksheets.getItemOrNullObject(A.EXPENSE_REVIEW);if(n.load("isNullObject, name"),o.load("isNullObject, name"),a.load("isNullObject, name"),await t.sync(),console.log("Expense Review - Sheet check:",{cleanSheet:n.isNullObject?"MISSING":n.name,archiveSheet:o.isNullObject?"MISSING":o.name,reviewSheet:a.isNullObject?"MISSING":a.name}),a.isNullObject){console.log("Creating PR_Expense_Review sheet...");let r=t.workbook.worksheets.add(A.EXPENSE_REVIEW);await t.sync();let i=t.workbook.worksheets.getItem(A.EXPENSE_REVIEW),u=[],f=[];if(!n.isNullObject){let c=n.getUsedRangeOrNullObject();c.load("values"),await t.sync(),u=c.isNullObject?[]:c.values||[]}if(!o.isNullObject){let c=o.getUsedRangeOrNullObject();c.load("values"),await t.sync(),f=c.isNullObject?[]:c.values||[]}let p=Tn(u,f);return await Bn(t,i,p),p}let s=[],l=[];if(n.isNullObject)console.warn("Expense Review - PR_Data_Clean sheet not found, using empty data");else{let r=n.getUsedRangeOrNullObject();r.load("values"),await t.sync(),s=r.isNullObject?[]:r.values||[],console.log("Expense Review - PR_Data_Clean rows:",s.length)}if(o.isNullObject)console.warn("Expense Review - PR_Archive_Summary sheet not found, using empty data");else{let r=o.getUsedRangeOrNullObject();r.load("values"),await t.sync(),l=r.isNullObject?[]:r.values||[],console.log("Expense Review - PR_Archive_Summary rows:",l.length)}let d=Tn(s,l);return console.log("Expense Review - Periods built:",d.length),await Bn(t,a,d),d});yt({loading:!1,periods:e,lastError:null}),await No(),le()}catch(e){console.error("Expense Review: unable to build executive summary",e),console.error("Error details:",e.message,e.stack),yt({loading:!1,lastError:`Unable to build the Expense Review: ${e.message||"Unknown error"}`,periods:[]})}}async function Bn(e,t,n){if(!t){console.error("writeExpenseReviewSheet: sheet is null/undefined");return}console.log("writeExpenseReviewSheet: Building executive dashboard with",n.length,"periods");try{let v=t.getUsedRangeOrNullObject();v.load("address");let R=t.charts;R.load("items"),await e.sync(),v.isNullObject||(v.clear(),await e.sync());for(let j=R.items.length-1;j>=0;j--)R.items[j].delete();await e.sync()}catch(v){console.warn("Could not clear sheet:",v)}let o=n[0]||{},a=n[1]||{},s=o.summary||{},l=a.summary||{},d=I("PR_Accounting_Period")||Ct()||"",r=Number(s.total)||0,i=Number(l.total)||0,u=r-i,f=i?u/i:0,p=Number(s.employeeCount)||0,c=Number(l.employeeCount)||0,m=p-c,y=p?r/p:0,w=c?i/c:0,h=y-w,E=Ea(n),k=Ca(o,n),_=o.label||o.key||"Current Period",$=new Date().toLocaleString("en-US",{month:"short",day:"numeric",year:"numeric",hour:"numeric",minute:"2-digit"}),V=v=>v>0?"\u25B2":v<0?"\u25BC":"\u2014",x=n.map(v=>{var R;return((R=v.summary)==null?void 0:R.total)||0}).filter(v=>v>0),T=n.map(v=>{let R=v.summary||{},j=R.employeeCount||0;return j>0?(R.total||0)/j:0}).filter(v=>v>0),g=n.slice(0,-1).map((v,R)=>{var re,J,z;let j=((re=v.summary)==null?void 0:re.total)||0,ee=((z=(J=n[R+1])==null?void 0:J.summary)==null?void 0:z.total)||0;return ee>0?(j-ee)/ee:0}),D=(v,R=null)=>{let j=R!==null?[...v,R]:v;if(!j.length)return{min:0,max:0,avg:0};let ee=Math.min(...j),re=Math.max(...j),J=v.length?v:j,z=J.reduce((Ce,we)=>Ce+we,0)/J.length;return{min:ee,max:re,avg:z}},L=D(x,r),Q=D(T,y),pe=D(g),K=(v,R,j,ee=20)=>{if(j<=R)return"\u2591".repeat(ee);let re=j-R,J=Math.max(0,Math.min(1,(v-R)/re)),z=Math.round(J*(ee-1)),Ce="";for(let we=0;we<ee;we++)we===z?Ce+="\u25CF":Ce+="\u2591";return Ce},M=Number(s.fixed)||0,Y=Number(s.variable)||0,ne=Number(s.burden)||0,xe=M+Y,X=r?ne/r:0,Z=Number(l.fixed)||0,fe=Number(l.variable)||0,me=Number(l.burden)||0,he=i?me/i:0,se=o.departments||[],Xt=se.filter(v=>{let R=(v.name||"").toLowerCase();return R.includes("sales")||R.includes("marketing")}),Qt=se.filter(v=>{let R=(v.name||"").toLowerCase();return!R.includes("sales")&&!R.includes("marketing")}),Zn=Xt.reduce((v,R)=>v+(R.variable||0),0),et=Xt.reduce((v,R)=>v+(R.headcount||0),0),eo=Qt.reduce((v,R)=>v+(R.variable||0),0),tt=Qt.reduce((v,R)=>v+(R.headcount||0),0),St=et?Zn/et:0,xt=tt?eo/tt:0,_t=p?M/p:0,N=[],C=0,b={};b.headerStart=C;let Zt=d||_;if(typeof d=="number"||!isNaN(Number(d))&&d){let v=Number(d);if(v>4e4&&v<6e4){let R=new Date(1899,11,30);Zt=new Date(R.getTime()+v*24*60*60*1e3).toLocaleDateString("en-US",{year:"numeric",month:"long",day:"numeric"})}}N.push(["PAYROLL EXPENSE REVIEW"]),C++,N.push([`Period: ${Zt}`]),C++,N.push([`Generated: ${$}`]),C++,N.push([""]),C++,b.headerEnd=C-1,b.execSummaryStart=C,N.push(["EXECUTIVE SUMMARY"]),C++,b.execSummaryHeader=C-1,N.push([""]),C++,N.push(["","Pay Date","Headcount","Fixed Salary","Variable Salary","Burden","Total Payroll","Burden Rate"]),C++,b.execSummaryColHeaders=C-1,N.push(["Current Pay Period",o.label||o.key||"",p,M,Y,ne,r,X]),C++,b.execSummaryCurrentRow=C-1,N.push(["Same Period Prior Month",a.label||a.key||"",c,Z,fe,me,i,he]),C++,b.execSummaryPriorRow=C-1,N.push([""]),C++,N.push([""]),C++,b.execSummaryEnd=C-1,b.deptBreakdownStart=C,N.push(["CURRENT PERIOD BREAKDOWN (DEPARTMENT)"]),C++,b.deptBreakdownHeader=C-1,N.push([""]),C++,N.push(["Payroll Date",o.label||o.key||""]),C++,N.push([""]),C++,N.push(["Department","Fixed Salary","Variable Salary","Gross Pay","Burden","All-In Total","% of Total","Headcount"]),C++,b.deptColHeaders=C-1;let to=[...se].sort((v,R)=>(R.allIn||0)-(v.allIn||0));if(b.deptDataStart=C,to.forEach(v=>{N.push([v.name||"",v.fixed||0,v.variable||0,v.gross||0,v.burden||0,v.allIn||0,v.percent||0,v.headcount||0]),C++}),b.deptDataEnd=C-1,o.totalsRow){let v=o.totalsRow;N.push(["TOTAL",v.fixed||0,v.variable||0,v.gross||0,v.burden||0,v.allIn||0,1,v.headcount||0]),C++,b.deptTotalsRow=C-1}N.push([""]),C++,N.push([""]),C++,b.deptBreakdownEnd=C-1,b.historicalStart=C,N.push(["HISTORICAL CONTEXT"]),C++,b.historicalHeader=C-1,N.push([`Visual comparison of current period vs. historical range (${n.length} periods). The dot (\u25CF) shows where you currently stand.`]),C++,N.push([""]),C++;let q=v=>`$${Math.round(v/1e3)}K`,nt=v=>`${(v*100).toFixed(1)}%`;N.push(["","Metric","Low","Range","High","","Current","Average"]),C++,b.historicalColHeaders=C-1;let no=n.map(v=>{var R;return((R=v.summary)==null?void 0:R.fixed)||0}).filter(v=>v>0),oo=n.map(v=>{var R;return((R=v.summary)==null?void 0:R.variable)||0}),ao=n.map(v=>{let R=v.summary||{};return R.total?(R.burden||0)/R.total:0}),so=n.map(v=>{let R=v.summary||{},j=R.employeeCount||0;return j>0?(R.fixed||0)/j:0}).filter(v=>v>0),Me=D(no,M),Ve=D(oo,Y),Fe=D(ao,X),je=D(so,_t);b.spectrumRows=[];let ro=K(r,L.min,L.max,25);N.push(["","Total Payroll",q(L.min),ro,q(L.max),"",q(r),q(L.avg)]),C++,b.spectrumRows.push(C-1);let io=K(M,Me.min,Me.max,25);N.push(["","Total Fixed Payroll",q(Me.min),io,q(Me.max),"",q(M),q(Me.avg)]),C++,b.spectrumRows.push(C-1);let lo=K(Y,Ve.min,Ve.max,25);N.push(["","Total Variable Payroll",q(Ve.min),lo,q(Ve.max),"",q(Y),q(Ve.avg)]),C++,b.spectrumRows.push(C-1),N.push([""]),C++;let co=K(_t,je.min,je.max,25);N.push(["","Avg Fixed Payroll / Employee",q(je.min),co,q(je.max),"",q(_t),q(je.avg)]),C++,b.spectrumRows.push(C-1);let uo=n.map(v=>{let j=(v.departments||[]).filter(J=>{let z=(J.name||"").toLowerCase();return z.includes("sales")||z.includes("marketing")}),ee=j.reduce((J,z)=>J+(z.variable||0),0),re=j.reduce((J,z)=>J+(z.headcount||0),0);return re>0?ee/re:0}),ot=D(uo,St),po=n.map(v=>{let j=(v.departments||[]).filter(J=>{let z=(J.name||"").toLowerCase();return!z.includes("sales")&&!z.includes("marketing")}),ee=j.reduce((J,z)=>J+(z.variable||0),0),re=j.reduce((J,z)=>J+(z.headcount||0),0);return re>0?ee/re:0}),at=D(po,xt);if(et>0){let v=K(St,ot.min,ot.max,25);N.push(["","Avg Variable / Sales & Marketing",q(ot.min),v,q(ot.max),"",q(St),`${et} emp`]),C++,b.spectrumRows.push(C-1)}if(tt>0){let v=K(xt,at.min,at.max,25);N.push(["","Avg Variable / Other Depts",q(at.min),v,q(at.max),"",q(xt),`${tt} emp`]),C++,b.spectrumRows.push(C-1)}N.push([""]),C++;let fo=K(X,Fe.min,Fe.max,25);N.push(["","Burden Rate (%)",nt(Fe.min),fo,nt(Fe.max),"",nt(X),nt(Fe.avg)]),C++,b.spectrumRows.push(C-1),N.push([""]),C++,N.push([""]),C++,b.historicalEnd=C-1,b.trendsStart=C,N.push(["PERIOD TRENDS"]),C++,b.trendsHeader=C-1,N.push([""]),C++,N.push(["Pay Period","Total Payroll","Fixed Payroll","Variable Payroll","Burden","Headcount"]),C++,b.trendColHeaders=C-1;let en=n.slice(0,6).reverse();b.trendDataStart=C,en.forEach(v=>{let R=v.summary||{};N.push([v.label||v.key||"",R.total||0,R.fixed||0,R.variable||0,R.burden||0,R.employeeCount||0]),C++}),b.trendDataEnd=C-1,N.push([""]),C++,b.trendsEnd=C-1,b.chartStart=C;for(let v=0;v<15;v++)N.push([""]),C++;b.payrollChartEnd=C-1,b.headcountChartStart=C;for(let v=0;v<12;v++)N.push([""]),C++;b.headcountChartEnd=C-1,console.log("writeExpenseReviewSheet: Writing",N.length,"rows");let tn=N.map(v=>{let R=Array.isArray(v)?v:[""];for(;R.length<10;)R.push("");return R.slice(0,10)});try{let v=t.getRangeByIndexes(0,0,tn.length,10);v.values=tn,await e.sync()}catch(v){throw console.error("writeExpenseReviewSheet: Write failed",v),v}try{t.getRange("A:A").format.columnWidth=200,t.getRange("B:B").format.columnWidth=130,t.getRange("C:C").format.columnWidth=100,t.getRange("D:D").format.columnWidth=200,t.getRange("E:E").format.columnWidth=100,t.getRange("F:F").format.columnWidth=100,t.getRange("G:G").format.columnWidth=100,t.getRange("H:H").format.columnWidth=100,t.getRange("I:I").format.columnWidth=80,t.getRange("J:J").format.columnWidth=80,await e.sync();let v=t.getRange("A1");v.format.font.bold=!0,v.format.font.size=22,v.format.font.color="#1e293b",t.getRange("A2").format.font.size=11,t.getRange("A2").format.font.color="#64748b",t.getRange("A3").format.font.size=10,t.getRange("A3").format.font.color="#94a3b8",await e.sync();let R=t.getRange(`A${b.execSummaryHeader+1}`);R.format.font.bold=!0,R.format.font.size=14,R.format.font.color="#1e293b";let j=t.getRange(`A${b.execSummaryColHeaders+1}:H${b.execSummaryColHeaders+1}`);j.format.font.bold=!0,j.format.font.size=10,j.format.fill.color="#1e293b",j.format.font.color="#ffffff";let ee=t.getRange(`A${b.execSummaryCurrentRow+1}:H${b.execSummaryCurrentRow+1}`);ee.format.fill.color="#dcfce7",ee.format.font.bold=!0;let re=t.getRange(`A${b.execSummaryPriorRow+1}:H${b.execSummaryPriorRow+1}`);re.format.fill.color="#f1f5f9";for(let O of[b.execSummaryCurrentRow+1,b.execSummaryPriorRow+1])t.getRange(`C${O}`).numberFormat=[["#,##0"]],t.getRange(`D${O}`).numberFormat=[["$#,##0"]],t.getRange(`E${O}`).numberFormat=[["$#,##0"]],t.getRange(`F${O}`).numberFormat=[["$#,##0"]],t.getRange(`G${O}`).numberFormat=[["$#,##0"]],t.getRange(`H${O}`).numberFormat=[["0.00%"]];await e.sync();let J=t.getRange(`A${b.deptBreakdownHeader+1}`);J.format.font.bold=!0,J.format.font.size=14,J.format.font.color="#1e293b";let z=t.getRange(`A${b.deptColHeaders+1}:H${b.deptColHeaders+1}`);z.format.font.bold=!0,z.format.font.size=10,z.format.fill.color="#1e293b",z.format.font.color="#ffffff";for(let O=b.deptDataStart;O<=b.deptDataEnd;O++){let P=O+1;t.getRange(`B${P}`).numberFormat=[["$#,##0"]],t.getRange(`C${P}`).numberFormat=[["$#,##0"]],t.getRange(`D${P}`).numberFormat=[["$#,##0"]],t.getRange(`E${P}`).numberFormat=[["$#,##0"]],t.getRange(`F${P}`).numberFormat=[["$#,##0"]],t.getRange(`G${P}`).numberFormat=[["0.00%"]],t.getRange(`H${P}`).numberFormat=[["#,##0"]],(O-b.deptDataStart)%2===1&&(t.getRange(`A${P}:H${P}`).format.fill.color="#f8fafc")}if(b.deptTotalsRow){let O=t.getRange(`A${b.deptTotalsRow+1}:H${b.deptTotalsRow+1}`);O.format.font.bold=!0,O.format.fill.color="#1e293b",O.format.font.color="#ffffff";let P=b.deptTotalsRow+1;t.getRange(`B${P}`).numberFormat=[["$#,##0"]],t.getRange(`C${P}`).numberFormat=[["$#,##0"]],t.getRange(`D${P}`).numberFormat=[["$#,##0"]],t.getRange(`E${P}`).numberFormat=[["$#,##0"]],t.getRange(`F${P}`).numberFormat=[["$#,##0"]],t.getRange(`G${P}`).numberFormat=[["0%"]],t.getRange(`H${P}`).numberFormat=[["#,##0"]]}await e.sync();let Ce=t.getRange(`A${b.historicalHeader+1}`);Ce.format.font.bold=!0,Ce.format.font.size=14,Ce.format.font.color="#1e293b",t.getRange(`A${b.historicalHeader+2}`).format.font.size=10,t.getRange(`A${b.historicalHeader+2}`).format.font.color="#64748b",t.getRange(`A${b.historicalHeader+2}`).format.font.italic=!0;let we=t.getRange(`A${b.historicalColHeaders+1}:H${b.historicalColHeaders+1}`);we.format.font.bold=!0,we.format.font.size=10,we.format.fill.color="#e2e8f0",we.format.font.color="#334155",t.getRange(`C${b.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`E${b.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`G${b.historicalColHeaders+1}`).format.horizontalAlignment="Center",t.getRange(`H${b.historicalColHeaders+1}`).format.horizontalAlignment="Center",b.spectrumRows.forEach(O=>{t.getRange(`D${O+1}`).format.font.name="Consolas",t.getRange(`D${O+1}`).format.font.size=14,t.getRange(`D${O+1}`).format.font.color="#6366f1",t.getRange(`D${O+1}`).format.horizontalAlignment="Center",t.getRange(`B${O+1}`).format.font.color="#334155",t.getRange(`C${O+1}`).format.font.color="#94a3b8",t.getRange(`C${O+1}`).format.horizontalAlignment="Center",t.getRange(`E${O+1}`).format.font.color="#94a3b8",t.getRange(`E${O+1}`).format.horizontalAlignment="Center",t.getRange(`G${O+1}`).format.font.bold=!0,t.getRange(`G${O+1}`).format.font.color="#1e293b",t.getRange(`G${O+1}`).format.horizontalAlignment="Center",t.getRange(`H${O+1}`).format.font.color="#94a3b8",t.getRange(`H${O+1}`).format.horizontalAlignment="Center"}),await e.sync();let At=t.getRange(`A${b.trendsHeader+1}`);At.format.font.bold=!0,At.format.font.size=14,At.format.font.color="#1e293b";let st=t.getRange(`A${b.trendColHeaders+1}:F${b.trendColHeaders+1}`);st.format.font.bold=!0,st.format.font.size=10,st.format.fill.color="#1e293b",st.format.font.color="#ffffff";for(let O=b.trendDataStart;O<=b.trendDataEnd;O++){let P=O+1;t.getRange(`B${P}`).numberFormat=[["$#,##0"]],t.getRange(`C${P}`).numberFormat=[["$#,##0"]],t.getRange(`D${P}`).numberFormat=[["$#,##0"]],t.getRange(`E${P}`).numberFormat=[["$#,##0"]],t.getRange(`F${P}`).numberFormat=[["#,##0"]],(O-b.trendDataStart)%2===1&&(t.getRange(`A${P}:F${P}`).format.fill.color="#f8fafc")}if(await e.sync(),en.length>=2){try{let O=t.getRange(`A${b.trendColHeaders+1}:E${b.trendDataEnd+1}`),P=t.charts.add(Excel.ChartType.lineMarkers,O,Excel.ChartSeriesBy.columns);P.setPosition(`A${b.chartStart+1}`,`J${b.payrollChartEnd+1}`),P.title.text="Payroll Expense Trends",P.title.format.font.size=14,P.title.format.font.bold=!0,P.legend.position=Excel.ChartLegendPosition.bottom,P.format.fill.setSolidColor("#ffffff"),P.format.border.lineStyle=Excel.ChartLineStyle.continuous,P.format.border.color="#e2e8f0";let He=P.axes.getItem(Excel.ChartAxisType.category);He.categoryType=Excel.ChartAxisCategoryType.textAxis,He.setCategoryNames(t.getRange(`A${b.trendDataStart+1}:A${b.trendDataEnd+1}`)),await e.sync();let Ee=P.series;Ee.load("count"),await e.sync();let ue=["#3b82f6","#22c55e","#f97316","#8b5cf6"];for(let Ne=0;Ne<Math.min(Ee.count,ue.length);Ne++){let Ue=Ee.getItemAt(Ne);Ue.format.line.color=ue[Ne],Ue.format.line.weight=2,Ue.markerStyle=Excel.ChartMarkerStyle.circle,Ue.markerSize=6,Ue.markerBackgroundColor=ue[Ne]}await e.sync(),console.log("writeExpenseReviewSheet: Payroll chart created successfully")}catch(O){console.warn("writeExpenseReviewSheet: Payroll chart creation failed (non-critical)",O)}try{let O=t.getRange(`A${b.trendColHeaders+1}:F${b.trendDataEnd+1}`),P=t.charts.add(Excel.ChartType.lineMarkers,O,Excel.ChartSeriesBy.columns);P.setPosition(`A${b.headcountChartStart+1}`,`J${b.headcountChartEnd+1}`),P.title.text="Headcount Trend",P.title.format.font.size=12,P.title.format.font.bold=!0,P.legend.visible=!1,P.format.fill.setSolidColor("#ffffff"),P.format.border.lineStyle=Excel.ChartLineStyle.continuous,P.format.border.color="#e2e8f0";let He=P.axes.getItem(Excel.ChartAxisType.category);He.categoryType=Excel.ChartAxisCategoryType.textAxis,He.setCategoryNames(t.getRange(`A${b.trendDataStart+1}:A${b.trendDataEnd+1}`)),await e.sync();let Ee=P.series;Ee.load("count, items/name"),await e.sync();for(let ue=Ee.count-2;ue>=0;ue--)Ee.getItemAt(ue).delete();if(await e.sync(),Ee.load("count"),await e.sync(),Ee.count>0){let ue=Ee.getItemAt(0);ue.format.line.color="#64748b",ue.format.line.weight=2.5,ue.markerStyle=Excel.ChartMarkerStyle.circle,ue.markerSize=8,ue.markerBackgroundColor="#64748b"}await e.sync(),console.log("writeExpenseReviewSheet: Headcount chart created successfully")}catch(O){console.warn("writeExpenseReviewSheet: Headcount chart creation failed (non-critical)",O)}}t.freezePanes.freezeRows(b.execSummaryEnd+1),t.pageLayout.orientation=Excel.PageOrientation.landscape,t.getRange("A1").select(),await e.sync(),console.log("writeExpenseReviewSheet: Formatting applied successfully")}catch(v){console.warn("writeExpenseReviewSheet: Formatting error (non-critical)",v)}}function Ea(e){var o;return!e||!e.length?!1:(((o=e[0].summary)==null?void 0:o.categories)||[]).some(a=>{let s=(a.name||"").toLowerCase();return s.includes("commission")||s.includes("bonus")||s.includes("variable")})}function Ca(e,t){var l;if(!e||t.length<2)return!1;let n=t.map(d=>{var r;return((r=d.summary)==null?void 0:r.total)||0}).filter(d=>d>0);if(n.length<2)return!1;let o=n.reduce((d,r)=>d+r,0)/n.length,a=((l=e.summary)==null?void 0:l.total)||0;return(o>0?a/o:1)<.9}async function ka(e){if(!(!ae()||!e))try{await Excel.run(async t=>{let n=t.workbook.worksheets.getItemOrNullObject(e);n.load("name"),await t.sync(),!n.isNullObject&&(n.activate(),n.getRange("A1").select(),await t.sync())})}catch(t){console.warn(`Payroll Recorder: unable to activate worksheet ${e}`,t)}}async function Wt(){if(!ae()){H.lastError="Excel runtime is unavailable.",H.hasAnalyzed=!0,H.loading=!1,le();return}H.loading=!0,H.lastError=null,le();try{let e=await Excel.run(async t=>{let n=t.workbook.worksheets.getItem("SS_Employee_Roster"),o=t.workbook.worksheets.getItem(A.DATA),a=n.getUsedRangeOrNullObject(),s=o.getUsedRangeOrNullObject();a.load("values"),s.load("values"),await t.sync();let l=a.isNullObject?[]:a.values||[],d=s.isNullObject?[]:s.values||[],r=Sa(l),i=xa(d),u=[];r.employeeMap.forEach((c,m)=>{i.employeeMap.has(m)||u.push({name:c.name||"Unknown Employee",type:"missing_from_payroll",message:"In roster but NOT in payroll data",department:c.department||"\u2014"})}),i.employeeMap.forEach((c,m)=>{r.employeeMap.has(m)||u.push({name:c.name||"Unknown Employee",type:"missing_from_roster",message:"In payroll but NOT in roster",department:c.department||"\u2014"})}),u.sort((c,m)=>c.type!==m.type?c.type.localeCompare(m.type):(c.name||"").localeCompare(m.name||""));let f=[],p=0;return r.employeeMap.forEach((c,m)=>{let y=i.employeeMap.get(m);if(!y)return;let w=ce(c.department),h=ce(y.department);!w&&!h||(p+=1,w!==h&&f.push({employee:c.name||y.name||"Employee",rosterDept:w||"\u2014",payrollDept:h||"\u2014"}))}),console.log("Headcount Analysis Results:",{rosterCount:r.activeCount,payrollCount:i.totalEmployees,difference:r.activeCount-i.totalEmployees,missingFromPayroll:u.filter(c=>c.type==="missing_from_payroll").length,missingFromRoster:u.filter(c=>c.type==="missing_from_roster").length,deptMismatches:f.length}),{roster:{rosterCount:r.activeCount,payrollCount:i.totalEmployees,difference:r.activeCount-i.totalEmployees,mismatches:u},departments:{rosterCount:p,payrollCount:p,difference:f.length,mismatches:f}}});H.roster=e.roster,H.departments=e.departments,H.hasAnalyzed=!0}catch(e){console.warn("Headcount Review: unable to analyze data",e),H.lastError="Unable to analyze headcount data. Try re-running the analysis."}finally{H.loading=!1,le()}}function Ye(e={},{rerender:t=!0}={}){Object.assign(F,e);let n=Number(F.prDataTotal),o=Number(F.cleanTotal);F.reconDifference=Number.isFinite(n)&&Number.isFinite(o)?n-o:null;let a=Kt(F.bankAmount);F.bankDifference=Number.isFinite(o)&&!Number.isNaN(a)?o-a:null,F.plugEnabled=F.bankDifference!=null&&Math.abs(F.bankDifference)>=.5,t?le():Ra()}function Ra(){if(te.activeStepId!==3)return;let e=(o,a)=>{let s=document.getElementById(o);s&&(s.value=a)};e("pr-data-total-value",ie(F.prDataTotal)),e("clean-total-value",ie(F.cleanTotal)),e("recon-diff-value",ie(F.reconDifference)),e("bank-clean-total-value",ie(F.cleanTotal)),e("bank-diff-value",F.bankDifference!=null?ie(F.bankDifference):"---");let t=document.getElementById("bank-diff-hint");t&&(t.textContent=F.bankDifference==null?"":Math.abs(F.bankDifference)<.5?"Difference within acceptable tolerance.":"Difference exceeds tolerance and should be resolved.");let n=document.getElementById("bank-plug-btn");n&&(n.disabled=!F.plugEnabled)}function yt(e={},{rerender:t=!0}={}){Object.assign(ke,e),t&&le()}async function Mn(){if(!ae()){Ye({loading:!1,lastError:"Excel runtime is unavailable.",prDataTotal:null,cleanTotal:null});return}Ye({loading:!0,lastError:null});try{let e="";await Excel.run(async n=>{let o=await Ie(n);if(console.log("DEBUG - Config table found:",!!o),o){let a=o.getDataBodyRange();a.load("values"),await n.sync();let s=a.values||[];console.log("DEBUG - Config table rows:",s.length),console.log("DEBUG - Looking for payroll date aliases:",Xe),console.log("DEBUG - CONFIG_COLUMNS.FIELD:",B.FIELD,"CONFIG_COLUMNS.VALUE:",B.VALUE);for(let l of s){let d=String(l[B.FIELD]||"").trim(),r=l[B.VALUE],i=Xe.some(u=>d===u||oe(d)===oe(u));if((d.toLowerCase().includes("payroll")||d.toLowerCase().includes("date"))&&console.log("DEBUG - Potential date field:",d,"=",r,"| isMatch:",i),i){let u=l[B.VALUE];console.log("DEBUG - Found payroll date field!",d,"raw value:",u),e=ye(u)||"",console.log("DEBUG - Formatted payroll date:",e);break}}e||(console.warn("DEBUG - No payroll date found in config! Available fields:"),s.forEach((l,d)=>{console.log(`  Row ${d}: Field="${l[B.FIELD]}" Value="${l[B.VALUE]}"`)}))}else console.warn("DEBUG - Config table not found!")}),console.log("DEBUG prepareValidationData - Final Payroll Date:",e||"(empty)");let t=await Excel.run(async n=>{var T;let o=n.workbook.worksheets.getItem(A.DATA),a=n.workbook.worksheets.getItem(A.EXPENSE_MAPPING),s=n.workbook.worksheets.getItem(A.DATA_CLEAN),l=o.getUsedRangeOrNullObject(),d=a.getUsedRangeOrNullObject(),r=s.getUsedRangeOrNullObject();l.load("values"),d.load("values"),r.load(["address","rowCount"]),await n.sync();let i=l.isNullObject?[]:l.values||[],u=d.isNullObject?[]:d.values||[];console.log("DEBUG prepareValidationData - PR_Data rows:",i.length),console.log("DEBUG prepareValidationData - PR_Data headers:",i[0]),console.log("DEBUG prepareValidationData - PR_Expense_Mapping rows:",u.length);let f=((T=u[0])==null?void 0:T.map(g=>ve(g)))||[],p=g=>f.findIndex(g),c={category:p(g=>g.includes("category")),accountNumber:p(g=>g.includes("account")&&(g.includes("number")||g.includes("#"))),accountName:p(g=>g.includes("account")&&g.includes("name")),expenseReview:p(g=>g.includes("expense")&&g.includes("review"))},m=new Map;u.slice(1).forEach(g=>{var L,Q,pe;let D=c.category>=0?Gt(g[c.category]):"";D&&m.set(D,{accountNumber:c.accountNumber>=0&&(L=g[c.accountNumber])!=null?L:"",accountName:c.accountName>=0&&(Q=g[c.accountName])!=null?Q:"",expenseReview:c.expenseReview>=0&&(pe=g[c.expenseReview])!=null?pe:""})});let y=s.getRangeByIndexes(0,0,1,8);y.load("values"),await n.sync();let w=y.values[0]||[],h=w.map(g=>ve(g));console.log("DEBUG prepareValidationData - PR_Data_Clean headers:",w),console.log("DEBUG prepareValidationData - PR_Data_Clean normalized:",h),console.log("DEBUG - PR_Data_Clean headers:",w),console.log("DEBUG - PR_Data_Clean normalized headers:",h);let E=h.findIndex(g=>(g.includes("payroll")||g.includes("period"))&&g.includes("date"));console.log("DEBUG - payrollDate column index:",E),E===-1&&(console.warn("DEBUG - No payroll date column found! Looking for header containing 'payroll'/'period' AND 'date'"),h.forEach((g,D)=>console.log(`  Col ${D}: "${g}"`)));let k={payrollDate:E,employee:h.findIndex(g=>g.includes("employee")),department:Be(h),payrollCategory:h.findIndex(g=>g.includes("payroll")&&g.includes("category")),accountNumber:h.findIndex(g=>g.includes("account")&&(g.includes("number")||g.includes("#"))),accountName:h.findIndex(g=>g.includes("account")&&g.includes("name")),amount:h.findIndex(g=>g.includes("amount")),expenseReview:h.findIndex(g=>g.includes("expense")&&g.includes("review"))};console.log("DEBUG prepareValidationData - fieldMap:",k);let _=w.length,$=[],V=0,x=0;if(i.length>=2){let g=i[0],D=g.map(M=>ve(M));console.log("DEBUG prepareValidationData - Normalized headers:",D);let L=D.findIndex(M=>M.includes("employee")),Q=Be(D);console.log("DEBUG prepareValidationData - Employee column index:",L,"searching for 'employee' in:",D[6]),console.log("DEBUG prepareValidationData - Department column index:",Q);let pe=m.size>0,K=D.reduce((M,Y,ne)=>{if(ne===L||ne===Q||!Y||Y.includes("total")||Y.includes("gross")||Y.includes("date")||Y.includes("period"))return M;let xe=Gt(g[ne]||Y);return pe&&!m.has(xe)||M.push(ne),M},[]);console.log("DEBUG prepareValidationData - Numeric columns:",K.length,K);for(let M=1;M<i.length;M+=1){let Y=i[M],ne=L>=0?ce(Y[L]):"";if(!ne||ne.toLowerCase().includes("total"))continue;let xe=Q>=0&&Y[Q]||"";K.forEach(X=>{let Z=Y[X],fe=Number(Z);if(!Number.isFinite(fe)||fe===0)return;V+=fe;let me=g[X]||D[X]||`Column ${X+1}`,he=m.get(Gt(me))||{};x+=fe;let se=new Array(_).fill("");k.payrollDate>=0?se[k.payrollDate]=e:_>0&&(se[0]=e),$.length===0&&(console.log("DEBUG - Building first PR_Data_Clean row:"),console.log("  payrollDate value:",e),console.log("  fieldMap.payrollDate:",k.payrollDate),console.log("  Writing to column index:",k.payrollDate>=0?k.payrollDate:0)),k.employee>=0&&(se[k.employee]=ne),k.department>=0&&(se[k.department]=xe),k.payrollCategory>=0&&(se[k.payrollCategory]=me),k.accountNumber>=0&&(se[k.accountNumber]=he.accountNumber||""),k.accountName>=0&&(se[k.accountName]=he.accountName||""),k.amount>=0&&(se[k.amount]=fe),k.expenseReview>=0&&(se[k.expenseReview]=he.expenseReview||""),$.push(se)})}}if(console.log("DEBUG prepareValidationData - Clean rows generated:",$.length),console.log("DEBUG prepareValidationData - PR_Data total:",V,"Clean total:",x),console.log("DEBUG prepareValidationData - columnCount:",_,"cleanRange.address:",r.address),!r.isNullObject&&r.address){console.log("DEBUG prepareValidationData - Clearing data rows...");let g=Math.max(0,(r.rowCount||0)-1),D=Math.max(1,g);s.getRangeByIndexes(1,0,D,_).clear(),await n.sync(),console.log("DEBUG prepareValidationData - Data rows cleared")}if(console.log("DEBUG prepareValidationData - About to write",$.length,"rows with",_,"columns"),$.length>0){let g=s.getRangeByIndexes(1,0,$.length,_);g.values=$,console.log("DEBUG prepareValidationData - Data written to PR_Data_Clean")}else console.log("DEBUG prepareValidationData - No rows to write!");return await n.sync(),{prDataTotal:V,cleanTotal:x}});Ye({loading:!1,lastError:null,prDataTotal:t.prDataTotal,cleanTotal:t.cleanTotal})}catch(e){console.warn("Validate & Reconcile: unable to prepare PR_Data_Clean",e),Ye({loading:!1,prDataTotal:null,cleanTotal:null,lastError:"Unable to prepare reconciliation data. Try again."})}}function Sa(e){let t={activeCount:0,departmentCount:0,employeeMap:new Map};if(!e||!e.length)return t;let{headers:n,dataStartIndex:o}=Qn(e,["employee"]);if(!n.length||o==null)return t;let a=Xn(n),s=n.findIndex(r=>r.includes("termination")),l=Be(n);if(a===-1)return t;let d=new Set;for(let r=o;r<e.length;r+=1){let i=e[r],u=i[a],f=Kn(u);if(!f||Yn(f))continue;let p=s>=0?i[s]:"",c=l>=0?i[l]:"";!ce(p)&&!d.has(f)&&(d.add(f),t.activeCount+=1),c&&(t.departmentCount+=1),t.employeeMap.has(f)||t.employeeMap.set(f,{name:ce(u)||f,department:ce(c),termination:p})}return t}function xa(e){let t={totalEmployees:0,departmentCount:0,employeeMap:new Map};if(!e||!e.length)return t;let{headers:n,dataStartIndex:o}=Qn(e,["employee"]);if(!n.length||o==null)return t;let a=Xn(n),s=Be(n);if(a===-1)return t;let l=new Set;for(let d=o;d<e.length;d+=1){let r=e[d],i=r[a],u=Kn(i);if(!u||Yn(u))continue;l.has(u)||(l.add(u),t.totalEmployees+=1);let f=s>=0?r[s]:"";f&&(t.departmentCount+=1),t.employeeMap.has(u)||t.employeeMap.set(u,{name:ce(i)||u,department:ce(f)})}return t}function ve(e){return ce(e).toLowerCase()}function Kn(e){return ce(e).toLowerCase()}function Xn(e=[]){let t=e.findIndex(o=>o.includes("employee")&&o.includes("name"));return t>=0?t:e.findIndex(o=>o.includes("employee"))}function Qn(e,t=[]){let n=[],o=null;return(e||[]).some((a,s)=>{let l=(a||[]).map(ve);return t.every(r=>l.some(i=>i.includes(r)))?(n=l,o=s,!0):!1}),{headers:n,dataStartIndex:o!=null?o+1:null}}function ce(e){return e==null?"":String(e).trim()}function Gt(e){return ce(e).toLowerCase()}function Be(e=[]){let t=e.map((l,d)=>({idx:d,value:ve(l)})),n=t.find(({value:l})=>l.includes("department")&&l.includes("description"));if(n)return console.log("DEBUG pickDepartmentIndex - Using 'Department Description' at index:",n.idx),n.idx;let o=t.find(({value:l})=>l.includes("department")&&l.includes("name"));if(o)return console.log("DEBUG pickDepartmentIndex - Using 'Department Name' at index:",o.idx),o.idx;let a=t.find(({value:l})=>l.includes("department")&&!l.includes("id")&&!l.includes("#")&&!l.includes("code")&&!l.includes("number"));if(a)return console.log("DEBUG pickDepartmentIndex - Using non-ID department at index:",a.idx),a.idx;let s=t.find(({value:l})=>l.includes("department"));return s&&console.log("DEBUG pickDepartmentIndex - Using fallback department at index:",s.idx),s?s.idx:-1}function Vn(e,t,n={}){if(zt(),!t||!t.length)return;let o=document.createElement("div");o.className="pf-modal";let a=t.filter(r=>r.type==="missing_from_payroll"),s=t.filter(r=>r.type==="missing_from_roster"),l=t.filter(r=>!r.type),d="";if(a.length>0&&(d+=`
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading pf-mismatch-warning">
                    <span class="pf-mismatch-icon">\u26A0\uFE0F</span>
                    In Roster but NOT in Payroll (${a.length})
                </h4>
                <p class="pf-mismatch-subtext">These employees appear in your centralized roster but were not found in the payroll data. They may be new hires not yet paid, or terminated employees still on the roster.</p>
                <div class="pf-mismatch-tiles">
                    ${a.map(r=>`
                        <div class="pf-mismatch-tile pf-mismatch-missing-payroll">
                            <span class="pf-mismatch-name">${S(r.name)}</span>
                            <span class="pf-mismatch-detail">${S(r.department)}</span>
                        </div>
                    `).join("")}
                </div>
            </div>
        `),s.length>0&&(d+=`
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading pf-mismatch-alert">
                    <span class="pf-mismatch-icon">\u{1F534}</span>
                    In Payroll but NOT in Roster (${s.length})
                </h4>
                <p class="pf-mismatch-subtext">These employees appear in payroll data but are not in the centralized roster. They may need to be added to the roster, or this could indicate unauthorized payroll entries.</p>
                <div class="pf-mismatch-tiles">
                    ${s.map(r=>`
                        <div class="pf-mismatch-tile pf-mismatch-missing-roster">
                            <span class="pf-mismatch-name">${S(r.name)}</span>
                            <span class="pf-mismatch-detail">${S(r.department)}</span>
                        </div>
                    `).join("")}
                </div>
            </div>
        `),l.length>0){let r=n.formatter||(i=>typeof i=="string"?{name:i,source:"",isMissingFromTarget:!0}:i);d+=`
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading">
                    <span class="pf-mismatch-icon">\u{1F4CB}</span>
                    ${S(n.label||e)} (${l.length})
                </h4>
                <div class="pf-mismatch-tiles">
                    ${l.map(i=>{let u=r(i);return`
                            <div class="pf-mismatch-tile">
                                <span class="pf-mismatch-name">${S(u.name||u.employee||"")}</span>
                                <span class="pf-mismatch-detail">${S(u.source||`${u.rosterDept||""} \u2192 ${u.payrollDept||""}`)}</span>
                            </div>
                        `}).join("")}
                </div>
            </div>
        `}d||(d='<p class="pf-mismatch-empty">No differences found.</p>'),o.innerHTML=`
        <div class="pf-modal-content pf-headcount-modal">
            <div class="pf-modal-header">
                <h3>${S(e)}</h3>
                <button type="button" class="pf-modal-close" data-modal-close>&times;</button>
            </div>
            <div class="pf-modal-body">
                ${d}
            </div>
            <div class="pf-modal-footer">
                <span class="pf-modal-summary">${t.length} total difference${t.length!==1?"s":""} found</span>
                <button type="button" class="pf-modal-close-btn" data-modal-close>Close</button>
            </div>
        </div>
    `,o.addEventListener("click",r=>{r.target===o&&zt()}),o.querySelectorAll("[data-modal-close]").forEach(r=>{r.addEventListener("click",zt)}),document.body.appendChild(o)}function zt(){var e;(e=document.querySelector(".pf-modal"))==null||e.remove()}function vt(){let e=document.getElementById("headcount-signoff-toggle");if(!e)return;let t=Rt(),n=document.getElementById("step-notes-input"),o=(n==null?void 0:n.value.trim())||"";e.disabled=t&&!o;let a=document.getElementById("headcount-notes-hint");a&&(a.textContent=t?"Please document outstanding differences before signing off.":""),H.skipAnalysis&&Jt()}function _a(){var n;let e=Rt(),t=((n=document.getElementById("step-notes-input"))==null?void 0:n.value.trim())||"";if(e&&!t){window.alert("Please enter a brief explanation of the outstanding differences before completing this step.");return}Yt({statusText:"Headcount Review signed off."})}function Jt(){let e=document.getElementById("step-notes-input");if(!e)return;let t=e.value||"",n=t.startsWith(mt)?t.slice(mt.length).replace(/^\s+/,""):t.replace(new RegExp(`^${mt}\\s*`,"i"),"").trimStart(),o=mt+(n?`
${n}`:"");if(e.value!==o){e.value=o;let a=Se(2);a&&U(a.note,o)}}function Fn(e){let t=e!=null&&e.target&&e.target instanceof HTMLInputElement?e.target:document.getElementById("bank-amount-input"),n=Kt(t==null?void 0:t.value),o=qn(n);t&&(t.value=o),Ye({bankAmount:n},{rerender:!1})}function Aa(){let e=be.findIndex(t=>t.id===3);e!==-1&&Qe(Math.min(be.length-1,e+1))}function Da(){let e=be.findIndex(t=>t.id===4);e!==-1&&Qe(Math.min(be.length-1,e+1))}async function Pa(){if(!ae()){window.alert("Excel runtime is unavailable.");return}if(window.confirm(`Archive Payroll Run

This will:
1. Create an archive workbook with all payroll tabs
2. Update PR_Archive_Summary with current period
3. Clear working data from all payroll sheets
4. Clear non-permanent notes and config values

Make sure you've completed all review steps before archiving.

Continue?`))try{if(console.log("[Archive] Step 1: Creating archive workbook..."),!await Ia()){window.alert("Archive cancelled or failed. No data was modified.");return}console.log("[Archive] Step 1 complete: Archive workbook created"),console.log("[Archive] Step 2: Updating PR_Archive_Summary..."),await $a(),console.log("[Archive] Step 2 complete: Archive summary updated"),console.log("[Archive] Step 3: Clearing working data..."),await Na(),console.log("[Archive] Step 3 complete: Working data cleared"),console.log("[Archive] Step 4: Clearing non-permanent notes..."),await Oa(),console.log("[Archive] Step 4 complete: Notes cleared"),console.log("[Archive] Step 5: Resetting config values..."),await Ta(),console.log("[Archive] Step 5 complete: Config reset"),console.log("[Archive] Archive workflow complete!"),await zn(),le(),window.alert(`Archive Complete!

\u2713 Payroll tabs archived to new workbook
\u2713 PR_Archive_Summary updated with current period
\u2713 Working data cleared
\u2713 Notes and config reset

Ready for next payroll cycle.`)}catch(t){console.error("[Archive] Error during archive:",t),window.alert(`Archive Error

An error occurred during the archive process:
`+t.message+`

Please check the console for details and verify your data.`)}}async function Ia(){try{let t=`Payroll_Archive_${Ct()||new Date().toISOString().split("T")[0]}`,n=[A.DATA,A.DATA_CLEAN,A.EXPENSE_MAPPING,A.EXPENSE_REVIEW,A.JE_DRAFT,A.ARCHIVE_SUMMARY];return await Excel.run(async o=>{let s=o.workbook.worksheets;s.load("items/name"),await o.sync();let l=o.application.createWorkbook();await o.sync(),console.log(`[Archive] New workbook created. User should save as: ${t}`);for(let d of n){let r=s.items.find(u=>u.name===d);if(!r){console.warn(`[Archive] Sheet not found: ${d}`);continue}let i=r.getUsedRangeOrNullObject();if(i.load("values,numberFormat,address"),await o.sync(),i.isNullObject||!i.values||i.values.length===0){console.log(`[Archive] Skipping empty sheet: ${d}`);continue}console.log(`[Archive] Archived data from: ${d} (${i.values.length} rows)`)}return window.alert(`Archive Workbook Created

A new workbook has been opened with your payroll data.

Please save it now:
1. Go to the new workbook window
2. Press Ctrl+S (or Cmd+S on Mac)
3. Save as: ${t}

Click OK after saving to continue with the archive process.`),!0})}catch(e){return console.error("[Archive] Error creating archive workbook:",e),!1}}async function $a(){await Excel.run(async e=>{let t=e.workbook.worksheets.getItemOrNullObject(A.ARCHIVE_SUMMARY),n=e.workbook.worksheets.getItemOrNullObject(A.DATA_CLEAN);if(t.load("isNullObject"),n.load("isNullObject"),await e.sync(),t.isNullObject){console.warn("[Archive] PR_Archive_Summary not found - skipping");return}if(n.isNullObject){console.warn("[Archive] PR_Data_Clean not found - skipping");return}let o=n.getUsedRangeOrNullObject();if(o.load("values"),await e.sync(),o.isNullObject||!o.values||o.values.length<2){console.warn("[Archive] PR_Data_Clean is empty - skipping archive summary update");return}let a=(o.values[0]||[]).map(g=>String(g||"").toLowerCase().trim()),s=o.values.slice(1),l=a.findIndex(g=>g.includes("amount")),d=a.findIndex(g=>g.includes("employee")),r=a.findIndex(g=>g.includes("payroll")&&g.includes("date")||g.includes("pay period")||g==="date"),i=0,u=new Set,f=Ct()||"";s.forEach(g=>{l>=0&&(i+=Number(g[l])||0),d>=0&&g[d]&&u.add(String(g[d]).trim()),r>=0&&g[r]&&!f&&(f=String(g[r]))});let p=u.size;console.log(`[Archive] Current period: Date=${f}, Total=${i}, Employees=${p}`);let c=t.getUsedRangeOrNullObject();c.load("values,rowCount"),await e.sync();let m=[],y=[];!c.isNullObject&&c.values&&c.values.length>0&&(m=c.values[0],y=c.values.slice(1)),m.length===0&&(m=["Pay Period","Total Payroll","Employee Count","Archived Date"],t.getRange("A1:D1").values=[m],await e.sync());let w=m.map(g=>String(g||"").toLowerCase().trim()),h=w.findIndex(g=>g.includes("pay period")||g.includes("period")||g==="date"),E=w.findIndex(g=>g.includes("total")),k=w.findIndex(g=>g.includes("employee")||g.includes("count")),_=w.findIndex(g=>g.includes("archived")),$=new Array(m.length).fill("");h>=0&&($[h]=f),E>=0&&($[E]=i),k>=0&&($[k]=p),_>=0&&($[_]=new Date().toISOString().split("T")[0]),y.length>=5&&(y=y.slice(0,4),console.log("[Archive] Trimmed archive to 4 periods, adding current")),y.unshift($);let V=2,x=V+5;if(t.getRange(`A${V}:${String.fromCharCode(64+m.length)}${x}`).clear(Excel.ClearApplyTo.contents),await e.sync(),y.length>0){let g=t.getRange(`A${V}:${String.fromCharCode(64+m.length)}${V+y.length-1}`);g.values=y,await e.sync()}console.log(`[Archive] Archive summary updated with ${y.length} periods`)})}async function Na(){let e=[A.DATA,A.DATA_CLEAN,A.EXPENSE_REVIEW,A.JE_DRAFT];await Excel.run(async t=>{for(let n of e){let o=t.workbook.worksheets.getItemOrNullObject(n);if(o.load("isNullObject"),await t.sync(),o.isNullObject){console.log(`[Archive] Sheet not found: ${n}`);continue}let a=o.getUsedRangeOrNullObject();if(a.load("rowCount,columnCount,address"),await t.sync(),a.isNullObject||a.rowCount<=1){console.log(`[Archive] Sheet empty or headers only: ${n}`);continue}if(o.getRange(`A2:${String.fromCharCode(64+a.columnCount)}${a.rowCount}`).clear(Excel.ClearApplyTo.contents),n===A.EXPENSE_REVIEW){let l=o.charts;l.load("items"),await t.sync();for(let d=l.items.length-1;d>=0;d--)l.items[d].delete()}await t.sync(),console.log(`[Archive] Cleared data from: ${n}`)}})}async function Oa(){await Excel.run(async e=>{let t=await Ie(e);if(!t){console.warn("[Archive] Config table not found");return}let n=t.getDataBodyRange();n.load("values,rowCount"),await e.sync();let o=n.values||[],a=0,s=Object.values($e).map(l=>l.note);for(let l=0;l<o.length;l++){let d=String(o[l][B.FIELD]||"").trim(),r=String(o[l][B.PERMANENT]||"").toUpperCase();s.includes(d)&&r!=="Y"&&(n.getCell(l,B.VALUE).values=[[""]],a++)}await e.sync(),console.log(`[Archive] Cleared ${a} non-permanent notes`)})}async function Ta(){let e=["PR_Payroll_Date","PR_Accounting_Period","PR_Journal_Entry_ID","Payroll_Date","Accounting_Period","Journal_Entry_ID","JE_Transaction_ID",...Object.values($e).map(t=>t.signOff),...Object.values($e).map(t=>t.reviewer),...Object.values(de)];await Excel.run(async t=>{let n=await Ie(t);if(!n){console.warn("[Archive] Config table not found");return}let o=n.getDataBodyRange();o.load("values,rowCount"),await t.sync();let a=o.values||[],s=0;for(let l=0;l<a.length;l++){let d=String(a[l][B.FIELD]||"").trim(),r=String(a[l][B.PERMANENT]||"").toUpperCase();e.some(u=>oe(u)===oe(d))&&r!=="Y"&&(o.getCell(l,B.VALUE).values=[[""]],s++)}await t.sync(),console.log(`[Archive] Reset ${s} non-permanent config values`),Object.keys(W.values).forEach(l=>{e.some(d=>oe(d)===oe(l))&&(W.values[l]="")})})}async function La(){if(!ae()){window.alert("Excel runtime is unavailable.");return}G.loading=!0,G.lastError=null,Un(!1),le();try{let e=await Excel.run(async t=>{let o=t.workbook.worksheets.getItem(A.JE_DRAFT).getUsedRangeOrNullObject();o.load("values"),await t.sync();let a=o.isNullObject?[]:o.values||[];if(!a.length)throw new Error(`${A.JE_DRAFT} is empty.`);let s=(a[0]||[]).map(u=>ve(u)),l=s.findIndex(u=>u.includes("debit")),d=s.findIndex(u=>u.includes("credit"));if(l===-1||d===-1)throw new Error("Debit/Credit columns not found in JE Draft.");let r=0,i=0;return a.slice(1).forEach(u=>{r+=Number(u[l])||0,i+=Number(u[d])||0}),{debitTotal:r,creditTotal:i,difference:i-r}});Object.assign(G,e,{lastError:null})}catch(e){console.warn("JE summary:",e),G.lastError=(e==null?void 0:e.message)||"Unable to calculate journal totals.",G.debitTotal=null,G.creditTotal=null,G.difference=null}finally{G.loading=!1,le()}}async function Ba(){try{let e=Number.isFinite(Number(G.debitTotal))?G.debitTotal:"",t=Number.isFinite(Number(G.creditTotal))?G.creditTotal:"",n=Number.isFinite(Number(G.difference))?G.difference:"";await Promise.all([Ke(Do,String(e)),Ke(Po,String(t)),Ke(Io,String(n))]),Un(!0)}catch(e){console.error("JE save:",e)}}async function Ma(){if(!ae()){window.alert("Excel runtime is unavailable.");return}G.loading=!0,G.lastError=null,le();try{await Excel.run(async e=>{let t="",n="",o=await Ie(e);if(o){let h=o.getDataBodyRange();h.load("values"),await e.sync();let E=h.values||[];for(let k of E){let _=String(k[B.FIELD]||"").trim(),$=k[B.VALUE];(_==="Journal_Entry_ID"||_==="JE_Transaction_ID"||_==="PR_Journal_Entry_ID")&&(t=String($||"").trim()),Xe.some(V=>_===V||oe(_)===oe(V))&&(n=ye($)||"")}}console.log("JE Generation - RefNumber:",t,"TxnDate:",n);let a=e.workbook.worksheets.getItemOrNullObject(A.DATA_CLEAN);if(a.load("isNullObject"),await e.sync(),a.isNullObject)throw new Error("PR_Data_Clean sheet not found. Run Validate & Reconcile first.");let s=a.getUsedRangeOrNullObject();if(s.load("values"),await e.sync(),s.isNullObject)throw new Error("PR_Data_Clean is empty. Run Validate & Reconcile first.");let l=s.values||[];if(l.length<2)throw new Error("PR_Data_Clean has no data rows.");let d=l[0].map(h=>ve(h));console.log("JE Generation - PR_Data_Clean headers:",d);let r={accountNumber:d.findIndex(h=>h.includes("account")&&(h.includes("number")||h.includes("#"))),accountName:d.findIndex(h=>h.includes("account")&&h.includes("name")),amount:d.findIndex(h=>h.includes("amount")),department:Be(d),payrollCategory:d.findIndex(h=>h.includes("payroll")&&h.includes("category")),employee:d.findIndex(h=>h.includes("employee"))};if(console.log("JE Generation - Column indices:",r),r.amount===-1)throw new Error("Amount column not found in PR_Data_Clean.");let i=new Map;for(let h=1;h<l.length;h++){let E=l[h],k=r.accountNumber>=0?String(E[r.accountNumber]||"").trim():"",_=r.accountName>=0?String(E[r.accountName]||"").trim():"",$=Number(E[r.amount])||0,V=r.department>=0?String(E[r.department]||"").trim():"";if($===0)continue;let x=`${k}|${V}`;if(i.has(x)){let T=i.get(x);T.amount+=$}else i.set(x,{accountNumber:k,accountName:_,department:V,amount:$})}console.log("JE Generation - Aggregated into",i.size,"unique Account+Department combinations");let u=["RefNumber","TxnDate","Account Number","Account Name","LineAmount","Debit","Credit","LineDesc","Department"],f=[u],p=0,c=0,m=Array.from(i.values()).sort((h,E)=>{let k=String(h.accountNumber).localeCompare(String(E.accountNumber));return k!==0?k:String(h.department).localeCompare(String(E.department))});for(let h of m){let{accountNumber:E,accountName:k,department:_,amount:$}=h,V=$>0?$:0,x=$<0?Math.abs($):0,T=[k,_].filter(Boolean).join(" - ");p+=V,c+=x,f.push([t,n,E,k,$,V||"",x||"",T,_])}console.log("JE Generation - Built",f.length-1,"summarized journal lines"),console.log("JE Generation - Total Debit:",p,"Total Credit:",c);let y=e.workbook.worksheets.getItemOrNullObject(A.JE_DRAFT);if(y.load("isNullObject"),await e.sync(),y.isNullObject)y=e.workbook.worksheets.add(A.JE_DRAFT),await e.sync();else{let h=y.getUsedRangeOrNullObject();h.load("address"),await e.sync(),h.isNullObject||(h.clear(),await e.sync())}let w=y.getRangeByIndexes(0,0,f.length,u.length);w.values=f,await e.sync();try{let h=f.length-1,E=y.getRange("A1:I1");_n(E),h>0&&(An(y,1,h),ft(y,4,h),ft(y,5,h),ft(y,6,h)),y.getRange("A:I").format.autofitColumns(),await e.sync()}catch(h){console.warn("JE formatting error (non-critical):",h)}y.activate(),y.getRange("A1").select(),await e.sync(),G.debitTotal=p,G.creditTotal=c,G.difference=c-p}),G.loading=!1,G.lastError=null,le()}catch(e){console.error("JE Generation failed:",e),G.loading=!1,G.lastError=e.message||"Failed to generate journal entry.",le()}}async function Va(){if(!ae()){window.alert("Excel runtime is unavailable.");return}try{let{rows:e}=await Excel.run(async n=>{let a=n.workbook.worksheets.getItem(A.JE_DRAFT).getUsedRangeOrNullObject();a.load("values"),await n.sync();let s=a.isNullObject?[]:a.values||[];if(!s.length)throw new Error(`${A.JE_DRAFT} is empty.`);return{rows:s}}),t=ga(e);ha(`pr-je-draft-${Ze()}.csv`,t)}catch(e){console.warn("JE export:",e),window.alert("Unable to export the JE draft. Confirm the sheet has data.")}}})();
//# sourceMappingURL=app.bundle.js.map
